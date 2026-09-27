#!/usr/bin/env python3
"""docx-inspect 探测引擎。

读取单个 Office/OOXML 产物，并向 stdout 写出「指标级」的 JSON 探测结果。
本脚本是【唯一】了解 zipfile / lxml 的地方；它刻意对 Profile、安全策略、
office-core 的 IR 形状一无所知。

与 TypeScript 适配器之间的协议：

  * stdout   : 恰好一个 JSON 对象（探测载荷，版本见 PROBE_VERSION）
  * stderr   : 供人阅读的诊断信息，永不被解析
  * 退出码   : 0 表示已产出载荷（即便产物不受支持也返回 0）；
               2 表示彻底无法产出载荷

安全性说明：
  * 探测过程从不把部件解压到磁盘，也从不解析（更不访问）关系目标，
    因此恶意文档无法借此产生写文件、联网或路径穿越行为；
  * XML 解析已加固，抵御 XXE 与实体膨胀攻击；
  * 每次读取都受调用方传入的 limits 约束（zip 炸弹防护）；
  * ZIP 中央目录里的 `file_size` 仅作为参考，真实读取按「硬字节上限」流式截断。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import zipfile
from typing import Any

# 探测协议版本：引擎与模块可独立演进，模块据此判断兼容性。
PROBE_VERSION = 1

try:  # lxml 是推荐后端（见模块 spec）。
    from lxml import etree as _lxml_etree  # type: ignore

    XML_BACKEND = "lxml"

    def _parse_xml(data: bytes) -> Any:
        """用安全配置解析 XML。"""
        parser = _lxml_etree.XMLParser(
            resolve_entities=False,  # 不展开实体，杜绝 XXE
            no_network=True,  # 禁止联网取 DTD
            load_dtd=False,  # 不加载 DTD
            huge_tree=False,  # 关闭超大文档模式，限制资源占用
            recover=False,  # 遇错直接失败，不静默容错
        )
        return _lxml_etree.fromstring(data, parser=parser)

except ImportError:  # pragma: no cover - 仅在缺少 lxml 的主机上走到
    import xml.etree.ElementTree as _stdlib_etree  # type: ignore

    XML_BACKEND = "stdlib"

    def _parse_xml(data: bytes) -> Any:
        """标准库回退实现。

        ElementTree 既不展开外部实体也不获取 DTD，
        因此对不可信输入而言同样是安全的。
        """
        return _stdlib_etree.fromstring(data)


# --------------------------------------------------------------------------- #
# XML 辅助函数（与后端无关）                                                    #
# --------------------------------------------------------------------------- #

# 匹配形如 `{namespace}local` 的标签，用于剥离命名空间前缀。
_NS_BRACE = re.compile(r"^\{[^}]*\}")


def local_name(tag: Any) -> str:
    """返回元素/属性标签的本地名（去掉命名空间）。"""
    if not isinstance(tag, str):
        return ""
    return _NS_BRACE.sub("", tag)


def attr_local(element: Any, name: str) -> str | None:
    """返回首个本地名匹配的属性值（忽略命名空间差异）。"""
    for key, value in element.attrib.items():
        if local_name(key) == name:
            return value
    return None


def findall_local(element: Any, name: str) -> list[Any]:
    """递归查找所有本地名匹配的后代元素。"""
    return [child for child in element.iter() if local_name(child.tag) == name]


# --------------------------------------------------------------------------- #
# Content-Type / 关系常量                                                       #
# --------------------------------------------------------------------------- #

CT_VBA_PROJECT = "application/vnd.ms-office.vbaproject"
CT_FLASH = "application/vnd.ms-office.flash"
CT_OLE_OBJECT = "application/vnd.openxmlformats-officedocument.oleobject"
CT_PACKAGE = "application/vnd.openxmlformats-officedocument.package"
CT_ALT_CHUNK = "application/vnd.openxmlformats-officedocument.wordprocessingml.altchunk"

# 主部件 Content-Type -> (documentKind, documentVariant, 产物媒体类型)
MAIN_DOCUMENT_TYPES: dict[str, tuple[str, str | None, str]] = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": (
        "wordprocessingml",
        "document",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": (
        "wordprocessingml",
        "template",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    ),
    "application/vnd.ms-word.document.macroenabled.main+xml": (
        "wordprocessingml",
        "macroEnabled",
        "application/vnd.ms-word.document.macroEnabled.12",
    ),
    "application/vnd.ms-word.template.macroenabledtemplate.main+xml": (
        "wordprocessingml",
        "macroEnabledTemplate",
        "application/vnd.ms-word.template.macroEnabled.12",
    ),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": (
        "spreadsheetml",
        "document",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ),
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml": (
        "presentationml",
        "document",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ),
}

# 各类嵌入对象所在的目录前缀（一律小写比较）。
EMBEDDING_DIRS = ("word/embeddings/", "xl/embeddings/", "ppt/embeddings/")
ACTIVEX_DIRS = ("word/activex/", "xl/activex/", "ppt/activex/")

# DDE 字段指令前缀。
DDE_PREFIXES = ("dde", "ddeauto")

# 正文部件：document.xml、header*.xml、footer*.xml。
BODY_PART_PATTERN = re.compile(r"word/(document|header\d*|footer\d*)\.xml$", re.IGNORECASE)


# --------------------------------------------------------------------------- #
# 小工具                                                                        #
# --------------------------------------------------------------------------- #


def sha256_of_file(path: str) -> str:
    """流式计算文件 SHA-256，避免一次性把大文件读入内存。"""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def detect_container(magic: bytes) -> str:
    """依据文件头魔数判断容器类型。

    注意：这里【只用魔数】，不信任扩展名——这正是「扩展名/MIME 与内容不符」
    能被发现的前提。
    """
    if magic[:4] in (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"):
        return "zip"
    # OLE/CFB 复合文档头（.doc/.xls 遗留格式，或加密后的 OOXML）。
    if magic[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "ole"
    if magic[:5].lower() == b"{\\rtf":
        return "rtf"
    return "unknown"


def guess_extension(path: str) -> str | None:
    """从文件名推断扩展名（小写）。"""
    base = path.replace("\\", "/").rsplit("/", 1)[-1]
    if "." not in base:
        return None
    return base.rsplit(".", 1)[-1].lower() or None


def resolve_relationship_source(rels_part: str) -> str:
    """由 `.rels` 路径反推其来源部件。

    例：`word/_rels/document.xml.rels` -> `word/document.xml`；
        包根 `_rels/.rels` -> `''`（空字符串代表包根）。
    """
    normalized = rels_part.replace("\\", "/")
    if normalized == "_rels/.rels":
        return ""
    if "/_rels/" not in normalized:
        return normalized
    prefix, _, remainder = normalized.partition("/_rels/")
    target = remainder[: -len(".rels")] if remainder.endswith(".rels") else remainder
    return f"{prefix}/{target}"


class LimitTracker:
    """资源预算跟踪器。

    只记录【第一个】被触碰的预算名：调用方据此决定是硬失败还是软化处理，
    而无需在每次越界时都做判断。
    """

    def __init__(self, limits: dict[str, Any]) -> None:
        self.max_entries = int(limits.get("maxArchiveEntries", 4096))
        self.max_entry_bytes = int(limits.get("maxEntryUncompressedBytes", 32 * 1024 * 1024))
        self.max_total_bytes = int(limits.get("maxTotalUncompressedBytes", 256 * 1024 * 1024))
        self.max_relationships = int(limits.get("maxRelationships", 4096))
        self.max_external = int(limits.get("maxExternalTargets", 512))
        self.first_hit: str | None = None
        self.total_bytes = 0

    def hit(self, name: str) -> None:
        """登记一次越界（仅保留首个）。"""
        if self.first_hit is None:
            self.first_hit = name

    def account_declared_size(self, declared: int) -> None:
        """按中央目录声明的大小记账。

        该数值可被攻击者伪造，但读取代价极低，适合作为第一道粗筛；
        真正的防护在 `read_entry` 的流式硬上限。
        """
        if declared > self.max_entry_bytes:
            self.hit("maxEntryUncompressedBytes")
        self.total_bytes += declared
        if self.total_bytes > self.max_total_bytes:
            self.hit("maxTotalUncompressedBytes")

    def account_actual_size(self, actual: int) -> None:
        """按真实读取到的字节数记账。"""
        if actual > self.max_entry_bytes:
            self.hit("maxEntryUncompressedBytes")


# --------------------------------------------------------------------------- #
# 探测实现                                                                      #
# --------------------------------------------------------------------------- #


def empty_result() -> dict[str, Any]:
    """构造字段齐全的空结果。

    所有字段都显式给出默认值，这样 TypeScript 侧的协议校验永远能通过，
    不必为「字段缺失」写分支。
    """
    return {
        "probeVersion": PROBE_VERSION,
        "xmlBackend": XML_BACKEND,
        "container": "unknown",
        "extension": None,
        "sizeBytes": 0,
        "sha256": "",
        "encrypted": None,
        "documentKind": "unknown",
        "documentVariant": None,
        "mediaType": None,
        "partCount": 0,
        "parts": [],
        "relationships": [],
        "macroIndicators": [],
        # 语义分析结果（可选增强，默认空；仅在开启 analyzeMacroCode 且装有 oletools 时填充）。
        "macroAnalyses": [],
        "externalReferences": [],
        "embeddedObjects": [],
        "ddeFields": [],
        "limitHit": None,
        "issues": [],
        "error": None,
    }


def probe(path: str, limits: dict[str, Any], flags: dict[str, Any]) -> dict[str, Any]:
    """执行一次完整探测，返回可 JSON 序列化的结果字典。

    流程：魔数判容器 → （非 ZIP 直接给出对应 issue）→ 枚举部件 →
    解析 Content_Types → 依次收集宏/嵌入对象/关系/DDE → 测量正文是否为空。
    """
    result = empty_result()
    result["extension"] = guess_extension(path)

    # 只读 8 字节即可判定容器类型，无需读取整个文件。
    with open(path, "rb") as handle:
        magic = handle.read(8)
    container = detect_container(magic)
    result["container"] = container

    # --- 非 ZIP 容器：给出明确问题并提前返回 ------------------------------- #
    if container == "ole":
        result["issues"].append(
            {
                "code": "LEGACY_OR_ENCRYPTED_CONTAINER",
                "message": (
                    "Artifact is an OLE/CFB container (legacy .doc or an encrypted OOXML "
                    "package); docx-inspect only handles the OOXML ZIP form."
                ),
            }
        )
        return result

    if container == "rtf":
        result["issues"].append(
            {"code": "NON_OOXML_CONTAINER", "message": "Artifact is an RTF stream."}
        )
        return result

    if container != "zip":
        result["issues"].append(
            {
                "code": "UNRECOGNIZED_CONTAINER",
                "message": "File magic matches neither ZIP nor an Office container.",
            }
        )
        return result

    tracker = LimitTracker(limits)

    try:
        with zipfile.ZipFile(path) as archive:
            # 按名称排序，保证同一份文档每次探测的部件顺序一致（输出可 diff）。
            infos = sorted(archive.infolist(), key=lambda info: info.filename)
            names = [info.filename for info in infos]

            if len(infos) > tracker.max_entries:
                tracker.hit("maxArchiveEntries")
                result["issues"].append(
                    {
                        "code": "LIMIT_ARCHIVE_ENTRIES",
                        "message": f"Archive has more than {tracker.max_entries} entries.",
                    }
                )

            # 能被 ZipFile 打开即说明这是一份可读的 OOXML 包，
            # 因此可以【确定地】回答「未加密」——而不是留在「无法判断」状态。
            # （OLE 容器才会提前返回，保持 encrypted=None。）
            result["encrypted"] = False

            # 加密的 OOXML 表现为一个 EncryptedPackage 条目。
            if any(name.endswith("EncryptedPackage") for name in names):
                result["encrypted"] = True
                result["issues"].append(
                    {
                        "code": "ENCRYPTED_PACKAGE",
                        "message": "Package declares an EncryptedPackage entry.",
                    }
                )

            # 记录部件清单（含声明大小），供 IR 与上游使用。
            for info in infos:
                tracker.account_declared_size(info.file_size)
                result["parts"].append(
                    {
                        "name": info.filename,
                        "sizeBytes": info.file_size,
                        "compressedSize": info.compress_size,
                    }
                )
            result["partCount"] = len(infos)

            # 解析 Content_Types 并确定主部件与文档变体。
            content_types = parse_content_types(archive, names, tracker, result)
            main_part = apply_document_kind(result, content_types)

            # 各项指标按能力开关分别收集，便于按需关闭以省开销。
            if flags.get("detectMacros", True):
                collect_macros(result, names, content_types)

                # 宏语义分析是【可选增强】：默认关闭，且依赖 oletools；
                # 缺失时由 analyze_macro_code 内部降级，不影响主流程。
                if flags.get("analyzeMacroCode", False):
                    analyze_macro_code(path, result)

            if flags.get("detectEmbeddedObjects", True) or flags.get("detectAltChunks", True):
                collect_embedded(result, names, content_types, flags)

            if flags.get("detectExternalLinks", True):
                collect_relationships(archive, names, tracker, result)

            if flags.get("detectDdeFields", True):
                collect_dde_fields(archive, names, tracker, result)

            # 仅对 WordprocessingML 做「正文是否为空」的测量。
            if main_part is not None and result["documentKind"] == "wordprocessingml":
                measure_body(archive, main_part, tracker, result)

    except zipfile.BadZipFile as error:
        # 无法当作 ZIP 打开：截断或损坏。
        result["error"] = {"kind": "BAD_ZIP", "message": str(error)}
        return result
    except (OSError, EOFError, ValueError, NotImplementedError) as error:
        result["error"] = {"kind": "READ_FAILED", "message": str(error)}
        return result

    # 汇总预算命中情况，作为一条问题上报（是否致命由 TS 侧决定）。
    result["limitHit"] = tracker.first_hit
    if tracker.first_hit is not None:
        result["issues"].append(
            {
                "code": "LIMIT_REACHED",
                "message": f"Probe stopped accounting after hitting {tracker.first_hit}.",
            }
        )

    return result


def read_entry(
    archive: zipfile.ZipFile,
    name: str,
    tracker: LimitTracker,
    result: dict[str, Any],
) -> bytes | None:
    """以「硬字节上限」流式读取单个条目。

    不信任 `file_size`：即使中央目录谎报很小的值，实际解压数据一旦超过上限
    也会立刻中止并丢弃，从而阻断 zip 炸弹。
    """
    try:
        info = archive.getinfo(name)
    except KeyError:
        return None

    # 声明大小已超限：直接跳过，连读都不必读。
    if info.file_size > tracker.max_entry_bytes:
        tracker.hit("maxEntryUncompressedBytes")
        result["issues"].append(
            {
                "code": "LIMIT_ENTRY_BYTES",
                "message": f"{name} exceeds the per-entry budget; not read.",
                "path": name,
            }
        )
        return None

    chunks: list[bytes] = []
    read_bytes = 0
    try:
        with archive.open(info) as stream:
            # 分块读取而非 read() 一把梭：这样才能在越界时及时止损。
            while True:
                chunk = stream.read(65536)
                if not chunk:
                    break
                read_bytes += len(chunk)
                if read_bytes > tracker.max_entry_bytes:
                    tracker.hit("maxEntryUncompressedBytes")
                    result["issues"].append(
                        {
                            "code": "LIMIT_ENTRY_BYTES",
                            "message": f"{name} exceeded the per-entry budget while reading.",
                            "path": name,
                        }
                    )
                    return None
                chunks.append(chunk)
    except (OSError, EOFError, zipfile.BadZipFile, RuntimeError) as error:
        result["issues"].append(
            {"code": "ENTRY_UNREADABLE", "message": f"{name}: {error}", "path": name}
        )
        return None

    tracker.account_actual_size(read_bytes)
    return b"".join(chunks)


def parse_content_types(
    archive: zipfile.ZipFile,
    names: list[str],
    tracker: LimitTracker,
    result: dict[str, Any],
) -> dict[str, Any]:
    """读取 `[Content_Types].xml`，返回 {'overrides': {}, 'defaults': {}}。"""
    content_types: dict[str, Any] = {"overrides": {}, "defaults": {}}
    # 大小写不敏感地查找该条目：不同生成器的大小写并不统一。
    entry_name = next((name for name in names if name.lower() == "[content_types].xml"), None)
    if entry_name is None:
        result["issues"].append(
            {
                "code": "MISSING_CONTENT_TYPES",
                "message": "Package has no [Content_Types].xml; not a valid OPC package.",
            }
        )
        return content_types

    data = read_entry(archive, entry_name, tracker, result)
    if data is None:
        return content_types
    try:
        root = _parse_xml(data)
    except Exception as error:  # noqa: BLE001 - 记为问题而非崩溃
        result["issues"].append(
            {"code": "XML_INVALID", "message": f"{entry_name}: {error}", "path": entry_name}
        )
        return content_types

    # Override 精确指定某个部件的类型；Default 按扩展名兜底。
    for element in root:
        kind = local_name(element.tag)
        if kind == "Override":
            # PartName 以 `/` 开头，去掉后与包内部件名对齐。
            part = (attr_local(element, "PartName") or "").lstrip("/")
            content_types["overrides"][part] = attr_local(element, "ContentType") or ""
        elif kind == "Default":
            extension = (attr_local(element, "Extension") or "").lower()
            content_types["defaults"][extension] = attr_local(element, "ContentType") or ""
    return content_types


def apply_document_kind(result: dict[str, Any], content_types: dict[str, Any]) -> str | None:
    """设置 kind/variant/mediaType，并返回主部件名（若识别到）。"""
    overrides: dict[str, str] = content_types["overrides"]
    main_part: str | None = None
    for part, content_type in overrides.items():
        entry = MAIN_DOCUMENT_TYPES.get(content_type.lower())
        if entry is None:
            continue
        kind, variant, media_type = entry
        result["documentKind"] = kind
        result["documentVariant"] = variant
        result["mediaType"] = media_type
        if main_part is None:
            main_part = part
        if kind == "wordprocessingml":
            # 若包中意外声明了多个主部件，优先返回 Word 主部件。
            return part
    return main_part


def collect_macros(
    result: dict[str, Any], names: list[str], content_types: dict[str, Any]
) -> None:
    """收集宏指标。

    三条互补的发现路径（任一命中即计入）：
      1. 部件名以 `vbaProject.bin` / `vbaData.xml` 结尾；
      2. Content-Type 显式声明 VBA 工程或宏表；
      3. 路径中出现宏表目录或 `xlMacrosheet.bin`。
    """
    overrides: dict[str, str] = content_types["overrides"]
    sizes = {part["name"]: part["sizeBytes"] for part in result["parts"]}
    seen: set[str] = set()

    def add(kind: str, part: str) -> None:
        key = f"{kind}\u0000{part}"
        if key in seen:
            return
        seen.add(key)
        result["macroIndicators"].append(
            {"kind": kind, "part": part, "sizeBytes": sizes.get(part, 0)}
        )

    for name in names:
        lowered = name.lower()
        if lowered.endswith("vbaproject.bin") or lowered.endswith("vbadata.xml"):
            add("vba", name)

    for part, content_type in overrides.items():
        lowered = content_type.lower()
        if lowered == CT_VBA_PROJECT:
            add("vba", part)
        elif "macrosheet" in lowered:
            add("xlm", part)

    for name in names:
        lowered = name.lower()
        if "/macrosheets/" in lowered or lowered.endswith("xlmacrosheet.bin"):
            add("xlm", name)


def analyze_macro_code(path: str, result: dict[str, Any]) -> None:
    """对已检出的宏做语义分析（可选增强能力）。

    工程范式抄自 oletools 自身的发行方式：核心能力保持零依赖，重能力放进可选增强，
    依赖缺席时【降级】而不是报错——这样本模块在没有 oletools 的环境里仍能独立运行。

    判定算法沿用 oletools 的 MacroRaptor（BSD 许可），不自研，原因有二：
      1. 结论可解释：A(自动执行) ∧ (W(写文件/内存) ∨ X(执行外部程序))。
         真正的恶意宏几乎必然同时命中 A 与 X/W；而「只有宏」本身并不构成证据。
      2. 三个要素各由一条正则表达，命中项可直接作为证据回传给上层审计。

    与 `collect_macros` 的分工：后者回答「有没有宏」，本函数回答「宏在干什么」。
    """
    # 没有检出宏部件：没有可分析的对象，直接返回（不产生任何告警）。
    if not result["macroIndicators"]:
        return

    try:
        # 延迟导入：只有真正需要语义分析时才付出加载成本。
        # 实测代价约 +580ms（其中 pyparsing 约占 +256ms），因此【绝不能】放到模块顶层导入。
        from oletools.mraptor import MacroRaptor
        from oletools.olevba import VBA_Parser
    except ImportError:
        # 依赖缺失属于【预期内】情况，不是错误：降级为「仅存在性」指标，
        # 并留下一条 issue，让上层知道本次能力被跳过了。
        result["issues"].append(
            {
                "code": "MACRO_ANALYSIS_UNAVAILABLE",
                "message": (
                    "oletools is not installed; macro code analysis was skipped. "
                    "Install it (pip install oletools) to enable this optional capability."
                ),
            }
        )
        return

    parser = None
    try:
        parser = VBA_Parser(path)
        # extract_macros 逐模块产出 (filename, stream_path, vba_filename, vba_code)。
        for entry in parser.extract_macros():
            vba_filename = entry[2]
            vba_code = entry[3]
            if not vba_code:
                # 个别流（如 dir）不含源码，跳过但不报错。
                continue
            raptor = MacroRaptor(vba_code)
            raptor.scan()
            result["macroAnalyses"].append(
                {
                    # 模块名优先取 VBA 内部文件名，缺失时退回流路径。
                    "module": vba_filename or entry[1] or "",
                    "autoExec": bool(raptor.autoexec),
                    "write": bool(raptor.write),
                    "execute": bool(raptor.execute),
                    "suspicious": bool(raptor.suspicious),
                    # 形如 "A-X" 的三字符标志，便于人读与日志检索。
                    "flags": raptor.get_flags(),
                    "matches": [str(item) for item in raptor.matches],
                }
            )
    except Exception as error:  # noqa: BLE001
        # 第三方库面对畸形输入可能抛出任意异常；语义分析失败【不应】让整次探测失败，
        # 因为主流程的存在性指标已经拿到了。
        result["issues"].append(
            {
                "code": "MACRO_ANALYSIS_FAILED",
                "message": f"{type(error).__name__}: {error}",
            }
        )
    finally:
        if parser is not None:
            try:
                parser.close()
            except Exception:  # noqa: BLE001
                pass

    # 排序：让输出与 olevba 的产出顺序无关，保证同一文档每次结果完全一致（可 diff）。
    result["macroAnalyses"].sort(key=lambda item: item["module"])


def collect_embedded(
    result: dict[str, Any],
    names: list[str],
    content_types: dict[str, Any],
    flags: dict[str, Any],
) -> None:
    """收集嵌入对象指标：OLE 嵌入、包对象、altChunk、ActiveX、Flash。"""
    overrides: dict[str, str] = content_types["overrides"]
    sizes = {part["name"]: part["sizeBytes"] for part in result["parts"]}
    detect_embedded = flags.get("detectEmbeddedObjects", True)
    detect_alt_chunks = flags.get("detectAltChunks", True)
    detect_activex = flags.get("detectActiveX", True)

    def add(part: str, kind: str, content_type: str | None) -> None:
        result["embeddedObjects"].append(
            {
                "part": part,
                "kind": kind,
                "mediaType": content_type,
                "sizeBytes": sizes.get(part, 0),
                "name": part.rsplit("/", 1)[-1],
            }
        )

    for name in names:
        lowered = name.lower()
        # 注意：overrides 的键是原始大小写，这里用原始 name 取值。
        content_type = overrides.get(name)
        normalized_type = (content_type or "").lower()

        # 路径优先：embeddings 目录下的部件一定是嵌入对象。
        if detect_embedded and any(lowered.startswith(directory) for directory in EMBEDDING_DIRS):
            add(name, classify_embedding(content_type, lowered), content_type)
            continue

        if detect_activex and any(lowered.startswith(directory) for directory in ACTIVEX_DIRS):
            add(name, "activeX", content_type)
            continue

        # altChunk：可通过 Content-Type 或部件名（afchunk）识别。
        if detect_alt_chunks and (normalized_type == CT_ALT_CHUNK or "afchunk" in lowered):
            add(name, "altChunk", content_type)
            continue

        # 兜底：由 Content-Type 判断 OLE 对象 / 包对象。
        if detect_embedded and normalized_type in (CT_OLE_OBJECT, CT_PACKAGE):
            add(name, "oleEmbedding" if normalized_type == CT_OLE_OBJECT else "package", content_type)


def classify_embedding(content_type: str | None, lowered_name: str) -> str:
    """判定 embeddings 目录下某个部件的具体种类。"""
    normalized = (content_type or "").lower()
    if normalized == CT_OLE_OBJECT:
        return "oleEmbedding"
    if normalized == CT_PACKAGE:
        return "package"
    if normalized == CT_FLASH:
        return "flash"
    if lowered_name.endswith(".swf"):
        return "flash"
    # 无 Content-Type 时按扩展名猜：`.bin` 通常是 OLE 嵌入二进制。
    if lowered_name.endswith(".bin"):
        return "oleEmbedding"
    return "package"


def collect_relationships(
    archive: zipfile.ZipFile,
    names: list[str],
    tracker: LimitTracker,
    result: dict[str, Any],
) -> None:
    """解析所有 `.rels`，记录关系并挑出外部引用。"""
    rels_parts = [name for name in names if name.lower().endswith(".rels")]
    if len(rels_parts) > tracker.max_relationships:
        tracker.hit("maxRelationships")
        rels_parts = rels_parts[: tracker.max_relationships]

    external_count = 0
    for rels_name in rels_parts:
        data = read_entry(archive, rels_name, tracker, result)
        if data is None:
            continue
        try:
            root = _parse_xml(data)
        except Exception as error:  # noqa: BLE001 - 单个 .rels 损坏不应中断整体探测
            result["issues"].append(
                {"code": "XML_INVALID", "message": f"{rels_name}: {error}", "path": rels_name}
            )
            continue

        source_part = resolve_relationship_source(rels_name)
        for element in findall_local(root, "Relationship"):
            rel_id = attr_local(element, "Id") or ""
            rel_type = attr_local(element, "Type") or ""
            target = attr_local(element, "Target") or ""
            # 规范中缺省值即 Internal；大小写不敏感地识别 External。
            raw_mode = attr_local(element, "TargetMode") or "Internal"
            target_mode = "External" if raw_mode.lower() == "external" else "Internal"

            # 全部关系都记录进 IR（上游可能关心内部结构）。
            result["relationships"].append(
                {
                    "sourcePart": source_part,
                    "id": rel_id,
                    "type": rel_type,
                    "target": target,
                    "targetMode": target_mode,
                }
            )

            # 仅外部引用进入 externalReferences（本模块的核心指标之一）。
            if target_mode != "External":
                continue

            external_count += 1
            # 超出上限后仍继续记录关系，但不再累积外部引用，避免内存膨胀。
            if external_count > tracker.max_external:
                tracker.hit("maxExternalTargets")
                continue

            result["externalReferences"].append(
                {
                    "sourcePart": source_part,
                    "relationshipId": rel_id,
                    "relationshipType": rel_type,
                    "target": target,
                    "category": classify_external(rel_type, target),
                }
            )


def classify_external(rel_type: str, target: str) -> str:
    """给外部关系目标分类。

    必须与 TypeScript 侧 `classifyExternalTarget` 保持完全一致，
    否则同一份文档在不同链路上会得到不同分类。
    """
    lowered = rel_type.lower()
    if lowered.endswith("/attachedtemplate"):
        return "attachedTemplate"
    if lowered.endswith("/hyperlink"):
        return "hyperlink"
    if lowered.endswith("/image"):
        return "externalImage"
    if lowered.endswith("/oleobject") or lowered.endswith("/package"):
        return "externalObject"
    if "mailmerge" in lowered or lowered.endswith("/datasource"):
        return "dataSource"
    if lowered.endswith("/framefile"):
        return "externalObject"
    if target.strip().lower().startswith(DDE_PREFIXES):
        return "dde"
    return "other"


def collect_dde_fields(
    archive: zipfile.ZipFile,
    names: list[str],
    tracker: LimitTracker,
    result: dict[str, Any],
) -> None:
    """在正文/页眉/页脚中查找 DDE/DDEAUTO 字段指令。

    DDE 字段由 `fldSimple@instr` 或分隔的 `instrText` 承载；
    这里只提取指令文本，不执行也不展开任何字段。
    """
    collected: list[str] = []
    for name in names:
        if not BODY_PART_PATTERN.search(name):
            continue
        data = read_entry(archive, name, tracker, result)
        if data is None:
            continue
        try:
            root = _parse_xml(data)
        except Exception:  # noqa: BLE001 - 坏掉的部件不应中断探测
            continue
        for element in root.iter():
            kind = local_name(element.tag)
            instruction: str | None = None
            if kind == "fldSimple":
                instruction = attr_local(element, "instr")
            elif kind == "instrText" and element.text:
                instruction = element.text
            if instruction is None:
                continue
            stripped = instruction.strip()
            if stripped.lower().startswith(DDE_PREFIXES):
                # 截断，避免超长指令把输出撑爆。
                collected.append(stripped[:512])
    # 排序去重，保证输出稳定。
    result["ddeFields"] = sorted(set(collected))


def measure_body(
    archive: zipfile.ZipFile,
    main_part: str,
    tracker: LimitTracker,
    result: dict[str, Any],
) -> None:
    """判断主文档部件是否「无可见内容」（既无文本，也无表格/图片/对象）。"""
    data = read_entry(archive, main_part, tracker, result)
    if data is None:
        return
    try:
        root = _parse_xml(data)
    except Exception:  # noqa: BLE001
        return

    texts: list[str] = []
    has_structure = False
    for element in root.iter():
        name = local_name(element.tag)
        if name == "t" and element.text:
            texts.append(element.text)
        elif name in ("drawing", "tbl", "object", "pict"):
            # 这些元素即使没有文本也算「有内容」。
            has_structure = True

    # 注意：空段落（w:p）会被 Word 写入，因此不能只看是否存在段落。
    if not has_structure and "".join(texts).strip() == "":
        result["issues"].append(
            {"code": "EMPTY_DOCUMENT", "message": "Main document part has no visible content."}
        )


# --------------------------------------------------------------------------- #
# 入口                                                                          #
# --------------------------------------------------------------------------- #


def build_parser() -> argparse.ArgumentParser:
    """构造命令行参数解析器。"""
    parser = argparse.ArgumentParser(description="docx-inspect probe engine")
    parser.add_argument("--path", required=True, help="Path to the artifact to probe")
    parser.add_argument(
        "--config",
        default="{}",
        help="JSON object with `limits` and `featureFlags` overrides",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """脚本入口：解析参数、执行探测、输出 JSON。"""
    args = build_parser().parse_args(argv)
    try:
        config = json.loads(args.config) if args.config else {}
        if not isinstance(config, dict):
            raise ValueError("--config must be a JSON object")
    except (json.JSONDecodeError, ValueError) as error:
        # 参数本身不合法：没有任何可产出的载荷，返回 2。
        print(f"invalid --config: {error}", file=sys.stderr)
        return 2

    limits = config.get("limits") or {}
    flags = config.get("featureFlags") or {}

    # 注意：即使文件不存在等异常，也要输出「字段齐全的结果对象」，
    # 让 TS 侧始终走统一的协议解析路径，而不是去猜 stderr。
    try:
        result = probe(args.path, limits, flags)
    except FileNotFoundError:
        result = empty_result()
        result["extension"] = guess_extension(args.path)
        result["error"] = {"kind": "NOT_FOUND", "message": "Artifact path does not exist."}
    except PermissionError as error:
        result = empty_result()
        result["error"] = {"kind": "PERMISSION_DENIED", "message": str(error)}
    except OSError as error:
        result = empty_result()
        result["error"] = {"kind": "IO_ERROR", "message": str(error)}

    # 文件级元信息（大小与摘要）在探测之后统一补全。
    try:
        result["sizeBytes"] = os.path.getsize(args.path)
    except OSError:
        result["sizeBytes"] = 0
    try:
        result["sha256"] = sha256_of_file(args.path)
    except OSError:
        result["sha256"] = ""

    # ensure_ascii=False：保留中文可读性（配合客户端设置的 UTF-8 编码）。
    json.dump(result, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
