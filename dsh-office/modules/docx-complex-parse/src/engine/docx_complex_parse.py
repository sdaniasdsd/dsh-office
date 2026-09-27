#!/usr/bin/env python3
"""docx-complex-parse 的解析引擎（Python 侧）。

两条独立的数据线，靠 `body_index` 会合：

  1. **结构线** —— 标准库 `zipfile` + `xml.etree.ElementTree` 直读 OOXML，
     拿的是文档里【写着的事实】：`tblGrid` 列数、`gridSpan`、`vMerge`、
     `gridBefore`/`gridAfter`。
  2. **版面线** —— `rdocx` 的版面模型，拿的是【算出来的事实】：页尺寸，以及
     每个正文块落在哪一页、页内哪个矩形。

为什么必须分开：rdocx 的 Python 绑定是纯高层 API（`Table.rows` / `Cell.text`），
**没有任何 gridSpan / vMerge / tblGrid 出口**，也无法读到合并结构。所以合并网格
只能靠自己读 XML；反过来，分页与页内几何只能靠版面模型，XML 里没有。

两者通过正文块序号对齐：rdocx 按文档顺序给每个正文块一个连续的 `body_index`，
本脚本枚举 `w:body` 直接子节点（`w:p` / `w:tbl`）时使用同一套顺序。

输出协议：stdout 上一份 JSON。成功时是原始观察结果，分三层：

  - **包级事实** —— 容器、格式、加密、部件与关系、宏/外部引用/嵌入对象指标。
    这一层必须由引擎实测：`indicators` 为空等于断言「没有宏」，是安全结论而非猜测。
  - **结构** —— 正文块与表格网格（上面第 1 条线）。
  - **版面** —— 页尺寸与块级页内矩形（上面第 2 条线）。

失败时是 `{"failure": {"reason": ...}}`——**引擎只报「原因」，不报错误码**，
错误码由 TS 侧的 mapper 决定，这样换引擎时上层错误语义不变。

协议版本与 TS 侧 `SUPPORTED_PARSE_VERSION` 必须一致。
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from typing import Any

# 随包发布的解释器是 Python embeddable 版，目录里带 python313._pth——那等价于
# isolated 模式：脚本所在目录【不会】自动进入 sys.path，PYTHONPATH 也会被忽略。
# observations 就在脚本旁边，只能靠显式插入；否则整个进程以 ModuleNotFoundError
# 非零退出，上层只看到一个没有失败信封的 ENGINE_FAILED。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from observations import read_observations  # noqa: E402  — 必须在 sys.path 修正之后

PARSE_VERSION = 1
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MAIN_DOCUMENT_PART = "word/document.xml"


# --------------------------------------------------------------------------- #
# 协议输出                                                                     #
# --------------------------------------------------------------------------- #


def emit(payload: dict[str, Any]) -> None:
    """把结果写到 stdout。ensure_ascii=False 保留中文原文，便于人读诊断。"""
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    sys.stdout.flush()


def fail(reason: str, message: str, **detail: Any) -> None:
    """报告一个「原因」，由 TS 侧翻译成错误码。"""
    emit({"failure": {"reason": reason, "message": message, "detail": detail}})
    sys.exit(0)


# --------------------------------------------------------------------------- #
# XML 小工具                                                                   #
# --------------------------------------------------------------------------- #


def local_name(tag: Any) -> str:
    """去掉命名空间前缀，只留本地名。"""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def children(element: Any) -> list[Any]:
    """直接子元素。"""
    return [node for node in element]


def direct_child(element: Any, name: str) -> Any | None:
    """按本地名找第一个直接子元素。"""
    for node in element:
        if local_name(node.tag) == name:
            return node
    return None


def direct_children(element: Any, name: str) -> list[Any]:
    """按本地名找全部直接子元素。"""
    return [node for node in element if local_name(node.tag) == name]


def attr(element: Any, name: str) -> str | None:
    """读取 `w:` 命名空间下的属性值。"""
    return element.get(f"{{{W_NS}}}{name}")


def parse_int(raw: str | None, fallback: int) -> int:
    """把属性值解析成非负整数；非法值回落到默认值。"""
    if raw is None:
        return fallback
    try:
        value = int(str(raw).strip())
    except (TypeError, ValueError):
        return fallback
    return value if value >= 0 else fallback


# --------------------------------------------------------------------------- #
# 文本提取                                                                     #
# --------------------------------------------------------------------------- #


def paragraph_text(paragraph: Any) -> str:
    """提取段落的可见文本。

    `w:del` 里的内容（修订删除）用 `w:delText` 承载，天然不会被 `w:t` 命中，
    因此删除稿不会污染文本。
    """
    pieces: list[str] = []
    for node in paragraph.iter():
        name = local_name(node.tag)
        if name == "t":
            pieces.append(node.text or "")
        elif name == "tab":
            pieces.append("\t")
        elif name in ("br", "cr"):
            pieces.append("\n")
    return "".join(pieces)


def cell_text(cell: Any) -> str:
    """单元格文本：段落之间用换行连接（表格里换行比空格更接近原意）。"""
    lines = [paragraph_text(p) for p in direct_children(cell, "p")]
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 结构线：正文块与表格网格                                                      #
# --------------------------------------------------------------------------- #


def count_columns(table: Any) -> int:
    """`w:tblGrid` / `w:gridCol` 声明的逻辑列数；缺失时为 0。"""
    grid = direct_child(table, "tblGrid")
    if grid is None:
        return 0
    return len(direct_children(grid, "gridCol"))


def read_row(row: Any, parse_merges: bool) -> dict[str, Any]:
    """读取一行：`gridBefore`/`gridAfter` 与逐格合并属性。"""
    tr_pr = direct_child(row, "trPr")
    grid_before = 0
    grid_after = 0
    if tr_pr is not None:
        before_node = direct_child(tr_pr, "gridBefore")
        after_node = direct_child(tr_pr, "gridAfter")
        grid_before = parse_int(attr(before_node, "val") if before_node is not None else None, 0)
        grid_after = parse_int(attr(after_node, "val") if after_node is not None else None, 0)

    cells: list[dict[str, Any]] = []
    for cell in direct_children(row, "tc"):
        tc_pr = direct_child(cell, "tcPr")
        grid_span = 1
        v_merge: str | None = None

        if parse_merges and tc_pr is not None:
            span_node = direct_child(tc_pr, "gridSpan")
            if span_node is not None:
                grid_span = max(1, parse_int(attr(span_node, "val"), 1))

            merge_node = direct_child(tc_pr, "vMerge")
            if merge_node is not None:
                # 省略 w:val 等价于 "continue"——这是规范里写明的默认值，
                # 也是「向下合并的延续格」最常见的写法。
                raw = attr(merge_node, "val")
                v_merge = "restart" if raw == "restart" else "continue"

        cells.append({"gridSpan": grid_span, "vMerge": v_merge, "text": cell_text(cell)})

    return {"gridBefore": grid_before, "gridAfter": grid_after, "cells": cells}


def read_table(table: Any, path: str, flags: dict[str, Any], limits: dict[str, Any],
               warnings: list[dict[str, Any]]) -> dict[str, Any]:
    """读取一个 `w:tbl`：网格声明 + 逐行逐格。"""
    parse_merges = bool(flags.get("parseMerges", True))
    rows: list[dict[str, Any]] = []
    header_rows = 0

    for row in direct_children(table, "tr"):
        tr_pr = direct_child(row, "trPr")
        if tr_pr is not None and direct_child(tr_pr, "tblHeader") is not None:
            header_rows += 1
        rows.append(read_row(row, parse_merges))

    # 嵌套表格不在本层展开：它的行列是另一张网格，混进外层会让坐标失去意义。
    for cell in _iter_cells(table):
        if direct_child(cell, "tbl") is not None:
            warnings.append({
                "code": "NESTED_TABLE_FLATTENED",
                "message": "A nested table was flattened to text; its grid is not modelled.",
                "path": path,
            })

    return {
        "kind": "table",
        "path": path,
        "tblGridColumns": count_columns(table),
        "headerRowCount": header_rows,
        "rows": rows,
    }


def _iter_cells(table: Any) -> list[Any]:
    found: list[Any] = []
    for row in direct_children(table, "tr"):
        found.extend(direct_children(row, "tc"))
    return found


def read_body(body: Any, flags: dict[str, Any], limits: dict[str, Any],
              warnings: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """按文档顺序枚举正文块。

    只认 `w:p` 与 `w:tbl` 两种直接子节点——这是与 rdocx `body_index` 对齐的
    关键：两侧必须对「第几个块」有一致的定义。
    """
    parse_tables = bool(flags.get("parseTables", True))
    max_blocks = limits.get("maxBlocks")

    blocks: list[dict[str, Any]] = []
    paragraph_count = 0
    table_count = 0

    for node in children(body):
        name = local_name(node.tag)
        if name == "p":
            paragraph_count += 1
            blocks.append({
                "kind": "paragraph",
                "path": f"/w:document/w:body/w:p[{paragraph_count}]",
                "text": paragraph_text(node),
            })
        elif name == "tbl":
            table_count += 1
            path = f"/w:document/w:body/w:tbl[{table_count}]"
            if parse_tables:
                blocks.append(read_table(node, path, flags, limits, warnings))
            else:
                # 关闭表格解析时仍保留块位——少一个块会让后续块的 body_index 全错位。
                blocks.append({
                    "kind": "table", "path": path,
                    "tblGridColumns": 0, "headerRowCount": 0, "rows": [],
                })
        else:
            continue

        if isinstance(max_blocks, int) and max_blocks > 0 and len(blocks) > max_blocks:
            fail("limit_exceeded", f"Block count exceeds maxBlocks={max_blocks}.",
                 limit="maxBlocks", value=len(blocks))

    return blocks


# --------------------------------------------------------------------------- #
# 版面线：rdocx                                                                #
# --------------------------------------------------------------------------- #


def build_layout(path: str, limits: dict[str, Any]) -> dict[str, Any]:
    """用 rdocx 取页尺寸与每个正文块的页内矩形。

    rdocx 缺失时是硬失败：它是本模块的必需依赖，没有它「来源坐标」就只剩结构路径。
    但 rdocx 已安装却无法排版某一份文档时，不能在这里直接终止进程；调用方需要
    根据 requirePageGeometry 决定是降级为结构坐标，还是将其视为硬失败。
    """
    try:
        import rdocx  # noqa: PLC0415  — 延迟导入，让「未安装」成为一句可读的失败原因
    except ImportError as exc:
        fail("runtime_missing", "The 'rdocx' package is required but not importable.",
             package="rdocx", error=str(exc))

    try:
        document = rdocx.Document(path)
        fragments = document.layout()
    except Exception as exc:  # noqa: BLE001
        # 文档能读、版面算不出来：这不是「文档坏了」，只是「拿不到页」。
        # 不要调用 fail()：它会抛 SystemExit，从而跳过 main() 中按 feature flag
        # 实施的降级逻辑。
        raise RuntimeError("The layout model could not compute page geometry.") from exc

    collected: list[dict[str, Any]] = []
    highest_page = 0
    for fragment in fragments:
        bounds = fragment.bounds
        page = int(fragment.physical_page)
        highest_page = max(highest_page, page)
        collected.append({
            "bodyIndex": int(fragment.body_index),
            "physicalPage": page,
            "displayedPage": int(fragment.displayed_page),
            "x": float(bounds.x),
            "y": float(bounds.y),
            "width": float(bounds.width),
            "height": float(bounds.height),
        })

    max_pages = limits.get("maxPages")
    if isinstance(max_pages, int) and max_pages > 0 and highest_page > max_pages:
        fail("limit_exceeded", f"Page count exceeds maxPages={max_pages}.",
             limit="maxPages", value=highest_page)

    pages: list[dict[str, Any]] = []
    for index in range(highest_page):
        try:
            page = document.layout_page(index)
            width = float(page.width)
            height = float(page.height)
        except Exception:  # noqa: BLE001
            continue
        pages.append({"index": index, "width": width, "height": height})

    # pages 与 fragments 必须同源：有页无片段说明版面模型没算出内容，宁可标为不可用。
    available = bool(pages) and bool(collected)
    return {
        "available": available,
        "unit": "pt",
        "pages": pages if available else [],
        "fragments": collected if available else [],
    }


# --------------------------------------------------------------------------- #
# 容器读取                                                                     #
# --------------------------------------------------------------------------- #


# --------------------------------------------------------------------------- #
# 包级事实：FormatProfile 与 FormatIR 需要的那一层                              #
# --------------------------------------------------------------------------- #
#
# 为什么必须由引擎提供而不是 TS 侧猜：`indicators.macros` 为空数组是一句
# 【安全断言】（「这个文档没有宏」）。猜错的代价是被策略层放行一个带宏文档，
# 因此这些事实只能来自对容器的实际读取。

# 主文档部件的媒体类型 → 粗粒度格式身份（office-core 的 FormatKind）。
MAIN_TYPE_TO_FORMAT = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "docx",
    "application/vnd.ms-word.document.macroEnabled.main+xml": "docm",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": "dotx",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml": "dotm",
    # 同一种格式在不同实现里出现过两种拼写，两种都要认。
    "application/vnd.ms-word.template.macroEnabled.main+xml": "dotm",
}


def source_part_for_rels(name: str) -> str:
    """由 `.rels` 部件名反推它的来源部件。

    `_rels/.rels` → `''`（包根）；`word/_rels/document.xml.rels` → `word/document.xml`。
    """
    directory, _, base = name.rpartition("/")
    stem = base[: -len(".rels")] if base.endswith(".rels") else base
    if directory == "_rels":
        return stem
    if directory.endswith("/_rels"):
        return f"{directory[: -len('/_rels')]}/{stem}"
    return stem


def resolve_target(source_part: str, target: str) -> str:
    """把关系的相对 target 解析成包内路径。

    OOXML 的 target 是相对来源部件所在目录的路径，且允许 `..`。不解析的话
    `embeddings/` 之外的对象会漏掉，影响 `embeddedObjects` 这份安全指标。
    """
    if target.startswith("/"):
        return target.lstrip("/")
    directory = source_part.rsplit("/", 1)[0] if "/" in source_part else ""
    combined = f"{directory}/{target}" if directory else target

    segments: list[str] = []
    for segment in combined.split("/"):
        if segment in ("", "."):
            continue
        if segment == "..":
            if segments:
                segments.pop()
            continue
        segments.append(segment)
    return "/".join(segments)


def read_content_types(archive: Any, names: list[str]) -> tuple[dict[str, str], dict[str, str]]:
    """读 `[Content_Types].xml` → （按部件的 Override 表，按扩展名的 Default 表）。

    缺失或损坏时返回空表：格式识别会因此退化为「靠部件名判断」，但不会中止解析。
    """
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    if "[Content_Types].xml" not in names:
        return overrides, defaults
    try:
        root = ET.fromstring(archive.read("[Content_Types].xml"))
    except Exception:  # noqa: BLE001
        return overrides, defaults

    for node in root:
        kind = local_name(node.tag)
        content_type = node.get("ContentType")
        if not content_type:
            continue
        if kind == "Override":
            part = node.get("PartName")
            if part:
                overrides[part.lstrip("/")] = content_type
        elif kind == "Default":
            extension = node.get("Extension")
            if extension:
                defaults[extension.lower()] = content_type
    return overrides, defaults


def content_type_for(part: str, overrides: dict[str, str], defaults: dict[str, str]) -> str | None:
    """按 Override 优先、扩展名兜底，求某个部件的媒体类型。"""
    if part in overrides:
        return overrides[part]
    if "." not in part:
        return None
    return defaults.get(part.rsplit(".", 1)[-1].lower())


def external_category(rel_type: str) -> str:
    """给外部引用分类，供策略层做差异化处理。"""
    if rel_type.endswith("/hyperlink"):
        return "hyperlink"
    if rel_type.endswith("/attachedTemplate"):
        return "remoteTemplate"
    if rel_type.endswith("/externalLink"):
        return "externalLink"
    if rel_type.endswith(("/image", "/video", "/audio")):
        return "remoteMedia"
    return "external"


def embedded_kind(rel_type: str) -> str | None:
    """把嵌入对象的类型 URI 归一到一小组标签；不是嵌入对象则返回 None。"""
    if rel_type.endswith("/oleObject"):
        return "oleEmbedding"
    if rel_type.endswith("/package"):
        return "package"
    if rel_type.endswith("/control"):
        return "activeX"
    if rel_type.endswith("/aFChunk"):
        return "altChunk"
    return None


def read_relationships(
    archive: Any,
    names: list[str],
    limits: dict[str, Any],
    overrides: dict[str, str],
    defaults: dict[str, str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """解析全部 `.rels`，返回（关系列表、外部引用、嵌入对象）。"""
    relationships: list[dict[str, Any]] = []
    external: list[dict[str, Any]] = []
    embedded: list[dict[str, Any]] = []
    max_relationships = limits.get("maxRelationships")

    for name in names:
        if not name.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(archive.read(name))
        except Exception:  # noqa: BLE001
            continue

        source_part = source_part_for_rels(name)
        for node in root:
            if local_name(node.tag) != "Relationship":
                continue

            rel_id = node.get("Id") or ""
            rel_type = node.get("Type") or ""
            target = node.get("Target") or ""
            is_external = (node.get("TargetMode") or "Internal") == "External"

            relationships.append({
                "sourcePart": source_part,
                "id": rel_id,
                "type": rel_type,
                "target": target,
                "targetMode": "External" if is_external else "Internal",
            })

            if isinstance(max_relationships, int) and max_relationships > 0 \
                    and len(relationships) > max_relationships:
                fail("limit_exceeded",
                     f"Relationships exceed maxRelationships={max_relationships}.",
                     limit="maxRelationships", value=len(relationships))

            if is_external:
                external.append({
                    "sourcePart": source_part,
                    "relationshipId": rel_id,
                    "relationshipType": rel_type,
                    "target": target,
                    "category": external_category(rel_type),
                })
                continue

            kind = embedded_kind(rel_type)
            if kind is not None:
                part = resolve_target(source_part, target)
                embedded.append({
                    "part": part,
                    "kind": kind,
                    "mediaType": content_type_for(part, overrides, defaults),
                    "sizeBytes": 0,  # 由 parts 表回填
                    "name": None,
                })

    return relationships, external, embedded


def read_macro_indicators(parts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """只回答「存在宏」，不判定其内容或风险。"""
    found: list[dict[str, Any]] = []
    for part in parts:
        if part["name"].lower().endswith("/vbaproject.bin"):
            found.append({"kind": "vba", "part": part["name"], "sizeBytes": part["sizeBytes"]})
    return found


def read_container(path: str, limits: dict[str, Any]) -> dict[str, Any]:
    """读取 zip 容器，产出解析所需的【全部】包级事实。

    这一步同时承担格式与安全的把关：非 ZIP、OLE（加密或旧版 .doc）在这里就被
    识别并拒绝，因此后续环节可以假定手里是一个结构完整的 zip。
    """
    try:
        with open(path, "rb") as handle:
            head = handle.read(8)
    except OSError as exc:
        fail("not_found", "The artifact could not be opened.", error=str(exc))

    if head.startswith(OLE_MAGIC):
        # 加密文档与 .doc 都是 OLE 复合文件：能识别，但读不了。
        fail("encrypted", "The artifact is an OLE compound file (encrypted or legacy .doc).")

    if not head.startswith(b"PK"):
        fail("format_mismatch", "The artifact is not a ZIP container.")

    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        fail("format_mismatch", "The artifact is not a readable ZIP container.", error=str(exc))

    with archive:
        infos = archive.infolist()

        max_entries = limits.get("maxArchiveEntries")
        if isinstance(max_entries, int) and max_entries > 0 and len(infos) > max_entries:
            fail("limit_exceeded", f"Archive entries exceed maxArchiveEntries={max_entries}.",
                 limit="maxArchiveEntries", value=len(infos))

        max_entry = limits.get("maxEntryUncompressedBytes")
        max_total = limits.get("maxTotalUncompressedBytes")
        total = 0
        parts: list[dict[str, Any]] = []
        for info in infos:
            # 目录条目不是部件，报出去只会让 parts 表出现零字节的空项。
            if info.is_dir():
                continue
            size = int(info.file_size)
            if isinstance(max_entry, int) and max_entry > 0 and size > max_entry:
                fail("limit_exceeded", "A single entry exceeds maxEntryUncompressedBytes.",
                     limit="maxEntryUncompressedBytes", entry=info.filename, value=size)
            total += size
            if isinstance(max_total, int) and max_total > 0 and total > max_total:
                fail("limit_exceeded", "Total uncompressed size exceeds maxTotalUncompressedBytes.",
                     limit="maxTotalUncompressedBytes", value=total)
            parts.append({
                "name": info.filename,
                "sizeBytes": size,
                "compressedSize": int(info.compress_size),
            })

        names = [part["name"] for part in parts]
        if MAIN_DOCUMENT_PART not in names:
            fail("format_mismatch", "The container has no word/document.xml part.")

        document_bytes = archive.read(MAIN_DOCUMENT_PART)

        overrides, defaults = read_content_types(archive, names)
        main_type = content_type_for(MAIN_DOCUMENT_PART, overrides, defaults)
        relationships, external, embedded = read_relationships(
            archive, names, limits, overrides, defaults)

    # 嵌入对象的字节数由 parts 表回填：关系本身不携带大小。
    size_by_part = {part["name"]: part["sizeBytes"] for part in parts}
    for item in embedded:
        item["sizeBytes"] = size_by_part.get(item["part"], 0)

    format_kind = MAIN_TYPE_TO_FORMAT.get(main_type or "", "unknown")
    signatures = ["magic:zip", f"part:{MAIN_DOCUMENT_PART}"]
    if format_kind != "unknown":
        signatures.append("content-types:wordprocessingml")

    return {
        "documentBytes": document_bytes,
        "parts": parts,
        "relationships": relationships,
        "indicators": {
            "macros": read_macro_indicators(parts),
            "externalReferences": external,
            "embeddedObjects": embedded,
        },
        "profile": {
            "container": "zip",
            "format": format_kind,
            # 走到这里就不可能是 OLE，因此「未加密」是可判定的结论而非猜测。
            "encrypted": False,
            "mediaType": main_type,
            "signatures": signatures,
        },
    }


# --------------------------------------------------------------------------- #
# 入口                                                                         #
# --------------------------------------------------------------------------- #


def main() -> None:
    parser = argparse.ArgumentParser(description="docx-complex-parse engine")
    parser.add_argument("--path", required=True, help="artifact 的本地路径")
    parser.add_argument("--config", default="{}", help="JSON 配置：limits 与 featureFlags")
    args = parser.parse_args()

    try:
        config = json.loads(args.config)
    except json.JSONDecodeError as exc:
        fail("bad_config", "The --config payload is not valid JSON.", error=str(exc))

    if not isinstance(config, dict):
        fail("bad_config", "The --config payload must be a JSON object.")

    limits = config.get("limits") or {}
    flags = config.get("featureFlags") or {}
    if not isinstance(limits, dict) or not isinstance(flags, dict):
        fail("bad_config", "limits and featureFlags must be JSON objects.")

    warnings: list[dict[str, Any]] = []

    # Disabled enforcement is still measured and reported by the TS mapper.
    if not flags.get("enforceLimits", True):
        limits = {}

    container = read_container(args.path, limits)

    try:
        root = ET.fromstring(container["documentBytes"])
    except Exception as exc:  # noqa: BLE001
        fail("parse_failed", "word/document.xml is not well-formed XML.",
             error=f"{type(exc).__name__}: {exc}")

    body = direct_child(root, "body")
    if body is None:
        fail("parse_failed", "word/document.xml has no w:body element.")

    blocks = read_body(body, flags, limits, warnings)
    try:
        observations = read_observations(body, flags, warnings, limits.get("maxFloatingObjects"))
    except OverflowError:
        fail("limit_exceeded", "Floating objects exceed maxFloatingObjects.",
             limit="maxFloatingObjects", value=limits["maxFloatingObjects"] + 1)

    require_geometry = bool(flags.get("requirePageGeometry", False))

    if not bool(flags.get("parsePageGeometry", True)):
        # 调用方明确不要页面几何：这是唯一能跳过 rdocx 的机会，必须尊重——
        # 版面分析是整个流程里最贵的一步，白跑一遍再丢掉纯属浪费。
        if require_geometry:
            fail("layout_unavailable", "Page geometry is required but disabled by feature flags.")
        layout = {"available": False, "unit": "pt", "pages": [], "fragments": []}
    else:
        try:
            layout = build_layout(args.path, limits)
        except SystemExit:
            raise
        except Exception as exc:  # noqa: BLE001
            if require_geometry:
                fail("layout_unavailable", "Page geometry is required but unavailable.",
                     error=f"{type(exc).__name__}: {exc}")
            warnings.append({
                "code": "PAGE_GEOMETRY_UNAVAILABLE",
                "message": "The layout model failed; coordinates degrade to structural paths.",
            })
            layout = {"available": False, "unit": "pt", "pages": [], "fragments": []}

    if layout["available"] and layout["fragments"]:
        covered = {fragment["bodyIndex"] for fragment in layout["fragments"]}
        missing = [index for index in range(len(blocks)) if index not in covered]
        if missing:
            # 结构与版面的块序号对不齐：说明两侧对「正文块」的定义有分歧，
            # 这会让坐标静默错位，必须显式告警而不是当没事。
            warnings.append({
                "code": "PAGE_BREAK_ESTIMATED",
                "message": (
                    "Layout covered "
                    f"{len(covered)} of {len(blocks)} body blocks; "
                    f"{len(missing)} block(s) have no page geometry."
                ),
                "path": "word/document.xml",
            })

    emit({
        "parseVersion": PARSE_VERSION,
        "mainPart": MAIN_DOCUMENT_PART,
        "profile": container["profile"],
        "parts": container["parts"],
        "relationships": container["relationships"],
        "indicators": container["indicators"],
        "blocks": blocks,
        **observations,
        "warnings": warnings,
        "layout": layout,
    })


if __name__ == "__main__":
    main()
