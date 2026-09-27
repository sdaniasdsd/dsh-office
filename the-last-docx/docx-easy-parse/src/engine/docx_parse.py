#!/usr/bin/env python3
"""docx-parse 解析引擎。

读取单个 DOCX/OOXML 产物，并把「结构级」解析结果以 JSON 写到 stdout。
本脚本是【唯一】了解 zipfile / lxml 的地方；它刻意对 Profile、安全策略、
office-core 的 IR 形状一无所知。

与 TypeScript 适配器之间的协议：

  * stdout   : 恰好一个 JSON 对象（解析载荷，版本见 PARSE_VERSION）
  * stderr   : 供人阅读的诊断信息，永不被解析
  * 退出码   : 0 表示已产出载荷（即便产物不受支持也返回 0）；
               2 表示彻底无法产出载荷

安全性说明：
  * 解析过程从不把部件解压到磁盘，也从不解析（更不访问）关系目标的内容，
    因此恶意文档无法借此产生写文件、联网或路径穿越行为；
  * XML 解析已加固，抵御 XXE 与实体膨胀攻击；
  * 每次读取都受调用方传入的 limits 约束（zip 炸弹防护）；
  * ZIP 中央目录里的 `file_size` 仅作为参考，真实读取按「硬字节上限」流式截断。

设计取舍（轻量优先，不做双 IR）：
  引擎直接产出与领域模型同形的结构（blocks / styles / relationships /
  comments / footnotes），TS 侧只做信任边界校验与派生（outline / counts），
  因此不存在「引擎中间模型 → 模块 IR」的二次搬运。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import zipfile
from typing import Any, Callable, Iterator

# 解析协议版本：引擎与模块可独立演进，模块据此判断兼容性。
PARSE_VERSION = 1

# --------------------------------------------------------------------------- #
# XML 后端（lxml 优先，缺失时回退标准库）                                       #
# --------------------------------------------------------------------------- #

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
# 命名空间与常量                                                                #
# --------------------------------------------------------------------------- #

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def W(name: str) -> str:
    """构造 wordprocessingml 命名空间的 Clark 记法标签。"""
    return f"{{{W_NS}}}{name}"


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

STYLE_TYPE_MAP = {
    "paragraph": "paragraph",
    "character": "character",
    "table": "table",
    "numbering": "numbering",
}

ALIGN_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "both": "both",
    "distribute": "distribute",
    "start": "left",
    "end": "right",
}

_HEADING_RE = re.compile(r"(?:heading|标题)\s*([1-9])", re.IGNORECASE)
_HEADING_ID_RE = re.compile(r"^heading([1-9])$", re.IGNORECASE)
_NS_BRACE = re.compile(r"^\{[^}]*\}")

# 分批读取的块大小。
_CHUNK = 1024 * 1024


# --------------------------------------------------------------------------- #
# XML 小工具（与后端无关）                                                      #
# --------------------------------------------------------------------------- #


def local_name(tag: Any) -> str:
    """返回元素/属性标签的本地名（去掉命名空间）。"""
    if not isinstance(tag, str):
        return ""
    return _NS_BRACE.sub("", tag)


def w_child(element: Any, name: str) -> Any | None:
    """返回首个位于 wordprocessingml 命名空间的直接子元素。"""
    if element is None:
        return None
    wanted = W(name)
    for child in element:
        if child.tag == wanted:
            return child
    return None


def w_children(element: Any, name: str) -> list[Any]:
    """返回全部位于 wordprocessingml 命名空间的直接子元素。"""
    if element is None:
        return []
    wanted = W(name)
    return [child for child in element if child.tag == wanted]


def w_attr(element: Any, name: str) -> str | None:
    """读取 w: 命名空间的属性；兼容未加命名空间的写法。"""
    if element is None:
        return None
    value = element.get(W(name))
    if value is None:
        value = element.get(name)
    return value


def w_attr_or(element: Any, name: str, default: str) -> str:
    value = w_attr(element, name)
    return default if value is None else value


def w_text(element: Any) -> str:
    """拼接元素下全部 `w:t` 的文本（不含其他命名空间的 t，如 a:t）。"""
    if element is None:
        return ""
    target = W("t")
    return "".join(node.text or "" for node in element.iter(target))


def find_local(element: Any, name: str) -> Any | None:
    """按本地名查找后代元素（忽略命名空间），用于 core.xml 等中性部件。"""
    if element is None:
        return None
    for node in element.iter():
        if local_name(node.tag) == name:
            return node
    return None


def to_int(value: str | None, default: int | None = None) -> int | None:
    if value is None:
        return default
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


# --------------------------------------------------------------------------- #
# 文件级工具                                                                    #
# --------------------------------------------------------------------------- #


def sha256_of_file(path: str) -> str:
    """流式计算文件 SHA-256，避免一次性把大文件读入内存。"""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(_CHUNK), b""):
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
    return posixpath.join(prefix, target)


def resolve_internal_target(source_part: str, target: str) -> str:
    """把内部关系目标解析成包内路径，用于「悬空关系」检测。"""
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part)
    joined = posixpath.join(base, target) if base else target
    return posixpath.normpath(joined)


# --------------------------------------------------------------------------- #
# 预算与读取                                                                    #
# --------------------------------------------------------------------------- #


class Budget:
    """集中跟踪资源预算与问题列表，确保超限行为一致且可审计。"""

    def __init__(self, limits: dict[str, Any]) -> None:
        self.limits = limits
        self.total_read = 0
        self.limit_hit: str | None = None
        self.issues: list[dict[str, Any]] = []

    def issue(self, code: str, message: str, path: str | None = None) -> None:
        entry: dict[str, Any] = {"code": code, "message": message}
        if path is not None:
            entry["path"] = path
        self.issues.append(entry)

    def note_limit(self, name: str, message: str, path: str | None = None) -> None:
        """记录首个被触发的预算；只保留第一个，便于上层做稳定判断。"""
        if self.limit_hit is None:
            self.limit_hit = name
        self.issue("LIMIT_REACHED", message, path)

    def add_read(self, size: int, path: str) -> None:
        self.total_read += size
        if self.total_read > self.limits["maxTotalUncompressedBytes"]:
            self.note_limit("LIMIT_TOTAL_BYTES", "total uncompressed bytes exceed budget", path)


def read_entry(archive: zipfile.ZipFile, name: str, budget: Budget) -> bytes | None:
    """读取单个条目，并按单条目硬上限截断。"""
    try:
        info = archive.getinfo(name)
    except KeyError:
        return None

    per_entry = budget.limits["maxEntryUncompressedBytes"]
    try:
        with archive.open(info) as handle:
            data = handle.read(per_entry + 1)
    except (zipfile.BadZipFile, OSError, RuntimeError, NotImplementedError) as exc:
        budget.issue("ENTRY_UNREADABLE", f"cannot read {name}: {exc}", name)
        return None

    if len(data) > per_entry:
        budget.note_limit("LIMIT_ENTRY_BYTES", f"entry {name} exceeds per-entry budget", name)
        data = data[:per_entry]
    budget.add_read(len(data), name)
    return data


# --------------------------------------------------------------------------- #
# 各部件解析                                                                    #
# --------------------------------------------------------------------------- #


def parse_content_types(root: Any) -> tuple[dict[str, str], dict[str, str]]:
    """解析 `[Content_Types].xml`，返回 (overrides, defaults)。

    overrides 键统一为小写、带前导斜杠的部件路径；defaults 键为小写扩展名。
    """
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    for node in root:
        tag = local_name(node.tag)
        if tag == "Override":
            part = (node.get("PartName") or "").strip()
            ctype = (node.get("ContentType") or "").strip()
            if part:
                overrides[part.lower()] = ctype
        elif tag == "Default":
            ext = (node.get("Extension") or "").strip().lower()
            ctype = (node.get("ContentType") or "").strip()
            if ext:
                defaults[ext] = ctype
    return overrides, defaults


def classify_document(overrides: dict[str, str]) -> tuple[str, str | None, str | None]:
    """由 Content-Types 判定文档族、变体与媒体类型。"""
    # 优先识别 WordprocessingML；找不到再看是否属于其他 OOXML 族。
    # Content-Type 大小写不固定（如 `macroEnabled`），因此查表前统一转小写。
    best: tuple[str, str | None, str | None] = ("unknown", None, None)
    for _part, ctype in sorted(overrides.items()):
        mapped = MAIN_DOCUMENT_TYPES.get(ctype.lower())
        if mapped is None:
            continue
        kind, variant, media = mapped
        if kind == "wordprocessingml":
            return kind, variant, media
        if best[0] == "unknown":
            best = (kind, variant, media)
    return best


def parse_relationships(root: Any, rels_part: str, budget: Budget) -> list[dict[str, Any]]:
    """解析单个 `.rels` 部件。"""
    out: list[dict[str, Any]] = []
    source_part = resolve_relationship_source(rels_part)
    limit = budget.limits["maxRelationships"]
    for node in root:
        if local_name(node.tag) != "Relationship":
            continue
        if len(out) >= limit:
            budget.note_limit("LIMIT_RELATIONSHIPS", "relationship budget exceeded", rels_part)
            break
        mode = node.get("TargetMode") or "Internal"
        out.append(
            {
                "sourcePart": source_part,
                "id": node.get("Id") or "",
                "type": node.get("Type") or "",
                "target": node.get("Target") or "",
                "targetMode": "External" if mode == "External" else "Internal",
            }
        )
    return out


def derive_heading(
    style_id: str | None, style_name: str | None, outline: int | None
) -> int | None:
    """推导标题层级（1 基）。

    优先级：显式 outlineLvl > 样式名（Heading N / 标题 N）> 样式 id（HeadingN）。
    """
    if outline is not None and outline >= 0:
        return outline + 1
    for candidate in (style_name, style_id):
        if not candidate:
            continue
        match = _HEADING_RE.search(candidate)
        if match:
            return int(match.group(1))
        match = _HEADING_ID_RE.match(candidate)
        if match:
            return int(match.group(1))
    return None


def parse_styles(root: Any, budget: Budget, flags: dict[str, Any]) -> list[dict[str, Any]]:
    """解析 `word/styles.xml`。"""
    styles: list[dict[str, Any]] = []
    if not flags["parseStyles"]:
        return styles
    limit = budget.limits["maxStyles"]
    for node in root:
        if local_name(node.tag) != "style":
            continue
        style_id = node.get(W("styleId")) or node.get("styleId")
        if not style_id:
            continue
        if len(styles) >= limit:
            budget.note_limit("LIMIT_STYLES", "style budget exceeded", "word/styles.xml")
            break
        type_val = (node.get(W("type")) or node.get("type") or "unknown").lower()
        name_node = w_child(node, "name")
        name = w_attr(name_node, "val")
        based_node = w_child(node, "basedOn")
        based = w_attr(based_node, "val")
        is_default = (node.get(W("default")) or node.get("default") or "") in (
            "1",
            "true",
            "on",
        )
        style_type = STYLE_TYPE_MAP.get(type_val, "unknown")
        heading = None
        if flags["parseHeadings"] and style_type == "paragraph":
            heading = derive_heading(style_id, name, None)
        styles.append(
            {
                "styleId": style_id,
                "name": name,
                "type": style_type,
                "basedOn": based,
                "isDefault": is_default,
                "headingLevel": heading,
            }
        )
    return styles


def parse_run(node: Any) -> dict[str, Any]:
    """解析一个 `w:r`，提取文本与基础字符格式。"""
    props = w_child(node, "rPr")
    bold = False
    italic = False
    underline = False
    if props is not None:
        b = w_child(props, "b")
        if b is not None:
            bold = (w_attr(b, "val") or "1") not in ("0", "false", "off")
        i = w_child(props, "i")
        if i is not None:
            italic = (w_attr(i, "val") or "1") not in ("0", "false", "off")
        u = w_child(props, "u")
        if u is not None:
            underline = (w_attr(u, "val") or "single") not in ("none", "0", "false", "off")
    return {"text": w_text(node), "bold": bold, "italic": italic, "underline": underline}


def collect_runs(paragraph: Any) -> list[Any]:
    """收集段落下的运行元素（含超链接内部的运行），保持出现顺序。"""
    runs: list[Any] = []
    for child in paragraph:
        if child.tag == W("r"):
            runs.append(child)
        elif child.tag == W("hyperlink"):
            for inner in child:
                if inner.tag == W("r"):
                    runs.append(inner)
    return runs


def parse_paragraph(
    node: Any, index: int, styles_by_id: dict[str, dict[str, Any]], flags: dict[str, Any]
) -> dict[str, Any]:
    """解析一个 `w:p`。"""
    props = w_child(node, "pPr")
    style_id: str | None = None
    alignment: str | None = None
    outline: int | None = None
    list_item = False
    if props is not None:
        style_node = w_child(props, "pStyle")
        style_id = w_attr(style_node, "val")
        jc = w_child(props, "jc")
        jc_val = w_attr(jc, "val")
        if jc_val is not None:
            alignment = ALIGN_MAP.get(jc_val, "unknown")
        ol = w_child(props, "outlineLvl")
        outline = to_int(w_attr(ol, "val"))
        list_item = w_child(props, "numPr") is not None

    style_name = None
    if style_id is not None:
        record = styles_by_id.get(style_id)
        if record is not None:
            style_name = record.get("name")

    heading_level = None
    if flags["parseHeadings"]:
        heading_level = derive_heading(style_id, style_name, outline)

    runs = [parse_run(run) for run in collect_runs(node)] if flags["parseParagraphs"] else []

    return {
        "index": index,
        "text": w_text(node),
        "styleId": style_id,
        "styleName": style_name,
        "outlineLevel": outline,
        "headingLevel": heading_level,
        "alignment": alignment,
        "listItem": list_item,
        "runs": runs,
    }


def parse_table(
    node: Any, index: int, budget: Budget, flags: dict[str, Any]
) -> dict[str, Any]:
    """解析一个 `w:tbl`（单元格扁平化，不做无界递归）。"""
    style_id = None
    tbl_pr = w_child(node, "tblPr")
    if tbl_pr is not None:
        style_node = w_child(tbl_pr, "tblStyle")
        style_id = w_attr(style_node, "val")

    grid = w_child(node, "tblGrid")
    declared_columns = len(w_children(grid, "gridCol")) if grid is not None else 0

    rows_nodes = w_children(node, "tr")
    cell_limit = budget.limits["maxTableCells"]
    cells: list[dict[str, Any]] = []
    columns = declared_columns
    truncated = False

    for row_index, row_node in enumerate(rows_nodes):
        column = 0
        for cell_node in w_children(row_node, "tc"):
            if len(cells) >= cell_limit:
                if not truncated:
                    budget.note_limit("LIMIT_TABLE_CELLS", "table cell budget exceeded")
                    truncated = True
                break
            cell_pr = w_child(cell_node, "tcPr")
            column_span = 1
            row_span = 1
            if cell_pr is not None:
                span_node = w_child(cell_pr, "gridSpan")
                column_span = to_int(w_attr(span_node, "val"), 1) or 1
                # 纵向合并：起始单元格记 1，续格也记 1（扁平模型不追踪跨度区间）。
                if w_child(cell_pr, "vMerge") is not None:
                    row_span = 1
            cells.append(
                {
                    "row": row_index,
                    "column": column,
                    "text": w_text(cell_node),
                    "paragraphCount": len(w_children(cell_node, "p")),
                    "columnSpan": column_span,
                    "rowSpan": row_span,
                }
            )
            column += column_span
        if column > columns:
            columns = column
        if truncated:
            break

    return {
        "index": index,
        "rows": len(rows_nodes),
        "columns": columns,
        "styleId": style_id,
        "cells": cells,
    }


# `w:body` 里允许出现、但不承载正文内容的块级元素（标记类）。
# 真实 Word 文档的 `w:body` 末尾恒定有一个 `w:sectPr`（页面尺寸/页边距），
# 还常见书签、批注锚点、校对标记与修订范围。它们都不含文本，必须静默忽略——
# 否则每一份真实文档都会被误报成「存在未解析内容」。
IGNORABLE_BODY_ELEMENTS = frozenset(
    {
        "sectPr",
        "bookmarkStart",
        "bookmarkEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "proofErr",
        "permStart",
        "permEnd",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart",
        "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart",
        "customXmlMoveToRangeEnd",
    }
)

# 单次解析最多报告多少种未解析元素；其余聚合为一条，避免刷屏。
MAX_UNPARSED_ELEMENTS_REPORTED = 20


def iter_body_blocks(container: Any) -> Iterator[tuple[str, Any]]:
    """按出现顺序迭代正文块：`w:p` / `w:tbl`，并下钻 `w:sdt` 内容控件。

    既不是块、也不属于标记类的元素按 `"unparsed"` 产出。迭代器只做分类，
    是否上报交由调用方决定（保持迭代器无副作用）。
    """
    for child in container:
        tag = child.tag
        if tag == W("p"):
            yield "p", child
        elif tag == W("tbl"):
            yield "tbl", child
        elif tag == W("sdt"):
            content = w_child(child, "sdtContent")
            if content is not None:
                yield from iter_body_blocks(content)
        elif local_name(tag) not in IGNORABLE_BODY_ELEMENTS:
            yield "unparsed", child


def parse_body(
    body: Any,
    styles_by_id: dict[str, dict[str, Any]],
    budget: Budget,
    flags: dict[str, Any],
) -> list[dict[str, Any]]:
    """解析正文，返回按文档顺序排列的块。"""
    blocks: list[dict[str, Any]] = []
    block_limit = budget.limits["maxBlocks"]
    paragraph_index = 0
    table_index = 0
    unparsed_names: list[str] = []

    for tag, node in iter_body_blocks(body):
        if tag == "unparsed":
            # 正文里确实有元素，但本模块不解析它。去重后统一上报，
            # 让调用方知道「抽取到的文本 != 文档全部文本」。
            name = local_name(node.tag)
            if name not in unparsed_names:
                unparsed_names.append(name)
            continue
        if len(blocks) >= block_limit:
            budget.note_limit("LIMIT_BLOCKS", "block budget exceeded", "word/document.xml")
            break
        if tag == "p":
            if not flags["parseParagraphs"]:
                continue
            blocks.append(
                {
                    "kind": "paragraph",
                    "paragraph": parse_paragraph(node, paragraph_index, styles_by_id, flags),
                }
            )
            paragraph_index += 1
        else:
            if not flags["parseTables"]:
                continue
            blocks.append({"kind": "table", "table": parse_table(node, table_index, budget, flags)})
            table_index += 1

    for name in unparsed_names[:MAX_UNPARSED_ELEMENTS_REPORTED]:
        budget.issue(
            "UNKNOWN_ELEMENT",
            f"body element '{name}' is not parsed and was skipped",
            "word/document.xml",
        )
    if len(unparsed_names) > MAX_UNPARSED_ELEMENTS_REPORTED:
        budget.issue(
            "UNKNOWN_ELEMENT",
            f"{len(unparsed_names) - MAX_UNPARSED_ELEMENTS_REPORTED} more unparsed body elements were not listed",
            "word/document.xml",
        )

    return blocks


def parse_comments(root: Any, budget: Budget, flags: dict[str, Any]) -> list[dict[str, Any]]:
    """解析 `word/comments.xml`。"""
    comments: list[dict[str, Any]] = []
    if not flags["parseComments"]:
        return comments
    limit = budget.limits["maxComments"]
    for node in root:
        if local_name(node.tag) != "comment":
            continue
        if len(comments) >= limit:
            budget.note_limit("LIMIT_COMMENTS", "comment budget exceeded", "word/comments.xml")
            break
        comments.append(
            {
                "id": w_attr_or(node, "id", ""),
                "author": w_attr(node, "author"),
                "initials": w_attr(node, "initials"),
                "date": w_attr(node, "date"),
                "text": w_text(node),
            }
        )
    return comments


def parse_notes(
    root: Any, kind: str, budget: Budget, flags: dict[str, Any]
) -> list[dict[str, Any]]:
    """解析 `word/footnotes.xml` 或 `word/endnotes.xml`。

    只保留用户可见的注释（`type` 缺省或 normal）；分隔符/续页符不进入 IR。
    """
    notes: list[dict[str, Any]] = []
    if not flags["parseFootnotes"]:
        return notes
    child_name = "footnote" if kind == "footnote" else "endnote"
    limit = budget.limits["maxFootnotes"]
    for node in root:
        if local_name(node.tag) != child_name:
            continue
        note_type = w_attr(node, "type")
        if note_type not in (None, "normal"):
            continue
        if len(notes) >= limit:
            budget.note_limit("LIMIT_FOOTNOTES", "note budget exceeded")
            break
        notes.append({"id": w_attr_or(node, "id", ""), "kind": kind, "text": w_text(node)})
    return notes


def parse_metadata(root: Any) -> dict[str, Any]:
    """解析 `docProps/core.xml`。"""
    fields = ("title", "creator", "lastModifiedBy", "created", "modified", "revision")
    metadata: dict[str, Any] = {}
    for local in fields:
        node = find_local(root, local)
        value = None
        if node is not None and node.text:
            value = node.text.strip() or None
        metadata[local] = value
    return metadata


def empty_metadata() -> dict[str, Any]:
    return {
        "title": None,
        "creator": None,
        "lastModifiedBy": None,
        "created": None,
        "modified": None,
        "revision": None,
    }


def detect_encryption(
    archive_names: set[str], container: str, magic_head: bytes
) -> bool | None:
    """判定是否加密：三态（True/False/None=无法判断）。"""
    if container == "zip":
        return "EncryptedPackage" in archive_names or "EncryptionInfo" in archive_names
    if container == "ole":
        # 真实加密 OOXML 是 OLE 容器，内部含 EncryptedPackage 流名。
        return True if b"EncryptedPackage" in magic_head else None
    return None


def note_sort_key(value: str) -> tuple[int, str]:
    """让脚注按「数值优先」稳定排序，非数值退化为字符串序。"""
    try:
        return (0, f"{int(value):010d}")
    except (TypeError, ValueError):
        return (1, value)


def parse_optional_part(
    archive: zipfile.ZipFile,
    name: str,
    budget: Budget,
    parser: Callable[[Any], list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    """读取并解析一个可选部件；缺失或损坏时返回空列表，绝不抛出。"""
    data = read_entry(archive, name, budget)
    if data is None:
        return []
    try:
        return parser(_parse_xml(data))
    except Exception as exc:  # noqa: BLE001 - 任何 XML 失败都降级为问题
        budget.issue("XML_INVALID", f"cannot parse {name}: {exc}", name)
        return []


# --------------------------------------------------------------------------- #
# 主流程                                                                        #
# --------------------------------------------------------------------------- #


def empty_payload() -> dict[str, Any]:
    """构造一个「结构完整但内容为空」的载荷，保证任何失败路径输出形状一致。"""
    return {
        "parseVersion": PARSE_VERSION,
        "xmlBackend": XML_BACKEND,
        "container": "unknown",
        "extension": None,
        "sizeBytes": 0,
        "sha256": "",
        "encrypted": None,
        "documentKind": "unknown",
        "documentVariant": None,
        "mediaType": None,
        "metadata": empty_metadata(),
        "blocks": [],
        "styles": [],
        "relationships": [],
        "comments": [],
        "footnotes": [],
        "limitHit": None,
        "issues": [],
        "error": None,
    }


# 正文解析器签名：接收已打开的归档与预算，返回 block 列表。
# 换引擎只需提供另一个同签名的实现（见 docx_parse_docling.py）。
BodyParser = Callable[
    [zipfile.ZipFile, Budget, dict[str, Any], dict[str, dict[str, Any]]],
    list[dict[str, Any]],
]


def parse_body_part(
    archive: zipfile.ZipFile,
    budget: Budget,
    flags: dict[str, Any],
    styles_by_id: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """默认（轻量）正文解析器：读取 word/document.xml 并走纯标准库解析。

    之所以把这一步抽成可替换的调用点：正文结构是全流程中唯一值得换实现的环节，
    OPC 层（Content-Types、样式、关系、批注、脚注、元数据）各引擎完全一致，
    因此换引擎时不应重复实现它们。
    """
    document_data = read_entry(archive, "word/document.xml", budget)
    if document_data is None:
        budget.issue("XML_INVALID", "missing word/document.xml", "word/document.xml")
        return []
    try:
        document_root = _parse_xml(document_data)
        body = find_local(document_root, "body")
        if body is None:
            return []
        return parse_body(body, styles_by_id, budget, flags)
    except Exception as exc:  # noqa: BLE001
        budget.issue(
            "XML_INVALID", f"cannot parse word/document.xml: {exc}", "word/document.xml"
        )
        return []


def build_payload(
    path: str,
    limits: dict[str, Any],
    flags: dict[str, Any],
    body_parser: "BodyParser | None" = None,
) -> dict[str, Any]:
    """读取并解析一个产物，返回协议载荷。"""
    payload = empty_payload()
    payload["extension"] = guess_extension(path)

    # --- 文件级信息 -------------------------------------------------------- //
    try:
        payload["sha256"] = sha256_of_file(path)
        payload["sizeBytes"] = os.path.getsize(path)
        with open(path, "rb") as handle:
            magic_head = handle.read(65536)
    except FileNotFoundError:
        payload["error"] = {"kind": "NOT_FOUND", "message": f"file not found: {path}"}
        return payload
    except PermissionError:
        payload["error"] = {"kind": "PERMISSION_DENIED", "message": f"permission denied: {path}"}
        return payload
    except OSError as exc:
        payload["error"] = {"kind": "IO_ERROR", "message": f"cannot read file: {exc}"}
        return payload

    container = detect_container(magic_head)
    payload["container"] = container

    if container != "zip":
        # OLE / RTF / 未知：不是本模块可解析的结构，但容器本身已可识别。
        payload["encrypted"] = detect_encryption(set(), container, magic_head)
        return payload

    budget = Budget(limits)
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        payload["error"] = {"kind": "BAD_ZIP", "message": f"not a valid zip archive: {exc}"}
        return payload

    with archive:
        archive_names = archive.namelist()
        name_set = set(archive_names)
        payload["encrypted"] = detect_encryption(set(archive_names), "zip", magic_head)

        if len(archive_names) > limits["maxArchiveEntries"]:
            budget.note_limit(
                "LIMIT_ARCHIVE_ENTRIES",
                f"archive has {len(archive_names)} entries, exceeding budget",
            )

        # --- Content-Types：判定文档族 ------------------------------------ //
        ct_data = read_entry(archive, "[Content_Types].xml", budget)
        overrides: dict[str, str] = {}
        if ct_data is not None:
            try:
                overrides, _defaults = parse_content_types(_parse_xml(ct_data))
            except Exception as exc:  # noqa: BLE001
                budget.issue(
                    "XML_INVALID",
                    f"cannot parse [Content_Types].xml: {exc}",
                    "[Content_Types].xml",
                )
        else:
            budget.issue("XML_INVALID", "missing [Content_Types].xml", "[Content_Types].xml")

        kind, variant, media = classify_document(overrides)
        payload["documentKind"] = kind
        payload["documentVariant"] = variant
        payload["mediaType"] = media

        # 非 WordprocessingML：不再深入解析正文（属于其他模块的范围）。
        if kind != "wordprocessingml":
            payload["issues"] = budget.issues
            payload["limitHit"] = budget.limit_hit
            return payload

        # --- 样式表 -------------------------------------------------------- //
        styles: list[dict[str, Any]] = []
        styles_data = read_entry(archive, "word/styles.xml", budget)
        if styles_data is not None:
            try:
                styles = parse_styles(_parse_xml(styles_data), budget, flags)
            except Exception as exc:  # noqa: BLE001
                budget.issue(
                    "XML_INVALID", f"cannot parse word/styles.xml: {exc}", "word/styles.xml"
                )
        elif flags["parseStyles"]:
            budget.issue("STYLES_MISSING", "word/styles.xml is absent", "word/styles.xml")
        styles_by_id = {style["styleId"]: style for style in styles}
        payload["styles"] = styles

        # --- 正文 ---------------------------------------------------------- //
        # 唯一可替换实现的环节：换引擎只需替换 body_parser，
        # OPC 层（样式/关系/批注/脚注/元数据）的处理完全复用。
        blocks = (body_parser or parse_body_part)(archive, budget, flags, styles_by_id)
        payload["blocks"] = blocks
        # 无任何可见内容（空段落不算内容）：结构合法但无信息，值得提示。
        has_visible = any(
            (block["kind"] == "paragraph" and block["paragraph"]["text"].strip() != "")
            or (block["kind"] == "table" and block["table"]["rows"] > 0)
            for block in blocks
        )
        if not has_visible:
            budget.issue(
                "EMPTY_DOCUMENT",
                "document body has no visible content",
                "word/document.xml",
            )

        # 断链样式引用：正文引用但样式表里没有。
        referenced = {
            block["paragraph"]["styleId"]
            for block in blocks
            if block["kind"] == "paragraph" and block["paragraph"]["styleId"]
        }
        missing = sorted(style_id for style_id in referenced if style_id not in styles_by_id)
        for style_id in missing[:20]:
            budget.issue(
                "STYLE_NOT_FOUND",
                f"paragraph references undefined style '{style_id}'",
                "word/styles.xml",
            )

        # --- 关系 ---------------------------------------------------------- //
        relationships: list[dict[str, Any]] = []
        if flags["parseRelationships"]:
            for rels_part in sorted(name for name in archive_names if name.endswith(".rels")):
                rels_data = read_entry(archive, rels_part, budget)
                if rels_data is None:
                    continue
                try:
                    relationships.extend(
                        parse_relationships(_parse_xml(rels_data), rels_part, budget)
                    )
                except Exception as exc:  # noqa: BLE001
                    budget.issue("XML_INVALID", f"cannot parse {rels_part}: {exc}", rels_part)
            relationships.sort(key=lambda rel: (rel["sourcePart"], rel["id"]))
            payload["relationships"] = relationships

            # 悬空内部关系：关系目标在包内不存在。
            dangling = 0
            for rel in relationships:
                if rel["targetMode"] != "Internal":
                    continue
                resolved = resolve_internal_target(rel["sourcePart"], rel["target"])
                if resolved and resolved not in name_set:
                    dangling += 1
                    if dangling <= 20:
                        budget.issue(
                            "DANGLING_RELATIONSHIP",
                            f"relationship '{rel['id']}' targets missing part '{resolved}'",
                            rel["sourcePart"] or "_rels/.rels",
                        )
            if dangling > 20:
                budget.issue(
                    "DANGLING_RELATIONSHIP",
                    f"{dangling - 20} more dangling relationships were not listed",
                )

        # --- 批注 ---------------------------------------------------------- //
        payload["comments"] = parse_optional_part(
            archive,
            "word/comments.xml",
            budget,
            lambda root: parse_comments(root, budget, flags),
        )

        # --- 脚注 / 尾注 --------------------------------------------------- //
        footnotes = parse_optional_part(
            archive,
            "word/footnotes.xml",
            budget,
            lambda root: parse_notes(root, "footnote", budget, flags),
        )
        footnotes.extend(
            parse_optional_part(
                archive,
                "word/endnotes.xml",
                budget,
                lambda root: parse_notes(root, "endnote", budget, flags),
            )
        )
        footnotes.sort(key=lambda note: (note_sort_key(note["kind"]), note_sort_key(note["id"])))
        payload["footnotes"] = footnotes

        # --- 元数据 -------------------------------------------------------- //
        if flags["extractMetadata"]:
            core_data = read_entry(archive, "docProps/core.xml", budget)
            if core_data is not None:
                try:
                    payload["metadata"] = parse_metadata(_parse_xml(core_data))
                except Exception as exc:  # noqa: BLE001
                    budget.issue(
                        "XML_INVALID", f"cannot parse docProps/core.xml: {exc}", "docProps/core.xml"
                    )

        payload["issues"] = budget.issues
        payload["limitHit"] = budget.limit_hit

    return payload


# --------------------------------------------------------------------------- #
# CLI                                                                           #
# --------------------------------------------------------------------------- #


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Parse a DOCX into structural JSON")
    parser.add_argument("--path", required=True, help="Path to the artifact")
    parser.add_argument("--config", required=True, help="JSON config: {limits, featureFlags}")
    args = parser.parse_args(argv)

    try:
        config = json.loads(args.config)
        limits = config["limits"]
        flags = config["featureFlags"]
    except (KeyError, ValueError, TypeError) as exc:
        sys.stderr.write(f"invalid --config payload: {exc}\n")
        return 2

    try:
        payload = build_payload(args.path, limits, flags)
    except Exception as exc:  # noqa: BLE001 - 兜底：任何未预期异常都不能污染协议
        payload = empty_payload()
        payload["error"] = {"kind": "READ_FAILED", "message": f"unexpected failure: {exc}"}

    # 强制 UTF-8 输出，避免 Windows 控制台默认编码把非 ASCII 内容写坏。
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except (AttributeError, ValueError):  # pragma: no cover
        pass
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
