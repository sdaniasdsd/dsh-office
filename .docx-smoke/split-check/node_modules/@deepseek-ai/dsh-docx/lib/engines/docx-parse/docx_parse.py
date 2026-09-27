#!/usr/bin/env python3
"""docx-parse 解析引擎。

读取单个 DOCX 产物，向其 stdout 输出「原始观察级」的 JSON。
本脚本是【唯一】了解 zipfile / lxml 的地方；它刻意对 Profile、安全策略、
office-core 的 IR 形状以及节点身份方案一无所知 —— 这些都由 TypeScript 侧决定。

与 TypeScript 适配器之间的协议：

  * stdout   : 恰好一个 JSON 对象（观察载荷，版本见 PARSE_VERSION）
  * stderr   : 供人阅读的诊断信息，永不被解析
  * 退出码   : 0 表示已产出载荷（即便产物不受支持也返回 0）；
               2 表示彻底无法产出载荷

它【只上报事实】：这个元素在哪、文本是什么、pStyle 是什么、有没有 paraId。
「它是不是标题」「级别几级」「语义 id 是什么」全部留给 TS 侧裁决。

安全性说明：
  * 从不把部件解压到磁盘，也从不解析（更不访问）任何关系目标，
    因此恶意文档无法借此产生写文件、联网或路径穿越行为；
  * XML 解析已加固，抵御 XXE 与实体膨胀攻击；
  * 每次读取都受调用方传入的 limits 约束（zip 炸弹防护）。
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

# 解析协议版本：引擎与模块可独立演进，模块据此判断兼容性。
PARSE_VERSION = 1

try:  # lxml 是推荐后端。
    from lxml import etree as _lxml_etree  # type: ignore

    XML_BACKEND = "lxml"

    def _parse_xml(data: bytes) -> Any:
        """用安全配置解析 XML。"""
        parser = _lxml_etree.XMLParser(
            resolve_entities=False,  # 不展开实体，杜绝 XXE
            no_network=True,  # 禁止联网取 DTD
            load_dtd=False,  # 不加载 DTD
            huge_tree=False,  # 关闭超大文档模式
            recover=False,  # 遇错直接失败，不静默容错
        )
        return _lxml_etree.fromstring(data, parser=parser)

except ImportError:  # pragma: no cover - 仅在缺少 lxml 的主机上走到
    import xml.etree.ElementTree as _stdlib_etree  # type: ignore

    XML_BACKEND = "stdlib"

    def _parse_xml(data: bytes) -> Any:
        """标准库回退实现（ElementTree 同样不展开外部实体、不取 DTD）。"""
        return _stdlib_etree.fromstring(data)


# --------------------------------------------------------------------------- #
# XML 辅助函数（与后端无关）                                                    #
# --------------------------------------------------------------------------- #

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


def children(element: Any) -> list[Any]:
    """返回直接子元素（跳过注释/处理指令等非元素节点）。"""
    return [child for child in element if isinstance(child.tag, str)]


def descendants(element: Any, name: str) -> list[Any]:
    """递归查找所有本地名匹配的后代元素。"""
    return [child for child in element.iter() if local_name(child.tag) == name]


def first_descendant(element: Any, name: str) -> Any | None:
    """查找首个本地名匹配的后代元素。"""
    for child in element.iter():
        if local_name(child.tag) == name:
            return child
    return None


def direct_child(element: Any, name: str) -> Any | None:
    """查找首个本地名匹配的直接子元素。"""
    for child in children(element):
        if local_name(child.tag) == name:
            return child
    return None


def text_of(element: Any) -> str:
    """拼接元素内所有 <w:t> 的文本。

    只取 `t`（text）：刻意排除 instrText（域指令）与 delText（已删除文本），
    因为它们不是「用户看到的正文」，混入会污染内容指纹。
    """
    parts: list[str] = []
    for node in element.iter():
        if local_name(node.tag) == "t" and node.text:
            parts.append(node.text)
    return "".join(parts)


# --------------------------------------------------------------------------- #
# 常量表                                                                        #
# --------------------------------------------------------------------------- #

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

# 形如 "Heading 2" / "heading2" / "标题 2" 的样式名。
_HEADING_NAME = re.compile(r"(?i)^(?:heading|标题)\s*([1-9])?\s*$")
# 从任意样式名中提取结尾数字，作为级别线索。
_TRAILING_DIGIT = re.compile(r"(\d+)\s*$")

# 正文中可承载块级内容的元素本地名。
BLOCK_TAGS = ("p", "tbl")


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
    """依据文件头魔数判断容器类型（不信任扩展名）。"""
    if magic[:4] in (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"):
        return "zip"
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


def parse_int(value: Any, fallback: int) -> int:
    """把属性值解析为整数，失败时回退。"""
    try:
        return int(str(value))
    except (TypeError, ValueError):
        return fallback


class LimitTracker:
    """资源预算跟踪器，只记录【第一个】被触碰的预算名。"""

    def __init__(self, limits: dict[str, Any]) -> None:
        self.max_entries = int(limits.get("maxArchiveEntries", 4096))
        self.max_entry_bytes = int(limits.get("maxEntryUncompressedBytes", 32 * 1024 * 1024))
        self.max_total_bytes = int(limits.get("maxTotalUncompressedBytes", 256 * 1024 * 1024))
        self.max_relationships = int(limits.get("maxRelationships", 4096))
        self.max_blocks = int(limits.get("maxBlocks", 50000))
        self.max_table_cells = int(limits.get("maxTableCells", 200000))
        self.max_annotations = int(limits.get("maxAnnotations", 10000))
        self.first_hit: str | None = None
        self.total_bytes = 0
        self.block_count = 0
        self.cell_count = 0
        self.annotation_count = 0

    def hit(self, name: str) -> None:
        if self.first_hit is None:
            self.first_hit = name

    def account_declared_size(self, declared: int) -> None:
        """按中央目录声明的大小记账（第一道粗筛）。"""
        if declared > self.max_entry_bytes:
            self.hit("maxEntryUncompressedBytes")
        self.total_bytes += declared
        if self.total_bytes > self.max_total_bytes:
            self.hit("maxTotalUncompressedBytes")

    def account_block(self) -> None:
        self.block_count += 1
        if self.block_count > self.max_blocks:
            self.hit("maxBlocks")

    def account_cell(self) -> None:
        self.cell_count += 1
        if self.cell_count > self.max_table_cells:
            self.hit("maxTableCells")

    def account_annotation(self) -> None:
        self.annotation_count += 1
        if self.annotation_count > self.max_annotations:
            self.hit("maxAnnotations")


# --------------------------------------------------------------------------- #
# 结果骨架                                                                      #
# --------------------------------------------------------------------------- #


def empty_result() -> dict[str, Any]:
    """构造字段齐全的空结果（保证 TS 侧协议校验永远能通过）。"""
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
        "partCount": 0,
        "parts": [],
        "relationships": [],
        "mainPart": "",
        "blocks": [],
        "annotations": [],
        "limitHit": None,
        "issues": [],
        "error": None,
    }


# --------------------------------------------------------------------------- #
# 样式解析                                                                      #
# --------------------------------------------------------------------------- #


class StyleTable:
    """styles.xml 的样式表，负责把 pStyle 解析为「可读名 + outlineLvl + 首行缩进」。"""

    def __init__(self) -> None:
        self.by_id: dict[str, dict[str, Any]] = {}
        self.readable = True
        self.default_style_id: str | None = None
        self.default_indent: dict[str, Any] | None = None

    def load(self, archive: zipfile.ZipFile, name: str, tracker: LimitTracker, result: dict[str, Any]) -> None:
        data = read_entry(archive, name, tracker, result)
        if data is None:
            self.readable = False
            result["issues"].append(
                {"code": "STYLES_UNREADABLE", "message": f"{name} could not be read.", "path": name}
            )
            return
        try:
            root = _parse_xml(data)
        except Exception as error:  # noqa: BLE001 - 记为问题而非崩溃
            self.readable = False
            result["issues"].append(
                {"code": "STYLES_UNREADABLE", "message": f"{name}: {error}", "path": name}
            )
            return

        for style in descendants(root, "style"):
            style_id = attr_local(style, "styleId")
            if not style_id:
                continue
            name_node = direct_child(style, "name")
            based_on = direct_child(style, "basedOn")
            outline_node = direct_child(style, "pPr")
            outline = None
            if outline_node is not None:
                level_node = direct_child(outline_node, "outlineLvl")
                if level_node is not None:
                    outline = parse_int(attr_local(level_node, "val"), 0)
            self.by_id[style_id] = {
                "name": attr_local(name_node, "val") if name_node is not None else None,
                "basedOn": attr_local(based_on, "val") if based_on is not None else None,
                "outlineLevel": outline,
                "indent": read_indent(outline_node),
            }
            if attr_local(style, "default") in ("1", "true", "on") and attr_local(style, "type") in (None, "paragraph"):
                self.default_style_id = style_id

        # docDefaults 是样式链的终点：任何样式都没声明缩进时，由它决定。
        for defaults in descendants(root, "docDefaults"):
            for default in descendants(defaults, "pPrDefault"):
                self.default_indent = read_indent(direct_child(default, "pPr"))

    def indent_for(self, style_id: str | None) -> dict[str, Any] | None:
        """沿 basedOn 解析有效首行缩进。

        样式表可读时「谁都没声明」是一个确定的事实：有效缩进为 0。样式表读不到时
        返回 None —— 那才是「不知道」，调用方不能把不知道当成没有。
        """
        if not self.readable:
            return None
        seen: set[str] = set()
        current = style_id or self.default_style_id
        while current and current not in seen:
            seen.add(current)
            entry = self.by_id.get(current)
            if entry is None:
                break
            if entry["indent"] is not None:
                return entry["indent"]
            current = entry["basedOn"]
        if self.default_indent is not None:
            return self.default_indent
        return {"firstLineTwips": 0, "firstLineChars": 0}

    def resolve_outline(self, style_id: str | None) -> int | None:
        """沿 basedOn 继承链解析 outlineLvl（带环检测）。"""
        seen: set[str] = set()
        current = style_id
        while current and current not in seen:
            seen.add(current)
            entry = self.by_id.get(current)
            if entry is None:
                return None
            if entry["outlineLevel"] is not None:
                return int(entry["outlineLevel"])
            current = entry["basedOn"]
        return None

    def resolve_name(self, style_id: str | None) -> str | None:
        if not style_id:
            return None
        entry = self.by_id.get(style_id)
        if entry is None:
            return None
        return entry["name"] or None

    def name_level(self, style_id: str | None, style_name: str | None) -> int | None:
        """从样式名/id 的数字后缀推断标题级别。"""
        for candidate in (style_name, style_id):
            if not candidate:
                continue
            match = _HEADING_NAME.match(candidate.strip())
            if match:
                return parse_int(match.group(1), 1)
            if re.match(r"(?i)^(?:heading|标题)", candidate.strip()):
                trailing = _TRAILING_DIGIT.search(candidate.strip())
                if trailing:
                    return parse_int(trailing.group(1), 1)
        return None

    def is_heading_style(self, outline_level: int | None, style_id: str | None, style_name: str | None) -> bool:
        """判断该样式是否属于标题样式族。"""
        if outline_level is not None:
            return True
        for candidate in (style_name, style_id):
            if candidate and re.match(r"(?i)^(?:heading|标题)\s*[1-9]?\s*$", candidate.strip()):
                return True
        return False


# --------------------------------------------------------------------------- #
# 正文解析                                                                      #
# --------------------------------------------------------------------------- #


def build_path(parent_path: str, local: str, index: int) -> str:
    """构造 XPath 风格的结构路径片段。"""
    return f"{parent_path}/w:{local}[{index}]"


def read_indent(ppr: Any) -> dict[str, Any] | None:
    """一个 pPr 里**直接声明**的首行缩进，原样上报（twips 与百分之一字符）。

    只上报首行缩进：悬挂缩进（`w:hanging`）属于列表标记的排布，创建计划用编号定义
    重建它，因此这里遇到悬挂缩进就不上报，而不是把负数硬写成 0。
    """
    if ppr is None:
        return None
    ind = direct_child(ppr, "ind")
    if ind is None:
        return None
    if attr_local(ind, "hanging") is not None or attr_local(ind, "hangingChars") is not None:
        return None
    entry: dict[str, Any] = {}
    twips = attr_local(ind, "firstLine")
    chars = attr_local(ind, "firstLineChars")
    if twips is not None:
        entry["firstLineTwips"] = parse_int(twips, 0)
    if chars is not None:
        entry["firstLineChars"] = parse_int(chars, 0)
    return entry or None


def run_formatting(run: Any) -> dict[str, Any]:
    """一个 run 的字符格式观察。

    只记录**直接可观测**的事实；缺省一律不写，而不是写成 False——OOXML 里
    「没有 w:b」表示继承，与「w:b w:val="0"」明确不加粗不是一回事。
    """
    entry: dict[str, Any] = {"text": text_of(run)}
    rpr = direct_child(run, "rPr")
    if rpr is None:
        return entry
    bold = direct_child(rpr, "b")
    if bold is not None:
        entry["bold"] = attr_local(bold, "val") not in ("0", "false", "off")
    size = direct_child(rpr, "sz")
    if size is not None:
        half_points = parse_int(attr_local(size, "val"), 0)
        if half_points > 0:
            # 以磅为单位对外暴露，与其它模块的长度口径一致。
            entry["size"] = half_points / 2
    fonts = direct_child(rpr, "rFonts")
    if fonts is not None:
        east = attr_local(fonts, "eastAsia")
        if east:
            entry["eastAsia"] = east
    color = direct_child(rpr, "color")
    if color is not None:
        value = attr_local(color, "val")
        if value and value != "auto":
            entry["color"] = value
    return entry


def paragraph_formatting(
    element: Any, ppr: Any, styles: "StyleTable", style_id: str | None
) -> dict[str, Any]:
    """段落对齐 + 首行缩进 + 逐 run 格式。仅在调用方开启 parseFormatting 时收集。

    缩进上报的是**解析过样式链之后**的有效值，而不是只看段落自己声明了什么：一份
    把缩进写在 Normal 样式里的文档，和一份逐段声明缩进的文档，段落上看到的字符串
    完全不同，但排出来是一样的，调用方要的是后者。
    """
    alignment = None
    if ppr is not None:
        jc = direct_child(ppr, "jc")
        if jc is not None:
            alignment = attr_local(jc, "val")
    runs = [
        run_formatting(child)
        for child in children(element)
        if local_name(child.tag) == "r"
    ]
    entry: dict[str, Any] = {"alignment": alignment, "runs": runs}
    indent = read_indent(ppr)
    if indent is None:
        indent = styles.indent_for(style_id)
    if indent is not None:
        entry["indent"] = indent
    return entry


def parse_paragraph(
    element: Any,
    part: str,
    path: str,
    ordinal: int,
    styles: StyleTable,
    flags: dict[str, Any],
) -> dict[str, Any]:
    """把一个 <w:p> 解析为原始段落观察。"""
    text = text_of(element)
    ppr = direct_child(element, "pPr")

    style_id = None
    inline_outline = None
    num_id = None
    num_level = None
    if ppr is not None:
        style_node = direct_child(ppr, "pStyle")
        if style_node is not None:
            style_id = attr_local(style_node, "val")
        outline_node = direct_child(ppr, "outlineLvl")
        if outline_node is not None:
            inline_outline = parse_int(attr_local(outline_node, "val"), 0)
        numpr = direct_child(ppr, "numPr")
        if numpr is not None:
            num_id_node = direct_child(numpr, "numId")
            ilvl_node = direct_child(numpr, "ilvl")
            if num_id_node is not None:
                num_id = attr_local(num_id_node, "val")
            if ilvl_node is not None:
                num_level = parse_int(attr_local(ilvl_node, "val"), 0)

    # outlineLvl 优先取段落直接声明；否则沿样式链解析。
    outline_level = inline_outline
    if outline_level is None and flags.get("parseHeadings", True):
        outline_level = styles.resolve_outline(style_id)

    style_name = styles.resolve_name(style_id) if flags.get("parseStyles", True) else None
    style_name_level = (
        styles.name_level(style_id, style_name) if flags.get("parseHeadings", True) else None
    )
    heading_style = (
        styles.is_heading_style(outline_level, style_id, style_name)
        if flags.get("parseHeadings", True)
        else False
    )

    comment_refs: list[str] = []
    footnote_refs: list[str] = []
    endnote_refs: list[str] = []
    for node in element.iter():
        name = local_name(node.tag)
        if name == "commentReference":
            ref = attr_local(node, "id")
            if ref is not None:
                comment_refs.append(ref)
        elif name == "footnoteReference":
            ref = attr_local(node, "id")
            if ref is not None:
                footnote_refs.append(ref)
        elif name == "endnoteReference":
            ref = attr_local(node, "id")
            if ref is not None:
                endnote_refs.append(ref)

    return {
        "pointer": f"{part}!{path}",
        "part": part,
        "path": path,
        "ordinal": ordinal,
        "paraId": attr_local(element, "paraId"),
        "text": text,
        "styleId": style_id,
        "styleName": style_name,
        "outlineLevel": outline_level,
        "styleNameLevel": style_name_level,
        "headingStyle": heading_style,
        "list": {"numId": num_id, "level": num_level or 0} if num_id is not None else None,
        # 表现层观察：默认不收集，因为语义视图刻意不承载「长什么样」。
        "formatting": paragraph_formatting(element, ppr, styles, style_id)
        if flags.get("parseFormatting")
        else None,
        "commentRefs": comment_refs,
        "footnoteRefs": footnote_refs,
        "endnoteRefs": endnote_refs,
    }


def parse_table(
    element: Any,
    part: str,
    path: str,
    ordinal: int,
    styles: StyleTable,
    flags: dict[str, Any],
    tracker: LimitTracker,
    result: dict[str, Any],
) -> dict[str, Any]:
    """把一个 <w:tbl> 解析为原始表格观察。"""
    grid = direct_child(element, "grid") or first_descendant(element, "tblGrid")
    grid_columns = 0
    if grid is not None:
        grid_columns = len([node for node in children(grid) if local_name(node.tag) == "gridCol"])

    rows_out: list[dict[str, Any]] = []
    row_ordinal = 0
    for row in children(element):
        if local_name(row.tag) != "tr":
            continue
        row_ordinal += 1
        cells_out: list[dict[str, Any]] = []
        cell_ordinal = 0
        for cell in children(row):
            if local_name(cell.tag) != "tc":
                continue
            cell_ordinal += 1
            tracker.account_cell()

            grid_span = 1
            v_merge = None
            tcpr = direct_child(cell, "tcPr")
            if tcpr is not None:
                span_node = direct_child(tcpr, "gridSpan")
                if span_node is not None:
                    grid_span = parse_int(attr_local(span_node, "val"), 1)
                vmerge_node = direct_child(tcpr, "vMerge")
                if vmerge_node is not None:
                    raw = attr_local(vmerge_node, "val")
                    v_merge = "restart" if raw == "restart" else "continue"

            paragraphs: list[dict[str, Any]] = []
            paragraph_ordinal = 0
            cell_path = build_path(build_path(path, "tr", row_ordinal), "tc", cell_ordinal)
            for paragraph in children(cell):
                name = local_name(paragraph.tag)
                if name == "p":
                    paragraph_ordinal += 1
                    paragraphs.append(
                        parse_paragraph(
                            paragraph,
                            part,
                            build_path(cell_path, "p", paragraph_ordinal),
                            paragraph_ordinal - 1,
                            styles,
                            flags,
                        )
                    )
                elif name == "tbl":
                    # 嵌套表格：本版本不展开，如实标记为未支持内容。
                    result["issues"].append(
                        {
                            "code": "UNSUPPORTED_ELEMENT",
                            "message": "Nested tables are not expanded in this version.",
                            "path": f"{part}!{cell_path}",
                        }
                    )

            cells_out.append(
                {
                    "gridSpan": grid_span,
                    "vMerge": v_merge,
                    "paragraphs": paragraphs,
                }
            )
        rows_out.append({"cells": cells_out})

    table_text = "\n".join(
        "".join(text_of(cell) for cell in children(row) if local_name(cell.tag) == "tc")
        for row in children(element)
        if local_name(row.tag) == "tr"
    )

    return {
        "kind": "table",
        "pointer": f"{part}!{path}",
        "part": part,
        "path": path,
        "ordinal": ordinal,
        "text": table_text,
        "gridColumns": grid_columns,
        "rows": rows_out,
    }


def parse_body(
    archive: zipfile.ZipFile,
    part: str,
    tracker: LimitTracker,
    flags: dict[str, Any],
    styles: StyleTable,
    result: dict[str, Any],
) -> list[dict[str, Any]]:
    """解析主文档部件，返回按文档顺序排列的正文块。"""
    data = read_entry(archive, part, tracker, result)
    if data is None:
        return []
    try:
        root = _parse_xml(data)
    except Exception as error:  # noqa: BLE001
        result["error"] = {"kind": "BAD_XML", "message": f"{part}: {error}"}
        return []

    body = first_descendant(root, "body")
    if body is None:
        result["issues"].append(
            {"code": "EMPTY_DOCUMENT", "message": "Main part has no <w:body>.", "path": part}
        )
        return []

    base_path = "/w:document/w:body"
    blocks: list[dict[str, Any]] = []
    # per-local 计数器：让路径与序号只按同名兄弟递增，得到稳定的 XPath 风格定位。
    counters: dict[str, int] = {}

    for child in children(body):
        name = local_name(child.tag)
        if name not in BLOCK_TAGS:
            # sectPr / bookmarkStart / sdt 等：本版本不建模为块。
            if name not in ("sectPr",):
                result["issues"].append(
                    {
                        "code": "UNSUPPORTED_ELEMENT",
                        "message": f"Body-level element <w:{name}> is not modelled as a block.",
                        "path": f"{part}!{base_path}",
                    }
                )
            continue

        counters[name] = counters.get(name, 0) + 1
        index = counters[name]
        path = build_path(base_path, name, index)
        tracker.account_block()

        if name == "p":
            block = parse_paragraph(child, part, path, index - 1, styles, flags)
            block["kind"] = "paragraph"
            blocks.append(block)
        else:
            if not flags.get("parseTables", True):
                continue
            blocks.append(parse_table(child, part, path, index - 1, styles, flags, tracker, result))

    return blocks


# --------------------------------------------------------------------------- #
# 关系 / 批注 / 脚注                                                            #
# --------------------------------------------------------------------------- #


def resolve_relationship_source(rels_part: str) -> str:
    """由 `.rels` 路径反推其来源部件。"""
    normalized = rels_part.replace("\\", "/")
    if normalized == "_rels/.rels":
        return ""
    if "/_rels/" not in normalized:
        return normalized
    prefix, _, remainder = normalized.partition("/_rels/")
    target = remainder[: -len(".rels")] if remainder.endswith(".rels") else remainder
    return f"{prefix}/{target}"


def collect_relationships(
    archive: zipfile.ZipFile,
    names: list[str],
    tracker: LimitTracker,
    result: dict[str, Any],
) -> None:
    """收集包内全部 `.rels` 关系。"""
    count = 0
    for name in names:
        if not name.lower().endswith(".rels"):
            continue
        data = read_entry(archive, name, tracker, result)
        if data is None:
            continue
        try:
            root = _parse_xml(data)
        except Exception as error:  # noqa: BLE001
            result["issues"].append(
                {"code": "XML_INVALID", "message": f"{name}: {error}", "path": name}
            )
            continue

        source_part = resolve_relationship_source(name)
        for node in children(root):
            if local_name(node.tag) != "Relationship":
                continue
            count += 1
            if count > tracker.max_relationships:
                tracker.hit("maxRelationships")
                continue
            target_mode = attr_local(node, "TargetMode") or "Internal"
            result["relationships"].append(
                {
                    "sourcePart": source_part,
                    "id": attr_local(node, "Id") or "",
                    "type": attr_local(node, "Type") or "",
                    "target": attr_local(node, "Target") or "",
                    "targetMode": "External" if target_mode == "External" else "Internal",
                }
            )


def collect_annotations(
    archive: zipfile.ZipFile,
    names: list[str],
    tracker: LimitTracker,
    flags: dict[str, Any],
    comment_anchors: dict[str, list[str]],
    footnote_anchors: dict[str, list[str]],
    endnote_anchors: dict[str, list[str]],
    result: dict[str, Any],
) -> None:
    """收集批注 / 脚注 / 尾注。"""
    specs = [
        ("comment", "word/comments.xml", "/w:comments/w:comment", comment_anchors, "parseComments"),
        ("footnote", "word/footnotes.xml", "/w:footnotes/w:footnote", footnote_anchors, "parseFootnotes"),
        ("endnote", "word/endnotes.xml", "/w:endnotes/w:endnote", endnote_anchors, "parseFootnotes"),
    ]
    for kind, part, base_path, anchors, flag_name in specs:
        if not flags.get(flag_name, True):
            continue
        entry = next((name for name in names if name.lower() == part), None)
        if entry is None:
            continue
        data = read_entry(archive, entry, tracker, result)
        if data is None:
            continue
        try:
            root = _parse_xml(data)
        except Exception as error:  # noqa: BLE001
            result["issues"].append(
                {"code": "XML_INVALID", "message": f"{entry}: {error}", "path": entry}
            )
            continue

        element_name = "comment" if kind == "comment" else ("footnote" if kind == "footnote" else "endnote")
        ordinal = 0
        for node in children(root):
            if local_name(node.tag) != element_name:
                continue
            reference = attr_local(node, "id")
            if reference is None:
                continue
            # 脚注/尾注的 id 0 是分隔符、-1 是延续分隔符，都不是真实注释。
            if kind != "comment" and reference in ("-1", "0"):
                continue
            ordinal += 1
            tracker.account_annotation()
            path = f"{base_path}[{ordinal}]"
            result["annotations"].append(
                {
                    "kind": kind,
                    "reference": reference,
                    "text": text_of(node),
                    "author": attr_local(node, "author"),
                    "date": attr_local(node, "date"),
                    "pointer": f"{entry}!{path}",
                    "part": entry,
                    "path": path,
                    "ordinal": ordinal - 1,
                    "anchorPointers": anchors.get(reference, []),
                }
            )


# --------------------------------------------------------------------------- #
# 条目读取（硬字节上限）                                                        #
# --------------------------------------------------------------------------- #


def read_entry(
    archive: zipfile.ZipFile,
    name: str,
    tracker: LimitTracker,
    result: dict[str, Any],
) -> bytes | None:
    """以「硬字节上限」流式读取单个条目，不信任中央目录声称的大小。"""
    try:
        info = archive.getinfo(name)
    except KeyError:
        return None

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

    return b"".join(chunks)


def parse_content_types(
    archive: zipfile.ZipFile,
    names: list[str],
    tracker: LimitTracker,
    result: dict[str, Any],
) -> dict[str, Any]:
    """读取 `[Content_Types].xml`。"""
    content_types: dict[str, Any] = {"overrides": {}, "defaults": {}}
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
    except Exception as error:  # noqa: BLE001
        result["issues"].append(
            {"code": "XML_INVALID", "message": f"{entry_name}: {error}", "path": entry_name}
        )
        return content_types

    for element in children(root):
        kind = local_name(element.tag)
        if kind == "Override":
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
        mapping = MAIN_DOCUMENT_TYPES.get(content_type)
        if mapping is None:
            continue
        kind, variant, media_type = mapping
        # 只把 WordprocessingML 主部件当作本模块的目标；其余仅用于判定 kind。
        if kind == "wordprocessingml" and main_part is None:
            main_part = part
            result["documentKind"] = kind
            result["documentVariant"] = variant
            result["mediaType"] = media_type
        elif result["documentKind"] == "unknown":
            result["documentKind"] = kind
            result["documentVariant"] = variant
            result["mediaType"] = media_type
    return main_part


# --------------------------------------------------------------------------- #
# 主流程                                                                        #
# --------------------------------------------------------------------------- #


def parse(path: str, limits: dict[str, Any], flags: dict[str, Any]) -> dict[str, Any]:
    """执行一次完整解析，返回可 JSON 序列化的结果字典。"""
    result = empty_result()
    result["extension"] = guess_extension(path)

    with open(path, "rb") as handle:
        magic = handle.read(8)
    result["container"] = detect_container(magic)

    if result["container"] == "ole":
        result["issues"].append(
            {
                "code": "LEGACY_OR_ENCRYPTED_CONTAINER",
                "message": (
                    "Artifact is an OLE/CFB container (legacy .doc or an encrypted OOXML "
                    "package); docx-parse only handles the OOXML ZIP form."
                ),
            }
        )
        return result

    if result["container"] == "rtf":
        result["issues"].append(
            {"code": "NON_OOXML_CONTAINER", "message": "Artifact is an RTF stream."}
        )
        return result

    if result["container"] != "zip":
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

            # 能被 ZipFile 打开即说明这是一份可读的 OOXML 包，因此可确定「未加密」。
            result["encrypted"] = False
            if any(name.endswith("EncryptedPackage") for name in names):
                result["encrypted"] = True
                result["issues"].append(
                    {
                        "code": "ENCRYPTED_PACKAGE",
                        "message": "Package declares an EncryptedPackage entry.",
                    }
                )

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

            content_types = parse_content_types(archive, names, tracker, result)
            main_part = apply_document_kind(result, content_types)
            # 兜底：未声明 Content-Type 时按约定路径找主部件。
            if main_part is None and "word/document.xml" in names:
                main_part = "word/document.xml"

            if main_part is not None:
                result["mainPart"] = main_part

            if flags.get("parseHeadings", True) or flags.get("parseStyles", True):
                styles = StyleTable()
                styles_entry = next((n for n in names if n.lower() == "word/styles.xml"), None)
                if styles_entry is None:
                    styles.readable = False
                    result["issues"].append(
                        {
                            "code": "STYLES_MISSING",
                            "message": "Package has no word/styles.xml; heading levels rely on inline data only.",
                        }
                    )
                else:
                    styles.load(archive, styles_entry, tracker, result)
            else:
                styles = StyleTable()
                styles.readable = False

            collect_relationships(archive, names, tracker, result)

            # 先扫正文：既得到块，也得到「注释 id -> 锚点块」的反向索引。
            blocks = parse_body(archive, main_part, tracker, flags, styles, result) if main_part else []
            result["blocks"] = blocks

            comment_anchors: dict[str, list[str]] = {}
            footnote_anchors: dict[str, list[str]] = {}
            endnote_anchors: dict[str, list[str]] = {}
            for block in blocks:
                _index_anchors(block, comment_anchors, footnote_anchors, endnote_anchors)

            collect_annotations(
                archive,
                names,
                tracker,
                flags,
                comment_anchors,
                footnote_anchors,
                endnote_anchors,
                result,
            )

    except zipfile.BadZipFile as error:
        result["error"] = {"kind": "BAD_ZIP", "message": str(error)}
        return result
    except (OSError, EOFError, ValueError, NotImplementedError) as error:
        result["error"] = {"kind": "READ_FAILED", "message": str(error)}
        return result

    if not result["blocks"] and result["error"] is None:
        result["issues"].append(
            {"code": "EMPTY_DOCUMENT", "message": "Document body contains no visible blocks."}
        )

    result["limitHit"] = tracker.first_hit
    if tracker.first_hit is not None:
        result["issues"].append(
            {
                "code": "LIMIT_REACHED",
                "message": f"Parsing stopped accounting after hitting {tracker.first_hit}.",
            }
        )

    return result


def _index_anchors(
    block: dict[str, Any],
    comments: dict[str, list[str]],
    footnotes: dict[str, list[str]],
    endnotes: dict[str, list[str]],
) -> None:
    """把块内的注释引用登记到「注释 id -> 块指针」的反向索引。

    对表格块，锚点登记到【单元格内段落的指针】，这样批注能精确落到那一格，
    而不是含糊地指向整张表。
    """
    if block["kind"] == "paragraph":
        _record_refs(block, comments, footnotes, endnotes)
        return

    for row in block["rows"]:
        for cell in row["cells"]:
            for paragraph in cell["paragraphs"]:
                _record_refs(paragraph, comments, footnotes, endnotes)


def _record_refs(
    paragraph: dict[str, Any],
    comments: dict[str, list[str]],
    footnotes: dict[str, list[str]],
    endnotes: dict[str, list[str]],
) -> None:
    for ref in paragraph["commentRefs"]:
        comments.setdefault(ref, []).append(paragraph["pointer"])
    for ref in paragraph["footnoteRefs"]:
        footnotes.setdefault(ref, []).append(paragraph["pointer"])
    for ref in paragraph["endnoteRefs"]:
        endnotes.setdefault(ref, []).append(paragraph["pointer"])


def main() -> int:
    parser = argparse.ArgumentParser(description="Parse a DOCX into raw structural observations")
    parser.add_argument("--path", required=True, help="Path to the artifact")
    parser.add_argument("--config", required=True, help="JSON config: {limits, featureFlags}")
    args = parser.parse_args()

    try:
        config = json.loads(args.config)
    except json.JSONDecodeError as error:
        print(f"invalid --config payload: {error}", file=sys.stderr)
        return 2

    limits = config.get("limits", {}) if isinstance(config, dict) else {}
    flags = config.get("featureFlags", {}) if isinstance(config, dict) else {}

    try:
        result = parse(args.path, limits, flags)
    except Exception as error:  # noqa: BLE001 - 兜底：任何未预期异常都变成可解析载荷
        result = empty_result()
        result["error"] = {"kind": "READ_FAILED", "message": f"{type(error).__name__}: {error}"}

    if result.get("sizeBytes") == 0:
        try:
            result["sizeBytes"] = os.path.getsize(args.path)
        except OSError:
            result["sizeBytes"] = 0
    if not result.get("sha256"):
        try:
            result["sha256"] = sha256_of_file(args.path)
        except OSError:
            result["sha256"] = ""

    # 先写 stderr（诊断），再写 stdout（协议载荷），避免二者交错。
    print(json.dumps(result, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
