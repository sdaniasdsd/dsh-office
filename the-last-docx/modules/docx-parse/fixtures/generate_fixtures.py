#!/usr/bin/env python3
"""为 docx-parse 生成测试用 fixture。

设计要点：
  * 只依赖标准库（zipfile），因此无需安装 python-docx 也能重建全部样本；
  * 每个样本都是「最小但结构正确」的 OOXML 包，便于精确断言；
  * 输出到 `fixtures/_generated/`（已在 .gitignore 中忽略），
    保证仓库里只保留「生成器」这一份可审计的源码，而不是二进制样本。

用法：
    python fixtures/generate_fixtures.py --out fixtures/_generated
"""

from __future__ import annotations

import argparse
import os
import zipfile
from typing import Iterable

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_XML = "application/xml"
CT_DOC_MAIN = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
CT_DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
CT_VBA = "application/vnd.ms-office.vbaproject"
CT_SHEET_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_HYPERLINK = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def content_types(overrides: dict[str, str]) -> str:
    defaults = "".join(
        f'<Default Extension="{ext}" ContentType="{ct}"/>'
        for ext, ct in sorted({"rels": CT_RELS, "xml": CT_XML}.items())
    )
    parts = "".join(
        f'<Override PartName="{name}" ContentType="{ct}"/>'
        for name, ct in sorted(overrides.items())
    )
    return f'{XML_DECL}<Types xmlns="{CT_NS}">{defaults}{parts}</Types>'


def relationships(entries: Iterable[tuple[str, str, str, bool]]) -> str:
    body = "".join(
        f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"'
        + (' TargetMode="External"' if external else "")
        + "/>"
        for rel_id, rel_type, target, external in entries
    )
    return f'{XML_DECL}<Relationships xmlns="{REL_NS}">{body}</Relationships>'


def document_xml(body: str) -> str:
    return (
        f'{XML_DECL}<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">'
        f"<w:body>{body}</w:body></w:document>"
    )


def paragraph(text: str, *, style_id: str | None = None, para_id: str | None = None) -> str:
    ppr = ""
    if style_id is not None:
        ppr = f'<w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>'
    para = f' w14:paraId="{para_id}"' if para_id is not None else ""
    return f"<w:p{para}>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def write_package(path: str, entries: dict[str, str | bytes]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in sorted(entries.items()):
            payload = data.encode("utf-8") if isinstance(data, str) else data
            archive.writestr(name, payload)


def root_rels() -> str:
    return relationships([("rId1", REL_OFFICE_DOCUMENT, "word/document.xml", False)])


def document_rels(extra: list[tuple[str, str, str, bool]] | None = None) -> str:
    entries: list[tuple[str, str, str, bool]] = [
        ("rId1", REL_STYLES, "styles.xml", False),
    ]
    if extra:
        entries.extend(extra)
    return relationships(entries)


# --------------------------------------------------------------------------- #
# 样本                                                                          #
# --------------------------------------------------------------------------- #


def build_clean(out: str) -> None:
    """单一普通段落。"""
    write_package(
        os.path.join(out, "clean.docx"),
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Hello docx-parse")),
        },
    )


def build_empty(out: str) -> None:
    """正文为空（只有必填的 sectPr）。"""
    write_package(
        os.path.join(out, "empty.docx"),
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml("<w:sectPr/>"),
        },
    )


def build_rich(out: str) -> None:
    """综合样本：标题、段落、表格、批注、脚注、样式继承链。

    这是双 IR 回归测试的主样本：它同时覆盖
      * 显式 outlineLvl（Heading1）、
      * 经 basedOn 继承得到 outlineLvl（Heading3 -> HeadingBase）、
      * 表格（含单元格内段落需各自登记锚点）、
      * 批注与脚注的锚点反查、
      * paraId 存在与缺失两种情形。
    """
    body = "".join(
        [
            paragraph("Chapter One", style_id="Heading1", para_id="1A2B3C4D"),
            paragraph("Intro text"),
            # 批注锚定在「批注段落」这一块上。
            '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Commented paragraph</w:t></w:r>'
            '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
            # 脚注锚定在「脚注段落」这一块上。
            '<w:p><w:r><w:t>Footnote paragraph</w:t></w:r>'
            '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
            "<w:tbl>"
            "<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
            "<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>"
            "<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:tbl>",
            paragraph("Section", style_id="Heading3"),
            paragraph("Closing line"),
        ]
    )
    styles = (
        f'{XML_DECL}<w:styles xmlns:w="{W_NS}">'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
        # Heading3 自身没有 outlineLvl，必须经 basedOn 链继承。
        '<w:style w:type="paragraph" w:styleId="Heading3">'
        '<w:name w:val="heading 3"/><w:basedOn w:val="HeadingBase"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="HeadingBase">'
        '<w:name w:val="Base Heading"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>'
        "</w:styles>"
    )
    comments = (
        f'{XML_DECL}<w:comments xmlns:w="{W_NS}">'
        '<w:comment w:id="1" w:author="Ada" w:date="2024-01-02T03:04:05Z">'
        "<w:p><w:r><w:t>A review note</w:t></w:r></w:p></w:comment></w:comments>"
    )
    footnotes = (
        f'{XML_DECL}<w:footnotes xmlns:w="{W_NS}">'
        # id 0 为分隔符、-1 为延续分隔符，二者都不应被当作真实脚注。
        '<w:footnote w:id="-1" w:type="separator"><w:p/></w:footnote>'
        '<w:footnote w:id="0" w:type="continuationSeparator"><w:p/></w:footnote>'
        '<w:footnote w:id="1"><w:p><w:r><w:t>A footnote</w:t></w:r></w:p></w:footnote>'
        "</w:footnotes>"
    )
    write_package(
        os.path.join(out, "rich.docx"),
        {
            "[Content_Types].xml": content_types(
                {
                    "/word/document.xml": CT_DOC_MAIN,
                    "/word/styles.xml": CT_STYLES,
                    "/word/comments.xml": CT_COMMENTS,
                    "/word/footnotes.xml": CT_FOOTNOTES,
                }
            ),
            "_rels/.rels": root_rels(),
            "word/_rels/document.xml.rels": document_rels(
                [("rId2", REL_COMMENTS, "comments.xml", False), ("rId3", REL_FOOTNOTES, "footnotes.xml", False)]
            ),
            "word/document.xml": document_xml(body),
            "word/styles.xml": styles,
            "word/comments.xml": comments,
            "word/footnotes.xml": footnotes,
        },
    )


def build_heading_by_name(out: str) -> None:
    """标题级别只能由样式名推断（没有任何 outlineLvl）。"""
    styles = (
        f'{XML_DECL}<w:styles xmlns:w="{W_NS}">'
        '<w:style w:type="paragraph" w:styleId="Heading2">'
        '<w:name w:val="Heading 2"/></w:style>'
        "</w:styles>"
    )
    write_package(
        os.path.join(out, "heading-by-name.docx"),
        {
            "[Content_Types].xml": content_types(
                {"/word/document.xml": CT_DOC_MAIN, "/word/styles.xml": CT_STYLES}
            ),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Named heading", style_id="Heading2")),
            "word/styles.xml": styles,
        },
    )


def build_external(out: str) -> None:
    """含一个外链关系的文档。"""
    write_package(
        os.path.join(out, "external.docx"),
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/_rels/document.xml.rels": document_rels(
                [("rId9", REL_HYPERLINK, "https://example.com/", True)]
            ),
            "word/document.xml": document_xml(paragraph("See the link")),
        },
    )


def build_macro(out: str) -> None:
    """启用宏的模板（.docm）。"""
    write_package(
        os.path.join(out, "macro.docm"),
        {
            "[Content_Types].xml": content_types(
                {"/word/document.xml": CT_DOCM_MAIN, "/word/vbaProject.bin": CT_VBA}
            ),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Macro enabled")),
            "word/vbaProject.bin": b"\x00\x01\x02macro",
        },
    )


def build_not_word(out: str) -> None:
    """合法的 OPC 包，但主部件是电子表格。"""
    write_package(
        os.path.join(out, "not-word.docx"),
        {
            "[Content_Types].xml": content_types({"/xl/workbook.xml": CT_SHEET_MAIN}),
            "_rels/.rels": relationships(
                [("rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml", False)]
            ),
            "xl/workbook.xml": f'{XML_DECL}<workbook/>',
        },
    )


def build_no_content_types(out: str) -> None:
    """ZIP 结构合法，但缺少 [Content_Types].xml。"""
    write_package(
        os.path.join(out, "no-content-types.zip"),
        {"word/document.xml": document_xml(paragraph("Orphan"))},
    )


def build_bomb(out: str) -> None:
    """解压炸弹：单个条目压缩后很小、解压后极大。

    测试时通过极小的 maxEntryUncompressedBytes 触发条目预算告警。
    """
    write_package(
        os.path.join(out, "bomb.docx"),
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("x" * 2_000_000)),
        },
    )


def build_plain_files(out: str) -> None:
    """非 OOXML 的杂项文件（用于容器识别分支）。"""
    os.makedirs(out, exist_ok=True)
    with open(os.path.join(out, "plain.bin"), "wb") as handle:
        handle.write(b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")
    # 合法但截断的 ZIP：从一份正常包上截掉尾部。
    full = os.path.join(out, "clean.docx")
    build_clean(out)
    with open(full, "rb") as handle:
        data = handle.read()
    with open(os.path.join(out, "truncated.docx"), "wb") as handle:
        handle.write(data[: len(data) // 2])
    # RTF 流。
    with open(os.path.join(out, "sample.rtf"), "wb") as handle:
        handle.write(b"{\\rtf1\\ansi Hello}")
    # OLE/CFB 容器（旧式 .doc 或加密包）。
    with open(os.path.join(out, "legacy.doc"), "wb") as handle:
        handle.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)


# --------------------------------------------------------------------------- #
# 黑盒样本：同源不同版的一对文档                                                #
# --------------------------------------------------------------------------- #

# v1 / v2 共用的样式表：一个带 outlineLvl 的标题样式 + 一个非标题样式。
_REVISION_STYLES = (
    f'{XML_DECL}<w:styles xmlns:w="{W_NS}">'
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
    '<w:style w:type="character" w:styleId="Emphasis">'
    '<w:name w:val="Emphasis"/></w:style>'
    "</w:styles>"
)


def _id_table(cells: list[list[tuple[str, str]]]) -> str:
    """构造一个单元格段落带 paraId 的表格。

    带 paraId 是刻意的：这样「只改了某一格」能落到那个单元格段落上，
    从而产生最小 diff，而不是含糊地指向整张表。
    """
    rows = ""
    for row in cells:
        tcs = "".join(
            f'<w:tc><w:p w14:paraId="{para_id}"><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>'
            for text, para_id in row
        )
        rows += f"<w:tr>{tcs}</w:tr>"
    return f"<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>{rows}</w:tbl>"


def _write_revision_doc(out: str, name: str, body: str) -> None:
    """写出一份带样式表的最小文档。"""
    write_package(
        os.path.join(out, name),
        {
            "[Content_Types].xml": content_types(
                {"/word/document.xml": CT_DOC_MAIN, "/word/styles.xml": CT_STYLES}
            ),
            "_rels/.rels": root_rels(),
            "word/_rels/document.xml.rels": document_rels(),
            "word/document.xml": document_xml(body),
            "word/styles.xml": _REVISION_STYLES,
        },
    )


def build_revision_pair(out: str) -> None:
    """一对同源文档（v1 -> v2），用于黑盒验证「实时新旧修改」。

    两版中所有段落都带 w14:paraId，因此节点身份不随位置漂移。
    这正是稳定身份方案发挥作用、也是推荐上游产出的形态。

    v2 相对 v1 的改动（刻意各覆盖一类）：
      * Revenue      : 纯文字改写（10 -> 12 percent）
      * Hiring resumed: 中途新增一段
      * Costs         : 一个字没改，只换了样式
      * 表格 135       : 只改了某一格
      * Quarterly Report / Outlook / 其余单元格: 完全没动
    """
    v1 = "".join(
        [
            paragraph("Quarterly Report", style_id="Heading1", para_id="1A2B0001"),
            paragraph("Revenue grew by 10 percent.", para_id="1A2B0002"),
            paragraph("Costs were flat.", para_id="1A2B0003"),
            paragraph("Outlook is stable.", para_id="1A2B0004"),
            _id_table(
                [
                    [("North", "1A2B0011"), ("120", "1A2B0012")],
                    [("South", "1A2B0013"), ("90", "1A2B0014")],
                ]
            ),
        ]
    )
    v2 = "".join(
        [
            paragraph("Quarterly Report", style_id="Heading1", para_id="1A2B0001"),
            paragraph("Revenue grew by 12 percent.", para_id="1A2B0002"),
            paragraph("Hiring resumed in Q3.", para_id="1A2B0009"),
            paragraph("Costs were flat.", style_id="Emphasis", para_id="1A2B0003"),
            paragraph("Outlook is stable.", para_id="1A2B0004"),
            _id_table(
                [
                    [("North", "1A2B0011"), ("135", "1A2B0012")],
                    [("South", "1A2B0013"), ("90", "1A2B0014")],
                ]
            ),
        ]
    )
    _write_revision_doc(out, "rev-v1.docx", v1)
    _write_revision_doc(out, "rev-v2.docx", v2)


def build_revision_drift_pair(out: str) -> None:
    """同样的修订场景，但段落【没有】paraId。

    用于黑盒暴露位置漂移的后果：id 只能从结构路径推导，
    中途插入一段会让后面所有段落的 id 全部错位，
    差异报告随之失真（看起来像「后面全被改了」）。
    """
    v1 = "".join(paragraph(text) for text in ("Alpha", "Beta", "Gamma"))
    v2 = "".join(paragraph(text) for text in ("Alpha", "New", "Beta", "Gamma"))
    _write_revision_doc(out, "rev-drift-v1.docx", v1)
    _write_revision_doc(out, "rev-drift-v2.docx", v2)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate docx-parse fixtures")
    parser.add_argument("--out", required=True, help="Output directory")
    args = parser.parse_args()

    out = args.out
    os.makedirs(out, exist_ok=True)

    for builder in (
        build_clean,
        build_empty,
        build_rich,
        build_heading_by_name,
        build_external,
        build_macro,
        build_not_word,
        build_no_content_types,
        build_bomb,
        build_plain_files,
        build_revision_pair,
        build_revision_drift_pair,
    ):
        builder(out)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
