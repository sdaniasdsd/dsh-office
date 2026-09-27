#!/usr/bin/env python3
"""为 docx-parse 生成测试用 fixture。

设计要点：
  * 只依赖标准库（zipfile），无需安装 python-docx 或 lxml 即可重建全部样本；
  * 每个样本都是「最小但结构正确」的 OOXML 包，便于对 IR 做精确断言；
  * 输出到 `fixtures/_generated/`（.gitignore 已忽略），仓库里只保留这一份
    可审计的生成器源码，而不是难以解释的二进制样本。

样本清单与覆盖场景：
  clean.docx            正常文档：标题 + 段落 + 表格 + 样式 + 关系 + 元数据
                        并含真实 Word 会有的标记类块级元素（sectPr/书签/校对标记）
  unknown-elements.docx 含公式段落：本模块不解析，须上报 UNKNOWN_BLOCK_SKIPPED
  headings.docx         多级标题：验证 outline 派生
  tables.docx           含合并单元格的表格
  empty.docx            只有空段落：EMPTY_DOCUMENT
  broken-style.docx     段落引用了未定义的样式：BROKEN_STYLE_REFERENCE
  no-styles.docx        缺少 word/styles.xml：STYLES_MISSING
  dangling.docx         内部关系指向不存在的部件：DANGLING_RELATIONSHIP
  comments.docx         含批注
  footnotes.docx        含脚注与尾注
  many-blocks.docx      60 个段落：用于触发 maxBlocks 预算
  macro.docm            宏启用文档变体
  encrypted.docx        ZIP 内含 EncryptionInfo，用于加密三态检测
  spreadsheet.docx      ZIP 但实为 xlsx：FORMAT_MISMATCH
  no-content-types.docx ZIP 缺 [Content_Types].xml：FORMAT_MISMATCH
  truncated.docx        被截断的 zip：PARSE_FAILED
  legacy.doc            OLE/CFB 魔数：UNSUPPORTED_CONTAINER
  sample.rtf            RTF 魔数：UNSUPPORTED_CONTAINER
  plain.bin             未知字节：FORMAT_MISMATCH

用法：
    python fixtures/generate_fixtures.py --out fixtures/_generated
"""

from __future__ import annotations

import argparse
import io
import os
import zipfile
from typing import Iterable, Sequence

# --- 命名空间与 Content-Type 常量 ----------------------------------------- #

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"

CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_XML = "application/xml"
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
CT_DOC_MAIN = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
CT_DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
CT_XLSX_MAIN = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)

REL_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_OFFICE_DOCUMENT = f"{REL_OFFICE}/officeDocument"
REL_STYLES = f"{REL_OFFICE}/styles"
REL_COMMENTS = f"{REL_OFFICE}/comments"
REL_FOOTNOTES = f"{REL_OFFICE}/footnotes"
REL_ENDNOTES = f"{REL_OFFICE}/endnotes"
REL_HYPERLINK = f"{REL_OFFICE}/hyperlink"
REL_IMAGE = f"{REL_OFFICE}/image"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

# 关系四元组：(id, type, target, is_external)
RelEntry = tuple[str, str, str, bool]

# 样式五元组：(styleId, name, type, basedOn, isDefault)
StyleEntry = tuple[str, str, str, "str | None", bool]

DEFAULT_STYLES: tuple[StyleEntry, ...] = (
    ("Normal", "Normal", "paragraph", None, True),
    ("Heading1", "heading 1", "paragraph", "Normal", False),
    ("Heading2", "heading 2", "paragraph", "Normal", False),
    ("Heading3", "heading 3", "paragraph", "Normal", False),
    ("BodyText", "Body Text", "paragraph", "Normal", False),
    ("TableGrid", "Table Grid", "table", None, False),
    ("Strong", "Strong", "character", None, False),
)


# --- XML 构造小工具 -------------------------------------------------------- #


def content_types(overrides: dict[str, str]) -> str:
    """构造 `[Content_Types].xml`。"""
    defaults = {"rels": CT_RELS, "xml": CT_XML}
    defaults_xml = "".join(
        f'<Default Extension="{ext}" ContentType="{ct}"/>'
        for ext, ct in sorted(defaults.items())
    )
    overrides_xml = "".join(
        f'<Override PartName="{part}" ContentType="{ct}"/>'
        for part, ct in sorted(overrides.items())
    )
    return f'{XML_DECL}<Types xmlns="{CT_NS}">{defaults_xml}{overrides_xml}</Types>'


def rels_xml(entries: Iterable[RelEntry]) -> str:
    """构造 `.rels` 内容。"""
    body = "".join(
        f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"'
        + (' TargetMode="External"' if external else "")
        + "/>"
        for rel_id, rel_type, target, external in entries
    )
    return f'{XML_DECL}<Relationships xmlns="{REL_NS}">{body}</Relationships>'


# 真实 Word 文档的 `w:body` 末尾【恒定】带一个 `w:sectPr`（页面尺寸/页边距），
# 它不承载正文内容。fixture 必须保留它，否则「未知块级元素」的误报风险
# 在测试里永远暴露不出来。
SECT_PR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'

# 数学命名空间：`m:oMathPara`（公式段落）是本模块不解析的正文内容。
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def document_xml(body: str) -> str:
    """构造 `word/document.xml`。"""
    return (
        f'{XML_DECL}<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}">'
        f"<w:body>{body}{SECT_PR}</w:body></w:document>"
    )


def p(
    text: str = "",
    *,
    style: str | None = None,
    outline: int | None = None,
    numbered: bool = False,
    align: str | None = None,
) -> str:
    """构造一个 `w:p`；text 为空时只输出空段落。"""
    props: list[str] = []
    if style:
        props.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        props.append(f'<w:jc w:val="{align}"/>')
    if outline is not None:
        props.append(f'<w:outlineLvl w:val="{outline}"/>')
    if numbered:
        props.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    props_xml = f'<w:pPr>{"".join(props)}</w:pPr>' if props else ""
    run_xml = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{props_xml}{run_xml}</w:p>"


def tc(text: str, *, span: int = 1, vmerge: bool = False) -> str:
    """构造一个 `w:tc`。"""
    props: list[str] = []
    if span != 1:
        props.append(f'<w:gridSpan w:val="{span}"/>')
    if vmerge:
        props.append('<w:vMerge w:val="restart"/>')
    props_xml = f'<w:tcPr>{"".join(props)}</w:tcPr>' if props else ""
    return f"<w:tc>{props_xml}{p(text)}</w:tc>"


def table(rows: Sequence[Sequence[str]], *, style: str | None = None) -> str:
    """构造一个 `w:tbl`。"""
    props_xml = f'<w:tblPr><w:tblStyle w:val="{style}"/></w:tblPr>' if style else ""
    columns = max((sum(1 for _ in row) for row in rows), default=0)
    grid_xml = f'<w:tblGrid>{"".join("<w:gridCol/>" for _ in range(columns))}</w:tblGrid>'
    rows_xml = "".join(f'<w:tr>{"".join(row)}</w:tr>' for row in rows)
    return f"<w:tbl>{props_xml}{grid_xml}{rows_xml}</w:tbl>"


def styles_xml(styles: Sequence[StyleEntry] = DEFAULT_STYLES) -> str:
    """构造 `word/styles.xml`。"""
    parts: list[str] = []
    for style_id, name, style_type, based_on, is_default in styles:
        attrs = f'w:type="{style_type}" w:styleId="{style_id}"'
        if is_default:
            attrs += ' w:default="1"'
        inner = f'<w:name w:val="{name}"/>'
        if based_on:
            inner += f'<w:basedOn w:val="{based_on}"/>'
        parts.append(f"<w:style {attrs}>{inner}</w:style>")
    return f'{XML_DECL}<w:styles xmlns:w="{W_NS}">{"".join(parts)}</w:styles>'


def comments_xml(entries: Sequence[tuple[str, str, str]]) -> str:
    """构造 `word/comments.xml`。entries 为 (id, author, text)。"""
    parts = "".join(
        f'<w:comment w:id="{cid}" w:author="{author}" w:date="2024-01-01T00:00:00Z">'
        f"{p(text)}</w:comment>"
        for cid, author, text in entries
    )
    return f'{XML_DECL}<w:comments xmlns:w="{W_NS}">{parts}</w:comments>'


def notes_xml(kind: str, entries: Sequence[tuple[str, str]]) -> str:
    """构造 footnotes/endnotes 部件。entries 为 (id, text)。"""
    tag = "footnote" if kind == "footnote" else "endnote"
    # 前置的分隔符注释：引擎应过滤掉 type 非 normal 的项。
    separator = f'<w:{tag} w:type="separator" w:id="-1">{p("")}</w:{tag}>'
    parts = "".join(
        f'<w:{tag} w:id="{nid}">{p(text)}</w:{tag}>' for nid, text in entries
    )
    return f'{XML_DECL}<w:{kind}s xmlns:w="{W_NS}">{separator}{parts}</w:{kind}s>'


def core_xml(title: str, creator: str) -> str:
    """构造 `docProps/core.xml`。"""
    return (
        f'{XML_DECL}<cp:coreProperties xmlns:cp="{CP_NS}" xmlns:dc="{DC_NS}">'
        f"<dc:title>{title}</dc:title>"
        f"<dc:creator>{creator}</dc:creator>"
        f"<cp:lastModifiedBy>{creator}</cp:lastModifiedBy>"
        "</cp:coreProperties>"
    )


# --- 包构造 ----------------------------------------------------------------- #


class Package:
    """按名字累积条目，最后写出一个连续 zip。"""

    def __init__(self, *, compress: bool = True) -> None:
        self.entries: dict[str, bytes] = {}
        self.overrides: dict[str, str] = {}
        self.compress = compress

    def add(self, name: str, data: "str | bytes") -> None:
        self.entries[name] = data.encode("utf-8") if isinstance(data, str) else data

    def override(self, part: str, content_type: str) -> None:
        self.overrides[part] = content_type

    def add_rels(self, rels_part: str, entries: Iterable[RelEntry]) -> None:
        self.add(rels_part, rels_xml(entries))

    def to_bytes(self) -> bytes:
        # Content-Types 必须最后写入（此前各步骤只登记 override）。
        self.entries["[Content_Types].xml"] = content_types(self.overrides).encode("utf-8")
        buffer = io.BytesIO()
        method = zipfile.ZIP_DEFLATED if self.compress else zipfile.ZIP_STORED
        # 排序写入：让同一份 fixture 每次生成都得到完全一致的字节。
        with zipfile.ZipFile(buffer, "w", method) as archive:
            for name in sorted(self.entries):
                archive.writestr(name, self.entries[name])
        return buffer.getvalue()


def build_word_package(
    body: str,
    *,
    main_ct: str = CT_DOC_MAIN,
    styles: "str | None" = None,
    comments: "str | None" = None,
    footnotes: "str | None" = None,
    endnotes: "str | None" = None,
    core: "str | None" = None,
    document_rels: Sequence[RelEntry] = (),
    extra_overrides: "dict[str, str] | None" = None,
    extra_entries: "dict[str, bytes] | None" = None,
) -> bytes:
    """组装一个最小可用（或刻意残缺）的 WordprocessingML 包。"""
    package = Package()
    package.override("/word/document.xml", main_ct)
    package.add("word/document.xml", document_xml(body))

    root_rels: list[RelEntry] = [("rId1", REL_OFFICE_DOCUMENT, "word/document.xml", False)]
    if core is not None:
        package.override("/docProps/core.xml", CT_CORE)
        package.add("docProps/core.xml", core)
        root_rels.append(("rIdCore", REL_CORE, "docProps/core.xml", False))
    package.add_rels("_rels/.rels", root_rels)

    document_rel_entries: list[RelEntry] = list(document_rels)
    if styles is not None:
        package.override("/word/styles.xml", CT_STYLES)
        package.add("word/styles.xml", styles)
        document_rel_entries.append(("rIdStyles", REL_STYLES, "styles.xml", False))
    if comments is not None:
        package.override("/word/comments.xml", CT_COMMENTS)
        package.add("word/comments.xml", comments)
        document_rel_entries.append(("rIdComments", REL_COMMENTS, "comments.xml", False))
    if footnotes is not None:
        package.override("/word/footnotes.xml", CT_FOOTNOTES)
        package.add("word/footnotes.xml", footnotes)
        document_rel_entries.append(("rIdFootnotes", REL_FOOTNOTES, "footnotes.xml", False))
    if endnotes is not None:
        package.override("/word/endnotes.xml", CT_ENDNOTES)
        package.add("word/endnotes.xml", endnotes)
        document_rel_entries.append(("rIdEndnotes", REL_ENDNOTES, "endnotes.xml", False))
    if document_rel_entries:
        package.add_rels("word/_rels/document.xml.rels", document_rel_entries)

    if extra_overrides:
        for part, content_type in extra_overrides.items():
            package.override(part, content_type)
    if extra_entries:
        for name, data in extra_entries.items():
            package.add(name, data)

    return package.to_bytes()


# --- 各样本 ----------------------------------------------------------------- #


def fixture_bytes() -> dict[str, bytes]:
    """返回 `文件名 -> 字节` 的映射。"""
    fixtures: dict[str, bytes] = {}

    # 正常文档：标题 + 两段正文（其一使用命名样式）+ 2x2 表格。
    # 刻意带上真实 Word 常见的标记类块级元素（校对标记、书签）：它们不含文本，
    # 必须被静默忽略，而不是报成「未解析的正文」（见 unknown-elements.docx 的反例）。
    fixtures["clean.docx"] = build_word_package(
        '<w:proofErr w:type="spellStart"/>'
        + p("Clean Fixture", style="Heading1")
        + '<w:proofErr w:type="spellEnd"/>'
        + '<w:bookmarkStart w:id="1" w:name="_Toc1"/>'
        + p("Hello world.", style="BodyText")
        + '<w:bookmarkEnd w:id="1"/>'
        + p("Second paragraph.")
        + table(
            [[tc("A1"), tc("B1")], [tc("A2"), tc("B2")]],
            style="TableGrid",
        ),
        styles=styles_xml(),
        core=core_xml("Clean Fixture", "docx-parse tests"),
    )

    # 含本模块不解析的正文内容：公式段落。它必须被上报为 UNKNOWN_BLOCK_SKIPPED，
    # 而不是静默丢弃——否则调用方无从得知「抽取到的文本 != 文档全部文本」。
    fixtures["unknown-elements.docx"] = build_word_package(
        p("Before equation.")
        + "<m:oMathPara><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>"
        + p("After equation."),
        styles=styles_xml(),
    )

    # 多级标题：用于验证 outline 的层级与顺序。
    fixtures["headings.docx"] = build_word_package(
        p("Chapter One", style="Heading1")
        + p("Intro text.")
        + p("Section 1.1", style="Heading2")
        + p("Body under section.")
        + p("Subsection 1.1.1", style="Heading3"),
        styles=styles_xml(),
    )

    # 合并单元格：跨两列的单元格 + 纵向合并起始格。
    fixtures["tables.docx"] = build_word_package(
        p("Merged cells")
        + table(
            [
                [tc("Spanned", span=2)],
                [tc("Top", vmerge=True), tc("Right")],
                [tc("Bottom"), tc("Right2")],
            ],
            style="TableGrid",
        ),
        styles=styles_xml(),
    )

    # 只有空段落：结构合法但没有可见内容。
    fixtures["empty.docx"] = build_word_package(p(), styles=styles_xml())

    # 引用了未定义的样式。
    fixtures["broken-style.docx"] = build_word_package(
        p("No such style", style="MissingStyle"), styles=styles_xml()
    )

    # 缺少 styles.xml，但段落仍引用样式。
    fixtures["no-styles.docx"] = build_word_package(p("Styled but unstyled", style="Heading1"))

    # 内部关系指向包内不存在的部件。
    fixtures["dangling.docx"] = build_word_package(
        p("Dangling relationship"),
        styles=styles_xml(),
        document_rels=[("rIdImg", REL_IMAGE, "media/image1.png", False)],
    )

    fixtures["comments.docx"] = build_word_package(
        p("Reviewed text", style="BodyText"),
        styles=styles_xml(),
        comments=comments_xml([("2", "Reviewer B", "Second comment"), ("1", "Reviewer A", "First comment")]),
    )

    fixtures["footnotes.docx"] = build_word_package(
        p("Text with notes"),
        styles=styles_xml(),
        footnotes=notes_xml("footnote", [("1", "First footnote"), ("2", "Second footnote")]),
        endnotes=notes_xml("endnote", [("1", "Only endnote")]),
    )

    # 60 个段落：配合极小的 maxBlocks 预算，触发 LIMIT_EXCEEDED / LIMIT_APPLIED。
    fixtures["many-blocks.docx"] = build_word_package(
        "".join(p(f"Paragraph {index}") for index in range(60)), styles=styles_xml()
    )

    # 宏启用变体。
    fixtures["macro.docm"] = build_word_package(
        p("Macro enabled", style="Heading1"), main_ct=CT_DOCM_MAIN, styles=styles_xml()
    )

    # 合成的「加密」样本：真实加密 OOXML 是 OLE 容器，这里用 ZIP 内含
    # EncryptionInfo 的方式，专门覆盖引擎的加密三态检测分支。
    fixtures["encrypted.docx"] = build_word_package(
        p("Encrypted placeholder"),
        styles=styles_xml(),
        extra_overrides={
            "/EncryptionInfo": "application/vnd.ms-office.encryptioninfo",
            "/EncryptedPackage": "application/vnd.ms-office.encryptedpackage",
        },
        extra_entries={
            "EncryptionInfo": b"\x04\x00\x04\x00",
            "EncryptedPackage": b"\x00" * 64,
        },
    )

    # ZIP 容器，但内容其实是 xlsx。
    spreadsheet = Package()
    spreadsheet.override("/xl/workbook.xml", CT_XLSX_MAIN)
    spreadsheet.add("xl/workbook.xml", f'{XML_DECL}<workbook/>')
    fixtures["spreadsheet.docx"] = spreadsheet.to_bytes()

    # ZIP 缺少 [Content_Types].xml：直接手写 zip，绕开 Package 的自动补全。
    no_types = io.BytesIO()
    with zipfile.ZipFile(no_types, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document_xml(p("No content types")))
    fixtures["no-content-types.docx"] = no_types.getvalue()

    # 被截断的 zip：只保留前 32 字节，足以被识别为 zip 魔数但无法解析。
    clean = fixtures["clean.docx"]
    fixtures["truncated.docx"] = clean[:32]

    # OLE/CFB 魔数（遗留 .doc 或加密 OOXML）。
    fixtures["legacy.doc"] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504

    # RTF 魔数。
    fixtures["sample.rtf"] = b"{\\rtf1\\ansi legacy rich text}"

    # 完全无法识别的字节。
    fixtures["plain.bin"] = b"this is definitely not an office document\n"

    return fixtures


def main(argv: "list[str] | None" = None) -> int:
    parser = argparse.ArgumentParser(description="Generate docx-parse fixtures")
    parser.add_argument("--out", required=True, help="Output directory")
    args = parser.parse_args(argv)

    os.makedirs(args.out, exist_ok=True)
    fixtures = fixture_bytes()
    for name, data in fixtures.items():
        with open(os.path.join(args.out, name), "wb") as handle:
            handle.write(data)
    # 打印清单，便于人工确认生成结果（vitest 会吞掉这段输出）。
    for name in sorted(fixtures):
        print(f"wrote {name} ({len(fixtures[name])} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
