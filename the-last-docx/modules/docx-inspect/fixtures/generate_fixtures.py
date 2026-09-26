#!/usr/bin/env python3
"""为 docx-inspect 生成测试用 fixture。

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

# 常用命名空间与 Content-Type 常量。
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_XML = "application/xml"
CT_DOC_MAIN = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
CT_DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml"
CT_VBA = "application/vnd.ms-office.vbaproject"
CT_OLE_OBJECT = "application/vnd.openxmlformats-officedocument.oleobject"
CT_ALT_CHUNK = "application/vnd.openxmlformats-officedocument.wordprocessingml.altchunk"

REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_ATTACHED_TEMPLATE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
)
REL_HYPERLINK = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def content_types(overrides: dict[str, str]) -> str:
    """构造 `[Content_Types].xml`。"""
    default_types = {
        "rels": CT_RELS,
        "xml": CT_XML,
    }
    defaults = "".join(
        f'<Default Extension="{ext}" ContentType="{ct}"/>'
        for ext, ct in sorted(default_types.items())
    )
    parts = "".join(
        f'<Override PartName="{name}" ContentType="{ct}"/>'
        for name, ct in sorted(overrides.items())
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CT_NS}">{defaults}{parts}</Types>'
    )


def relationships(entries: Iterable[tuple[str, str, str, bool]]) -> str:
    """构造 `.rels` 内容。

    entries 为 (id, type, target, is_external) 四元组序列。
    """
    body = "".join(
        f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"'
        + (' TargetMode="External"' if external else "")
        + "/>"
        for rel_id, rel_type, target, external in entries
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL_NS}">{body}</Relationships>'


def document_xml(body: str) -> str:
    """把正文片段包进 `<w:document>`。"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>'
    )


def paragraph(text: str) -> str:
    """构造一个普通段落。"""
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def write_package(path: str, entries: dict[str, str | bytes]) -> None:
    """把 entries 写成一个 ZIP 包（值可以是 str 或 bytes）。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            payload = data.encode("utf-8") if isinstance(data, str) else data
            archive.writestr(name, payload)


def root_rels() -> str:
    """包根关系：指向主文档。"""
    return relationships([("rId1", REL_OFFICE_DOCUMENT, "word/document.xml", False)])


def build_clean(path: str) -> None:
    """正常文档：一段可见文字。"""
    write_package(
        path,
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Hello docx-inspect")),
        },
    )


def build_empty(path: str) -> None:
    """边界：结构合法但没有任何可见内容（只有一个空段落）。"""
    write_package(
        path,
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml("<w:p/>"),
        },
    )


def build_macro_enabled(path: str) -> None:
    """宏文档（.docm）：包含 vbaProject.bin，且主部件声明宏启用类型。"""
    write_package(
        path,
        {
            "[Content_Types].xml": content_types(
                {
                    "/word/document.xml": CT_DOCM_MAIN,
                    "/word/vbaProject.bin": CT_VBA,
                }
            ),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Macro enabled")),
            # 内容本身无关紧要：本模块只做「存在性」探测。
            # 刻意不含 `MZ` 头的伪 PE 特征，避免被杀软误判为可执行体。
            "word/vbaProject.bin": b"\x00\x01fake-vba-project",
        },
    )


def build_external(path: str) -> None:
    """外链文档：远程模板、超链接、外链图片，外加一个 DDE 字段。

    注意：这里刻意【不】使用 `cmd.exe` / `.exe` 之类载荷。
    Windows Defender 会把「DDEAUTO + 可执行文件路径」识别为漏洞利用特征并隔离样本，
    导致 fixture 在测试机上不可读。本模块只需要「存在 DDE 字段」这一事实，
    因此用一个无害的服务器/主题组合即可，测试结果与真实恶意样本完全一致。
    """
    doc_rels = relationships(
        [
            ("rId1", REL_ATTACHED_TEMPLATE, "https://evil.example/template.dotm", True),
            ("rId2", REL_HYPERLINK, "http://evil.example/click", True),
            ("rId3", REL_IMAGE, "https://evil.example/pixel.png", True),
        ]
    )
    dde = (
        '<w:p><w:fldSimple w:instr="DDEAUTO lines.example.invalid &quot;topic&quot;">'
        "<w:r><w:t>field</w:t></w:r></w:fldSimple></w:p>"
    )
    write_package(
        path,
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/_rels/document.xml.rels": doc_rels,
            "word/document.xml": document_xml(paragraph("External refs") + dde),
        },
    )


def build_embedded(path: str) -> None:
    """嵌入对象文档：OLE 嵌入 + ActiveX + altChunk。"""
    write_package(
        path,
        {
            "[Content_Types].xml": content_types(
                {
                    "/word/document.xml": CT_DOC_MAIN,
                    "/word/embeddings/oleObject1.bin": CT_OLE_OBJECT,
                    "/word/afchunk1.html": CT_ALT_CHUNK,
                }
            ),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Embedded")),
            "word/embeddings/oleObject1.bin": b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1fake-ole",
            # ActiveX 部件不需要 Content-Type：路径即可判定。
            "word/activex/activeX1.bin": b"\x00fake-activex",
            "word/afchunk1.html": "<html><body>chunk</body></html>",
        },
    )


def build_encrypted(path: str) -> None:
    """加密标记文档：含 EncryptedPackage 条目。

    注意：真实加密 OOXML 是 OLE 容器；此处只造一个 ZIP 形态的最小样本，
    用于验证「加密标记」这条路是否被正确识别。
    """
    write_package(
        path,
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Encrypted")),
            "EncryptedPackage": b"\x00\x01\x02encrypted-payload",
        },
    )


def build_bomb(path: str, size_bytes: int = 4 * 1024 * 1024) -> None:
    """解压炸弹样本：一个高度可压缩的超大条目。

    配合测试里调小的 `maxEntryUncompressedBytes` 使用，
    验证「流式读取 + 硬字节上限」是否真的阻断了解压。
    """
    write_package(
        path,
        {
            "[Content_Types].xml": content_types({"/word/document.xml": CT_DOC_MAIN}),
            "_rels/.rels": root_rels(),
            "word/document.xml": document_xml(paragraph("Bomb")),
            "word/bomb.bin": b"\x00" * size_bytes,
        },
    )


def build_not_word(path: str) -> None:
    """非 Word 的 OOXML 包：声明为 SpreadsheetML，应被判定为格式不匹配。"""
    spreadsheet_main = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    write_package(
        path,
        {
            "[Content_Types].xml": content_types({"/xl/workbook.xml": spreadsheet_main}),
            "_rels/.rels": relationships(
                [
                    (
                        "rId1",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                        "xl/workbook.xml",
                        False,
                    )
                ]
            ),
            "xl/workbook.xml": '<?xml version="1.0"?><workbook/>',
        },
    )


def build_truncated(path: str) -> None:
    """损坏样本：ZIP 魔数 + 垃圾字节，无法作为 ZIP 打开。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(b"PK\x03\x04")
        handle.write(b"this is not a real zip archive" * 4)


def build_rtf(path: str) -> None:
    """RTF 样本：容器可识别，但超出本模块范围。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(b"{\\rtf1\\ansi Hello}")


def build_legacy_ole(path: str) -> None:
    """遗留 OLE 样本（.doc）：OLE 魔数 + 填充。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        handle.write(b"\x00" * 64)


def build_plain(path: str) -> None:
    """完全无关的二进制样本：既不是 ZIP 也不是 Office 容器。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(b"just some bytes, not a document at all")


def build_no_content_types(path: str) -> None:
    """缺 `[Content_Types].xml` 的 ZIP：不是合法 OPC 包。"""
    write_package(path, {"readme.txt": "no content types here"})


BUILDERS = {
    "clean.docx": build_clean,
    "empty.docx": build_empty,
    "macro.docm": build_macro_enabled,
    "external.docx": build_external,
    "embedded.docx": build_embedded,
    "encrypted.docx": build_encrypted,
    "bomb.docx": build_bomb,
    "not-word.docx": build_not_word,
    "truncated.docx": build_truncated,
    "legacy.doc": build_legacy_ole,
    "plain.bin": build_plain,
    "no-content-types.zip": build_no_content_types,
    "sample.rtf": build_rtf,
}


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate docx-inspect test fixtures")
    parser.add_argument("--out", required=True, help="Output directory")
    args = parser.parse_args()

    for filename, builder in BUILDERS.items():
        builder(os.path.join(args.out, filename))
        print(f"generated {filename}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
