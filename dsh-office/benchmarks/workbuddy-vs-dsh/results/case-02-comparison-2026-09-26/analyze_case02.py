from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W, "w14": W14, "r": R, "pr": PR}


def qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def node_text(node: etree._Element, mode: str = "final") -> str:
    pieces: list[str] = []
    for item in node.iter():
        if item.tag == qn(W, "delText"):
            if mode in {"original", "all"}:
                pieces.append(item.text or "")
        elif item.tag == qn(W, "t"):
            in_insertion = any(a.tag == qn(W, "ins") for a in item.iterancestors())
            if mode == "original" and in_insertion:
                continue
            pieces.append(item.text or "")
        elif item.tag == qn(W, "tab"):
            pieces.append("\t")
        elif item.tag == qn(W, "br"):
            pieces.append("\n")
    return "".join(pieces)


def para_record(p: etree._Element) -> dict:
    return {
        "para_id": p.get(qn(W14, "paraId")),
        "original": node_text(p, "original"),
        "final": node_text(p, "final"),
        "comment_start_ids": [x.get(qn(W, "id")) for x in p.xpath(".//w:commentRangeStart", namespaces=NS)],
        "comment_end_ids": [x.get(qn(W, "id")) for x in p.xpath(".//w:commentRangeEnd", namespaces=NS)],
        "comment_ref_ids": [x.get(qn(W, "id")) for x in p.xpath(".//w:commentReference", namespaces=NS)],
    }


def table_records(document: etree._Element) -> list[dict]:
    records: list[dict] = []
    for table_index, table in enumerate(document.xpath("//w:body/w:tbl", namespaces=NS)):
        rows = []
        for row_index, row in enumerate(table.xpath("./w:tr", namespaces=NS)):
            cells = []
            for cell in row.xpath("./w:tc", namespaces=NS):
                cells.append({"original": node_text(cell, "original"), "final": node_text(cell, "final")})
            rows.append(
                {
                    "row": row_index,
                    "is_repeat_header": bool(row.xpath("./w:trPr/w:tblHeader", namespaces=NS)),
                    "cells": cells,
                }
            )
        records.append({"table": table_index, "rows": rows})
    return records


def analyze(path: Path) -> dict:
    raw = path.read_bytes()
    with ZipFile(path) as zf:
        names = set(zf.namelist())
        document = etree.fromstring(zf.read("word/document.xml"))
        rels = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
        comments = etree.fromstring(zf.read("word/comments.xml")) if "word/comments.xml" in names else None
        footnotes = etree.fromstring(zf.read("word/footnotes.xml")) if "word/footnotes.xml" in names else None

        revisions = []
        for kind in ("del", "ins"):
            for item in document.xpath(f"//w:{kind}", namespaces=NS):
                revisions.append(
                    {
                        "kind": kind,
                        "id": item.get(qn(W, "id")),
                        "author": item.get(qn(W, "author")),
                        "date": item.get(qn(W, "date")),
                        "text": node_text(item, "all"),
                    }
                )

        comment_records = []
        if comments is not None:
            paragraphs = document.xpath("//w:p", namespaces=NS)
            for comment in comments.xpath("//w:comment", namespaces=NS):
                cid = comment.get(qn(W, "id"))
                anchors = []
                for p in paragraphs:
                    ids = [x.get(qn(W, "id")) for x in p.xpath(".//w:commentRangeStart", namespaces=NS)]
                    if cid in ids:
                        anchors.append(para_record(p))
                comment_records.append(
                    {
                        "id": cid,
                        "author": comment.get(qn(W, "author")),
                        "date": comment.get(qn(W, "date")),
                        "text": node_text(comment, "all"),
                        "anchors": anchors,
                    }
                )

        footnote_records = []
        if footnotes is not None:
            for note in footnotes.xpath("//w:footnote", namespaces=NS):
                footnote_records.append({"id": note.get(qn(W, "id")), "text": node_text(note, "all")})

        settings = zf.read("word/settings.xml") if "word/settings.xml" in names else b""
        body_paragraphs = [para_record(p) for p in document.xpath("//w:body/w:p", namespaces=NS)]
        return {
            "path": str(path),
            "sha256": hashlib.sha256(raw).hexdigest(),
            "zip_ok": zf.testzip() is None,
            "parts": sorted(names),
            "body_paragraphs": body_paragraphs,
            "tables": table_records(document),
            "revisions": revisions,
            "comments": comment_records,
            "footnotes": footnote_records,
            "footnote_refs": [x.get(qn(W, "id")) for x in document.xpath("//w:footnoteReference", namespaces=NS)],
            "hyperlinks": [
                {"id": rel.get("Id"), "target": rel.get("Target"), "mode": rel.get("TargetMode")}
                for rel in rels.xpath("//pr:Relationship", namespaces=NS)
                if rel.get("Type", "").endswith("/hyperlink")
            ],
            "track_revisions": b"trackRevisions" in settings,
            "comment_anchor_counts": {
                "start": len(document.xpath("//w:commentRangeStart", namespaces=NS)),
                "end": len(document.xpath("//w:commentRangeEnd", namespaces=NS)),
                "ref": len(document.xpath("//w:commentReference", namespaces=NS)),
            },
        }


def main() -> None:
    args = list(sys.argv[1:])
    output = None
    if "--out" in args:
        index = args.index("--out")
        output = Path(args[index + 1])
        del args[index : index + 2]
    if not args:
        raise SystemExit("usage: analyze_case02.py [--out FILE] FILE [FILE ...]")
    result = {Path(p).parent.name: analyze(Path(p)) for p in args}
    encoded = json.dumps(result, ensure_ascii=False, indent=2)
    if output:
        output.write_text(encoded + "\n", encoding="utf-8")
    else:
        print(encoded)


if __name__ == "__main__":
    main()
