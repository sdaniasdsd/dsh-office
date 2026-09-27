"""Build the deterministic multi-page DOCX input for case 02.

Requires the bundled Python runtime with python-docx and lxml. The generated
source contains stable paragraph ids, a real Word comment, a real footnote,
an external hyperlink, and a repeating-header table that spans pages.
"""

from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "fixtures" / "case-02-project-status.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "w14": W14, "r": R}


def set_cell_fill(cell, color: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shade = properties.find(qn("w:shd"))
    if shade is None:
        shade = OxmlElement("w:shd")
        properties.append(shade)
    shade.set(qn("w:fill"), color)


def set_cell_margins(cell, top=90, start=110, bottom=90, end=110) -> None:
    properties = cell._tc.get_or_add_tcPr()
    margins = properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.append(margins)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table) -> None:
    properties = table._tbl.tblPr
    borders = properties.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        properties.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        item = borders.find(qn(f"w:{edge}"))
        if item is None:
            item = OxmlElement(f"w:{edge}")
            borders.append(item)
        item.set(qn("w:val"), "single")
        item.set(qn("w:sz"), "4")
        item.set(qn("w:space"), "0")
        item.set(qn("w:color"), "D7DEE8")


def format_table(table, widths, header_rows=(0,), font_size=8.5) -> None:
    table.autofit = False
    table.allow_autofit = False
    table.style = "Table Grid"
    set_table_borders(table)
    for row_index, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        no_split = OxmlElement("w:cantSplit")
        tr_pr.append(no_split)
        for col_index, cell in enumerate(row.cells):
            cell.width = Inches(widths[col_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if row_index in header_rows:
                set_cell_fill(cell, "24476B")
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(1)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(font_size)
                    if row_index in header_rows:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
    for row_index in header_rows:
        tr_pr = table.rows[row_index]._tr.get_or_add_trPr()
        repeat = OxmlElement("w:tblHeader")
        repeat.set(qn("w:val"), "true")
        tr_pr.append(repeat)


def add_hyperlink(paragraph, text: str, url: str) -> None:
    relationship_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    props.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.append(underline)
    run.append(props)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_page_field(paragraph, instruction: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def set_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(36, 48, 61)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.12
    for name, size in (("Title", 24), ("Heading 1", 16), ("Heading 2", 12)):
        style = document.styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = name != "Title"
        style.font.color.rgb = RGBColor(24, 45, 66)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8 if name != "Title" else 0)
        style.paragraph_format.space_after = Pt(5)


def build_docx() -> Path:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    set_styles(document)
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)
    section.header_distance = Inches(0.34)
    section.footer_distance = Inches(0.32)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.paragraph_format.space_after = Pt(0)
    header.add_run("BLUEBIRD  /  DELIVERY OFFICE").font.size = Pt(8)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(0)
    footer.paragraph_format.space_after = Pt(0)
    footer.add_run("INTERNAL REVIEW  ·  PAGE ").font.size = Pt(8)
    add_page_field(footer, "PAGE")
    footer.add_run(" OF ").font.size = Pt(8)
    add_page_field(footer, "NUMPAGES")

    title = document.add_paragraph("Project delivery and acceptance status", style="Title")
    # The built-in Word Title style can carry a themed bottom border.
    for properties in (document.styles["Title"]._element.pPr, title._p.pPr):
        if properties is not None:
            border = properties.find(qn("w:pBdr"))
            if border is not None:
                properties.remove(border)
    subtitle = document.add_paragraph("Bluebird client platform migration  |  Q3 2026")
    subtitle.paragraph_format.space_after = Pt(3)
    for run in subtitle.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(85, 101, 119)
    meta = document.add_paragraph("Report date: 26 September 2026    |    Version: 1.4    |    Owner: Lin Zhou")
    meta.paragraph_format.space_after = Pt(12)
    for run in meta.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(102, 112, 122)

    document.add_heading("1. Executive summary", level=1)
    document.add_paragraph(
        "The migration remains within the approved release window. Current tracking shows "
        "18 work items: 16 completed on schedule; 2 remain open."
    )
    document.add_paragraph(
        "The data migration package is staged for client review. The acceptance-testing stream "
        "is separately awaiting client sign-off; its status is not changed by the migration receipt."
    )
    document.add_heading("Status snapshot", level=2)
    snapshot = document.add_table(rows=1, cols=3)
    for cell, text in zip(snapshot.rows[0].cells, ("Measure", "Current", "Evidence")):
        cell.text = text
    for row in (
        ("Total work items", "18", "Approved delivery register"),
        ("Completed on schedule", "16", "Daily workstream log"),
        ("Open items", "2", "Acceptance and remediation trackers"),
        ("Release target", "30 Sep 2026", "Approved migration plan"),
    ):
        cells = snapshot.add_row().cells
        for cell, text in zip(cells, row):
            cell.text = text
    format_table(snapshot, (2.35, 1.25, 3.25), header_rows=(0,), font_size=9)
    document.add_heading("Review notes", level=2)
    document.add_paragraph(
        "The release gate remains open until client acceptance testing is complete. "
        "A receipt for one stream does not close the full acceptance gate."
    )
    p = document.add_paragraph()
    p.add_run("Reference: ")
    add_hyperlink(p, "transfer runbook", "https://example.com/bluebird/transfer-runbook")
    p.add_run(" and the controlled release checklist.")

    document.add_page_break()
    document.add_heading("2. Workstream detail", level=1)
    document.add_paragraph(
        "The table records current owners, due dates, status and source evidence. The same status "
        "wording may appear in separate workstreams; update only the row named in an approved change."
    )
    detail = document.add_table(rows=2, cols=5)
    merged = detail.rows[0].cells[0].merge(detail.rows[0].cells[4])
    merged.text = "Delivery workstreams  ·  controlled status register"
    for cell, text in zip(detail.rows[1].cells, ("Workstream", "Owner", "Due", "Status", "Evidence / note")):
        cell.text = text
    detail_rows = (
        (
            "Data migration",
            "Jia Lin",
            "26 Sep 2026",
            "Awaiting client sign-off",
            "The transfer package is staged in the client repository. Receipt is recorded after the client verifies the release manifest.",
        ),
        (
            "Acceptance testing",
            "Rui Chen",
            "27 Sep 2026",
            "Awaiting client sign-off",
            "The acceptance suite is complete; the client has not yet returned the signed test record.",
        ),
        (
            "API integration",
            "Min Zhou",
            "23 Sep 2026",
            "Complete",
            "Health-check output and interface test evidence are retained with the build record.",
        ),
        (
            "User training",
            "Anne Wu",
            "25 Sep 2026",
            "Complete",
            "Four sessions were delivered. Attendance records are filed in the project evidence folder.",
        ),
        (
            "Ledger reconciliation",
            "Han Li",
            "28 Sep 2026",
            "In progress",
            "The finance team is matching the final ledger export against the approved migration totals.",
        ),
        (
            "Access review",
            "Jia Lin",
            "28 Sep 2026",
            "Complete",
            "Role membership was reviewed against the client-approved access matrix.",
        ),
        (
            "Rollback rehearsal",
            "Min Zhou",
            "29 Sep 2026",
            "In progress",
            "The operations team is rehearsing restoration from the signed release package.",
        ),
        (
            "Operations handover",
            "Anne Wu",
            "30 Sep 2026",
            "Not started",
            "The handover checklist is prepared and will be completed after the release gate is approved.",
        ),
        (
            "Evidence archive",
            "Rui Chen",
            "30 Sep 2026",
            "In progress",
            "Logs, approvals and transfer checksums are collected in the controlled evidence archive.",
        ),
        (
            "Security sign-off",
            "Han Li",
            "29 Sep 2026",
            "In progress",
            "The security reviewer is checking the final access report, transfer audit trail and client-side retention settings before the release gate closes.",
        ),
        (
            "Cutover communications",
            "Anne Wu",
            "29 Sep 2026",
            "Not started",
            "The notification list and approved customer message are prepared. The coordinator will record delivery time and the client response in the release log.",
        ),
        (
            "Post-release monitoring",
            "Min Zhou",
            "30 Sep 2026",
            "Not started",
            "The team will monitor transfer completion, service health and reconciliation results during the agreed observation window and escalate exceptions through the change ticket.",
        ),
        (
            "Training evidence",
            "Rui Chen",
            "30 Sep 2026",
            "Complete",
            "Attendance, session materials and the approved operating guide are linked from the evidence archive for the receiving operations team.",
        ),
    )
    for row in detail_rows:
        cells = detail.add_row().cells
        for cell, text in zip(cells, row):
            cell.text = text
    format_table(detail, (1.42, 0.82, 0.93, 1.42, 2.26), header_rows=(0, 1), font_size=8.1)

    document.add_heading("3. Cutover and controls", level=1)
    document.add_paragraph(
        "The deployment window starts at 20:00 local time. The change coordinator will confirm "
        "the final go / no-go decision with the client before any production transfer begins."
    )
    document.add_paragraph(
        "The rollback package is verified and the transfer runbook is ready. [[FN1]]"
    )
    document.add_paragraph(
        "All transfer files remain in the client-controlled repository; local copies are "
        "non-authoritative."
    )
    document.add_paragraph(
        "Operations will confirm restart ownership at handover. The release coordinator will "
        "record the decision and retain the corresponding change ticket."
    )
    controls = document.add_table(rows=1, cols=3)
    for cell, text in zip(controls.rows[0].cells, ("Control", "Evidence", "Owner")):
        cell.text = text
    for row in (
        ("Transfer integrity", "Release manifest and SHA-256 record", "Jia Lin"),
        ("Client authorization", "Signed acceptance and change ticket", "Rui Chen"),
        ("Rollback readiness", "Rehearsal log and restore checkpoint", "Min Zhou"),
        ("Operational ownership", "Named on-call contact and handover record", "Anne Wu"),
    ):
        cells = controls.add_row().cells
        for cell, text in zip(cells, row):
            cell.text = text
    format_table(controls, (1.6, 3.45, 1.8), header_rows=(0,), font_size=9)
    document.add_heading("Closeout conditions", level=2)
    for item in (
        "Client acceptance record is returned and linked to the change ticket.",
        "Transfer checksum is verified against the release manifest.",
        "Open remediation item has an owner and an agreed completion date.",
    ):
        paragraph = document.add_paragraph(style="List Bullet")
        paragraph.add_run(item)

    with NamedTemporaryFile(suffix=".docx", dir=OUTPUT.parent, delete=False) as temporary:
        temporary_path = Path(temporary.name)
    document.save(temporary_path)
    patch_ooxml(temporary_path, OUTPUT)
    temporary_path.unlink(missing_ok=True)
    return OUTPUT


def patch_ooxml(source: Path, target: Path) -> None:
    with ZipFile(source) as archive:
        parts = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}

    parser = etree.XMLParser(remove_blank_text=False)
    document = etree.fromstring(parts["word/document.xml"], parser)
    for number, paragraph in enumerate(document.xpath(".//w:p", namespaces=NS), start=1):
        paragraph.set(f"{{{W14}}}paraId", f"{number:08X}")

    footnote_marker = document.xpath('.//w:t[contains(text(), "[[FN1]]")]', namespaces=NS)
    if len(footnote_marker) != 1:
        raise ValueError("Expected exactly one [[FN1]] marker")
    marker_text = footnote_marker[0]
    run = marker_text.getparent()
    marker_text.text = (marker_text.text or "").replace("[[FN1]]", "").rstrip()
    reference_run = etree.Element(f"{{{W}}}r")
    reference = etree.SubElement(reference_run, f"{{{W}}}footnoteReference")
    reference.set(f"{{{W}}}id", "1")
    run.addnext(reference_run)

    comment_paragraphs = document.xpath(
        './/w:p[.//w:t[contains(text(), "All transfer files remain in the client-controlled repository")]]',
        namespaces=NS,
    )
    if len(comment_paragraphs) != 1:
        raise ValueError("Expected one paragraph for the existing comment anchor")
    paragraph = comment_paragraphs[0]
    first_content = 1 if len(paragraph) and paragraph[0].tag == f"{{{W}}}pPr" else 0
    start = etree.Element(f"{{{W}}}commentRangeStart")
    start.set(f"{{{W}}}id", "0")
    paragraph.insert(first_content, start)
    end = etree.Element(f"{{{W}}}commentRangeEnd")
    end.set(f"{{{W}}}id", "0")
    paragraph.append(end)
    comment_reference_run = etree.SubElement(paragraph, f"{{{W}}}r")
    comment_reference = etree.SubElement(comment_reference_run, f"{{{W}}}commentReference")
    comment_reference.set(f"{{{W}}}id", "0")
    parts["word/document.xml"] = etree.tostring(
        document, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    comments = etree.Element(f"{{{W}}}comments", nsmap={"w": W})
    comment = etree.SubElement(comments, f"{{{W}}}comment")
    comment.set(f"{{{W}}}id", "0")
    comment.set(f"{{{W}}}author", "Release reviewer")
    comment.set(f"{{{W}}}date", "2026-09-24T09:00:00Z")
    comment_paragraph = etree.SubElement(comment, f"{{{W}}}p")
    comment_run = etree.SubElement(comment_paragraph, f"{{{W}}}r")
    comment_text = etree.SubElement(comment_run, f"{{{W}}}t")
    comment_text.text = "Retain the signed receipt with the audit package."
    parts["word/comments.xml"] = etree.tostring(
        comments, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    footnotes = etree.Element(f"{{{W}}}footnotes", nsmap={"w": W})
    for note_id, note_type, note_text in (
        ("-1", "separator", None),
        ("0", "continuationSeparator", None),
        ("1", None, "Hash values are compared as lowercase hexadecimal strings."),
    ):
        note = etree.SubElement(footnotes, f"{{{W}}}footnote")
        note.set(f"{{{W}}}id", note_id)
        if note_type:
            note.set(f"{{{W}}}type", note_type)
        note_paragraph = etree.SubElement(note, f"{{{W}}}p")
        if note_type:
            note_run = etree.SubElement(note_paragraph, f"{{{W}}}r")
            etree.SubElement(note_run, f"{{{W}}}separator" if note_type == "separator" else f"{{{W}}}continuationSeparator")
        else:
            note_run = etree.SubElement(note_paragraph, f"{{{W}}}r")
            text = etree.SubElement(note_run, f"{{{W}}}t")
            text.text = note_text
    parts["word/footnotes.xml"] = etree.tostring(
        footnotes, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    rels = etree.fromstring(parts["word/_rels/document.xml.rels"], parser)
    existing_ids = {relationship.get("Id", "") for relationship in rels}
    next_id = 1
    while f"rId{next_id}" in existing_ids:
        next_id += 1
    for relationship_type, target_part in (
        ("comments", "comments.xml"),
        ("footnotes", "footnotes.xml"),
    ):
        relationship = etree.SubElement(rels, f"{{{PKG_REL}}}Relationship")
        relationship.set("Id", f"rId{next_id}")
        relationship.set(
            "Type", f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{relationship_type}"
        )
        relationship.set("Target", target_part)
        next_id += 1
    parts["word/_rels/document.xml.rels"] = etree.tostring(
        rels, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    content_types = etree.fromstring(parts["[Content_Types].xml"], parser)
    for part_name, content_type in (
        ("/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
        ("/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"),
    ):
        override = etree.SubElement(content_types, f"{{{CT}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
    parts["[Content_Types].xml"] = etree.tostring(
        content_types, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    settings = etree.fromstring(parts["word/settings.xml"], parser)
    if not settings.xpath("./w:trackRevisions", namespaces=NS):
        settings.append(etree.Element(f"{{{W}}}trackRevisions"))
    parts["word/settings.xml"] = etree.tostring(
        settings, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


if __name__ == "__main__":
    print(build_docx())
