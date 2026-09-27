"""Build the deterministic synthetic labor-contract fixture for case 04."""

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
OUTPUT = ROOT / "fixtures" / "case-04-synthetic-labor-contract.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "w14": W14}


def set_font(run, size: float = 10, bold: bool | None = None, color: str = "263442") -> None:
    run.font.name = "Arial"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.font.bold = bold


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shade = tc_pr.find(qn("w:shd"))
    if shade is None:
        shade = OxmlElement("w:shd")
        tc_pr.append(shade)
    shade.set(qn("w:fill"), fill)


def set_cell_padding(cell, top=95, start=110, bottom=95, end=110) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        item = margins.find(qn(f"w:{edge}"))
        if item is None:
            item = OxmlElement(f"w:{edge}")
            margins.append(item)
        item.set(qn("w:w"), str(value))
        item.set(qn("w:type"), "dxa")


def set_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        item = borders.find(qn(f"w:{edge}"))
        if item is None:
            item = OxmlElement(f"w:{edge}")
            borders.append(item)
        item.set(qn("w:val"), "single")
        item.set(qn("w:sz"), "4")
        item.set(qn("w:space"), "0")
        item.set(qn("w:color"), "C9D3DE")


def format_table(table, widths: tuple[float, ...], header_rows=(0,), font_size=9) -> None:
    table.autofit = False
    table.style = "Table Grid"
    set_borders(table)
    for row_index, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        for column, cell in enumerate(row.cells):
            cell.width = Inches(widths[column])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_padding(cell)
            if row_index in header_rows:
                set_cell_shading(cell, "24476B")
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(1)
                paragraph.paragraph_format.line_spacing = 1.03
                for run in paragraph.runs:
                    set_font(run, font_size, row_index in header_rows, "FFFFFF" if row_index in header_rows else "263442")
    for row_index in header_rows:
        row = table.rows[row_index]._tr
        tr_pr = row.get_or_add_trPr()
        if tr_pr.find(qn("w:tblHeader")) is None:
            marker = OxmlElement("w:tblHeader")
            marker.set(qn("w:val"), "true")
            tr_pr.append(marker)


def add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, start)
    paragraph._p.append(end)


def add_internal_link(paragraph, label: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    props.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.append(underline)
    run.append(props)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_external_link(paragraph, label: str, url: str) -> None:
    rel_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    props.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.append(underline)
    run.append(props)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_page_field(paragraph, instruction: str, fallback: str = "1") -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = fallback
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def add_heading(document: Document, text: str, level: int, bookmark: str, bookmark_id: int) -> None:
    p = document.add_heading(text, level=level)
    add_bookmark(p, bookmark, bookmark_id)


def add_clause(document: Document, label: str, text: str, note_marker: bool = False):
    p = document.add_paragraph()
    p.paragraph_format.keep_together = True
    p.paragraph_format.space_after = Pt(5)
    lead = p.add_run(label + "  ")
    set_font(lead, 10, True)
    body = p.add_run(text)
    set_font(body, 10)
    if note_marker:
        marker = p.add_run(" [[FN-1]]")
        set_font(marker, 8)
    return p


def add_named_table(document: Document, headers: tuple[str, ...], rows: tuple[tuple[str, ...], ...], widths: tuple[float, ...], font_size=9, header_rows=(0,)):
    table = document.add_table(rows=1, cols=len(headers))
    for cell, value in zip(table.rows[0].cells, headers):
        cell.text = value
    for record in rows:
        cells = table.add_row().cells
        for cell, value in zip(cells, record):
            cell.text = value
    format_table(table, widths, header_rows=header_rows, font_size=font_size)
    return table


def set_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor(38, 52, 66)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.08
    for name, size in (("Title", 23), ("Heading 1", 15), ("Heading 2", 11.5)):
        style = document.styles[name]
        style.font.name = "Arial"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.size = Pt(size)
        style.font.bold = name != "Title"
        style.font.color.rgb = RGBColor(24, 45, 66)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(9 if name != "Title" else 0)
        style.paragraph_format.space_after = Pt(5)
        if name == "Title" and style._element.pPr is not None:
            border = style._element.pPr.find(qn("w:pBdr"))
            if border is not None:
                style._element.pPr.remove(border)


def build_docx() -> Path:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    set_styles(doc)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)
    section.header_distance = Inches(0.3)
    section.footer_distance = Inches(0.3)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font(header.add_run("澄海数字科技（虚构）  /  人力资源部"), 8, False, "68788A")
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(footer.add_run("劳动合同测试样本  ·  第 "), 8, False, "68788A")
    add_page_field(footer, "PAGE")
    set_font(footer.add_run(" 页 / 共 "), 8, False, "68788A")
    add_page_field(footer, "NUMPAGES")
    set_font(footer.add_run(" 页"), 8, False, "68788A")

    title = doc.add_paragraph("劳动合同", style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(title.runs[0], 23, True, "18324A")
    if title._p.pPr is not None:
        border = title._p.pPr.find(qn("w:pBdr"))
        if border is not None:
            title._p.pPr.remove(border)
    subtitle = doc.add_paragraph("固定期限  /  研发与数据服务岗位")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(subtitle.runs[0], 11, False, "5B6B7B")
    meta = doc.add_paragraph("合同编号：CH-RD-2026-0418     版本：1.3     签署地：上海市")
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(meta.runs[0], 9, False, "68788A")

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note.paragraph_format.space_before = Pt(14)
    note.paragraph_format.space_after = Pt(14)
    set_font(note.add_run("合成测试样本  ·  虚构主体与信息  ·  不可直接用于签署"), 9, True, "8A4B19")

    add_heading(doc, "甲方（用人单位）", 1, "party_a", 1)
    add_named_table(
        doc,
        ("项目", "合同记载"),
        (
            ("单位名称", "上海澄海数字科技有限公司（虚构）"),
            ("统一社会信用代码", "91310000MA1F000000（测试数据）"),
            ("住所", "上海市浦东新区云桥路 168 号澄海中心 8 层（虚构）"),
            ("法定代表人 / 委托代理人", "周明远（虚构）"),
            ("人力资源联系邮箱", "hr@example.invalid"),
        ),
        (2.2, 4.65),
        font_size=9,
    )
    add_heading(doc, "乙方（劳动者）", 1, "party_b", 2)
    add_named_table(
        doc,
        ("项目", "合同记载"),
        (
            ("姓名", "林知夏（虚构）"),
            ("身份证件号码", "310101199401010000（测试数据）"),
            ("户籍 / 通讯地址", "上海市杨浦区平凉路 100 号 502 室（虚构）"),
            ("个人联系邮箱", "employee@example.invalid"),
            ("紧急联系人及电话", "林女士  13800000000（虚构）"),
        ),
        (2.2, 4.65),
        font_size=9,
    )
    intro = doc.add_paragraph(
        "甲乙双方在平等自愿、协商一致的基础上订立本合同。双方确认已阅读合同正文及附件，"
        "并理解正文条款与附件中相互引用的具体内容。本文仅用于软件编辑能力测试，所有主体和数据均为虚构。"
    )
    intro.paragraph_format.space_before = Pt(8)
    for run in intro.runs:
        set_font(run, 10)

    toc = doc.add_paragraph()
    toc.paragraph_format.space_before = Pt(4)
    toc.paragraph_format.space_after = Pt(3)
    set_font(toc.add_run("合同导航："), 9, True)
    for label, anchor in (("合同期限", "clause_2"), ("岗位与地点", "clause_4"), ("劳动报酬", "clause_6"), ("保密与知识产权", "clause_9"), ("附件 A", "annex_a"), ("附件 B", "annex_b")):
        add_internal_link(toc, label, anchor)
        set_font(toc.add_run("  |  "), 9, False, "8794A0")

    doc.add_page_break()
    add_heading(doc, "第一章  合同基础与期限", 1, "chapter_1", 10)
    add_heading(doc, "第一条  合同依据与文件构成", 2, "clause_1", 11)
    add_clause(doc, "1.1", "本合同由甲方与乙方订立，约定劳动关系建立、岗位安排、劳动报酬、工作时间、休息休假、社会保险、劳动保护、培训、保密与合同终止等事项。")
    add_clause(doc, "1.2", "本合同正文、附件 A《薪酬与发放说明》、附件 B《岗位与办公安排》及附件 C《交接与资产清单》共同构成合同文本。各附件经双方确认后，与正文具有同等的文件完整性；正文与附件存在表述差异时，应先由双方书面确认具体适用内容。")
    add_clause(doc, "1.3", "甲方依法制定并向乙方公示或告知的规章制度，作为日常管理依据。规章制度不得被解释为对本合同中明确约定事项的单方改写。")

    add_heading(doc, "第二条  合同期限与试用期", 2, "clause_2", 12)
    add_clause(doc, "2.1", "本合同为固定期限劳动合同，期限自 2026 年 10 月 1 日起至 2029 年 9 月 30 日止。劳动关系自合同起始日起建立。")
    add_clause(doc, "2.2", "试用期为六个月，自 2026 年 10 月 1 日起至 2027 年 3 月 31 日止。试用期计入合同期限。甲方应围绕本合同约定岗位职责、合理工作要求及已告知的考核标准开展评价，并向乙方说明评价结果。")
    add_clause(doc, "2.3", "试用期内，乙方依法享有劳动报酬、休息休假、社会保险及劳动保护等权利。试用期工资及其构成见第六条与附件 A；试用期届满后的岗位安排不因培养计划或阶段性绩效回顾自动改变。")
    add_clause(doc, "2.4", "合同期满前，双方可就续订事宜进行协商。续订或变更应采用书面形式。合同期满、续订协商、工作交接的办理不影响依法应履行的程序。")

    add_heading(doc, "第二章  工作内容与工作安排", 1, "chapter_2", 20)
    add_heading(doc, "第三条  岗位职责与工作成果", 2, "clause_3", 21)
    add_clause(doc, "3.1", "乙方岗位为高级运维分析师，职级为 P4，所属部门为云平台组，直接汇报对象为平台平台主管。岗位职责包括生产运行分析、变更风险复核、服务指标跟踪、故障复盘材料整理及跨团队协作。具体职责见附件 B。")
    add_clause(doc, "3.2", "甲方可根据业务需要，在不违反法律法规且与乙方协商一致的基础上调整工作任务。岗位职责的实质变更、薪酬标准的变更或工作地点的变更，应依法协商并书面确认。")
    add_clause(doc, "3.3", "乙方按合理工作要求提交分析报告、运行记录和交接材料。工作成果的验收依据以项目说明、任务单或双方确认的交付清单为准；内部计划中的三个月培养回顾只是培训安排，不是劳动合同试用期。")
    add_clause(doc, "3.4", "乙方有权了解与本人岗位有关的流程、信息系统权限、安全要求和绩效评价标准。甲方应为乙方提供履行岗位职责所需的必要资源，并对可能影响劳动安全的事项进行说明。")

    add_heading(doc, "第四条  工作地点与出差安排", 2, "clause_4", 22)
    add_clause(doc, "4.1", "乙方日常工作地点为上海市浦东新区云桥路 168 号澄海中心 8 层（甲方上海总部）。该地址作为合同约定的常规办公地点记载于附件 B。")
    add_clause(doc, "4.2", "因项目需要安排的临时出差，应说明任务、地点、期限及费用报销方式。甲方不得以应急集合点或客户现场登记地址替代本合同记载的常规办公地点。")
    add_clause(doc, "4.3", "甲方可结合业务安排远程办公或在同一城市内的临时协作地点工作，并应明确工作时间、信息安全和考勤要求。常规办公地点变更应由双方协商并留存书面记录。")

    add_heading(doc, "第五条  工作时间、休息休假", 2, "clause_5", 23)
    add_clause(doc, "5.1", "甲方实行标准工时制，原则上每日工作八小时、每周工作五日。具体上下班时间、休息日安排及考勤方式，依依法制定并公示的制度执行。确需依法实行其他工时制度的，应履行相应程序。")
    add_clause(doc, "5.2", "甲方因工作需要安排延长工作时间或休息日、法定节假日工作的，应依法履行审批或协商程序，并依法安排补休或支付相应报酬。乙方不得因拒绝未经确认的非紧急加班而被视为违反岗位职责。")
    add_clause(doc, "5.3", "乙方依法享有法定节假日、年休假、婚假、产假、陪产假、病假及其他依法规定的假期。休假申请、证明材料和交接要求按甲方已公示制度办理；制度中的两日跨年结转上限属于假期管理规则，不代表乙方的年休假天数。")
    add_clause(doc, "5.4", "乙方应在合理范围内配合工作交接和紧急故障响应。甲方应建立值班安排并明确轮值周期、响应方式和补偿安排；本条不构成无边界的待命义务。")

    add_heading(doc, "第三章  劳动报酬与社会保障", 1, "chapter_3", 30)
    add_heading(doc, "第六条  工资结构、支付与扣缴情形", 2, "clause_6", 31)
    add_clause(doc, "6.1", "乙方转正后的税前月基本工资为人民币 18,800 元（大写：壹万捌仟捌佰元整）。基本工资按月支付，甲方于次月 10 日前支付至乙方指定账户；遇法定休息日或节假日，按依法确定的发薪安排办理。")
    add_clause(doc, "6.2", "乙方试用期税前月工资为转正后基本工资的百分之八十，即人民币 15,040 元。该金额按本合同现有基本工资计算。基本工资、试用期工资及附件 A 所列构成应相互一致；未在本合同约定的项目不得被默认包含于基本工资。")
    add_clause(doc, "6.3", "甲方可根据实际出勤和依法应扣项目代扣代缴个人所得税及个人承担的社会保险、公积金部分，并向乙方提供工资清单。工资支付记录应列明基本工资、津贴、依法计发的加班工资及其他实际支付项目。")
    add_clause(doc, "6.4", "交通津贴为税前每月人民币 800 元，按考勤和报销制度发放。该津贴与第六条约定的月基本工资分别列示；不得因基本工资调整而自动改变津贴数额。")
    add_clause(doc, "6.5", "绩效奖金依据经告知的考核周期、指标、计算规则及实际考核结果确定。年度目标奖金参考额为人民币 18,800 元，属于目标参考金额，不是保证支付金额，也不得与月基本工资混同。")
    add_clause(doc, "6.6", "任何工资结构调整均应依法协商并书面确认。工资条、预算表或内部成本表中的同名数字不当然构成本合同变更；附件 A 的薪酬表仅复述本条中明确约定的组成。")

    add_heading(doc, "第七条  社会保险与福利", 2, "clause_7", 32)
    add_clause(doc, "7.1", "甲方按国家和工作地规定为乙方办理社会保险及住房公积金相关手续，双方依法承担各自应缴部分。缴费基数、比例及办理时间依适用规定和主管部门要求执行。")
    add_clause(doc, "7.2", "乙方应如实提供办理参保所需资料。甲方不得要求乙方放弃依法参加社会保险的权利，也不得以商业保险、现金补贴或其他福利替代依法应办理的社会保险。")
    add_clause(doc, "7.3", "餐饮、体检、培训和办公设备属于依制度或实际安排提供的福利或工作条件，具体项目见附件 A 或甲方已公示制度。培训预算上限人民币 18,800 元是年度部门预算参考，不是乙方工资或个人培训服务费。")

    add_heading(doc, "第四章  劳动保护与信息管理", 1, "chapter_4", 40)
    add_heading(doc, "第八条  劳动保护、职业健康与设备", 2, "clause_8", 41)
    add_clause(doc, "8.1", "甲方应提供符合岗位需要的办公环境、劳动保护条件及必要的安全培训。涉及客户生产环境、远程接入或敏感数据处理的，甲方应提供适用的操作规范和必要授权。")
    add_clause(doc, "8.2", "乙方应遵守已告知的安全操作规范，合理使用设备和访问权限。发现可能危及人身安全、信息安全或业务连续性的风险时，应通过规定渠道及时报告。")
    add_clause(doc, "8.3", "甲方配发的电脑、门禁卡、令牌和测试设备仅供履职使用。乙方应按附件 C 登记领用与归还情况；日常办公备用设备编号 SH-168-04 与合同约定的常规办公地点无关。")
    add_clause(doc, "8.4", "涉及个人信息和客户数据时，双方按最小必要原则处理。乙方仅可在工作授权范围内访问和使用相关资料，不得通过个人邮箱、个人网盘或未批准的外部工具传输。")

    add_heading(doc, "第九条  保密义务与成果归属", 2, "clause_9", 42)
    add_clause(doc, "9.1", "保密信息包括甲方或客户依法要求保护的未公开技术资料、账号凭据、系统配置、经营数据和个人信息。公开信息、乙方能够证明在接收前已合法持有的信息，以及依法必须披露的信息不因本条而当然成为保密信息。")
    add_clause(doc, "9.2", "乙方应按照授权范围访问、复制、存储和传输保密信息，并在权限终止或合同解除后配合归还载体、移交工作文件和删除依法应删除的本地副本。甲方应提供合理的归还渠道，并对依法需要留存的劳动管理资料进行权限控制。", note_marker=True)
    add_clause(doc, "9.3", "乙方在履行岗位职责过程中形成的工作成果及其权利归属，依适用法律法规、岗位职责、项目安排及双方另行签署的有效文件确定。甲方不得将乙方入职前独立形成且已书面列明的作品当然视为职务成果。")
    add_clause(doc, "9.4", "如需另行约定竞业限制，应以单独书面协议明确人员范围、期限、地域、经济补偿及违约责任，并依法执行。本合同不以概括条款自动扩大竞业限制范围。")
    add_clause(doc, "9.5", "双方可依《中华人民共和国劳动合同法》及相关规定处理劳动关系事项。可查阅全国人大网站公布的法律文本：")
    law = doc.add_paragraph()
    law.paragraph_format.space_after = Pt(6)
    add_external_link(law, "中华人民共和国劳动合同法（全国人大网站）", "https://www.npc.gov.cn/c2/c30834/201905/t20190521_296651.html")

    add_heading(doc, "第五章  变更、解除、终止与争议处理", 1, "chapter_5", 50)
    add_heading(doc, "第十条  合同变更与日常通知", 2, "clause_10", 51)
    add_clause(doc, "10.1", "本合同内容变更应由双方协商一致并以书面形式确认。甲方规章制度更新不当然变更本合同中明确约定的期限、岗位、工作地点和工资标准。")
    add_clause(doc, "10.2", "双方应保证本合同首页所列通讯地址、电子邮箱和联系电话真实有效。向本合同列明的地址寄送通知的，应保留寄送记录；电子通知应发送至本合同记载邮箱并保留发送记录。重要变更应采用可核验方式确认收悉。")
    add_clause(doc, "10.3", "因地址、邮箱或电话号码变更导致通知未能送达的，变更方应及时书面告知对方。紧急联系人仅用于必要的安全联络，不作为合同通知送达地址。")

    add_heading(doc, "第十一条  合同解除与终止", 2, "clause_11", 52)
    add_clause(doc, "11.1", "双方解除或终止劳动合同，应依适用法律法规规定的事由、程序、通知期限和经济补偿规则办理。任何一方不得以格式条款预先排除法定权利或免除依法不得免除的义务。")
    add_clause(doc, "11.2", "乙方提出解除、甲方提出解除或双方协商解除时，应对交接事项、工资结算、证明出具和资料归还进行记录。工作交接应限于合理范围，并明确交接接收人及完成时间。")
    add_clause(doc, "11.3", "合同期满、主体变化或其他依法导致合同终止的情形，按法律规定办理。双方对是否续订尚未达成书面一致的，不得把内部审批状态记作已续订协议。")

    add_heading(doc, "第十二条  争议处理与文本效力", 2, "clause_12", 53)
    add_clause(doc, "12.1", "因履行本合同发生争议，双方可先协商；协商不成的，依法向有管辖权的劳动人事争议仲裁机构申请仲裁。对仲裁裁决不服的，依法向有管辖权的人民法院提起诉讼。")
    add_clause(doc, "12.2", "本合同未尽事项，依适用法律法规和双方后续书面约定处理。合同部分条款被依法认定无效，不影响其他条款的效力，但应依法对相关事项作相应处理。")
    add_clause(doc, "12.3", "本合同正文共十二条，另含三份附件。正文与附件共同组成同一份合同文件。双方确认签署页、附件页及变更页均应妥善保存；如使用电子签署，应按双方认可并可验证的方式留存签署记录。")

    doc.add_page_break()
    add_heading(doc, "附件 A  薪酬与发放说明", 1, "annex_a", 60)
    p = doc.add_paragraph("本附件用于复述正文第六条约定的工资构成。基本工资与津贴、奖金参考值、培训预算分别列示，不能只凭数字相同推定项目相同。")
    for run in p.runs:
        set_font(run, 9.5)
    add_named_table(
        doc,
        ("项目", "试用期", "试用期后", "依据 / 说明"),
        (
            ("月基本工资（税前）", "15,040 元", "18,800 元", "正文第 6.1、6.2 条"),
            ("交通津贴", "按制度核算", "800 元 / 月", "正文第 6.4 条；单列项目"),
            ("年度目标奖金参考额", "不适用", "18,800 元 / 年", "正文第 6.5 条；非保证支付"),
            ("年度培训预算参考上限", "按部门计划", "18,800 元 / 年", "正文第 7.3 条；部门预算，不是工资"),
            ("工资支付日", "次月 10 日前", "次月 10 日前", "遇法定休息日依支付安排办理"),
            ("发放账户", "乙方指定账户", "乙方指定账户", "账户信息另行安全提交，不写入公开副本"),
        ),
        (1.55, 1.32, 1.35, 2.7),
        font_size=8.6,
    )
    add_heading(doc, "A.1  工资核算说明", 2, "annex_a_1", 61)
    add_clause(doc, "A.1.1", "试用期工资按正文约定的比例计算。当前试用期工资为转正后月基本工资的百分之八十；该金额仅随经双方书面确认的基本工资标准变更而调整。")
    add_clause(doc, "A.1.2", "交通津贴以实际约定的每月 800 元列示。年度目标奖金参考额及培训预算上限虽然均为 18,800 元，但用途、支付条件和法律性质不同，不属于月基本工资。")
    add_clause(doc, "A.1.3", "工资清单应区分应发项目、扣缴情形和实发金额。发生缺勤、依法代扣或其他需要核算的情形时，甲方应提供可理解的项目明细。")

    add_heading(doc, "附件 B  岗位与办公安排", 1, "annex_b", 70)
    add_named_table(
        doc,
        ("字段", "约定内容", "补充说明"),
        (
            ("岗位名称", "高级运维分析师", "对应正文第 3.1 条"),
            ("岗位职级", "P4", "岗位等级，不是薪酬金额"),
            ("所属部门", "云平台组", "内部组织名称"),
            ("常规办公地点", "上海市浦东新区云桥路 168 号澄海中心 8 层", "与正文第 4.1 条保持一致"),
            ("临时应急集合点", "上海市浦东新区云桥路 168 号澄海中心 1 层北门", "仅为应急集合点，不是合同常规工作地点"),
            ("设备交接地点", "上海市杨浦区控江路 560 号仓储点", "仅用于资产收发记录"),
            ("直接汇报对象", "平台平台主管", "内部汇报关系"),
        ),
        (1.5, 3.2, 2.22),
        font_size=8.7,
    )
    add_heading(doc, "B.1  岗位主要职责", 2, "annex_b_1", 71)
    for item in (
        "维护生产服务运行指标，按既定周期检查异常趋势并记录处理过程。",
        "参与变更评审，核对影响范围、回退条件、验证步骤及责任人。",
        "协助整理故障复盘、运维交接和服务改进建议，不代替审批人作最终风险批准。",
        "按授权范围处理客户资料，并遵循正文第八条、第九条的信息安全要求。",
        "配合跨团队协作和合理的临时任务安排；重大岗位变化仍依正文第 3.2 条办理。",
    ):
        p = doc.add_paragraph(style="List Bullet")
        set_font(p.add_run(item), 9.5)
    add_heading(doc, "B.2  培训安排", 2, "annex_b_2", 72)
    add_clause(doc, "B.2.1", "入职后的三个月培养回顾安排为岗位熟悉、权限培训和阶段反馈。该安排不构成劳动合同试用期长度的记载；合同试用期以正文第 2.2 条为准。")
    add_clause(doc, "B.2.2", "培训记录应载明课程主题、日期和参与情况。部门年度培训预算参考上限为人民币 18,800 元，该数值不是乙方应支付的培训费或违约金。")

    add_heading(doc, "附件 C  交接与资产清单", 1, "annex_c", 80)
    add_named_table(
        doc,
        ("序号", "物品 / 权限", "编号 / 范围", "领用确认", "归还确认"),
        (
            ("1", "办公电脑", "CH-LT-2048", "待填写", "待填写"),
            ("2", "门禁卡", "SH-168-04", "待填写", "待填写"),
            ("3", "多因素认证令牌", "MFA-77821", "待填写", "待填写"),
            ("4", "测试环境只读权限", "OPS-READ-P4", "待填写", "待填写"),
            ("5", "客户生产权限（如获批准）", "按审批记录开通", "待填写", "待填写"),
            ("6", "纸质资料或其他载体", "以交接单为准", "待填写", "待填写"),
        ),
        (0.5, 1.65, 1.75, 1.48, 1.48),
        font_size=8.5,
    )
    add_heading(doc, "C.1  工作交接记录", 2, "annex_c_1", 81)
    add_clause(doc, "C.1.1", "离岗或岗位交接时，双方应记录文件存放位置、待办事项、风险状态和接收人。交接记录用于说明已移交事项，不替代依法办理的解除、终止或工资结算程序。")
    add_clause(doc, "C.1.2", "账号停用、设备归还和文件移交应分别记录完成时间。不得要求乙方提供个人账号密码；对工作账号应通过甲方管理流程进行权限移交或重置。")
    doc.add_page_break()
    add_heading(doc, "C.2  通知与文件签收记录", 2, "annex_c_2", 82)
    add_named_table(
        doc,
        ("文件名称", "交付日期", "交付方式", "接收人签字 / 电子记录"),
        (
            ("合同正文及附件", "____年__月__日", "纸质 / 电子", "________________________"),
            ("规章制度告知记录", "____年__月__日", "纸质 / 系统", "________________________"),
            ("岗位职责说明", "____年__月__日", "纸质 / 系统", "________________________"),
            ("薪酬与账户信息说明", "____年__月__日", "单独安全交付", "________________________"),
        ),
        (1.7, 1.28, 1.6, 2.28),
        font_size=8.5,
    )

    add_heading(doc, "签署页", 1, "signature_page", 90)
    p = doc.add_paragraph("甲乙双方确认已阅读本合同正文及附件，签署内容以双方最终确认的书面文本为准。")
    for run in p.runs:
        set_font(run, 10)
    add_named_table(
        doc,
        ("甲方（盖章）", "乙方（签名）"),
        (
            ("单位名称：上海澄海数字科技有限公司（虚构）", "姓名：林知夏（虚构）"),
            ("法定代表人 / 委托代理人：____________", "签名：____________________________"),
            ("签署日期：____年____月____日", "签署日期：____年____月____日"),
            ("签署地点：上海市", "签署地点：上海市"),
        ),
        (3.38, 3.38),
        font_size=9,
    )
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    set_font(p.add_run("目录和正文交叉引用："), 9, True)
    add_internal_link(p, "返回第六条工资结构", "clause_6")
    set_font(p.add_run("  ·  "), 9, False, "8794A0")
    add_internal_link(p, "转至附件 A 薪酬表", "annex_a")

    with NamedTemporaryFile(suffix=".docx", dir=OUTPUT.parent, delete=False) as temporary:
        temp_path = Path(temporary.name)
    doc.save(temp_path)
    patch_ooxml(temp_path, OUTPUT)
    temp_path.unlink(missing_ok=True)
    return OUTPUT


def patch_ooxml(source: Path, target: Path) -> None:
    with ZipFile(source) as archive:
        parts = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    parser = etree.XMLParser(remove_blank_text=False)
    document = etree.fromstring(parts["word/document.xml"], parser)
    for index, paragraph in enumerate(document.xpath(".//w:p", namespaces=NS), start=1):
        paragraph.set(f"{{{W14}}}paraId", f"{index:08X}")

    note_nodes = document.xpath('.//w:t[contains(text(), "[[FN-1]]")]', namespaces=NS)
    if len(note_nodes) != 1:
        raise ValueError("Expected exactly one [[FN-1]] marker")
    text_node = note_nodes[0]
    text_node.text = (text_node.text or "").replace("[[FN-1]]", "").rstrip()
    note_run = etree.Element(f"{{{W}}}r")
    footnote_ref = etree.SubElement(note_run, f"{{{W}}}footnoteReference")
    footnote_ref.set(f"{{{W}}}id", "1")
    text_node.getparent().getparent().addnext(note_run)

    # Existing review comment is deliberately attached to one full sentence.
    target_paragraphs = document.xpath(
        './/w:p[.//w:t[contains(text(), "本合同正文、附件 A")]]', namespaces=NS
    )
    if len(target_paragraphs) != 1:
        raise ValueError("Expected one existing comment anchor paragraph")
    paragraph = target_paragraphs[0]
    start_index = 1 if len(paragraph) and paragraph[0].tag == f"{{{W}}}pPr" else 0
    for tag in ("commentRangeStart", "commentRangeEnd"):
        marker = etree.Element(f"{{{W}}}{tag}")
        marker.set(f"{{{W}}}id", "0")
        if tag == "commentRangeStart":
            paragraph.insert(start_index, marker)
        else:
            paragraph.append(marker)
    reference_run = etree.SubElement(paragraph, f"{{{W}}}r")
    reference = etree.SubElement(reference_run, f"{{{W}}}commentReference")
    reference.set(f"{{{W}}}id", "0")
    parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)

    comments = etree.Element(f"{{{W}}}comments", nsmap={"w": W})
    comment = etree.SubElement(comments, f"{{{W}}}comment")
    comment.set(f"{{{W}}}id", "0")
    comment.set(f"{{{W}}}author", "合同审阅人")
    comment.set(f"{{{W}}}date", "2026-09-01T02:00:00Z")
    comment_p = etree.SubElement(comment, f"{{{W}}}p")
    comment_r = etree.SubElement(comment_p, f"{{{W}}}r")
    etree.SubElement(comment_r, f"{{{W}}}t").text = "请确认附件 A 与正文的工资项目仅作一致性复述，不扩大工资定义。"
    parts["word/comments.xml"] = etree.tostring(comments, xml_declaration=True, encoding="UTF-8", standalone=True)

    footnotes = etree.Element(f"{{{W}}}footnotes", nsmap={"w": W})
    for note_id, note_type, note_text in (
        ("-1", "separator", None),
        ("0", "continuationSeparator", None),
        ("1", None, "本段为虚构测试条款。实际文件中的保密范围与期限应以有效约定及适用法律为准。"),
    ):
        note = etree.SubElement(footnotes, f"{{{W}}}footnote")
        note.set(f"{{{W}}}id", note_id)
        if note_type:
            note.set(f"{{{W}}}type", note_type)
        p = etree.SubElement(note, f"{{{W}}}p")
        run = etree.SubElement(p, f"{{{W}}}r")
        if note_type:
            etree.SubElement(run, f"{{{W}}}separator" if note_type == "separator" else f"{{{W}}}continuationSeparator")
        else:
            etree.SubElement(run, f"{{{W}}}t").text = note_text
    parts["word/footnotes.xml"] = etree.tostring(footnotes, xml_declaration=True, encoding="UTF-8", standalone=True)

    rels = etree.fromstring(parts["word/_rels/document.xml.rels"], parser)
    existing = {item.get("Id", "") for item in rels}
    next_id = 1
    while f"rId{next_id}" in existing:
        next_id += 1
    for rel_type, target_part in (("comments", "comments.xml"), ("footnotes", "footnotes.xml")):
        relationship = etree.SubElement(rels, f"{{{PKG_REL}}}Relationship")
        relationship.set("Id", f"rId{next_id}")
        relationship.set("Type", f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}")
        relationship.set("Target", target_part)
        next_id += 1
    parts["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    types = etree.fromstring(parts["[Content_Types].xml"], parser)
    for part_name, content_type in (
        ("/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
        ("/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"),
    ):
        override = etree.SubElement(types, f"{{{CT}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
    parts["[Content_Types].xml"] = etree.tostring(types, xml_declaration=True, encoding="UTF-8", standalone=True)

    settings = etree.fromstring(parts["word/settings.xml"], parser)
    if not settings.xpath("./w:trackRevisions", namespaces=NS):
        settings.append(etree.Element(f"{{{W}}}trackRevisions"))
    parts["word/settings.xml"] = etree.tostring(settings, xml_declaration=True, encoding="UTF-8", standalone=True)

    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)


if __name__ == "__main__":
    print(build_docx())
