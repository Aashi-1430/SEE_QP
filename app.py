import tempfile
import os
import base64
import html
import io
import random
import re
import textwrap
from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether
)
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.shared import Inches, Pt, Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


st.set_page_config(page_title="SVIET Question Paper Generator", page_icon="📄", layout="wide")

COLLEGE_DEFAULT = "Swami Vivekanand Institute of Engineering & Technology"

UMC_LINE = (

    "NOTE: Disclosure of Identity by writing Mobile No. or Marking of passing request "
    "on any paper of Answer Sheet will lead to UMC against the Student."
)

MARK_OPTIONS = [2, 4, 5, 8, 10]


# =========================================================
# LOAD EXTERNAL CSS FILE
# =========================================================



from pathlib import Path
CSS_FILE = Path(__file__).resolve().parent / "style.css"
if CSS_FILE.exists():

    with open(CSS_FILE, "r", encoding="utf-8") as file:

        custom_css = file.read()

    st.markdown(

        f"<style>{custom_css}</style>",

        unsafe_allow_html=True

    )

else:

    st.warning("style.css file not found. Please keep it in the same folder as app.py.")

def step_header(number, title, subtitle=""):

    sub = f'<div class="step-sub">{html.escape(subtitle)}</div>' if subtitle else ""

    st.markdown(

        f'<div class="step-head"><div class="step-num">{number}</div>'

        f'<div><div class="step-title">{html.escape(title)}</div>{sub}</div></div>',

        unsafe_allow_html=True,

    )

LOGO_FILE = "Swami_Vivekanand_Institute_Logo.png"  # keep this file in the same folder as this script

LOGO_FOLDERS = (Path(__file__).resolve().parent, Path.cwd())

def find_logo():

    """Find the logo next to this script (exact name first, then any image with 'logo' in its name)."""

    for base in LOGO_FOLDERS:

        exact = base / LOGO_FILE

        if exact.is_file():

            return exact

    for base in LOGO_FOLDERS:

        if base.is_dir():

            for path in sorted(base.iterdir()):

                if path.is_file() and path.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp") and "logo" in path.name.lower():

                    return path

    return None

@st.cache_data(show_spinner=False)

def encode_logo(path_str, modified_time):

    """Downscale the logo and return it as base64 (cached per file and modification time)."""

    path = Path(path_str)

    try:

        from PIL import Image



        img = Image.open(path).convert("RGBA")

        img.thumbnail((420, 420))

        buf = io.BytesIO()

        img.save(buf, format="PNG", optimize=True)

        return base64.b64encode(buf.getvalue()).decode()

    except Exception:

        return base64.b64encode(path.read_bytes()).decode()

_logo_path = find_logo()

_logo_b64 = ""

if _logo_path is not None:

    try:

        _logo_b64 = encode_logo(str(_logo_path), _logo_path.stat().st_mtime)

    except Exception:

        _logo_b64 = ""

_logo_html = (

    f'<div class="hero-logo"><img src="data:image/png;base64,{_logo_b64}" alt="SVIET logo"></div>'

    if _logo_b64

    else ""

)
st.markdown(

    '<div class="hero">'

    '<div class="hero-text">'

    '<div class="badge">Examination Cell</div>'

    "<h1>SVIET Question Paper Generator</h1>"



    "</div>"

    f"{_logo_html}"

    "</div>",

    unsafe_allow_html=True,

)

if not _logo_b64:

    st.caption(f"Logo not found. Place '{LOGO_FILE}' in: {LOGO_FOLDERS[0]}")

def normalize(value):

    return re.sub(r"\s+", " ", str(value)).strip()

def format_question_parts(question):

    """

    Split question subparts (a), (b), (c) etc. onto separate lines.

    """

    if not question:

        return ""

    question = str(question).strip()

    pattern = r'\s+(?=\([a-z]\)\s*)'

    parts = re.split(pattern, question, flags=re.IGNORECASE)

    if len(parts) > 1 and re.match(
        r'^\([a-z]\)\s*', parts[1], re.IGNORECASE):
        return "\n".join(part.strip() for part in parts if part.strip())


    return question

def looks_like_heading(text):

    value = normalize(text).strip(" :-")

    if not value:

        return True

    return bool(re.match(

        r"^(unit\s+[ivx\d]+|module\s+\d+|chapter\s+\d+|question\s*(bank|text)?|questions?|s\.?\s*no\.?|serial\s*no\.?)$",

        value, flags=re.IGNORECASE

    ))


def unique(items):

    result, seen = [], set()

    for item in items:

        text = normalize(item)

        if len(text) < 3 or looks_like_heading(text):

            continue

        key = text.casefold()

        if key not in seen:

            seen.add(key)

            result.append(text)

    return result





def read_docx(upload):

    upload.seek(0)

    doc = Document(upload)

    items = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    for table in doc.tables:

        for row in table.rows:

            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]

            if cells:

                items.append(max(cells, key=len))

    return unique(items)





def read_excel(upload):

    upload.seek(0)

    workbook = pd.ExcelFile(upload)

    items = []

    for sheet_name in workbook.sheet_names:

        frame = pd.read_excel(workbook, sheet_name=sheet_name, header=None).fillna("")

        for _, row in frame.iterrows():

            cells = [normalize(v) for v in row.tolist() if normalize(v)]

            if cells:

                items.append(max(cells, key=len))

    return unique(items)





def read_bank(upload):

    name = upload.name.lower()

    if name.endswith(".docx"):

        return read_docx(upload)

    if name.endswith((".xlsx", ".xls")):

        return read_excel(upload)

    raise ValueError("Upload a Word (.docx) or Excel (.xlsx/.xls) file.")





def add_paragraph(doc, text="", *, bold=False, size=10, align=None, space_after=0, italic=False):

    p = doc.add_paragraph()

    if align is not None:

        p.alignment = align

    p.paragraph_format.space_after = Pt(space_after)

    run = p.add_run(str(text))

    run.bold = bold

    run.italic = italic

    run.font.name = "Times New Roman"

    run.font.size = Pt(size)

    return p





def number_to_words(n):

    words = {

        0: "ZERO", 1: "ONE", 2: "TWO", 3: "THREE", 4: "FOUR", 5: "FIVE",

        6: "SIX", 7: "SEVEN", 8: "EIGHT", 9: "NINE", 10: "TEN",

        11: "ELEVEN", 12: "TWELVE", 13: "THIRTEEN", 14: "FOURTEEN",

        15: "FIFTEEN", 16: "SIXTEEN", 17: "SEVENTEEN", 18: "EIGHTEEN",

        19: "NINETEEN", 20: "TWENTY"

    }

    return words.get(int(n), str(int(n)))





def add_centered_heading(doc, text):

    add_paragraph(doc, text, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)



def add_formatted_question(doc, number, question, marks=None):

    parts = format_question_parts(question).splitlines()

    parts = [part.strip() for part in parts if part.strip()]



    if not parts:

        parts = [str(question).strip()]



    # First line with the main question number

    p = doc.add_paragraph()

    p.paragraph_format.space_after = Pt(2)



    run = p.add_run(f"{number}. ")

    run.font.name = "Times New Roman"

    run.font.size = Pt(10)



    run = p.add_run(parts[0])

    run.font.name = "Times New Roman"

    run.font.size = Pt(10)



    # Display every remaining subpart in a separate paragraph

    for part in parts[1:]:

        sub_p = doc.add_paragraph()

        sub_p.paragraph_format.left_indent = Inches(0.30)

        sub_p.paragraph_format.space_after = Pt(2)



        run = sub_p.add_run(part)

        run.font.name = "Times New Roman"

        run.font.size = Pt(10)



    # Space after the complete question

    p.paragraph_format.space_after = Pt(6)



    return p





def build_docx(meta, section_a, section_b, section_c, marks_b, marks_c, instructions):

    doc = Document()

    sec = doc.sections[0]



    # A4 Landscape booklet sheet

    sec.orientation = WD_ORIENT.LANDSCAPE

    sec.page_width = Mm(297)

    sec.page_height = Mm(210)



    sec.top_margin = Inches(0.45)

    sec.bottom_margin = Inches(0.45)

    sec.left_margin = Inches(0.55)

    sec.right_margin = Inches(0.55)



    normal = doc.styles["Normal"]

    normal.font.name = "Times New Roman"

    normal.font.size = Pt(10)



    # Two-column layout

    cols = sec._sectPr.xpath('./w:cols')

    if cols:

        cols_element = cols[0]

    else:

        cols_element = OxmlElement('w:cols')

        sec._sectPr.append(cols_element)



    cols_element.set(qn('w:num'), '2')

    cols_element.set(qn('w:space'), '500')

    cols_element.set(qn('w:equalWidth'), '1')





    # --------------------------------------------------

    # Header: Roll Number, Total Questions and Pages

    # --------------------------------------------------



    total_questions = len(section_a) + len(section_b) + len(section_c)



    header_table = doc.add_table(rows=2, cols=2)

    header_table.autofit = False



    header_table.columns[0].width = Inches(2.5)

    header_table.columns[1].width = Inches(2.5)



    # First row: Roll Number across the complete header width

    top_cell = header_table.cell(0, 0).merge(header_table.cell(0, 1))

    top_cell.width = Inches(5.0)



    p = top_cell.paragraphs[0]

    p.paragraph_format.space_after = Pt(0)



    r = p.add_run("Roll No. __________________")

    r.font.name = "Times New Roman"

    r.font.size = Pt(10)



    # Second row, left: Total number of questions

    left_cell = header_table.cell(1, 0)

    left_cell.width = Inches(2.5)



    p = left_cell.paragraphs[0]

    p.paragraph_format.space_after = Pt(0)



    r = p.add_run(f"Total No. of Questions: {total_questions}")

    r.font.name = "Times New Roman"

    r.font.size = Pt(9)



    # Second row, right: Total number of pages

    right_cell = header_table.cell(1, 1)

    right_cell.width = Inches(2.5)



    p = right_cell.paragraphs[0]

    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p.paragraph_format.space_after = Pt(0)



    r = p.add_run("No. of Pages: ")

    r.font.name = "Times New Roman"

    r.font.size = Pt(9)



    # Automatic total page count

    fld_char = OxmlElement('w:fldSimple')

    fld_char.set(qn('w:instr'), 'NUMPAGES')



    run = OxmlElement('w:r')

    rPr = OxmlElement('w:rPr')



    font = OxmlElement('w:rFonts')

    font.set(qn('w:ascii'), 'Times New Roman')

    font.set(qn('w:hAnsi'), 'Times New Roman')

    rPr.append(font)



    run.append(rPr)



    text = OxmlElement('w:t')

    text.text = "1"

    run.append(text)



    fld_char.append(run)

    p._p.append(fld_char)



    # Institute and examination details

    add_paragraph(doc, meta["college"], bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)

    add_paragraph(doc, meta["degree_semester"], bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)

    add_paragraph(doc, meta["subject"], bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)

    add_paragraph(doc, f"Subject Code: {meta['subject_code']}", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)



    if meta["m_code"].strip():

        add_paragraph(doc, f"M. Code: {meta['m_code']}", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)



    add_paragraph(doc, f"Date of Examination: {meta['exam_date']}", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)



   # Time and Maximum Marks

    time_table = doc.add_table(rows=1, cols=2)

    time_table.autofit = False



    time_table.columns[0].width = Inches(2.5)

    time_table.columns[1].width = Inches(2.5)



    left_cell = time_table.cell(0, 0)

    right_cell = time_table.cell(0, 1)



    left_cell.width = Inches(2.5)

    right_cell.width = Inches(2.5)



    p = left_cell.paragraphs[0]

    p.paragraph_format.space_after = Pt(0)

    r = p.add_run(f"Time: {meta['duration']}")

    r.font.name = "Times New Roman"

    r.font.size = Pt(10)



    p = right_cell.paragraphs[0]

    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p.paragraph_format.space_after = Pt(0)

    r = p.add_run(f"Max. Marks: {meta['maximum_marks']}")

    r.font.name = "Times New Roman"

    r.font.size = Pt(10)



    # Exam instructions

    add_paragraph(

        doc,

        "INSTRUCTIONS TO CANDIDATES:",

        bold=True,

        size=10,

        space_after=2,

    )



    for instruction in instructions.splitlines():

        if instruction.strip():

            add_paragraph(

                doc,

                instruction,

                size=10,

                space_after=2,

            )





    # Sections and questions

    add_centered_heading(doc, "SECTION-A")

    add_paragraph(doc, "1. Write short notes on:", size=10, space_after=3)



    sub_labels = "abcdefghijklmnopqrstuvwxyz"



    for index, question in enumerate(section_a):

        label = sub_labels[index] if index < len(sub_labels) else str(index + 1)

        add_paragraph(doc, f"({label}) {question}", size=10, space_after=5)



    add_centered_heading(doc, "SECTION-B")

    q_number = 2



    for question in section_b:

        add_formatted_question(doc, q_number, question, marks_b)

        q_number += 1



    add_centered_heading(doc, "SECTION-C")



    for question in section_c:

        add_formatted_question(doc, q_number, question, marks_c)

        q_number += 1



    # UMC warning at the end

    add_paragraph(

        doc,

        UMC_LINE,

        bold=True,

        size=8,

        align=WD_ALIGN_PARAGRAPH.CENTER,

        space_after=0,

    )

        # --------------------------------------------------

    # Footer: Page X out of Y

    # --------------------------------------------------



    footer = sec.footer



    footer.is_linked_to_previous = False



    footer_p = footer.paragraphs[0]

    footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT



    footer_p.paragraph_format.space_before = Pt(0)

    footer_p.paragraph_format.space_after = Pt(0)



    run = footer_p.add_run("Page ")

    run.font.name = "Times New Roman"

    run.font.size = Pt(8)



    # Current page number

    fld_page = OxmlElement("w:fldSimple")

    fld_page.set(qn("w:instr"), "PAGE")



    run_page = OxmlElement("w:r")

    text_page = OxmlElement("w:t")

    text_page.text = "1"

    run_page.append(text_page)

    fld_page.append(run_page)

    footer_p._p.append(fld_page)



    run = footer_p.add_run(" out of ")

    run.font.name = "Times New Roman"

    run.font.size = Pt(8)



    # Total number of pages

    fld_total = OxmlElement("w:fldSimple")

    fld_total.set(qn("w:instr"), "NUMPAGES")



    run_total = OxmlElement("w:r")

    text_total = OxmlElement("w:t")

    text_total.text = "1"

    run_total.append(text_total)

    fld_total.append(run_total)

    footer_p._p.append(fld_total)



    settings = doc.settings.element

    update_fields = OxmlElement('w:updateFields')

    update_fields.set(qn('w:val'), 'true')

    settings.append(update_fields)

    output = io.BytesIO()

    doc.save(output)



    return output.getvalue()







class NumberedCanvas(canvas.Canvas):

    """Add question count, total pages and Page X out of Y."""



    def __init__(self, *args, total_questions=0, **kwargs):

        canvas.Canvas.__init__(self, *args, **kwargs)

        self.saved_page_states = []

        self.total_questions = total_questions



    def showPage(self):

        self.saved_page_states.append(dict(self.__dict__))

        self._startPage()



    def save(self):

        total_pages = len(self.saved_page_states)



        page_width, page_height = landscape(A4)



        margin = 0.55 * inch

        column_gap = 0.35 * inch



        column_width = (

            page_width - (2 * margin) - column_gap

        ) / 2



        for state in self.saved_page_states:

            self.__dict__.update(state)



            self.saveState()



            self.setFont("Times-Roman", 9)



            # First line: Roll Number

            self.drawString(

                margin,

                page_height - 0.32 * inch,

                "Roll No. __________________"

            )



            # Second line, left: Total questions

            self.drawString(

                margin,

                page_height - 0.49 * inch,

                f"Total No. of Questions: {self.total_questions}"

            )



            # Second line, right: Total pages (left column only)

            self.drawRightString(

                margin + column_width,

                page_height - 0.49 * inch,

                f"No. of Pages: {total_pages}"

            )



            # Footer: Page X out of Y

            self.setFont("Times-Roman", 8)



            self.drawRightString(

                page_width - margin,

                0.25 * inch,

                f"Page {self._pageNumber} out of {total_pages}"

            )



            self.restoreState()



            canvas.Canvas.showPage(self)



        canvas.Canvas.save(self)



def build_pdf(meta, section_a, section_b, section_c,

              marks_b, marks_c, instructions):



    buffer = io.BytesIO()



    page_width, page_height = landscape(A4)



    margin = 0.55 * inch

    column_gap = 0.35 * inch



    column_width = (

        page_width - (2 * margin) - column_gap

    ) / 2



    left_frame = Frame(

        margin,

        margin + 0.2 * inch,

        column_width,

        page_height - (2 * margin) - 0.2 * inch,

        leftPadding=0,

        rightPadding=0,

        topPadding=0,

        bottomPadding=0,

        id="left_column"

    )



    right_frame = Frame(

        margin + column_width + column_gap,

        margin + 0.2 * inch,

        column_width,

        page_height - (2 * margin) - 0.2 * inch,

        leftPadding=0,

        rightPadding=0,

        topPadding=0,

        bottomPadding=0,

        id="right_column"

    )



    doc = BaseDocTemplate(

        buffer,

        pagesize=landscape(A4),

        leftMargin=margin,

        rightMargin=margin,

        topMargin=margin,

        bottomMargin=margin

    )



    doc.addPageTemplates([

        PageTemplate(

            id="TwoColumns",

            frames=[left_frame, right_frame]

        )

    ])



    # --------------------------------------------------

    # PDF styles

    # --------------------------------------------------



    styles = getSampleStyleSheet()



    college_style = ParagraphStyle(

        "College",

        parent=styles["Normal"],

        fontName="Times-Bold",

        fontSize=12,

        leading=15,

        alignment=TA_CENTER,

        spaceAfter=3

    )



    heading_style = ParagraphStyle(

        "Heading",

        parent=styles["Normal"],

        fontName="Times-Bold",

        fontSize=10,

        leading=12,

        alignment=TA_CENTER,

        spaceAfter=3

    )



    normal_style = ParagraphStyle(

        "NormalCustom",

        parent=styles["Normal"],

        fontName="Times-Roman",

        fontSize=9.5,

        leading=12,

        alignment=TA_LEFT,

        spaceAfter=3

    )



    instruction_style = ParagraphStyle(

        "Instructions",

        parent=normal_style,

        fontSize=9,

        leading=11,

        spaceAfter=2

    )



    section_style = ParagraphStyle(

        "Section",

        parent=normal_style,

        fontName="Times-Bold",

        fontSize=10,

        leading=13,

        alignment=TA_CENTER,

        spaceBefore=7,

        spaceAfter=5

    )



    question_style = ParagraphStyle(

        "Question",

        parent=normal_style,

        fontSize=9.5,

        leading=12,

        spaceAfter=5

    )



    subpart_style = ParagraphStyle(

        "Subpart",

        parent=question_style,

        leftIndent=18,

        spaceAfter=3

    )



    umc_style = ParagraphStyle(

        "UMC",

        parent=normal_style,

        fontName="Times-Bold",

        fontSize=8,

        leading=10,

        alignment=TA_CENTER,

        spaceBefore=8

    )



    story = []



    # --------------------------------------------------

    # Examination header

    # --------------------------------------------------

    story.append(Paragraph(

        escape(meta["college"]), college_style

    ))



    story.append(Paragraph(

        escape(meta["degree_semester"]), heading_style

    ))



    story.append(Paragraph(

        escape(meta["subject"]), heading_style

    ))



    story.append(Paragraph(

        f"Subject Code: {escape(meta['subject_code'])}",

        heading_style

    ))



    if meta["m_code"].strip():

        story.append(Paragraph(

            f"M. Code: {escape(meta['m_code'])}",

            heading_style

        ))



    story.append(Paragraph(

        f"Date of Examination: {escape(meta['exam_date'])}",

        heading_style

    ))



    # Time and maximum marks



    time_table = Table(

        [[

            f"Time: {escape(meta['duration'])}",

            f"Max. Marks: {meta['maximum_marks']}"

        ]],

        colWidths=[column_width / 2] * 2

    )



    time_table.setStyle(TableStyle([

        ("FONTNAME", (0, 0), (-1, -1), "Times-Roman"),

        ("FONTSIZE", (0, 0), (-1, -1), 9),

        ("ALIGN", (0, 0), (0, 0), "LEFT"),

        ("ALIGN", (1, 0), (1, 0), "RIGHT"),

        ("TOPPADDING", (0, 0), (-1, -1), 5),

        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),

    ]))



    story.append(time_table)



    # --------------------------------------------------

    # Instructions

    # --------------------------------------------------



    story.append(Paragraph(

        "<b>INSTRUCTIONS TO CANDIDATES:</b>",

        instruction_style

    ))



    for instruction in instructions.splitlines():

        if instruction.strip():

            story.append(Paragraph(

                escape(instruction),

                instruction_style

            ))



    # --------------------------------------------------

    # Section A

    # --------------------------------------------------



    story.append(Paragraph("SECTION-A", section_style))



    story.append(Paragraph(

        "1. Write short notes on:",

        question_style

    ))



    sub_labels = "abcdefghijklmnopqrstuvwxyz"



    for index, question in enumerate(section_a):

        label = (

            sub_labels[index]

            if index < len(sub_labels)

            else str(index + 1)

        )



        story.append(Paragraph(

            f"({label}) {escape(str(question))}",

            question_style

        ))



    # --------------------------------------------------

    # Section B

    # --------------------------------------------------



    story.append(Paragraph("SECTION-B", section_style))



    q_number = 2



    def add_pdf_question(number, question):

        parts = format_question_parts(question).splitlines()

        parts = [part.strip() for part in parts if part.strip()]



        if not parts:

            parts = [str(question).strip()]



        elements = [

            Paragraph(

                f"{number}. {escape(parts[0])}",

                question_style

            )

        ]



        for part in parts[1:]:

            elements.append(Paragraph(

                escape(part),

                subpart_style

            ))



        return KeepTogether(elements)



    for question in section_b:

        story.append(add_pdf_question(q_number, question))

        q_number += 1



    # --------------------------------------------------

    # Section C

    # --------------------------------------------------



    story.append(Paragraph("SECTION-C", section_style))



    for question in section_c:

        story.append(add_pdf_question(q_number, question))

        q_number += 1



    # --------------------------------------------------

    # UMC warning

    # --------------------------------------------------



    story.append(Spacer(1, 8))



    story.append(Paragraph(

        escape(UMC_LINE),

        umc_style

    ))





    # --------------------------------------------------

    # Generate PDF with automatic question and page count

    # --------------------------------------------------



    total_questions = len(section_a) + len(section_b) + len(section_c)



    doc.build(

        story,

        canvasmaker=lambda *args, **kwargs: NumberedCanvas(

            *args,

            total_questions=total_questions,

            **kwargs

        )

    )



    buffer.seek(0)

    return buffer.getvalue()

step_header(1, "Exam details", "College, course, subject and examination particulars")

with st.expander("Exam details", expanded=True):

    left, right = st.columns(2)

    with left:

        college = st.text_input("College / institute", COLLEGE_DEFAULT)

        degree_semester = st.text_input("Course / semester heading", "B. Tech. (CSE) (Sem.–1)")

        subject = st.text_input("Subject name", "Engineering Physics")

        subject_code = st.text_input("Subject code", "BPH-101-26")

    with right:

        m_code = st.text_input("M. Code", "93794")

        exam_date = st.date_input("Date of examination", value=date.today())

        duration = st.text_input("Time / duration", "3 Hrs.")

        maximum_marks = st.number_input("Maximum marks", min_value=1, max_value=300, value=60, step=1)



# NEW: Editable instructions area

    instructions = st.text_area(

        "Exam Instructions",

        value=(

            "1. SECTION-A is COMPULSORY.\n"

            "2. SECTION-B contains the questions provided in the Section B question bank.\n"

            "3. SECTION-C contains the questions provided in the Section C question bank.\n"

            "4. Marks for each question are indicated with the respective section heading."

        ),

        height=150,

        help=(

            "Edit these instructions before generating the question paper. "

            "Each line will appear as a separate instruction in the generated question paper."

        ),

    )

step_header(2, "Upload question banks", "Word (.docx) or Excel (.xlsx / .xls) files are supported")

col_a, col_b, col_c = st.columns(3)
with col_a:
    short_upload = st.file_uploader("Section A — 2-mark questions", type=["docx", "xlsx", "xls"], key="short_bank")
with col_b:
    section_b_upload = st.file_uploader("Section B — questions", type=["docx", "xlsx", "xls"], key="section_b_bank")
with col_c:
    section_c_upload = st.file_uploader("Section C — questions", type=["docx", "xlsx", "xls"], key="section_c_bank")


def load_if_changed(upload, state_key, display_name):

    if upload is None:

        st.session_state.pop(state_key, None)

        st.session_state.pop(state_key + "_fingerprint", None)

        return []

    fingerprint = f"{upload.name}:{upload.size}"

    if st.session_state.get(state_key + "_fingerprint") != fingerprint:

        try:

            st.session_state[state_key] = read_bank(upload)

            st.session_state[state_key + "_fingerprint"] = fingerprint

            st.session_state.pop("selected_paper", None)

        except Exception as exc:

            st.error(f"Couldn't read {display_name}: {exc}")

            return []

    return st.session_state.get(state_key, [])





short_questions = load_if_changed(short_upload, "short_questions", "the Section A question bank")
section_b_questions = load_if_changed(section_b_upload, "section_b_questions", "the Section B question bank")
section_c_questions = load_if_changed(section_c_upload, "section_c_questions", "the Section C question bank")

chips = []
if short_upload:
    chips.append(f'<span class="chip">Section A bank: {len(short_questions)} questions loaded</span>')
if section_b_upload:
    chips.append(f'<span class="chip">Section B bank: {len(section_b_questions)} questions loaded</span>')
if section_c_upload:
    chips.append(f'<span class="chip">Section C bank: {len(section_c_questions)} questions loaded</span>')
if chips:
    st.markdown(f'<div class="chip-row">{"".join(chips)}</div>', unsafe_allow_html=True)


step_header(3, "Question counts and marks", "Decide how many questions each section should contain")
ca, cb, cc = st.columns(3)
with ca:
    count_a = st.number_input("Section A questions", min_value=0, max_value=200, value=10, step=1)
with cb:
    count_b = st.number_input("Section B questions", min_value=0, max_value=200, value=4, step=1)
with cc:
    count_c = st.number_input("Section C questions", min_value=0, max_value=200, value=4, step=1)
mb, mc = st.columns(2)
with mb:
    marks_b = st.selectbox("Marks per Section B question", MARK_OPTIONS, index=3)
with mc:
    marks_c = st.selectbox("Marks per Section C question", MARK_OPTIONS, index=3)

if int(count_a) > len(short_questions):
    st.info(f"Section A requested {int(count_a)} questions; only {len(short_questions)} are available. All available questions will be used.")
if int(count_b) > len(section_b_questions):
    st.info(f"Section B requested {int(count_b)} questions; only {len(section_b_questions)} are available. All available questions will be used.")
if int(count_c) > len(section_c_questions):
    st.info(f"Section C requested {int(count_c)} questions; only {len(section_c_questions)} are available. All available questions will be used.")


def choose_paper():
    a = random.sample(short_questions, min(int(count_a), len(short_questions)))
    b = random.sample(section_b_questions, min(int(count_b), len(section_b_questions)))
    c = random.sample(section_c_questions, min(int(count_c), len(section_c_questions)))
    st.session_state["selected_paper"] = {"A": a, "B": b, "C": c}


st.write("")

act1, act2 = st.columns(2)

with act1:

    if st.button("🎲 Generate random paper", type="primary", use_container_width=True):

        if not short_questions and not section_b_questions and not section_c_questions:

            st.error("Upload at least one question bank first.")

        else:

            choose_paper()

            st.rerun()

with act2:

    if st.button("🔀 Shuffle all sections", use_container_width=True):

        if not short_questions and not section_b_questions and not section_c_questions:

            st.error("Upload at least one question bank first.")

        else:

            choose_paper()

            st.rerun()



selected = st.session_state.get("selected_paper")

if selected is not None:

    step_header(4, "Preview and shuffle", "Review the paper, reshuffle any section, then download")



    esc = html.escape

    preview_html = textwrap.dedent(f"""

        <div class="paper">

            <div class="p-college">{esc(college)}</div>

            <div class="p-line">{esc(degree_semester)}</div>

            <div class="p-line" style="font-size:1.1rem;">{esc(subject)}</div>

            <div class="p-meta">

                <span><b>Subject Code:</b> {esc(subject_code) or '—'}</span>

                <span><b>M. Code:</b> {esc(m_code) or '—'}</span>

                <span><b>Date:</b> {exam_date.strftime('%d-%m-%Y')}</span>

                <span><b>Time:</b> {esc(duration)}</span>

                <span><b>Max. Marks:</b> {maximum_marks}</span>

            </div>

        </div>

        """).strip()

    st.html(preview_html)


    for key, title, requested in [("A", "SECTION-A — Write short notes on:", int(count_a)), ("B", f"SECTION-B ({marks_b} marks each)", int(count_b)), ("C", f"SECTION-C ({marks_c} marks each)", int(count_c))]:

        head_col, shuffle_col = st.columns([5, 1])

        with head_col:

            st.markdown(f'<div class="section-bar">{esc(title)}</div>', unsafe_allow_html=True)

        with shuffle_col:

            st.markdown('<div style="height:22px"></div>', unsafe_allow_html=True)

            if st.button("Shuffle", key=f"shuffle_{key}", use_container_width=True):

                if key == "A":
                    selected["A"] = random.sample(short_questions, min(int(count_a), len(short_questions)))
                elif key == "B":
                    selected["B"] = random.sample(section_b_questions, min(requested, len(section_b_questions)))
                else:
                    selected["C"] = random.sample(section_c_questions, min(requested, len(section_c_questions)))
                st.session_state["selected_paper"] = selected

                st.rerun()

        questions = selected.get(key, [])



        if not questions:

            st.markdown(

                '<div class="empty-note">No questions selected for this section.</div>',

                unsafe_allow_html=True

            )

        else:

            for i, question in enumerate(questions, start=1):

                mark_badge = ""



                mark_badge = ""



                formatted_question = html.escape(

                    format_question_parts(question)

                ).replace("\n", "<br>")



                question_html = (

                    f'<div class="q-item">'

                    f'{mark_badge}'

                    f'<span class="q-no">{i}.</span>'

                    f'{formatted_question}'

                    f'</div>'

                )



                st.markdown(

                    question_html,

                    unsafe_allow_html=True

                )

    meta = {

        "college": college,

        "degree_semester": degree_semester,

        "subject": subject,

        "subject_code": subject_code,

        "m_code": m_code,

        "exam_date": exam_date.strftime("%d-%m-%Y"),

        "duration": duration,

        "maximum_marks": int(maximum_marks),

    }

    st.write("")

    # Prepare filename

    safe_subject = re.sub(

        r"[^A-Za-z0-9_-]+", "_", subject

    ).strip("_") or "SVIET_Exam"

    word_col, pdf_col = st.columns(2)

    # DIRECT WORD DOWNLOAD

    with word_col:

        try:

            word_data = build_docx(

                meta,

                selected.get("A", []),

                selected.get("B", []),

                selected.get("C", []),

                int(marks_b),

                int(marks_c),

                instructions

            )

            st.download_button(

                "⬇️ Download Word",

                data=word_data,

                file_name=f"{safe_subject}_Question_Paper.docx",

                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",

                use_container_width=True,

            )

        except Exception as word_error:

            st.error(f"Word generation failed: {word_error}")



    # DIRECT PDF DOWNLOAD

    with pdf_col:

        try:

            pdf_data = build_pdf(

                meta,

                selected.get("A", []),

                selected.get("B", []),

                selected.get("C", []),

                int(marks_b),

                int(marks_c),


                instructions

            )



            st.download_button(

                "📄 Download PDF",

                data=pdf_data,

                file_name=f"{safe_subject}_Question_Paper.pdf",

                mime="application/pdf",

                use_container_width=True,

            )



        except Exception as pdf_error:

            st.error(f"PDF generation failed: {pdf_error}")
