from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from io import BytesIO

app = Flask(__name__)

ASSETS = "assets"

SUBJECT_ICONS = {
    "Math": "Math Icon.png",
    "English": "English Icon.png",
    "Science": "Science Icon.png",
    "Spanish": "Spanish Icon.png",
    "Art and Design": "Art and Design Icon.png",
    "Guided Reading": "Guided Reading Icon.png",
    "Design and Technology": "Design and Technology Icon.png",
    "Homework": "Homework Icon.png",
    "Religious Education": "Religious Education Icon.png",
    "Geography": "Geography Icon.png",
    "History": "History Icon.png",
    "Music": "Music Icon.png"
}


def format_name(name):
    """
    Converts:
    'Surname, Firstname'
    into:
    'Firstname Surname'
    """

    name = str(name).strip()

    if "," in name:
        surname, firstname = name.split(",", 1)

        surname = surname.strip()
        firstname = firstname.strip()

        return f"{firstname} {surname}"

    return name

def chunk_list(data, size):
    """Split list into chunks of 8"""
    for i in range(0, len(data), size):
        yield data[i:i + size]


from docx import Document
from docx.shared import Cm, Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from io import BytesIO
import os


def chunk_list(data, size):
    """Splits students into groups of 8 (1 page each)"""
    for i in range(0, len(data), size):
        yield data[i:i + size]
from docx.oxml.ns import qn

FONT_NAME = "NTPreCursivefk"

def apply_font(run, size):
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    run.bold = True

    run._element.rPr.rFonts.set(
        qn('w:eastAsia'),
        FONT_NAME
    )

def build_docx(students, year_group, subject):
    doc = Document()

    # =========================
    # PAGE SETUP (A4)
    # =========================
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)

    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(0.8)
    section.right_margin = Cm(0.8)

    LABELS_PER_PAGE = 8
    ROWS = 4
    COLS = 2

    label_width = Cm(9.8)
    label_height = Cm(6.7)

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    # =========================
    # MULTI-PAGE LOOP (IMPORTANT FIX)
    # =========================
    pages = list(chunk_list(students, LABELS_PER_PAGE))

    for page_index, page_students in enumerate(pages):

        table = doc.add_table(rows=ROWS, cols=COLS)
        table.autofit = False

        # enforce row height stability
        for row in table.rows:
            row.height = label_height
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # =========================
        # BUILD 8 LABELS PER PAGE
        # =========================
        for r in range(ROWS):
            for c in range(COLS):

                idx = r * COLS + c
                cell = table.cell(r, c)
                cell.width = label_width

                # skip empty cells (last page safety)
                if idx >= len(page_students):
                    continue

                student = page_students[idx]
                student = format_name(student)
                
# =========================
# LOGO
# =========================
p_logo = cell.paragraphs[0]

p_logo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

# move logo 2mm right
p_logo.paragraph_format.left_indent = Mm(2)

run_logo = p_logo.add_run()

try:
    run_logo.add_picture(
        logo_path,
        width=Cm(1.8)
    )
except:
    pass

                # =========================
                # STUDENT NAME
                # =========================
                p1 = cell.add_paragraph()

p1.paragraph_format.left_indent = Mm(2)
p1.paragraph_format.right_indent = Mm(2)

p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

               r1 = p1.add_run(student)

apply_font(r1, 18)

                r1.font.size = Pt(18)

                # =========================
                # SUBJECT
                # =========================
                p2 = cell.add_paragraph()

p2.paragraph_format.left_indent = Mm(2)
p2.paragraph_format.right_indent = Mm(2)

p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                r2 = p2.add_run(subject)

apply_font(r2, 16)

                # =========================
                # YEAR GROUP
                # =========================
                p3 = cell.add_paragraph()

p3.paragraph_format.left_indent = Mm(2)
p3.paragraph_format.right_indent = Mm(2)

p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

              r3 = p3.add_run(f"Year {year_group}")

apply_font(r3, 16)

                # =========================
                # ICON
                # =========================
               p_icon = cell.add_paragraph()

# keep icon inside safe area
p_icon.paragraph_format.right_indent = Mm(2)

p_icon.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                run_icon = p_icon.add_run()
                try:
                    run_icon.add_picture(icon_path, width=Cm(1.6))
                except:
                    pass

        # =========================
        # PAGE BREAK (CRITICAL FIX)
        # =========================
        if page_index < len(pages) - 1:
            doc.add_page_break()

    # =========================
    # RETURN FILE
    # =========================
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    file = request.files["file"]
    year_group = request.form["year_group"]
    subject = request.form["subject"]

    df = pd.read_excel(file)
    students = df.iloc[:, 0].dropna().tolist()

    docx_file = build_docx(students, year_group, subject)

    return send_file(
        docx_file,
        as_attachment=True,
        download_name=f"{subject}_Year{year_group}_Labels.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
