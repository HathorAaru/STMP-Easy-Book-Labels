from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy
import os
from io import BytesIO

app = Flask(__name__)

ASSETS = "assets"

# =========================
# SUBJECT ICONS
# =========================
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
    "Music": "Music Icon.png",
    "Brass Band": "Brass Band Icon.png"
}

# =========================
# FORMAT NAME
# =========================
def format_name(name):
    name = str(name).strip()
    if "," in name:
        surname, firstname = name.split(",", 1)
        return f"{firstname.strip()} {surname.strip()}"
    return name


# =========================
# CHUNK LIST
# =========================
def chunk_list(data, size):
    for i in range(0, len(data), size):
        yield data[i:i + size]


# =========================
# CLEAR CELL SAFELY
# =========================
def clear_cell(cell):
    cell._element.clear_content()


# =========================
# FIX TABLE LAYOUT
# =========================
def fix_table(table):
    tbl = table._tbl
    tblPr = tbl.tblPr

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


# =========================
# BUILD LABEL
# =========================
def build_label(cell, student, year_group, subject, logo_path, icon_path):

    clear_cell(cell)

    # --- logo ---
    p = cell.paragraphs[0]
    run = p.add_run()
    try:
        run.add_picture(logo_path, width=Mm(18))
    except:
        pass

    # --- name ---
    p1 = cell.add_paragraph()
    p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run(student)
    r1.bold = True
    r1.font.size = Pt(14)

    # --- subject ---
    p2 = cell.add_paragraph()
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    p2.add_run(subject).font.size = Pt(12)

    # --- year ---
    p3 = cell.add_paragraph()
    p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(0)
    p3.add_run(f"Year {year_group}").font.size = Pt(12)

    # --- icon ---
    p4 = cell.add_paragraph()
    p4.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after = Pt(0)
    run_icon = p4.add_run()

    try:
        run_icon.add_picture(icon_path, width=Mm(14))
    except:
        pass


# =========================
# BUILD DOCX
# =========================
def build_docx(students, year_group, subject):

    template = Document(os.path.join(ASSETS, "label_template.docx"))
    base_table = template.tables[0]

    fix_table(base_table)

    rows = len(base_table.rows)
    cols = len(base_table.columns)
    per_page = rows * cols  # MUST be 8 in your template

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    pages = list(chunk_list(students, per_page))

    # NEW document (clean safe start)
    doc = Document()

    for page_index, page_students in enumerate(pages):

        table = deepcopy(base_table._tbl)
        doc._body._element.append(table)

        table_obj = doc.tables[-1]
        fix_table(table_obj)

        cells = [c for r in table_obj.rows for c in r.cells]

        for i in range(per_page):
            cell = cells[i]

            if i < len(page_students):
                build_label(
                    cell,
                    format_name(page_students[i]),
                    year_group,
                    subject,
                    logo_path,
                    icon_path
                )
            else:
                clear_cell(cell)

        # lock row height (prevents overlap into gaps)
        for row in table_obj.rows:
            row.height = Cm(3.8)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # page break between tables
        if page_index < len(pages) - 1:
            doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =========================
# ROUTES
# =========================
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


# =========================
# RUN
# =========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
