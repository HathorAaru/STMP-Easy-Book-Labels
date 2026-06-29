from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
# CLEAN CELL COMPLETELY (IMPORTANT FIX)
# =========================
def clear_cell(cell):
    tc = cell._tc
    tc.clear_content()  # removes ALL hidden paragraphs safely


# =========================
# FIX TABLE STRUCTURE
# =========================
def fix_table_layout(table):
    tbl = table._tbl
    tblPr = tbl.tblPr

    # fixed layout
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


# =========================
# CHUNK LIST
# =========================
def chunk_list(data, size):
    for i in range(0, len(data), size):
        yield data[i:i + size]


# =========================
# BUILD ONE LABEL CELL
# =========================
def build_label(cell, student, year_group, subject, logo_path, icon_path):

    clear_cell(cell)

    # --- logo ---
    p_logo = cell.paragraphs[0]
    run_logo = p_logo.add_run()
    try:
        run_logo.add_picture(logo_path, width=Mm(18))
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

    r2 = p2.add_run(subject)
    r2.font.size = Pt(12)

    # --- year ---
    p3 = cell.add_paragraph()
    p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(0)

    r3 = p3.add_run(f"Year {year_group}")
    r3.font.size = Pt(12)

    # --- icon ---
    p_icon = cell.add_paragraph()
    p_icon.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p_icon.paragraph_format.space_before = Pt(0)
    p_icon.paragraph_format.space_after = Pt(0)

    run_icon = p_icon.add_run()
    try:
        run_icon.add_picture(icon_path, width=Mm(14))
    except:
        pass


# =========================
# BUILD DOCX
# =========================
def build_docx(students, year_group, subject):

    doc = Document(os.path.join(ASSETS, "label_template.docx"))

    table = doc.tables[0]
    fix_table_layout(table)

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    rows = len(table.rows)
    cols = len(table.columns)
    labels_per_page = rows * cols  # MUST be 8 in your template

    pages = list(chunk_list(students, labels_per_page))

    # IMPORTANT: prevent template ghost pages
    doc._body.clear_content()
    doc._body._element.append(table._tbl)

    for page_index, page_students in enumerate(pages):

        if page_index > 0:
            doc.add_page_break()
            table = doc.add_table(rows=rows, cols=cols)
            fix_table_layout(table)

        cells = [c for r in table.rows for c in r.cells]

        # force exact label fill
        for i in range(labels_per_page):
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

        # lock row heights (prevents “bleeding between columns”)
        for row in table.rows:
            row.height = Cm(3.8)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

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
