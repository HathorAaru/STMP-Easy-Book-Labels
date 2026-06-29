from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
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
# CHUNK LIST (8 per page)
# =========================
def chunk_list(data, size):
    for i in range(0, len(data), size):
        yield data[i:i + size]


# =========================
# CLEAR CELL CONTENT SAFELY
# =========================
def clear_cell(cell):
    cell.text = ""


# =========================
# BUILD SINGLE LABEL
# =========================
def build_label(cell, student, year_group, subject, logo_path, icon_path):

    clear_cell(cell)

    # Logo
    p0 = cell.paragraphs[0]
    run0 = p0.add_run()
    try:
        run0.add_picture(logo_path, width=Mm(18))
    except:
        pass

    # Name
    p1 = cell.add_paragraph()
    p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r1 = p1.add_run(student)
    r1.bold = True
    r1.font.size = Pt(14)

    # Subject
    p2 = cell.add_paragraph()
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r2 = p2.add_run(subject)
    r2.font.size = Pt(12)

    # Year
    p3 = cell.add_paragraph()
    p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r3 = p3.add_run(f"Year {year_group}")
    r3.font.size = Pt(12)

    # Icon
    p4 = cell.add_paragraph()
    p4.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_icon = p4.add_run()

    try:
        run_icon.add_picture(icon_path, width=Mm(14))
    except:
        pass


# =========================
# MAIN DOCX BUILDER
# =========================
def build_docx(students, year_group, subject):

    template = Document(os.path.join(ASSETS, "label_template.docx"))
    base_table = template.tables[0]

    rows = len(base_table.rows)
    cols = len(base_table.columns)
    per_page = rows * cols  # MUST be 8

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    doc = Document()

    pages = list(chunk_list(students, per_page))

    for page_index, page_students in enumerate(pages):

        # ✅ TRUE TEMPLATE DUPLICATION (safe)
        page_table = doc.add_table(rows=rows, cols=cols)
        page_table.style = base_table.style

        # copy structure feel (widths/layout are inherited from style)
        idx = 0

        for r in range(rows):
            for c in range(cols):

                cell = page_table.cell(r, c)

                if idx < len(page_students):
                    build_label(
                        cell,
                        format_name(page_students[idx]),
                        year_group,
                        subject,
                        logo_path,
                        icon_path
                    )
                else:
                    clear_cell(cell)

                idx += 1

        # lock row height (prevents overflow into gaps)
        for row in page_table.rows:
            row.height = Mm(38)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # page break between duplicated template pages
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

    df = pd.read_excel(file, engine="openpyxl")
    students = df.iloc[:, 0].dropna().astype(str).tolist()

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
