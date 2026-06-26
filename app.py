from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt, Mm
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
    "Music": "Music Icon.png"
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
# SPLIT INTO PAGES
# =========================
def chunk_list(data, size):
    for i in range(0, len(data), size):
        yield data[i:i + size]


# =========================
# CELL MARGINS
# =========================
def set_cell_margins(cell, top=80, bottom=80, left=80, right=80):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcMar = OxmlElement("w:tcMar")

    for name, value in [("top", top), ("bottom", bottom), ("start", left), ("end", right)]:
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)

    tcPr.append(tcMar)


# =========================
# BUILD DOCX
# =========================
def build_docx(students, year_group, subject):

    doc = Document()

    # =========================
    # A4 PAGE SETUP
    # =========================
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)

    # UPDATED MARGINS (YOUR REQUEST)
    section.top_margin = Mm(12)
    section.bottom_margin = Mm(12)

    section.left_margin = Mm(5)
    section.right_margin = Mm(5)

    # =========================
    # LABEL SPECS (AVERY L7165)
    # =========================
    LABEL_W = Mm(99.06)
    LABEL_H = Mm(67.73)

    ROWS = 4
    COLS = 2
    LABELS_PER_PAGE = 8

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    pages = list(chunk_list(students, LABELS_PER_PAGE))

    for page_index, page_students in enumerate(pages):

        table = doc.add_table(rows=ROWS, cols=COLS)
        table.autofit = False

        # FORCE ROW HEIGHTS
        for row in table.rows:
            row.height = LABEL_H
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        for r in range(ROWS):
            for c in range(COLS):

                idx = r * COLS + c
                cell = table.cell(r, c)

                # clear cell content
                cell.text = ""

                # prevents overflow into adjacent labels
                set_cell_margins(cell, top=70, bottom=70, left=70, right=70)

                if idx >= len(students):
                    continue

                student = format_name(students[idx])

                # =========================
                # LOGO (TOP LEFT, FIXED)
                # =========================
                p_logo = cell.paragraphs[0]
                p_logo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                run_logo = p_logo.add_run()
                try:
                    run_logo.add_picture(logo_path, width=Mm(18))
                except:
                    pass

                # =========================
                # STUDENT NAME (WRAPS NATURALLY)
                # =========================
                p1 = cell.add_paragraph()
                p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                r1 = p1.add_run(student)
                r1.bold = True
                r1.font.size = Pt(16)

                # =========================
                # SUBJECT
                # =========================
                p2 = cell.add_paragraph()
                p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                r2 = p2.add_run(subject)
                r2.font.size = Pt(14)

                # =========================
                # YEAR GROUP
                # =========================
                p3 = cell.add_paragraph()
                p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                r3 = p3.add_run(f"Year {year_group}")
                r3.font.size = Pt(14)

                # =========================
                # SPACER
                # =========================
                cell.add_paragraph()

                # =========================
                # ICON (BOTTOM RIGHT, FIXED)
                # =========================
                p_icon = cell.add_paragraph()
                p_icon.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                run_icon = p_icon.add_run()
                try:
                    run_icon.add_picture(icon_path, width=Mm(14))
                except:
                    pass

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
# RUN APP
# =========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
