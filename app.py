from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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
# CHUNK INTO PAGES (8 labels)
# =========================
def chunk_list(data, size=8):
    for i in range(0, len(data), size):
        yield data[i:i + size]


# =========================
# BUILD DOCX USING TEMPLATE
# =========================
def build_docx(students, year_group, subject):

    template_path = os.path.join(ASSETS, "label_template.docx")
    base_doc = Document(template_path)

    template_table = base_doc.tables[0]

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    # Create fresh document
    doc = Document()

    pages = list(chunk_list(students, 8))

    for page_index, chunk in enumerate(pages):

        # create new table with SAME structure as template
        table = doc.add_table(rows=len(template_table.rows),
                               cols=len(template_table.columns))

        table.autofit = False

        cells = [cell for row in table.rows for cell in row.cells]

        for i, cell in enumerate(cells):

            cell.text = ""

            if i >= len(chunk):
                continue

            student = format_name(chunk[i])

            # =====================
            # LOGO
            # =====================
            p_logo = cell.paragraphs[0]
            run_logo = p_logo.add_run()
            try:
                run_logo.add_picture(logo_path, width=Mm(18))
            except:
                pass

            # =====================
            # NAME
            # =====================
            p1 = cell.add_paragraph()
            p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            r1 = p1.add_run(student)
            r1.bold = True
            r1.font.size = Pt(16)

            # =====================
            # SUBJECT
            # =====================
            p2 = cell.add_paragraph()
            p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p2.add_run(subject).font.size = Pt(14)

            # =====================
            # YEAR
            # =====================
            p3 = cell.add_paragraph()
            p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p3.add_run(f"Year {year_group}").font.size = Pt(14)

            # =====================
            # ICON
            # =====================
            p_icon = cell.add_paragraph()
            p_icon.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            run_icon = p_icon.add_run()
            try:
                run_icon.add_picture(icon_path, width=Mm(14))
            except:
                pass

        # add page break
        if page_index != len(pages) - 1:
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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
