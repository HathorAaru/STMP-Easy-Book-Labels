from copy import deepcopy
from docx import Document
from io import BytesIO
import os
from docx.shared import Cm, Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def build_docx(students, year_group, subject):

    template_path = os.path.join(ASSETS, "label_template.docx")
    base_doc = Document(template_path)

    # Take first table as template
    template_table = base_doc.tables[0]

    logo_path = os.path.join(ASSETS, "STMP Logo.png")
    icon_path = os.path.join(ASSETS, SUBJECT_ICONS.get(subject, ""))

    def fill_table(table, student_chunk):
        cells = [cell for row in table.rows for cell in row.cells]

        for i, cell in enumerate(cells):

            cell.text = ""

            if i >= len(student_chunk):
                continue

            student = format_name(student_chunk[i])

            # logo
            p_logo = cell.paragraphs[0]
            run_logo = p_logo.add_run()
            try:
                run_logo.add_picture(logo_path, width=Mm(18))
            except:
                pass

            # name
            p1 = cell.add_paragraph()
            p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            r1 = p1.add_run(student)
            r1.bold = True
            r1.font.size = Pt(16)

            # subject
            p2 = cell.add_paragraph()
            p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p2.add_run(subject).font.size = Pt(14)

            # year
            p3 = cell.add_paragraph()
            p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p3.add_run(f"Year {year_group}").font.size = Pt(14)

            # icon
            p_icon = cell.add_paragraph()
            p_icon.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            run_icon = p_icon.add_run()
            try:
                run_icon.add_picture(icon_path, width=Mm(18))
            except:
                pass

    # split students into chunks of 8
    chunks = [students[i:i + 8] for i in range(0, len(students), 8)]

    doc = Document()  # new document (we rebuild pages)

    for page_index, chunk in enumerate(chunks):

        # clone template table
        table = deepcopy(template_table)
        doc._body._element.append(table._element)

        fill_table(table, chunk)

        # add page break after each page except last
        if page_index != len(chunks) - 1:
            doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
