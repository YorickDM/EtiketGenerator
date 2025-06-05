import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def parse_ranges(range_str):
    parts = re.split(r",\s*", range_str)
    result = []
    for part in parts:
        if "-" in part:
            start, end = map(int, part.split("-"))
            result.extend(range(start, end + 1))
        else:
            try:
                result.append(int(part))
            except ValueError:
                pass
    return result

def set_cell_spacing(paragraph, afstand_voor=None):
    p = paragraph._element
    pPr = p.get_or_add_pPr()

    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(afstand_voor) if afstand_voor is not None else "0")
    spacing.set(qn('w:after'), "0")
    spacing.set(qn('w:line'), "240")
    spacing.set(qn('w:lineRule'), "auto")
    pPr.append(spacing)

    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), "0")
    ind.set(qn('w:firstLine'), "0")
    pPr.append(ind)

def generate_box_labels(naam, ToegangsNmr, groups):
    labels = []
    for group in groups:
        cleaned = group.strip()
        if cleaned:
            label = [
                "Stadsarchief Amsterdam",
                f"{ToegangsNmr}".strip(),
                f"{naam}".strip(),
                f"{cleaned}"
            ]
            labels.append(label)
    return create_docx_table(labels)

def create_docx_table(labels):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.08)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.39)
    section.right_margin = Inches(0.39)

    def add_label_table(label_block):
        table = doc.add_table(rows=9, cols=3)
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblCellMar = OxmlElement('w:tblCellMar')
        for side in ['left', 'right']:
            mar = OxmlElement(f'w:{side}')
            mar.set(qn('w:w'), '15')
            mar.set(qn('w:type'), 'dxa')
            tblCellMar.append(mar)
        for side in ['top', 'bottom']:
            mar = OxmlElement(f'w:{side}')
            mar.set(qn('w:w'), '0')
            mar.set(qn('w:type'), 'dxa')
            tblCellMar.append(mar)
        tblPr.append(tblCellMar)

        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        widths = [Inches(2.4805), Inches(2.4805), Inches(2.4805)]
        for row in table.rows:
            row.height = Inches(1.26)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            for i, cell in enumerate(row.cells):
                cell.width = widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                if cell.paragraphs:
                    p = cell.paragraphs[0]._element
                    p.getparent().remove(p)

        idx = 0
        for row in table.rows:
            for cell in row.cells:
                if idx < len(label_block):
                    content = label_block[idx]
                    for i, line in enumerate(content):
                        if line.strip():
                            para = cell.add_paragraph()
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = para.add_run(line)
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
                            spacing_before = 226 if i == 0 else None
                            if i == 3 or "Inventaris" in line or i == 1:
                                run.bold = True
                                run.font.size = Pt(12)
                            para.paragraph_format.space_after = Pt(0)
                            para.paragraph_format.line_spacing_rule = 1
                            set_cell_spacing(para, afstand_voor=spacing_before)
                    idx += 1

    # Split labels into blocks of 27 and add a table for each
    for i in range(0, len(labels), 27):
        if i > 0:
            doc.add_page_break()
        add_label_table(labels[i:i+27])

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

import csv

def load_toegangstitels(csv_path="ToegangenLijst.csv"):
    mapping = {}
    try:
        with open(csv_path, newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                nummer = row.get("toegangsnummer")
                titel = row.get("titel")
                if nummer and titel:
                    mapping[nummer.strip()] = titel.strip()
    except Exception as e:
        print("CSV leesfout:", e)
    return mapping

TOEGANGSTITELS = load_toegangstitels()

def main():
    st.title("📋 Stadsarchief Label Generator")

    option = st.radio("Wat wil je doen?", [
        "📁 Omslagetiketten maken",
        "📦 Doosetiketten maken"
    ])

    if option == "📁 Omslagetiketten maken":
        toegangsnummer = st.text_input("Toegangsnummer", "")
        titel_default = TOEGANGSTITELS.get(toegangsnummer.strip(), "")
        titel = st.text_input("Archiefnaam", titel_default)
        van = st.number_input("Inventarisnummer vanaf", min_value=0, value=0)
        tot = st.number_input("Inventarisnummer t/m", min_value=0, value=1)
        a_nummers_input = st.text_input("A-nummers (optioneel, bijv. 68A, 99B)", "")

        inventarisnummers = [str(n) for n in range(van, tot + 1)]
        if a_nummers_input.strip():
            extra = [x.strip() for x in a_nummers_input.split(',') if x.strip()]
            inventarisnummers.extend(extra)

        if st.button("🎫 Genereer etiketten (.docx)"):
            labels = [
                ["Stadsarchief Amsterdam", f"{toegangsnummer}", f"{titel}", f"{str(num)}"]
                for num in inventarisnummers
            ]
            docx_file = create_docx_table(labels)
            st.download_button("⬇️ Download als DOCX", docx_file, file_name="omslagetiketten" + toegangsnummer + ".docx")

    elif option == "📦 Doosetiketten maken":
        ToegangsNmr = st.text_input("Toegangsnummer", "")
        naam_default = TOEGANGSTITELS.get(ToegangsNmr.strip(), "")
        naam = st.text_input("Archiefnaam", naam_default)

        invoer = st.text_area("Inventarisnummers (Scheid etiketten met een nieuwe regel, gebruik ',' voor losse nummers en '-' voor reeksen)", "")

        if st.button("📦 Genereer doosetiketten (.docx)"):
            groups = [grp.strip() for grp in invoer.strip().split("\n") if grp.strip()]
            docx_file = generate_box_labels(naam, ToegangsNmr, groups)
            st.download_button("⬇️ Download als DOCX", docx_file, file_name="doosetiketten" + ToegangsNmr + ".docx")

    st.markdown("####### Made by Yorick de Man" + " V0.3")

if __name__ == "__main__":
    main()
