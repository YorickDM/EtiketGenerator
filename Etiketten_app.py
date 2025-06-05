import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import csv

def parse_ranges(range_str):
    parts = re.split(r",\s*", range_str)
    result = []
    for part in parts:
        if "-" in part:
            try:
                start, end = map(int, part.split("-"))
                result.extend([n for n in range(start, end + 1) if n > 0])
            except ValueError:
                pass
        else:
            try:
                n = int(part)
                if n > 0:
                    result.append(n)
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

    for i in range(0, len(labels), 27):
        label_block = labels[i:i+27]
        add_label_table(label_block)
        if i + 27 < len(labels):
            doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_box_labels(titel, nummer, groepen):
    labels = []
    for groep in groepen:
        nummers = parse_ranges(groep)
        if not nummers:
            continue
        eerste = nummers[0]
        laatste = nummers[-1]
        if eerste == laatste:
            formatted = f"{eerste}"
        else:
            formatted = f"{eerste}–{laatste}"
        labels.append([
            "Stadsarchief Amsterdam",
            f"{nummer}",
            f"{titel}",
            formatted
        ])
    return create_docx_table(labels)

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
        multi_ui = st.toggle("Meerdere toegangen toevoegen?", value=False)

        if not multi_ui:
            if 'vorige_enkel_toegang' not in st.session_state:
                st.session_state['vorige_enkel_toegang'] = ''
            if 'invoer_enkel' not in st.session_state:
                st.session_state['invoer_enkel'] = ''

            toegangsnummer = st.text_input("Toegangsnummer", "")
            if toegangsnummer != st.session_state['vorige_enkel_toegang']:
                st.session_state['invoer_enkel'] = ''

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

            st.session_state['vorige_enkel_toegang'] = toegangsnummer

        else:
            st.subheader("Meerdere toegangen invoeren")

            if "omslagen" not in st.session_state:
                st.session_state["omslagen"] = []

            if "prev_omslag_toegang" not in st.session_state:
                st.session_state["prev_omslag_toegang"] = ""

            toegang_nmr = st.text_input("Toegangsnummer", "")

            if toegang_nmr != st.session_state["prev_omslag_toegang"]:
                st.session_state["van_multi"] = 0
                st.session_state["tot_multi"] = 1
                st.session_state["a_multi"] = ""
                st.session_state["prev_omslag_toegang"] = toegang_nmr

            toegang_titel_default = TOEGANGSTITELS.get(toegang_nmr.strip(), "")

            with st.form("omslag_form"):
                toegang_titel = st.text_input("Archiefnaam", toegang_titel_default)
                van = st.number_input("Inventarisnummer vanaf", min_value=0, value=st.session_state.get("van_multi", 0), key="van_multi")
                tot = st.number_input("Inventarisnummer t/m", min_value=0, value=st.session_state.get("tot_multi", 1), key="tot_multi")
                a_nummers_input = st.text_input("A-nummers (optioneel, bijv. 68A, 99B)", value=st.session_state.get("a_multi", ""), key="a_multi")

                toevoegen = st.form_submit_button("➕ Voeg toe aan lijst")

                if toevoegen and toegang_nmr:
                    nummers = [str(n) for n in range(int(van), int(tot) + 1)]
                    if a_nummers_input.strip():
                        extra = [x.strip() for x in a_nummers_input.split(',') if x.strip()]
                        nummers.extend(extra)

                    st.session_state["omslagen"].append({
                        "toegangsnummer": toegang_nmr,
                        "titel": toegang_titel,
                        "nummers": nummers
                    })

            if st.session_state["omslagen"]:
                st.write("### Toegevoegde etiketten")
                for i, omslag in enumerate(st.session_state["omslagen"]):
                    st.write(f"{i+1}. {omslag['toegangsnummer']} - {omslag['titel']}: {', '.join(omslag['nummers'])}")

                if st.button("🎫 Genereer alle etiketten (.docx)"):
                    labels = []
                    for omslag in st.session_state["omslagen"]:
                        for num in omslag["nummers"]:
                            labels.append([
                                "Stadsarchief Amsterdam",
                                omslag["toegangsnummer"],
                                omslag["titel"],
                                str(num)
                            ])
                    docx_file = create_docx_table(labels)
                    st.download_button("⬇️ Download als DOCX", docx_file, file_name="omslagetiketten_meerdere.docx")

                if st.button("🗑️ Verwijder alles"):
                    st.session_state["omslagen"] = []

    elif option == "📦 Doosetiketten maken":
        multi_ui = st.toggle("Meerdere toegangen toevoegen?", value=False)

        if not multi_ui:
            if 'vorige_toegang' not in st.session_state:
                st.session_state['vorige_toegang'] = ''
            if 'invoer_veld' not in st.session_state:
                st.session_state['invoer_veld'] = ''

            ToegangsNmr = st.text_input("Toegangsnummer", "")
            if ToegangsNmr != st.session_state['vorige_toegang']:
                st.session_state['invoer_veld'] = ''

            naam_default = TOEGANGSTITELS.get(ToegangsNmr.strip(), "")
            naam = st.text_input("Archiefnaam", naam_default)

            invoer = st.text_area(
                "Inventarisnummers (Scheid etiketten met een nieuwe regel, gebruik ',' voor losse nummers en '-' voor reeksen)",
                value=st.session_state.get('invoer_veld', ""),
                key="invoer_veld"
            )

            if st.button("📦 Genereer doosetiketten (.docx)"):
                groups = [grp.strip() for grp in invoer.strip().split("\n") if grp.strip()]
                if not groups:
                    st.warning("⚠️ Vul minstens één inventarisgroep in.")
                else:
                    docx_file = generate_box_labels(naam, ToegangsNmr, groups)
                    st.download_button("⬇️ Download als DOCX", docx_file, file_name="doosetiketten" + ToegangsNmr + ".docx")

            st.session_state['vorige_toegang'] = ToegangsNmr

        else:
            st.subheader("Meerdere toegangen invoeren")

            if 'toegangen' not in st.session_state:
                st.session_state['toegangen'] = []

            # Zet standaardwaarden in session_state
            if "prev_toegang_nmr" not in st.session_state:
                st.session_state["prev_toegang_nmr"] = ""

            # Laat gebruiker een toegangsnr invullen (maar maar één keer!)
            toegang_nmr = st.text_input("Toegangsnummer", "")

            # Controleer of toegang_nmr is gewijzigd, en reset zo nodig invoerveld
            if toegang_nmr != st.session_state["prev_toegang_nmr"]:
                st.session_state["invoer_multi"] = ""  # leeg maken
                st.session_state["prev_toegang_nmr"] = toegang_nmr

            # Toon toegangstitel op basis van huidig toegang_nmr
            toegang_titel_default = TOEGANGSTITELS.get(toegang_nmr.strip(), "")

            with st.form(key="doosetiket_form"):
                toegang_titel = st.text_input("Archiefnaam", toegang_titel_default)

                invoer = st.text_area(
                    "Inventarisnummers (Scheid etiketten met een nieuwe regel, gebruik ',' voor losse nummers en '-' voor reeksen)",
                    value=st.session_state.get("invoer_multi", ""),
                    key="invoer_multi"
                )

                toevoegen = st.form_submit_button("➕ Voeg toe aan lijst")

                if toevoegen and toegang_nmr and invoer:
                    groups = [grp.strip() for grp in invoer.strip().split("\n") if grp.strip()]
                    st.session_state["toegangen"].append((toegang_titel, toegang_nmr, groups))
                    del st.session_state["invoer_multi"]  # verwijder sleutel i.p.v. leegmaken
                    st.rerun()


            if st.session_state['toegangen']:
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("🗑️ Lijst leegmaken"):
                        st.session_state['toegangen'] = []
                        st.rerun()
                with col2:
                    if st.button("📦 Genereer gecombineerde doosetiketten (.docx)"):
                        alle_labels = []
                        for titel, nummer, groepen in st.session_state['toegangen']:
                            for groep in groepen:
                                label = [
                                    "Stadsarchief Amsterdam",
                                    f"{nummer}".strip(),
                                    f"{titel}".strip(),
                                    f"{groep.strip()}"
                                ]
                                alle_labels.append(label)

                        docx_file = create_docx_table(alle_labels)
                        st.download_button("⬇️ Download gecombineerde DOCX", docx_file, file_name="doosetiketten_gecombineerd.docx")

                st.markdown("### Toegevoegde etiketten")
                for i, (titel, nummer, groepen) in enumerate(st.session_state['toegangen']):
                    st.markdown(f"**{nummer} - {titel}**")
                    st.markdown("<ul>" + "".join([f"<li>{g}</li>" for g in groepen]) + "</ul>", unsafe_allow_html=True)
       
    st.markdown("##### Made by Yorick de Man" + " V0.5")
    
if __name__ == "__main__":
    main()
