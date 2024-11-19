import streamlit as st
import re
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO

# Funksjon for å hente tags fra Excel
def get_tags_from_excel(excel_file, column_name='Tag'):
    try:
        df = pd.read_excel(excel_file)
        if column_name in df.columns:
            tags = df[column_name].dropna().tolist()
            return tags
        else:
            raise KeyError(f"Kolonnen '{column_name}' ble ikke funnet i Excel-filen.")
    except Exception as e:
        raise ValueError(f"Feil ved lesing av Excel: {e}")

# Funksjon for å justere markeringsrute størrelse
def adjust_rectangle(rect, adjustment=6):
    x0, y0, x1, y1 = rect
    x0 -= adjustment
    y0 -= adjustment
    x1 += adjustment
    y1 += adjustment
    return fitz.Rect(x0, y0, x1, y1)

# Funksjon for PyMuPDF
def mark_text_with_pymupdf(input_pdf, tags, match_strictness, rect_adjustment=2):
    doc = input_pdf
    tags_found = set()  # Bruk et sett for å unngå duplikater
    tags_not_found = tags.copy()  # Initialisering av tags_not_found

    # Regex-mønster for nøyaktighetsnivå
    if match_strictness == "Streng":
        # For strenge søk, let etter spesifikke mønstre som ser ut som:
        # 4 sifre + J-PL- + en spesifikk struktur, f.eks. 1234J-PL-001-l-001-TC02-00
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        # For moderate søk, let etter noe som har et tall etterfulgt av L og en kode
        pattern = re.compile(r'(\d{2}-L-\d{4})', re.DOTALL)
    else:
        # For tolerant søk, let etter bare 4 sifre
        pattern = re.compile(r'(\d{4})', re.DOTALL)

    for page_num in range(doc.page_count):
        page = doc[page_num]

        # Bruk regex på sidens tekst for å finne de spesifikke taggene
        page_text = page.get_text("text")  # Hent all tekst fra siden
        for tag in tags:
            matches = pattern.findall(page_text)  # Finn matchende tags med regex
            if matches:
                for match in matches:
                    if match not in tags_found:
                        tags_found.add(match)

        # Søk etter tekstforekomster med "search_for" og marker dem
        for tag in tags:
            text_instances = page.search_for(tag)
            for inst in text_instances:
                rect = adjust_rectangle(inst, rect_adjustment)
                annotation = page.add_rect_annot(rect)
                annotation.set_colors(stroke=(1, 0, 0))  # Rød farge for markeringen
                annotation.update()
                if tag not in tags_found:
                    tags_found.add(tag)

    # Lag en rapport over tagger som ikke ble funnet
    if tags_not_found:
        report_page = doc.new_page()
        report_text = f"Tags som ikke ble funnet ({len(tags_not_found)} tags):\n"
        for tag in tags_not_found:
            report_text += f"{tag}\n"
        report_page.insert_text((50, 50), report_text, fontsize=12)

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf, list(tags_found)

# Streamlit app-grensesnitt
st.title("PDF Markeringsapp")

pdf_files = st.file_uploader("Last opp PDF-filer", type="pdf", accept_multiple_files=True)
excel_file = st.file_uploader("Last opp Excel-fil med tags", type="xlsx")
match_strictness = st.selectbox("Velg nøyaktighetsnivå", ("Streng", "Moderat", "Tolerant"))
method = st.radio("Velg metode for markering", ("PyMuPDF"))
start_button = st.button("Start søket")

if pdf_files and excel_file and start_button:
    try:
        # Hent tagger fra Excel
        tags = get_tags_from_excel(excel_file)
        
        # Valider om tagger finnes
        if not tags:
            st.error("Excel-filen inneholder ingen tagger. Vennligst sjekk filen.")
        else:
            st.write("Tags:", ", ".join(tags))

            merged_pdf = fitz.open()
            for pdf_file in pdf_files:
                merged_pdf.insert_pdf(fitz.open(stream=pdf_file.read(), filetype="pdf"))

            # Velg prosessering basert på valgt metode
            if method == "PyMuPDF":
                result_pdf, found_tags = mark_text_with_pymupdf(merged_pdf, tags, match_strictness)

            st.write("Funnet tags:", ", ".join(found_tags))
            st.download_button(
                label="Last ned merket PDF",
                data=result_pdf,
                file_name="marked_tags.pdf",
                mime="application/pdf",
            )
    except ValueError as ve:
        st.error(f"Feil ved lesing av Excel: {ve}")
    except Exception as e:
        st.error(f"En uventet feil oppstod: {e}")
