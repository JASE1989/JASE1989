import streamlit as st
import re
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO
import numpy as np

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
    tags_found = set()
    tags_not_found = tags.copy()

    # Regex-mønster for nøyaktighetsnivå
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{2}-L-\d{4})', re.DOTALL)
    else:
        pattern = re.compile(r'(\d{4})', re.DOTALL)

    for page_num in range(doc.page_count):
        page = doc[page_num]
        page_text = page.get_text("text")

        # Søk med regex og legg funn til i tags_found
        regex_matches = set()
        for match in pattern.findall(page_text):
            regex_matches.add(match)

        # Kombiner regex-funn og tags for markering
        for tag in tags + list(regex_matches):
            text_instances = page.search_for(tag)
            if text_instances:
                tags_found.add(tag)
                if tag in tags_not_found:
                    tags_not_found.remove(tag)

                # Marker tekst i PDF
                for inst in text_instances:
                    rect = adjust_rectangle(inst, rect_adjustment)
                    annotation = page.add_rect_annot(rect)
                    annotation.set_colors(stroke=(1, 0, 0))  # Rød markering
                    annotation.update()

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
start_button = st.button("Start søket")

if pdf_files and excel_file and start_button:
    try:
        tags = get_tags_from_excel(excel_file)
        if not tags:
            st.error("Excel-filen inneholder ingen tagger.")
        else:
            st.write("Tags:", ", ".join(tags))
            merged_pdf = fitz.open()
            for pdf_file in pdf_files:
                merged_pdf.insert_pdf(fitz.open(stream=pdf_file.read(), filetype="pdf"))

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
