import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO

# Funksjon for å hente tags fra Excel
def get_tags_from_excel(file, sheet_name="Sheet1", column_name="Tags"):
    df = pd.read_excel(file, sheet_name=sheet_name)
    tags = df[column_name].dropna().tolist()  # Fjern tomme rader og konverter til liste
    return tags

# Funksjon for å markere tags i PDF
def mark_text_with_red_ring(input_pdf, tags):
    doc = fitz.open(stream=input_pdf, filetype="pdf")  # Åpne PDF fra bytes
    
    for page_num in range(doc.page_count):
        page = doc[page_num]
        for tag in tags:
            text_instances = page.search_for(tag)
            for inst in text_instances:
                square = page.add_rect_annot(inst)
                square.set_colors(stroke=(1, 0, 0))  # Rød farge (RGB: 1, 0, 0)
                square.update()
    
    # Lagre den redigerte PDF-en i minnet
    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf

# Streamlit app-grensesnitt
st.title("PDF Markeringsapp")

# Last opp PDF og Excel
pdf_file = st.file_uploader("Last opp PDF-fil", type="pdf")
excel_file = st.file_uploader("Last opp Excel-fil med tags", type="xlsx")

if pdf_file and excel_file:
    # Hent tags fra Excel
    tags = get_tags_from_excel(excel_file)
    st.write("Tags hentet fra Excel:", tags)

    # Marker tags i PDF
    if st.button("Marker tags i PDF"):
        output_pdf = mark_text_with_red_ring(pdf_file.read(), tags)
        st.success("Tags markert i PDF!")

        # Last ned den redigerte PDF-en
        st.download_button(
            label="Last ned PDF med markerte tags",
            data=output_pdf,
            file_name="output_red_ring.pdf",
            mime="application/pdf"
        )
