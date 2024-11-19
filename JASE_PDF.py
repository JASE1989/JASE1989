import streamlit as st
import re
import fitz  # PyMuPDF
import easyocr
import pandas as pd
from io import BytesIO
import cv2
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

# Funksjon for å fjerne skjevhet i bilder før OCR
def remove_skew(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.bitwise_not(gray)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    
    coords = np.column_stack(np.where(thresh > 0))
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle

    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    
    return rotated

# Funksjon for å rotere vertikal tekst for OCR
def rotate_vertical_text(image):
    return cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)

# Funksjon for å justere markeringsrute størrelse
def adjust_rectangle(rect, adjustment=4):
    x0, y0, x1, y1 = rect
    x0 -= adjustment
    y0 -= adjustment
    x1 += adjustment
    y1 += adjustment
    return fitz.Rect(x0, y0, x1, y1)

# Cache EasyOCR-leser
@st.cache_resource
def load_easyocr_reader():
    return easyocr.Reader(['no', 'en'])

# Funksjon for OCR og markering
def mark_text_with_easyocr(input_pdf, tags, match_strictness, rect_adjustment=2):
    reader = load_easyocr_reader()
    doc = input_pdf
    tags_found = []
    tags_not_found = tags.copy()  # Initialisering av tags_not_found

    # Regex-mønster for nøyaktighetsnivå
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{2}-L-\d{4})', re.DOTALL)
    else:
        pattern = re.compile(r'(\d{4})', re.DOTALL)

    for page_num in range(doc.page_count):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        
        # Fjern skjevhet før OCR
        img = remove_skew(img)

        # OCR på originalbilde og rotert bilde
        ocr_result = reader.readtext(img, detail=1, paragraph=False)
        rotated_img = rotate_vertical_text(img)
        ocr_result_vertical = reader.readtext(rotated_img, detail=1, paragraph=False)

        # Kombiner resultater og søk etter tagger
        all_text = ' '.join([text for _, text, conf in ocr_result if conf > 0.4] +
                            [text for _, text, conf in ocr_result_vertical if conf > 0.4]).replace("\n", " ")

        for tag in tags:
            if re.search(tag, all_text, re.IGNORECASE):
                tags_found.append(tag)
                if tag in tags_not_found:
                    tags_not_found.remove(tag)

    # Legg til en ny side med rapport om tagger som ikke ble funnet
    if tags_not_found:
        report_page = doc.new_page()
        report_text = f"Tags som ikke ble funnet:\n{', '.join(tags_not_found)}"
        report_page.insert_text((50, 50), report_text, fontsize=12)

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf, tags_found

## Funksjon for PyMuPDF
def mark_text_with_pymupdf(input_pdf, tags, match_strictness, rect_adjustment=2):
    doc = input_pdf
    tags_found = []
    tags_not_found = tags.copy()  # Initialisering av tags_not_found

    for page_num in range(doc.page_count):
        page = doc[page_num]

        for tag in tags:
            text_instances = page.search_for(tag)
            for inst in text_instances:
                rect = adjust_rectangle(inst, rect_adjustment)
                annotation = page.add_rect_annot(rect)
                annotation.set_colors(stroke=(1, 0, 0))
                annotation.update()
                if tag not in tags_found:
                    tags_found.append(tag)

    # Legg til en ny side med rapport om tagger som ikke ble funnet
    if tags_not_found:
        report_page = doc.new_page()
        report_text = f"Tags som ikke ble funnet:\n{', '.join(tags_not_found)}"
        report_page.insert_text((50, 50), report_text, fontsize=12)

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf, tags_found

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
            if method == "OCR":
                result_pdf, found_tags = mark_text_with_easyocr(merged_pdf, tags, match_strictness)
            else:
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
