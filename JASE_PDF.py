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
    df = pd.read_excel(excel_file)
    if column_name in df.columns:
        tags = df[column_name].dropna().tolist()
        return tags
    else:
        raise KeyError(f"Kolonnen '{column_name}' ble ikke funnet i Excel-filen.")

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

# Funksjon for å rotere PDF-side for vertikal tekst
def rotate_page_for_vertical_text(page):
    matrix = fitz.Matrix(0, -1, 1, 0)  # 90-graders rotasjon
    page.set_rotation(90)  # Oppdaterer siden internt
    return page

# Funksjon for å justere markeringsrute størrelse
def adjust_rectangle(rect, adjustment=4):
    x0, y0, x1, y1 = rect
    x0 -= adjustment
    y0 -= adjustment
    x1 += adjustment
    y1 += adjustment
    return fitz.Rect(x0, y0, x1, y1)

# Funksjon for OCR og å markere tags i PDF
def mark_text_with_easyocr(input_pdf, tags, match_strictness, rect_adjustment=2):
    doc = input_pdf
    reader = easyocr.Reader(['no', 'en'])
    tags_found = []
    tags_not_found = tags.copy()

    # Regex-mønster for nøyaktighetsnivå
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{2}-L-\d{4})', re.DOTALL)
    else:
        pattern = re.compile(r'(\d{4})', re.DOTALL)

    marked_tags = set()
    all_found_tags = set()

    for page_num in range(doc.page_count):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        
        # Fjern skjevhet før OCR
        img = remove_skew(img)

        # OCR på originalbilde
        ocr_result = reader.readtext(img, detail=1, paragraph=False)

        # OCR på rotert bilde for vertikal tekst
        rotated_img = rotate_vertical_text(img)
        ocr_result_vertical = reader.readtext(rotated_img, detail=1, paragraph=False)

        # Kombiner OCR-resultater
        all_text = ' '.join([text for _, text, conf in ocr_result if conf > 0.4] +
                            [text for _, text, conf in ocr_result_vertical if conf > 0.4]).replace("\n", " ")

        # Søk etter tagger i kombinert OCR-resultat
        for match in pattern.findall(all_text):
            clean_text = ''.join(e for e in match.lower() if e.isalnum())
            for tag in tags:
                clean_tag = ''.join(e for e in tag.lower() if e.isalnum())
                if clean_tag in clean_text and tag not in marked_tags:
                    for bbox, text, conf in ocr_result + ocr_result_vertical:
                        if clean_tag in text.replace(" ", "").lower():
                            rect = fitz.Rect(bbox[0][0], bbox[0][1], bbox[2][0], bbox[2][1])
                            rect = adjust_rectangle(rect, rect_adjustment)
                            square = page.add_rect_annot(rect)
                            if square:
                                square.set_colors(stroke=(1, 0, 0))
                                square.update()

                    marked_tags.add(tag)
                    all_found_tags.add(tag)
                    if tag not in tags_found:
                        tags_found.append(tag)
                    if tag in tags_not_found:
                        tags_not_found.remove(tag)

    # Legg til rapport om manglende tags
    add_missing_tags_report(doc, tags_not_found)

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf, all_found_tags

# Funksjon for PyMuPDF og å markere tags i PDF
def mark_text_with_pymupdf(input_pdf, tags, match_strictness, rect_adjustment=2):
    doc = input_pdf
    tags_found = []
    tags_not_found = tags.copy()

    # Regex for PyMuPDF basert på match_strictness
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{2}-L-\d{4})', re.DOTALL)
    else:
        pattern = re.compile(r'(\d{4})', re.DOTALL)

    marked_tags = set()
    all_found_tags = set()

    for page_num in range(doc.page_count):
        page = doc[page_num]

        # Rotér side for vertikal tekst (valgfritt)
        rotated_page = rotate_page_for_vertical_text(page)

        for tag in tags:
            if tag not in marked_tags:
                text_instances = rotated_page.search_for(tag)
                for inst in text_instances:
                    rect = adjust_rectangle(inst, rect_adjustment)
                    annotation = rotated_page.add_rect_annot(rect)
                    if annotation:
                        annotation.set_colors(stroke=(1, 0, 0))
                        annotation.update()
                        marked_tags.add(tag)
                        all_found_tags.add(tag)
                    if tag not in tags_found:
                        tags_found.append(tag)
                    if tag in tags_not_found:
                        tags_not_found.remove(tag)

    # Legg til rapport om manglende tags
    add_missing_tags_report(doc, tags_not_found)

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf, all_found_tags

# Funksjon for å legge til rapport om manglende tags i PDF
def add_missing_tags_report(doc, tags_not_found):
    if tags_not_found:
        report_page = doc.new_page()
        report_page.insert_text((50, 50), "Rapport over manglende tags:", fontsize=12, color=(1, 0, 0))
        missing_tags_text = f"Antall manglende tags: {len(tags_not_found)}\n\n"
        missing_tags_text += "\n".join(tags_not_found)
        report_page.insert_text((50, 100), missing_tags_text, fontsize=10, color=(0, 0, 0))

# Streamlit app-grensesnitt
st.title("PDF Markeringsapp med OCR eller PyMuPDF")

pdf_files = st.file_uploader("Last opp PDF-filer", type="pdf", accept_multiple_files=True)
excel_file = st.file_uploader("Last opp Excel-fil med tags", type="xlsx")
match_strictness = st.selectbox("Velg nøyaktighetsnivå", ("Streng", "Moderat", "Tolerant"))

# Legg til valg for metode
method = st.radio("Velg metode for markering", ("OCR", "PyMuPDF"))

start_button = st.button("Start søket")

if pdf_files and excel_file and start_button:
    try:
        # Hent tags fra Excel
        tags = get_tags_from_excel(excel_file)

        # Vis tags fra Excel-filen
        st.write("### Tags i Excel-filen:")
        st.write(", ".join(tags))

        # Kombiner PDF-filer
        merged_pdf = fitz.open()  # Start med en tom PDF
        for pdf_file in pdf_files:
            merged_pdf.insert_pdf(fitz.open(stream=pdf_file.read(), filetype="pdf"))

        if method == "OCR":
            result_pdf, found_tags = mark_text_with_easyocr(merged_pdf, tags, match_strictness)
        else:
            result_pdf, found_tags = mark_text_with_pymupdf(merged_pdf, tags, match_strictness)

        st.write("### Funnet tags:")
        st.write(", ".join(found_tags))

        st.download_button(
            label="Last ned merket PDF",
            data=result_pdf,
            file_name="marked_tags.pdf",
            mime="application/pdf",
        )

    except Exception as e:
        st.error(f"En feil oppstod: {e}")
