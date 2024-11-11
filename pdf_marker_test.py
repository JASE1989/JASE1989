import streamlit as st
import re
import fitz  # PyMuPDF
import easyocr
import pandas as pd
from io import BytesIO
import streamlit as st
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

# Funksjon for OCR og å markere tags i PDF med forbedringer
def mark_text_with_easyocr(input_pdf, tags, match_strictness):
    doc = fitz.open(stream=input_pdf, filetype="pdf")
    reader = easyocr.Reader(['no', 'en'])
    tags_found = []
    tags_not_found = tags.copy()

    # Regex-mønster for nøyaktighetsnivå
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{4}J-PL-\d+-L-\d+)-TC02-00', re.DOTALL)
    else:  # Tolerant
        pattern = re.compile(r'(\d{4}J-PL-\d+-L-(\d+)-TC02-00)', re.DOTALL)

    for page_num in range(doc.page_count):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        
        # Fjern skjevhet (deskew) før OCR
        img = remove_skew(img)

        # OCR på originalbilde
        ocr_result = reader.readtext(img, detail=1, paragraph=False)

        # OCR på rotert bilde for vertikal tekst
        rotated_img = rotate_vertical_text(img)
        ocr_result_vertical = reader.readtext(rotated_img, detail=1, paragraph=False)

        # Kombiner OCR-resultater fra både horisontalt og vertikalt tekst
        all_text = ' '.join([text for _, text, conf in ocr_result if conf > 0.4] +
                            [text for _, text, conf in ocr_result_vertical if conf > 0.4]).replace("\n", " ")

        # Søk etter tagger i kombinert OCR-resultat
        for match in pattern.findall(all_text):
            clean_text = ''.join(e for e in match.lower() if e.isalnum())
            for tag in tags:
                clean_tag = ''.join(e for e in tag.lower() if e.isalnum())
                if clean_tag in clean_text:
                    # Marker teksten med rødt rektangel i PDF-en
                    for bbox, text, conf in ocr_result + ocr_result_vertical:
                        if clean_tag in text.replace(" ", "").lower():
                            rect = fitz.Rect(bbox[0][0], bbox[0][1], bbox[2][0], bbox[2][1])
                            square = page.add_rect_annot(rect)
                            square.set_colors(stroke=(1, 0, 0))
                            square.update()

                    if tag not in tags_found:
                        tags_found.append(tag)
                    if tag in tags_not_found:
                        tags_not_found.remove(tag)

    if tags_not_found:
        st.write(f"Følgende tags ble ikke funnet i PDF-en: {', '.join(tags_not_found)}")
    else:
        st.write("Alle tags ble funnet i PDF-en.")

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf



# Funksjon for å markere tekst med PyMuPDF (uten OCR)
def mark_text_with_pymupdf(input_pdf, tags, match_strictness):
    doc = fitz.open(stream=input_pdf, filetype="pdf")
    marked = False
    tags_found = []
    tags_not_found = tags.copy()

    # Regex for PyMuPDF basert på match_strictness
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{4}J-PL-\d+-L-\d+)-TC02-00', re.DOTALL)
    else:
        pattern = re.compile(r'(\d{4}J-PL-\d+-L-(\d+)-TC02-00)', re.DOTALL)

    for page_num in range(doc.page_count):
        page = doc[page_num]

        for tag in tags:
            text_instances = page.search_for(tag)
            for inst in text_instances:
                rect = fitz.Rect(inst)
                square = page.add_rect_annot(rect)
                square.set_colors(stroke=(1, 0, 0))
                square.update()

                if tag not in tags_found:
                    tags_found.append(tag)
                if tag in tags_not_found:
                    tags_not_found.remove(tag)

                marked = True

    if not marked:
        st.write("Ingen tagger ble markert på denne PDF-en.")

    if tags_not_found:
        st.write(f"Følgende tags ble ikke funnet i PDF-en: {', '.join(tags_not_found)}")
    else:
        st.write("Alle tags ble funnet i PDF-en.")

    if tags_not_found:
        report_page = doc.new_page()
        report_page.insert_text((50, 50), "Rapport over manglende tags:", fontsize=12)

        missing_tags_text = f"Antall manglende tags: {len(tags_not_found)}\n\n"
        missing_tags_text += "\n".join(tags_not_found)
        report_page.insert_text((50, 100), missing_tags_text, fontsize=10)

    output_pdf = BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf

# Streamlit app-grensesnitt
st.title("PDF Markeringsapp med OCR eller PyMuPDF")

pdf_file = st.file_uploader("Last opp PDF-fil", type="pdf")
excel_file = st.file_uploader("Last opp Excel-fil med tags", type="xlsx")

# Velg metode (OCR eller PyMuPDF)
method = st.radio("Velg metode for tekstutvinning", ('OCR (EasyOCR)', 'PyMuPDF'))

# Velg nøyaktighetsnivå
match_strictness = st.selectbox("Velg nøyaktighetsnivå", ("Streng", "Moderat", "Tolerant"))

if pdf_file and excel_file:
    tags = get_tags_from_excel(excel_file)
    st.write("Tags hentet fra Excel:", tags)

    if st.button("Marker tags i PDF"):
        if method == 'OCR (EasyOCR)':
            output_pdf = mark_text_with_easyocr(pdf_file.read(), tags, match_strictness)
            st.success("Tags markert i PDF med OCR!")
        else:
            output_pdf = mark_text_with_pymupdf(pdf_file.read(), tags, match_strictness)
            st.success("Tags markert i PDF med PyMuPDF!")

        st.download_button(
            label="Last ned redigert PDF",
            data=output_pdf,
            file_name="marked_pdf.pdf",
            mime="application/pdf"
        )
