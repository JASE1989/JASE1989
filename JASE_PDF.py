def mark_text_with_pymupdf(input_pdf, tags, match_strictness, rect_adjustment=2):
    doc = input_pdf
    tags_found = set()  # Bruk et sett for å unngå duplikater
    tags_not_found = tags.copy()  # Initialisering av tags_not_found

    # Regex-mønster for nøyaktighetsnivå
    if match_strictness == "Streng":
        pattern = re.compile(r'\d{4}J-PL-(\d+-l-\d+)-TC02-00', re.DOTALL)
    elif match_strictness == "Moderat":
        pattern = re.compile(r'(\d{2}-L-\d{4})', re.DOTALL)
    elif match_strictness == "Tolerant":
        # For tolerant søk, let etter de siste 4 sifrene (YYYY)
        pattern = re.compile(r'(\d{4})(?!\d)', re.DOTALL)

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
