import re
from docx import Document

# =========================
# REGEX
# =========================

FULL_NAME_REGEX = r"\b[A-ZÅÄÖ][a-zåäö]+ [A-ZÅÄÖ][a-zåäö]+\b"
INITIAL_NAME_REGEX = r"\b[A-Z](?:-[A-Z])?\.?\s?[A-ZÅÄÖ][a-zåäö]+\b"
HYPHEN_NAME_REGEX = r"\b[A-ZÅÄÖ][a-zåäö]+-[A-ZÅÄÖ][a-zåäö]+\b"

LAST_FIRST_REGEX = r"\b[A-ZÅÄÖ][a-zåäö]+,\s?[A-ZÅÄÖ][a-zåäö]+\b"
INITIAL_DOT_REGEX = r"\b[A-Z]\.\s?[A-ZÅÄÖ][a-zåäö]+\b"
LAST_INITIAL_REGEX = r"\b[A-ZÅÄÖ][a-zåäö]+\s[A-Z]\b"

EMAIL_REGEX = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
PERSONNUMMER_REGEX = r"\b(19|20)?\d{6}[- ]?\d{4}\b"


# =========================
# NAMNDETEKTION
# =========================

def detect_persons(text):

    persons = set()

    persons.update(re.findall(FULL_NAME_REGEX, text))
    persons.update(re.findall(INITIAL_NAME_REGEX, text))
    persons.update(re.findall(HYPHEN_NAME_REGEX, text))
    persons.update(re.findall(LAST_FIRST_REGEX, text))
    persons.update(re.findall(INITIAL_DOT_REGEX, text))
    persons.update(re.findall(LAST_INITIAL_REGEX, text))

    return persons


# =========================
# TEXTANONYMISERING
# =========================

def anonymize_text(text, persons):

    text = re.sub(PERSONNUMMER_REGEX, "[PERSONNUMMER]", text)
    text = re.sub(EMAIL_REGEX, "[EMAIL]", text)

    for name in sorted(persons, key=len, reverse=True):

        escaped = re.escape(name)

        text = re.sub(rf"\b{escaped}\b", "[PERSON]", text)

        parts = re.split(r"[ ,]", name)

        if parts:

            first = parts[0]

            text = re.sub(rf"\b{first}\b", "[PERSON]", text)

    return text


# =========================
# PARAGRAFER
# =========================

def anonymize_paragraph(paragraph, persons):

    original = paragraph.text

    if not original:
        return

    anonymized = anonymize_text(original, persons)

    if anonymized != original:

        for run in paragraph.runs:
            run.text = ""

        if paragraph.runs:
            paragraph.runs[0].text = anonymized
        else:
            paragraph.add_run(anonymized)


# =========================
# TABELLER
# =========================

def process_tables(doc, persons):

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for paragraph in cell.paragraphs:

                    anonymize_paragraph(paragraph, persons)


# =========================
# HEADER / FOOTER
# =========================

def process_headers_footers(doc, persons):

    for section in doc.sections:

        header_list = [
            section.header,
            section.first_page_header,
            section.even_page_header
        ]

        footer_list = [
            section.footer,
            section.first_page_footer,
            section.even_page_footer
        ]

        # HEADER
        for header in header_list:

            if header:

                for paragraph in header.paragraphs:

                    anonymize_paragraph(paragraph, persons)

                for table in header.tables:

                    for row in table.rows:

                        for cell in row.cells:

                            for paragraph in cell.paragraphs:

                                anonymize_paragraph(paragraph, persons)

        # FOOTER
        for footer in footer_list:

            if footer:

                for paragraph in footer.paragraphs:

                    anonymize_paragraph(paragraph, persons)

                for table in footer.tables:

                    for row in table.rows:

                        for cell in row.cells:

                            for paragraph in cell.paragraphs:

                                anonymize_paragraph(paragraph, persons)


# =========================
# RENSNING
# =========================

def remove_comments(doc):

    try:

        comments_part = doc.part._comments_part

        if comments_part:

            comments_part._element.clear()

    except:
        pass


def remove_track_changes(doc):

    try:

        body = doc.part._element.body

        for element in body.xpath(".//w:ins | .//w:del"):

            element.getparent().remove(element)

    except:
        pass


def clean_metadata(doc):

    props = doc.core_properties

    props.author = ""
    props.last_modified_by = ""
    props.title = ""
    props.subject = ""
    props.comments = ""


# =========================
# SCANNA NAMN
# =========================

def scan_document_for_persons(doc):

    persons = set()

    for paragraph in doc.paragraphs:
        persons.update(detect_persons(paragraph.text))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    persons.update(detect_persons(paragraph.text))

    for section in doc.sections:

        header_list = [
            section.header,
            section.first_page_header,
            section.even_page_header
        ]

        footer_list = [
            section.footer,
            section.first_page_footer,
            section.even_page_footer
        ]

        for header in header_list:

            if header:

                for paragraph in header.paragraphs:
                    persons.update(detect_persons(paragraph.text))

                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                persons.update(detect_persons(paragraph.text))

        for footer in footer_list:

            if footer:

                for paragraph in footer.paragraphs:
                    persons.update(detect_persons(paragraph.text))

                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                persons.update(detect_persons(paragraph.text))

    return persons


# =========================
# HUVUDFUNKTION
# =========================

def anonymize_docx(input_stream, output_stream, persons):

    doc = Document(input_stream)

    clean_metadata(doc)

    remove_comments(doc)

    remove_track_changes(doc)

    for paragraph in doc.paragraphs:

        anonymize_paragraph(paragraph, persons)

    process_tables(doc, persons)

    process_headers_footers(doc, persons)

    doc.save(output_stream)