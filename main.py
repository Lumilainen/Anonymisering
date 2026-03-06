import re
from docx import Document

FULL_NAME_REGEX = r"\b[A-ZÅÄÖ][a-zåäö\-]+ [A-ZÅÄÖ][a-zåäö\-]+\b"
INITIAL_NAME_REGEX = r"\b[A-Z](?:-[A-Z])?\.?\s?[A-ZÅÄÖ][a-zåäö\-]+\b"
EMAIL_REGEX = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
PERSONNUMMER_REGEX = r"\b(19|20)?\d{6}[- ]?\d{4}\b"


def detect_persons(text):

    persons = set()

    persons.update(re.findall(FULL_NAME_REGEX, text))
    persons.update(re.findall(INITIAL_NAME_REGEX, text))

    return persons


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

        for paragraph in section.header.paragraphs:
            persons.update(detect_persons(paragraph.text))

        for paragraph in section.footer.paragraphs:
            persons.update(detect_persons(paragraph.text))

    return persons


def anonymize_text(text, persons):

    text = re.sub(PERSONNUMMER_REGEX, "[PERSONNUMMER]", text)
    text = re.sub(EMAIL_REGEX, "[EMAIL]", text)

    for name in sorted(persons, key=len, reverse=True):

        escaped = re.escape(name)

        text = re.sub(rf"\b{escaped}\b", "[PERSON]", text)

    return text


def anonymize_paragraph(paragraph, persons):

    original = paragraph.text

    if not original:
        return

    anonymized = anonymize_text(original, persons)

    if anonymized != original:

        # säker rensning av runs
        for run in paragraph.runs:
            run.text = ""

        if paragraph.runs:
            paragraph.runs[0].text = anonymized
        else:
            paragraph.add_run(anonymized)


def process_tables(doc, persons):

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for paragraph in cell.paragraphs:

                    anonymize_paragraph(paragraph, persons)


def process_headers_footers(doc, persons):

    for section in doc.sections:

        for paragraph in section.header.paragraphs:

            anonymize_paragraph(paragraph, persons)

        for paragraph in section.footer.paragraphs:

            anonymize_paragraph(paragraph, persons)


def anonymize_docx(input_stream, output_stream, persons):

    doc = Document(input_stream)

    for paragraph in doc.paragraphs:

        anonymize_paragraph(paragraph, persons)

    process_tables(doc, persons)

    process_headers_footers(doc, persons)

    doc.save(output_stream)