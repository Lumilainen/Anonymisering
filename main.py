import re
import spacy
from docx import Document
from io import BytesIO

# säker laddning av spaCy-modell
MODEL = "sv_core_news_sm"

try:
    nlp = spacy.load(MODEL)
except OSError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "spacy", "download", MODEL])
    nlp = spacy.load(MODEL)


PERSON_REGEX = r"\b[A-ZÅÄÖ][a-zåäö]+ [A-ZÅÄÖ][a-zåäö]+\b"


def detect_persons(text):

    persons = set()

    doc = nlp(text)

    for ent in doc.ents:
        if ent.label_ == "PER":
            persons.add(ent.text)

    regex_matches = re.findall(PERSON_REGEX, text)

    for match in regex_matches:
        persons.add(match)

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

        header = section.header
        footer = section.footer

        for paragraph in header.paragraphs:
            persons.update(detect_persons(paragraph.text))

        for paragraph in footer.paragraphs:
            persons.update(detect_persons(paragraph.text))

    return persons


def anonymize_text(text, persons):

    for name in sorted(persons, key=len, reverse=True):

        escaped = re.escape(name)

        text = re.sub(rf"\b{escaped}\b", "[PERSON]", text)

    return text


def anonymize_paragraph(paragraph, persons):

    text = paragraph.text

    if not text:
        return

    anonymized = anonymize_text(text, persons)

    if anonymized != text:

        for run in paragraph.runs[::-1]:
            paragraph._element.remove(run._element)

        paragraph.add_run(anonymized)


def process_tables(doc, persons):

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for paragraph in cell.paragraphs:

                    anonymize_paragraph(paragraph, persons)


def process_headers_footers(doc, persons):

    for section in doc.sections:

        header = section.header
        footer = section.footer

        for paragraph in header.paragraphs:
            anonymize_paragraph(paragraph, persons)

        for paragraph in footer.paragraphs:
            anonymize_paragraph(paragraph, persons)


def anonymize_docx(input_stream, output_stream, persons):

    doc = Document(input_stream)

    for paragraph in doc.paragraphs:
        anonymize_paragraph(paragraph, persons)

    process_tables(doc, persons)

    process_headers_footers(doc, persons)

    doc.save(output_stream)