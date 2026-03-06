import streamlit as st
from docx import Document
from io import BytesIO
import spacy
import subprocess
import sys

# säker laddning av spaCy-modell
MODEL_NAME = "sv_core_news_sm"

try:
    nlp = spacy.load(MODEL_NAME)
except:
    subprocess.check_call([sys.executable, "-m", "spacy", "download", MODEL_NAME])
    nlp = spacy.load(MODEL_NAME)

from main import anonymize_docx, scan_document_for_persons


st.set_page_config(
    page_title="Word anonymiserare",
    page_icon="🔒",
    layout="centered"
)

st.title("🔒 Dokumentanonymisering")

st.write(
"""
Ladda upp ett Word-dokument (.docx) för att anonymisera personuppgifter.

**Integritet**
- Dokument lagras inte
- Bearbetning sker endast i minnet
- Filer raderas automatiskt efter nedladdning
"""
)

uploaded_file = st.file_uploader(
    "Ladda upp dokument",
    type=["docx"]
)

if uploaded_file:

    file_bytes = uploaded_file.read()

    doc_stream = BytesIO(file_bytes)

    doc = Document(doc_stream)

    persons = sorted(scan_document_for_persons(doc))

    if persons:

        st.subheader("Identifierade personer")

        selected_persons = []

        for person in persons:

            checked = st.checkbox(person, value=True)

            if checked:
                selected_persons.append(person)

    else:

        selected_persons = []

        st.info("Inga personer identifierades")

    st.subheader("Manuell anonymisering")

    manual_name = st.text_input(
        "Lägg till namn som ska anonymiseras"
    )

    if manual_name:

        if manual_name not in selected_persons:
            selected_persons.append(manual_name)

    if st.button("Starta anonymisering"):

        with st.spinner("Anonymiserar dokument..."):

            input_stream = BytesIO(file_bytes)
            output_stream = BytesIO()

            anonymize_docx(
                input_stream,
                output_stream,
                selected_persons
            )

        st.success("Dokument anonymiserat")

        st.download_button(
            "Ladda ner anonymiserat dokument",
            data=output_stream.getvalue(),
            file_name="anonymiserad_fil.docx"
        )