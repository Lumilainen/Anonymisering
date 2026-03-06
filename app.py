import streamlit as st
from docx import Document
from io import BytesIO

from main import (
    anonymize_docx,
    scan_document_for_persons,
    collect_all_text,
    count_occurrences
)

st.set_page_config(
    page_title="Word anonymiserare",
    page_icon="🔒"
)

st.title("🔒 Dokumentanonymisering")

st.write("""
Ladda upp ett Word-dokument (.docx) för att anonymisera personuppgifter.

Integritet:
- Dokument lagras inte
- Bearbetning sker endast i minnet
- Filer raderas efter nedladdning
""")

uploaded_file = st.file_uploader(
    "Ladda upp dokument",
    type=["docx"]
)

if uploaded_file:

    file_bytes = uploaded_file.read()

    doc_stream = BytesIO(file_bytes)

    doc = Document(doc_stream)

    persons = sorted(scan_document_for_persons(doc))

    full_text = collect_all_text(doc)

    counts = count_occurrences(full_text, persons)

    selected_persons = []

    if persons:

        st.subheader("Identifierade personer")

        for person in persons:

            label = f"{person} ({counts.get(person,0)} träffar)"

            checked = st.checkbox(label, value=True)

            if checked:
                selected_persons.append(person)

    else:

        st.info("Inga personer identifierades")

    st.subheader("Lägg till namn manuellt")

    manual_names = st.text_area(
        "Skriv flera namn separerade med komma eller radbrytning"
    )

    if manual_names:

        names = []

        for line in manual_names.split("\n"):

            parts = line.split(",")

            for p in parts:

                p = p.strip()

                if p:
                    names.append(p)

        selected_persons.extend(names)

    if st.button("Starta anonymisering"):

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