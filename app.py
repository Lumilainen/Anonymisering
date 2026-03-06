import streamlit as st
from docx import Document
from io import BytesIO

from main import anonymize_docx, scan_document_for_persons

st.set_page_config(
    page_title="Word anonymiserare",
    page_icon="🔒"
)

st.title("🔒 Dokumentanonymisering")

st.write("""
Ladda upp ett eller flera Word-dokument (.docx) för att anonymisera personuppgifter.

Integritet:
- Dokument lagras inte
- Bearbetning sker endast i minnet
- Filer raderas efter nedladdning
""")

uploaded_files = st.file_uploader(
    "Ladda upp dokument",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:

    for uploaded_file in uploaded_files:

        st.divider()

        st.subheader(uploaded_file.name)

        file_bytes = uploaded_file.read()

        doc_stream = BytesIO(file_bytes)

        doc = Document(doc_stream)

        persons = sorted(scan_document_for_persons(doc))

        selected_persons = []

        if persons:

            st.write("Identifierade personer")

            for person in persons:

                checked = st.checkbox(
                    person,
                    value=True,
                    key=f"{uploaded_file.name}_{person}"
                )

                if checked:
                    selected_persons.append(person)

        else:

            st.info("Inga personer identifierades")

        manual_names = st.text_area(
            "Lägg till namn manuellt",
            key=f"manual_{uploaded_file.name}"
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

        if st.button(
            f"Anonymisera {uploaded_file.name}",
            key=f"button_{uploaded_file.name}"
        ):

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
                file_name=f"anon_{uploaded_file.name}"
            )