import streamlit as st
from docx import Document
from io import BytesIO
import zipfile

from main import anonymize_docx, scan_document_for_persons

st.set_page_config(
    page_title="Word anonymiserare",
    page_icon="🔒"
)

st.title("🔒 Dokumentanonymisering")

st.write("""
Ladda upp ett eller flera Word-dokument (.docx) för att anonymisera personuppgifter.

Integritet och dataskydd

• Dokument behandlas endast under den aktiva sessionen
• Ingen permanent lagring sker i systemet
• Bearbetning sker i serverns arbetsminne
• Inga data skickas till externa tjänster eller AI-API
• Anonymiserade dokument genereras direkt och laddas ned av användaren

Observera:
Ladda inte upp dokument som innehåller känsliga personuppgifter
(t.ex. hälsodata eller personnummer).

Sådana uppgifter kräver enligt GDPR en högre skyddsnivå och ska
normalt inte behandlas i denna typ av verktyg.
""")

uploaded_files = st.file_uploader(
    "Ladda upp dokument",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:

    all_persons = set()

    file_data = []

    for uploaded_file in uploaded_files:

        file_bytes = uploaded_file.read()

        doc_stream = BytesIO(file_bytes)

        doc = Document(doc_stream)

        persons = scan_document_for_persons(doc)

        all_persons.update(persons)

        file_data.append((uploaded_file.name, file_bytes))

    st.subheader("Identifierade personer i alla dokument")

    selected_persons = []

    for person in sorted(all_persons):

        if st.checkbox(person, value=True):

            selected_persons.append(person)

    manual_names = st.text_area(
        "Lägg till namn manuellt (komma eller radbrytning)"
    )

    if manual_names:

        for line in manual_names.split("\n"):

            parts = line.split(",")

            for p in parts:

                p = p.strip()

                if p:
                    selected_persons.append(p)

    if st.button("Anonymisera alla dokument"):

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:

            for filename, file_bytes in file_data:

                input_stream = BytesIO(file_bytes)

                output_stream = BytesIO()

                anonymize_docx(
                    input_stream,
                    output_stream,
                    selected_persons
                )

                zip_file.writestr(
                    f"anon_{filename}",
                    output_stream.getvalue()
                )

        st.success("Alla dokument anonymiserade")

        st.download_button(
            "Ladda ner alla anonymiserade dokument",
            data=zip_buffer.getvalue(),
            file_name="anonymiserade_dokument.zip"
        )
