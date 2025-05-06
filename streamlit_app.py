import streamlit as st
import pandas as pd
import os
import zipfile
import tempfile
from docxtpl import DocxTemplate

st.set_page_config(page_title="Generatore Certificati DOCX", layout="centered")
st.title("📄 Generatore massivo di certificati - Solo Word (.docx)")

# Upload dei file
excel_file = st.file_uploader("📥 Carica il file Excel (.xlsx)", type=["xlsx"])
word_template = st.file_uploader("📄 Carica il template Word (.docx)", type=["docx"])

if excel_file and word_template:
    df = pd.read_excel(excel_file)
    st.success("Excel caricato correttamente. Colonne trovate:")
    st.write(df.columns.tolist())

    filename_field = st.selectbox("📝 Seleziona il campo da usare per rinominare i file", df.columns)

    if st.button("🚀 Genera certificati"):
        with st.spinner("Creazione documenti in corso..."):
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "certificati_word.zip")
            docx_files = []

            for _, row in df.iterrows():
                context = row.to_dict()
                doc = DocxTemplate(word_template)
                doc.render(context)

                # Costruisce il nome file
                base_name = str(row[filename_field]).strip().replace(" ", "_")
                filename = f"{base_name}.docx"
                filepath = os.path.join(temp_dir, filename)

                doc.save(filepath)
                docx_files.append(filepath)

            # Crea lo zip
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for f in docx_files:
                    zipf.write(f, arcname=os.path.basename(f))

            st.success("Tutti i file Word sono stati creati.")
            with open(zip_path, "rb") as f:
                st.download_button("⬇️ Scarica il pacchetto ZIP", f, file_name="certificati_word.zip")
