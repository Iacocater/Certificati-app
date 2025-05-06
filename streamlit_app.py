import streamlit as st
import pandas as pd
import os
import zipfile
import tempfile
from docxtpl import DocxTemplate
from pathlib import Path
import subprocess

st.set_page_config(page_title="Generatore Certificati", layout="centered")
st.title("📄 Generatore massivo di certificati CISL Medici")

excel_file = st.file_uploader("📥 Carica il file Excel", type=["xlsx"])
word_template = st.file_uploader("📄 Carica il template Word (.docx)", type=["docx"])

if excel_file and word_template:
    df = pd.read_excel(excel_file)
    st.success("Excel caricato. Campi trovati:")
    st.write(df.columns.tolist())

    filename_field = st.selectbox("📝 Seleziona il campo da usare per rinominare i certificati", df.columns)

    if st.button("🚀 Genera certificati"):
        with st.spinner("Generazione certificati in corso..."):
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "certificati.zip")

            docx_files = []
            pdf_files = []

            for _, row in df.iterrows():
                context = row.to_dict()
                doc = DocxTemplate(word_template)
                doc.render(context)

                filename_base = str(row[filename_field]).strip().replace(" ", "_")
                docx_path = os.path.join(temp_dir, f"{filename_base}.docx")
                doc.save(docx_path)
                docx_files.append(docx_path)

                # Conversione a PDF tramite LibreOffice
                subprocess.run([
                    "soffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, docx_path
                ], check=True)
                pdf_files.append(os.path.join(temp_dir, f"{filename_base}.pdf"))

            # Zippiamo tutto
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file in docx_files + pdf_files:
                    zipf.write(file, arcname=os.path.basename(file))

            with open(zip_path, "rb") as f:
                st.success("🎉 Tutti i certificati sono stati generati.")
                st.download_button("⬇️ Scarica il pacchetto ZIP", f, file_name="certificati.zip")
