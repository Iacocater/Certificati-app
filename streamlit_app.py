import streamlit as st
import pandas as pd
import os
import re
import zipfile
import tempfile
from docxtpl import DocxTemplate
from jinja2 import Environment, Undefined

st.set_page_config(page_title="Generatore Certificati DOCX", layout="centered")
st.title("📄 Generatore massivo di certificati - Solo Word (.docx)")

# --- Utility ---
def safe_str(x):
    """Converte valori (anche NaN) in stringa pulita."""
    if pd.isna(x):
        return ""
    return str(x).strip()

def sanitize_filename(name: str, fallback: str = "documento") -> str:
    """Rende il nome file valido su Windows/macOS/Linux."""
    name = safe_str(name)
    if not name:
        name = fallback
    # sostituisce caratteri vietati
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    # riduce spazi e underscore multipli
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"_+", "_", name)
    return name[:120]  # evita nomi troppo lunghi

class BlankUndefined(Undefined):
    """Se un campo manca nel template, lo rende vuoto invece di crashare."""
    def __str__(self):
        return ""
    def __unicode__(self):
        return ""

# Upload file
excel_file = st.file_uploader("📥 Carica il file Excel (.xlsx)", type=["xlsx"])
word_template = st.file_uploader("📄 Carica il template Word (.docx)", type=["docx"])

if excel_file and word_template:
    # Legge Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Errore lettura Excel: {e}")
        st.stop()

    if df.empty:
        st.warning("Il file Excel è vuoto.")
        st.stop()

    st.success("Excel caricato correttamente. Colonne trovate:")
    st.write(df.columns.tolist())

    filename_field = st.selectbox("📝 Seleziona il campo da usare per rinominare i file", df.columns)

    # Opzione utile: mostra anteprima della prima riga (per vedere che CodiceFiscale ecc. ci sia)
    with st.expander("🔎 Anteprima prima riga (debug)"):
        st.json({c: safe_str(df.loc[df.index[0], c]) for c in df.columns})

    if st.button("🚀 Genera certificati"):
        with st.spinner("Creazione documenti in corso..."):
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "certificati_word.zip")

            # Salva il template caricato su disco (Streamlit Cloud-friendly)
            template_path = os.path.join(temp_dir, "template.docx")
            try:
                with open(template_path, "wb") as f:
                    f.write(word_template.getbuffer())
            except Exception as e:
                st.error(f"Errore nel salvataggio del template: {e}")
                st.stop()

            docx_files = []
            errors = []

            # Ambiente Jinja: campi mancanti -> stringa vuota
            jinja_env = Environment(undefined=BlankUndefined)

            for idx, row in df.iterrows():
                # Contesto: tutte stringhe (evita problemi con date/float/NaN)
                context = {col: safe_str(row[col]) for col in df.columns}

                # Nome file
                base_name = sanitize_filename(context.get(filename_field, ""), fallback=f"riga_{idx+1}")
                out_path = os.path.join(temp_dir, f"{base_name}.docx")

                try:
                    doc = DocxTemplate(template_path)
                    doc.render(context, jinja_env)
                    doc.save(out_path)
                    docx_files.append(out_path)
                except Exception as e:
                    errors.append((idx, base_name, str(e)))

            # Se tutto fallisce, fermati e mostra errori
            if not docx_files:
                st.error("Nessun documento generato. Vedi errori qui sotto.")
                for idx, base_name, err in errors[:10]:
                    st.write(f"- Riga {idx+1} ({base_name}): {err}")
                st.stop()

            # ZIP
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for p in docx_files:
                    zipf.write(p, arcname=os.path.basename(p))

            if errors:
                st.warning(f"Generati {len(docx_files)} file. Alcune righe hanno dato errore: {len(errors)}.")
                with st.expander("📌 Dettaglio errori (prime 20 righe)"):
                    for idx, base_name, err in errors[:20]:
                        st.write(f"- Riga {idx+1} ({base_name}): {err}")
            else:
                st.success("✅ Tutti i file Word sono stati creati senza errori.")

            with open(zip_path, "rb") as f:
                st.download_button("⬇️ Scarica il pacchetto ZIP", f, file_name="certificati_word.zip")
