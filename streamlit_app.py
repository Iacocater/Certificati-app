import streamlit as st
import pandas as pd
import os
import re
import time
import tempfile
import zipfile
from io import BytesIO
from docxtpl import DocxTemplate
from jinja2 import Environment, Undefined

st.set_page_config(page_title="Generatore Certificati DOCX", layout="centered")
st.title("📄 Generatore massivo di certificati - Solo Word (.docx)")

# ---------- Utility ----------
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def sanitize_filename(name: str, fallback: str = "documento") -> str:
    name = safe_str(name)
    if not name:
        name = fallback
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"_+", "_", name)
    name = name.strip("._ ")
    return name[:120] if name else fallback

class BlankUndefined(Undefined):
    def __str__(self):
        return ""
    def __unicode__(self):
        return ""

jinja_env = Environment(undefined=BlankUndefined)

# ---------- UI Upload ----------
excel_file = st.file_uploader("📥 Carica il file Excel (.xlsx)", type=["xlsx"])
word_template = st.file_uploader("📄 Carica il template Word (.docx)", type=["docx"])

if excel_file and word_template:
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Errore lettura Excel: {e}")
        st.stop()

    if df.empty:
        st.warning("Il file Excel è vuoto.")
        st.stop()

    st.success("Excel caricato. Colonne trovate:")
    st.write(df.columns.tolist())

    filename_field = st.selectbox("📝 Campo per rinominare i file", df.columns)

    # Parametri stabilità
    batch_size = st.number_input("📦 Batch (righe per blocco)", min_value=10, max_value=500, value=80, step=10)
    sleep_ms = st.number_input("⏱️ Micro-pausa (ms) ogni riga (anti-freeze)", min_value=0, max_value=50, value=5, step=1)

    with st.expander("🔎 Anteprima prima riga (debug)"):
        st.json({c: safe_str(df.loc[df.index[0], c]) for c in df.columns})

    if st.button("🚀 Genera certificati"):
        # Salva template su disco temporaneo (più affidabile in Cloud)
        workdir = tempfile.mkdtemp()
        template_path = os.path.join(workdir, "template.docx")
        with open(template_path, "wb") as f:
            f.write(word_template.getbuffer())

        total = len(df)
        progress = st.progress(0)
        status = st.empty()
        error_box = st.empty()

        # ZIP in memoria (poi download)
        zip_buffer = BytesIO()

        errors = []
        used_names = {}

        t0 = time.time()

        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
            # Processa a blocchi
            for start in range(0, total, int(batch_size)):
                end = min(start + int(batch_size), total)
                chunk = df.iloc[start:end]

                for i, (_, row) in enumerate(chunk.iterrows(), start=start + 1):
                    context = {col: safe_str(row[col]) for col in df.columns}

                    base = sanitize_filename(context.get(filename_field, ""), fallback=f"riga_{i}")

                    # evita nomi duplicati
                    if base in used_names:
                        used_names[base] += 1
                        base_out = f"{base}_{used_names[base]}"
                    else:
                        used_names[base] = 1
                        base_out = base

                    try:
                        doc = DocxTemplate(template_path)
                        doc.render(context, jinja_env)

                        # Salva DOCX in memoria e scrivi dentro lo zip
                        mem = BytesIO()
                        doc.save(mem)
                        mem.seek(0)
                        zipf.writestr(f"{base_out}.docx", mem.read())

                    except Exception as e:
                        errors.append((i, base_out, str(e)))

                    # aggiornamento UI (anti-freeze)
                    if sleep_ms:
                        time.sleep(sleep_ms / 1000.0)

                    progress.progress(i / total)
                    status.write(f"Generazione: {i}/{total}")

        zip_buffer.seek(0)

        elapsed = time.time() - t0
        if errors:
            st.warning(f"Completato con errori: generati {total - len(errors)}/{total} documenti. Tempo: {elapsed:.1f}s")
            with st.expander("📌 Errori (prime 50 righe)"):
                for i, name, err in errors[:50]:
                    st.write(f"- Riga {i} ({name}): {err}")
        else:
            st.success(f"✅ Tutti i documenti creati. Tempo: {elapsed:.1f}s")

        st.download_button(
            "⬇️ Scarica ZIP",
            data=zip_buffer,
            file_name="certificati_word.zip",
            mime="application/zip",
        )
