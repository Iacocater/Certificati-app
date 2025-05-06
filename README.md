# Generatore Certificati CISL Medici

Applicazione Streamlit per generare automaticamente certificati personalizzati da un file Excel + template Word.

## Funzionalità
- Upload Excel con dati assicurati
- Upload template Word con segnaposto (es. {{Nome}}, {{CodiceFiscale}})
- Scelta del campo per il nome del file
- Generazione automatica di DOCX e PDF
- Download ZIP con tutti i documenti

## Requisiti
- Python 3.8+
- LibreOffice installato (per la conversione PDF)

## Avvio locale

```bash
streamlit run streamlit_app.py
```
