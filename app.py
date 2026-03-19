import streamlit as st
from processor import processa_excel

st.set_page_config(page_title="Divisore Prezzi Carburanti", page_icon="⛽", layout="centered")

st.title("⛽ Divisore Prezzi Carburanti")
st.markdown("Carica il file sorgente `.xls` e scarica il risultato con tutti i fogli divisi.")

file_sorgente = st.file_uploader("📂 File sorgente (dati prezzi)", type=["xls", "xlsx"])

with st.expander("ℹ️ Struttura attesa del file"):
    st.markdown("""
    **File sorgente** (`.xls` o `.xlsx`):
    - Riga 8 = intestazione colonne
    - Riga 9+ = dati
    - Colonne: Codice gestore, Comune PDV, Indirizzo PDV, Insegna, Comune conc., Indirizzo conc., Distanza, Gasolio, Benzina, GPL, Metano

    **Mapping e PDV** sono già caricati internamente dall'applicazione (`mapping.csv` e `pdv_selezionati.csv`).
    Per aggiornarli, sostituire i file CSV nella cartella del progetto su GitHub.
    """)

if file_sorgente:
    if st.button("🚀 Elabora", type="primary"):
        with st.spinner("Elaborazione in corso..."):
            try:
                src_bytes = file_sorgente.read()
                filename  = file_sorgente.name
                output_bytes, n_fogli = processa_excel(src_bytes, filename=filename)
                st.success(f"✅ Completato! Creati {n_fogli} fogli.")
                st.download_button(
                    label="⬇️ Scarica Excel risultato",
                    data=output_bytes,
                    file_name="risultato_diviso.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Errore: {e}")
                st.exception(e)
