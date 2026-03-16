import streamlit as st
import pandas as pd
import io
from processor import processa_excel

st.set_page_config(page_title="Divisore Prezzi Carburanti", page_icon="⛽", layout="centered")

st.title("⛽ Divisore Prezzi Carburanti")
st.markdown("Carica il file Excel sorgente e scarica il risultato con tutti i fogli divisi per codice.")

uploaded_file = st.file_uploader("📂 Carica il file Excel", type=["xlsx"])

if uploaded_file:
    st.success(f"File caricato: **{uploaded_file.name}**")

    with st.expander("ℹ️ Struttura attesa del file"):
        st.markdown("""
        Il file deve contenere questi fogli:
        - **Foglio 1** (sorgente): colonna 1 = Codice Gestore, colonna 2 = Codice Carburante, colonna 4+ = dati
        - **Mapping**: col.1 = codice, col.2 = nome, col.3 = categoria colore, col.4 = distanza max, col.5-8 = flag Gasolio/Benzina/GPL/Metano
        - **PDV_selezionati** *(opzionale)*: colonna 1 = codici da evidenziare in giallo
        """)

    if st.button("🚀 Elabora file", type="primary"):
        with st.spinner("Elaborazione in corso..."):
            try:
                output_bytes = processa_excel(uploaded_file.read())
                st.success("✅ Elaborazione completata!")
                st.download_button(
                    label="⬇️ Scarica Excel risultato",
                    data=output_bytes,
                    file_name="risultato_diviso.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Errore durante l'elaborazione: {e}")
                st.exception(e)
