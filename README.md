# ⛽ Divisore Prezzi Carburanti

Web app che sostituisce la macro Excel VBA per dividere i dati di prezzi carburanti in fogli separati per codice gestore.

## Struttura del progetto

```
prezzi_carburanti/
├── app.py           ← interfaccia Streamlit
├── processor.py     ← logica di elaborazione (equivalente della macro VBA)
├── requirements.txt
└── README.md
```

---

## Come avviare in locale

### 1. Installa le dipendenze
```bash
pip install -r requirements.txt
```

### 2. Avvia l'app
```bash
streamlit run app.py
```

Si aprirà automaticamente nel browser su `http://localhost:8501`

---

## Come condividere con altri (Streamlit Community Cloud)

1. Crea un account su [streamlit.io](https://streamlit.io)
2. Carica questa cartella su un repository GitHub (anche privato)
3. Su Streamlit Cloud → "New app" → seleziona il repo → `app.py` come file principale
4. Clicca "Deploy" → ottieni un link pubblico (es. `https://tuonome-prezzi.streamlit.app`)
5. Chiunque apre il link, carica il suo Excel, scarica il risultato — senza installare nulla

---

## Struttura attesa del file Excel in input

| Foglio | Descrizione |
|--------|-------------|
| **Foglio 1** (sorgente) | Col.1 = Codice Gestore, Col.2 = Codice Carburante, Col.4+ = dati (distanza, prezzi...) |
| **Mapping** | Col.1 = codice, Col.2 = nome, Col.3 = categoria colore (CAD/CCN/CIA/CNO/PAC 2000A), Col.4 = distanza max, Col.5-8 = flag 1/0 per Gasolio/Benzina/GPL/Metano |
| **PDV_selezionati** *(opzionale)* | Col.1 = codici PDV da evidenziare in giallo |

---

## Cosa fa l'app (equivalenza con la macro VBA)

| Funzionalità | VBA | Python |
|---|---|---|
| Legge Mapping + colori + distanze + flags | ✅ | ✅ |
| Legge PDV selezionati | ✅ | ✅ |
| Divide dati per coppia Codice1+Codice2 | ✅ | ✅ |
| Ordina fogli per nome mapping | ✅ | ✅ |
| Colori tab (CAD/CCN/CIA/CNO/PAC 2000A) | ✅ | ✅ |
| Crea foglio Indice con hyperlink colorati | ✅ | ✅ |
| Rimuove colonne carburanti con flag=0 | ✅ | ✅ |
| Filtra righe con prezzi tutti a 0 | ✅ | ✅ |
| Filtra righe oltre soglia distanza | ✅ | ✅ |
| Ordina per distanza crescente | ✅ | ✅ |
| Evidenzia PDV selezionati in giallo | ✅ | ✅ |
| Grafico a colonne per ogni foglio | ✅ | ✅ |
| Media locale nel grafico | ✅ | ✅ |
| Formato prezzi 0,000 | ✅ | ✅ |
| AutoFit colonne | ✅ | ✅ |
