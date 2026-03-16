import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.label import DataLabel
from openpyxl.worksheet.hyperlink import Hyperlink

# ── Colori tab per categoria ──────────────────────────────────────────────────
COLORI_CATEGORIA = {
    "CAD":      "FF9999",
    "CCN":      "99FF99",
    "CIA":      "FFFF99",
    "CNO":      "99CCFF",
    "PAC 2000A":"FFC000",
}

COLORI_CARBURANTI = {
    "Gasolio": "ED7D31",
    "Benzina": "70AD47",
    "GPL":     "4472C4",
    "Metano":  "E9967A",
}


def _pulisci_nome_foglio(nome: str) -> str:
    for ch in r"/\:*?[]'":
        nome = nome.replace(ch, "-")
    return nome[:31]


def _leggi_mapping(wb) -> tuple[dict, dict, dict, dict]:
    """Legge il foglio Mapping e restituisce 4 dizionari."""
    ws = wb["Mapping"]
    mapping, colori, distanze, carburanti = {}, {}, {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = str(row[0]).strip() if row[0] else ""
        if not k:
            continue
        mapping[k]    = str(row[1]).strip() if row[1] else ""
        colori[k]     = str(row[2]).strip().upper() if row[2] else ""
        distanze[k]   = row[3] if row[3] is not None else ""
        flags = [
            int(row[4]) if row[4] is not None else 1,
            int(row[5]) if row[5] is not None else 1,
            int(row[6]) if row[6] is not None else 1,
            int(row[7]) if row[7] is not None else 1,
        ]
        carburanti[k] = flags   # [Gasolio, Benzina, GPL, Metano]
    return mapping, colori, distanze, carburanti


def _leggi_pdv(wb) -> set:
    if "PDV_selezionati" not in wb.sheetnames:
        return set()
    ws = wb["PDV_selezionati"]
    pdv = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            pdv.add(str(row[0]).strip())
    return pdv


def _leggi_sorgente(wb) -> pd.DataFrame:
    ws = wb.worksheets[0]
    data = list(ws.values)
    df = pd.DataFrame(data[1:], columns=data[0])
    return df


def _normalizza_prezzi(val):
    if val is None:
        return 0.0
    s = str(val).replace("'", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def processa_excel(file_bytes: bytes) -> bytes:
    wb_in = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)

    mapping, colori, distanze, carburanti = _leggi_mapping(wb_in)
    pdv_selezionati = _leggi_pdv(wb_in)
    df_source = _leggi_sorgente(wb_in)

    # Colonne: index 0 = codice1, 1 = codice2, 3+ = dati (primaColonnaDati = 4 → index 3)
    col_names = df_source.columns.tolist()
    col_codice1 = col_names[0]
    col_codice2 = col_names[1]
    prima_col_dati = 3   # index base-0

    df_source[col_codice1] = df_source[col_codice1].astype(str).str.strip()
    df_source[col_codice2] = df_source[col_codice2].astype(str).str.strip()

    # Estrai coppie univoche ordinate per nome mapping
    coppie = df_source[[col_codice1, col_codice2]].drop_duplicates()
    coppie = coppie[coppie[col_codice1].str.len() > 0]
    coppie["_nome_sort"] = coppie[col_codice1].map(
        lambda k: mapping.get(k, "ZZ_NonMappato_" + k)
    )
    coppie = coppie.sort_values("_nome_sort").drop(columns="_nome_sort")

    # ── Nuovo workbook output ────────────────────────────────────────────────
    from openpyxl import Workbook
    wb_out = Workbook()
    wb_out.remove(wb_out.active)   # rimuoviamo il foglio default

    ws_indice = wb_out.create_sheet("Indice", 0)
    ws_indice["A1"] = "Indice Fogli Generati"
    ws_indice["A1"].font = Font(bold=True)
    riga_indice = 2

    nomi_fogli_creati = []

    for _, coppia in coppie.iterrows():
        v1 = coppia[col_codice1]
        v2 = coppia[col_codice2]

        nome_da_mapping = mapping.get(v1, f"NonMappato_{v1}")
        nome_foglio = _pulisci_nome_foglio(nome_da_mapping)

        # Evita duplicati nel nome foglio
        base_nome = nome_foglio
        suffisso = 1
        while nome_foglio in nomi_fogli_creati:
            nome_foglio = f"{base_nome[:28]}_{suffisso}"
            suffisso += 1
        nomi_fogli_creati.append(nome_foglio)

        # Filtra righe sorgente
        df_filt = df_source[
            (df_source[col_codice1] == v1) & (df_source[col_codice2] == v2)
        ].copy()

        # Prendi solo le colonne dati (dalla 4a in poi)
        df_dati = df_filt.iloc[:, prima_col_dati:].copy()

        # ── Flags carburanti ────────────────────────────────────────────────
        flags = carburanti.get(v1, [1, 1, 1, 1])
        nomi_carb = ["Gasolio", "Benzina", "GPL", "Metano"]

        # Colonne prezzi: posizioni 5,6,7,8 nel foglio dest (index 4,5,6,7 in df_dati)
        # Le prime 4 col dati sono: col4, col5 del sorgente = le prime 2 colonne di df_dati
        # Struttura attesa: [col4_sorgente, col5_sorgente=distanza, col6=Gasolio, col7=Benzina, col8=GPL, col9=Metano]
        # Adattiamo: colonna distanza = indice 1 in df_dati (base-0), prezzi da indice 2 in poi
        COL_DISTANZA_IDX = 1   # seconda colonna dei dati = distanza

        # Normalizza prezzi
        for ci in range(2, min(6, len(df_dati.columns))):
            df_dati.iloc[:, ci] = df_dati.iloc[:, ci].apply(_normalizza_prezzi)

        if COL_DISTANZA_IDX < len(df_dati.columns):
            df_dati.iloc[:, COL_DISTANZA_IDX] = pd.to_numeric(
                df_dati.iloc[:, COL_DISTANZA_IDX], errors="coerce"
            ).fillna(0)

        # Rimuovi colonne carburanti non attive (da destra)
        col_offset = 2  # i prezzi partono dall'indice 2 di df_dati
        cols_to_drop = []
        for ci, (flag, nome) in enumerate(zip(flags, nomi_carb)):
            col_idx = col_offset + ci
            if flag == 0 and col_idx < len(df_dati.columns):
                cols_to_drop.append(df_dati.columns[col_idx])
        df_dati = df_dati.drop(columns=cols_to_drop)

        # Colonne prezzi rimaste
        flag_attivi = [(i, n) for i, (f, n) in enumerate(zip(flags, nomi_carb)) if f == 1]
        num_prezzi = len(flag_attivi)

        # Filtro: rimuovi righe con somma prezzi = 0
        if num_prezzi > 0:
            idx_prezzi_dati = [col_offset + ci for ci in range(num_prezzi) if col_offset + ci < len(df_dati.columns)]
            df_dati["_somma"] = df_dati.iloc[:, idx_prezzi_dati].sum(axis=1)
            df_dati = df_dati[df_dati["_somma"] > 0].drop(columns="_somma")

        # Filtro distanza
        soglia = distanze.get(v1, "")
        if soglia != "" and COL_DISTANZA_IDX < len(df_dati.columns):
            try:
                soglia_f = float(soglia)
                df_dati = df_dati[df_dati.iloc[:, COL_DISTANZA_IDX] <= soglia_f]
            except (ValueError, TypeError):
                pass

        # Ordina per distanza
        if COL_DISTANZA_IDX < len(df_dati.columns):
            df_dati = df_dati.sort_values(df_dati.columns[COL_DISTANZA_IDX]).reset_index(drop=True)

        # ── Scrivi foglio ────────────────────────────────────────────────────
        ws = wb_out.create_sheet(nome_foglio)

        # Colore tab
        cat = colori.get(v1, "")
        if cat in COLORI_CATEGORIA:
            ws.sheet_properties.tabColor = COLORI_CATEGORIA[cat]
            cell_indice = ws_indice.cell(row=riga_indice, column=1)
            cell_indice.fill = PatternFill("solid", fgColor=COLORI_CATEGORIA[cat])

        # Link indice
        cell_indice = ws_indice.cell(row=riga_indice, column=1, value=nome_foglio)
        cell_indice.hyperlink = f"#{nome_foglio}!A1"
        cell_indice.font = Font(color="0000FF", underline="single")
        riga_indice += 1

        # Intestazione
        headers = list(df_dati.columns)
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font = Font(bold=True)

        # Dati
        for ri, row_data in enumerate(df_dati.itertuples(index=False), start=2):
            pdv_key = str(row_data[2]).strip() if len(row_data) > 2 else ""  # colonna 3 = PDV
            is_pdv = pdv_key in pdv_selezionati
            for ci, val in enumerate(row_data, start=1):
                cell = ws.cell(row=ri, column=ci, value=val)
                # Formato prezzi
                if ci >= col_offset + 1 and ci <= col_offset + num_prezzi:
                    cell.number_format = "0.000"
                if is_pdv:
                    cell.fill = PatternFill("solid", fgColor="FFFF00")

        # Autofit colonne
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 40)

        # ── Grafico ─────────────────────────────────────────────────────────
        ultima_riga = df_dati.shape[0] + 1   # +1 per header

        # Righe gialle (PDV selezionati) o tutte
        righe_grafico = []
        for ri in range(2, ultima_riga + 1):
            if ws.cell(row=ri, column=1).fill.fgColor.rgb == "FFFF00":
                righe_grafico.append(ri)
        if not righe_grafico:
            righe_grafico = list(range(2, ultima_riga + 1))

        if ultima_riga > 1 and num_prezzi > 0 and righe_grafico:
            chart = BarChart()
            chart.type = "col"
            chart.title = f"Confronto Prezzi [{nome_foglio}]"
            chart.grouping = "clustered"
            chart.overlap = 0
            chart.gapWidth = 100
            chart.width = 25
            chart.height = 14

            # Etichette asse X: col 1 + col 3 (nome + PDV)
            cats_vals = []
            for ri in righe_grafico:
                lbl = f"{ws.cell(ri,1).value} - {ws.cell(ri,3).value}, {ws.cell(ri,2).value}"
                cats_vals.append(lbl)

            for ci_s, (orig_idx, nome_carb) in enumerate(flag_attivi):
                col_excel = col_offset + 1 + ci_s   # colonna Excel (1-based)
                colore_hex = COLORI_CARBURANTI.get(nome_carb, "4472C4")

                valori = []
                for ri in righe_grafico:
                    v = ws.cell(row=ri, column=col_excel).value or 0
                    valori.append(round(float(v), 3) if v else 0)

                # Media (escludi distanza 0)
                vals_no_zero = [
                    v for ri, v in zip(righe_grafico, valori)
                    if v > 0 and (ws.cell(ri, COL_DISTANZA_IDX + 1).value or 0) != 0
                ]
                media = round(sum(vals_no_zero) / len(vals_no_zero), 3) if vals_no_zero else 0

                # Serie dati + media
                serie_vals = valori + [media]
                serie_lbl  = cats_vals + ["MEDIA"]

                # Scrivi i dati temporanei in righe nascoste in fondo al foglio
                riga_temp_base = ultima_riga + 5 + ci_s * 2
                ws.cell(row=riga_temp_base, column=1, value=f"_chart_serie_{nome_carb}")
                for ci_t, (lbl, val) in enumerate(zip(serie_lbl, serie_vals), start=2):
                    ws.cell(row=riga_temp_base, column=ci_t, value=lbl)
                    ws.cell(row=riga_temp_base + 1, column=ci_t, value=val)
                ws.row_dimensions[riga_temp_base].hidden = True
                ws.row_dimensions[riga_temp_base + 1].hidden = True

                n_punti = len(serie_vals)
                ser = Series(
                    Reference(ws, min_col=2, max_col=n_punti + 1, min_row=riga_temp_base + 1),
                    title=nome_carb
                )
                ser.graphicalProperties.solidFill = colore_hex
                chart.series.append(ser)

            # Posiziona il grafico
            col_chart = get_column_letter(len(headers) + 2)
            ws.add_chart(chart, f"{col_chart}1")

    # Autofit indice
    ws_indice.column_dimensions["A"].width = 35

    # ── Salva in bytes ───────────────────────────────────────────────────────
    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output.read()
