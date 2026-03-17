import pandas as pd
import io
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, Series

COLORI_CATEGORIA = {
    "CAD":       "FF9999",
    "CCN":       "99FF99",
    "CIA":       "FFFF99",
    "CNO":       "99CCFF",
    "PAC 2000A": "FFC000",
}

COLORI_CARBURANTI = {
    "Gasolio": "ED7D31",
    "Benzina": "70AD47",
    "GPL":     "4472C4",
    "Metano":  "E9967A",
}

NOMI_CARBURANTI = ["Gasolio", "Benzina", "GPL", "Metano"]


def _pulisci_nome_foglio(nome: str) -> str:
    for ch in r"/\:*?[]'":
        nome = nome.replace(ch, "-")
    return nome[:31]


def _leggi_mapping(wb):
    ws = wb["Mapping"]
    mapping, colori, distanze, carburanti = {}, {}, {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = str(row[0]).strip() if row[0] else ""
        if not k:
            continue
        mapping[k]  = str(row[1]).strip() if row[1] else ""
        colori[k]   = str(row[2]).strip().upper() if row[2] else ""
        distanze[k] = row[3] if row[3] is not None else ""
        carburanti[k] = [
            int(row[4]) if row[4] is not None else 1,
            int(row[5]) if row[5] is not None else 1,
            int(row[6]) if row[6] is not None else 1,
            int(row[7]) if row[7] is not None else 1,
        ]
    return mapping, colori, distanze, carburanti


def _leggi_pdv(wb) -> set:
    if "PDV_selezionati" not in wb.sheetnames:
        return set()
    ws = wb["PDV_selezionati"]
    return {str(row[0]).strip() for row in ws.iter_rows(min_row=2, values_only=True) if row[0]}


def _parse_distanza(val) -> float:
    """Converte '14,9' oppure 14.9 in float."""
    if val is None:
        return 0.0
    s = str(val).replace("'", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_prezzo(val) -> float:
    if val is None or str(val).strip() in ("-", "", "None"):
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

    # ── Leggi foglio sorgente ────────────────────────────────────────────────
    # Struttura colonne (0-based):
    # 0=Codice gestore, 1=Comune PDV, 2=Indirizzo PDV
    # 3=Insegna, 4=Comune conc., 5=Indirizzo conc., 6=Distanza
    # 7=Gasolio, 8=Benzina, 9=GPL, 10=Metano
    ws_src = wb_in.worksheets[0]
    rows = list(ws_src.values)
    data_rows = rows[1:]  # salta intestazione

    IDX_CODICE   = 0
    IDX_COMUNE   = 1
    IDX_DISTANZA = 6
    IDX_PREZZI   = [7, 8, 9, 10]

    # Estrai coppie univoche ordinate per nome mapping
    coppie_viste = {}
    for r in data_rows:
        k1 = str(r[IDX_CODICE]).strip() if r[IDX_CODICE] else ""
        k2 = str(r[IDX_COMUNE]).strip() if r[IDX_COMUNE] else ""
        chiave = f"{k1}|{k2}"
        if k1 and k2 and chiave not in coppie_viste:
            coppie_viste[chiave] = (k1, k2)

    coppie_ordinate = sorted(
        coppie_viste.values(),
        key=lambda x: mapping.get(x[0], "ZZ_" + x[0])
    )

    # ── Nuovo workbook output ────────────────────────────────────────────────
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    ws_indice = wb_out.create_sheet("Indice", 0)
    ws_indice["A1"] = "Indice Fogli Generati"
    ws_indice["A1"].font = Font(bold=True)
    riga_indice = 2
    nomi_fogli_creati = []

    for (v1, v2) in coppie_ordinate:
        nome_da_mapping = mapping.get(v1, f"NonMappato_{v1}")
        nome_foglio = _pulisci_nome_foglio(nome_da_mapping)
        base = nome_foglio
        suf = 1
        while nome_foglio in nomi_fogli_creati:
            nome_foglio = f"{base[:28]}_{suf}"; suf += 1
        nomi_fogli_creati.append(nome_foglio)

        # Filtra righe per questa coppia
        righe_filtrate = [
            r for r in data_rows
            if str(r[IDX_CODICE]).strip() == v1 and str(r[IDX_COMUNE]).strip() == v2
        ]

        flags = carburanti.get(v1, [1, 1, 1, 1])
        carb_attivi = [(i, NOMI_CARBURANTI[i]) for i, f in enumerate(flags) if f == 1]

        # Soglia distanza
        soglia = distanze.get(v1, "")
        try:
            soglia_f = float(str(soglia).replace(",", ".")) if soglia != "" else None
        except (ValueError, TypeError):
            soglia_f = None

        # Costruisci e filtra righe elaborate
        righe_elaborate = []
        for r in righe_filtrate:
            dist = _parse_distanza(r[IDX_DISTANZA])
            prezzi = [_parse_prezzo(r[IDX_PREZZI[i]]) for i in range(4)]
            somma_attivi = sum(prezzi[i] for i, f in enumerate(flags) if f == 1)

            if somma_attivi == 0:
                continue
            if soglia_f is not None and dist > soglia_f:
                continue

            riga_out = [
                str(r[3]).strip() if r[3] else "",   # Insegna
                str(r[4]).strip() if r[4] else "",   # Comune concorrente
                str(r[5]).strip() if r[5] else "",   # Indirizzo concorrente
                dist,                                 # Distanza (numero)
            ]
            for i, _ in carb_attivi:
                riga_out.append(round(prezzi[i], 3) if prezzi[i] > 0 else None)

            righe_elaborate.append(riga_out)

        # Ordina per distanza
        righe_elaborate.sort(key=lambda x: x[3])

        # ── Scrivi foglio ────────────────────────────────────────────────────
        ws = wb_out.create_sheet(nome_foglio)

        cat = colori.get(v1, "")
        colore_hex = COLORI_CATEGORIA.get(cat, None)
        if colore_hex:
            ws.sheet_properties.tabColor = colore_hex

        # Voce indice
        cell_idx = ws_indice.cell(row=riga_indice, column=1, value=nome_foglio)
        cell_idx.hyperlink = f"#{nome_foglio}!A1"
        cell_idx.font = Font(color="0000FF", underline="single")
        if colore_hex:
            cell_idx.fill = PatternFill("solid", fgColor=colore_hex)
        riga_indice += 1

        # Intestazione
        intestazione = ["Insegna", "Comune", "Indirizzo", "Distanza (km)"] + \
                       [n for _, n in carb_attivi]
        for ci, h in enumerate(intestazione, 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)

        # Righe dati
        for ri, riga in enumerate(righe_elaborate, 2):
            pdv_key = riga[0]
            is_pdv = pdv_key in pdv_selezionati
            for ci, val in enumerate(riga, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                if ci >= 5:
                    cell.number_format = "0.000"
                if is_pdv:
                    cell.fill = PatternFill("solid", fgColor="FFFF00")

        # Autofit colonne
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 40)

        # ── Grafico ──────────────────────────────────────────────────────────
        ultima_riga = len(righe_elaborate) + 1
        if ultima_riga > 1 and carb_attivi:
            righe_grafico = [
                ri for ri in range(2, ultima_riga + 1)
                if ws.cell(ri, 1).fill.fgColor.rgb in ("FFFF00", "FFFFFF00")
            ]
            if not righe_grafico:
                righe_grafico = list(range(2, ultima_riga + 1))

            n_punti = len(righe_grafico) + 1  # +1 per MEDIA
            riga_base_temp = ultima_riga + 3

            # Riga etichette (nascosta)
            ws.cell(riga_base_temp, 1, "_lbl")
            for ci, ri in enumerate(righe_grafico, 2):
                ws.cell(riga_base_temp, ci, f"{ws.cell(ri,1).value} - {ws.cell(ri,2).value}")
            ws.cell(riga_base_temp, len(righe_grafico) + 2, "MEDIA")
            ws.row_dimensions[riga_base_temp].hidden = True

            chart = BarChart()
            chart.type = "col"
            chart.title = f"Confronto Prezzi [{nome_foglio}]"
            chart.grouping = "clustered"
            chart.gapWidth = 100
            chart.width = 25
            chart.height = 14

            for ci_s, (orig_idx, nome_carb) in enumerate(carb_attivi):
                col_excel = 5 + ci_s
                colore_carb = COLORI_CARBURANTI.get(nome_carb, "4472C4")

                valori = [ws.cell(ri, col_excel).value or 0 for ri in righe_grafico]
                vals_media = [
                    v for ri, v in zip(righe_grafico, valori)
                    if v and float(v) > 0 and (ws.cell(ri, 4).value or 0) != 0
                ]
                media = round(sum(vals_media) / len(vals_media), 3) if vals_media else 0

                riga_val = riga_base_temp + 1 + ci_s
                ws.cell(riga_val, 1, nome_carb)
                for ci_t, v in enumerate(valori + [media], 2):
                    ws.cell(riga_val, ci_t, round(float(v), 3) if v else 0)
                ws.row_dimensions[riga_val].hidden = True

                ser = Series(
                    Reference(ws, min_col=2, max_col=n_punti + 1, min_row=riga_val),
                    title=nome_carb
                )
                ser.graphicalProperties.solidFill = colore_carb
                chart.series.append(ser)

            col_chart = get_column_letter(len(intestazione) + 2)
            ws.add_chart(chart, f"{col_chart}1")

    ws_indice.column_dimensions["A"].width = 35

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output.read()
