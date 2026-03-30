import re
import io
from collections import defaultdict

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(page_title="PDF to Excel Transposer", layout="wide")
st.title("PDF to Excel - Trasposizione ordini")

st.write(
    "Carica uno o più PDF ordine. "
    "Lo script estrae CODICE, COLORE, DESCRIZIONE, PREZZO WHS, PREZZO RTL "
    "e le quantità per taglia, poi genera un Excel."
)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def parse_price_line(text: str):
    """
    Esempio:
    'Prima data di cons. 19.03.26 Ultima data di cons. 19.03.26 Prezzo EUR 45,45 / 109,00 Sconto 4%'
    """
    m = re.search(r"Prezzo\s+EUR\s+([\d.,]+)\s*/\s*([\d.,]+)", text, re.IGNORECASE)
    if not m:
        return None, None
    return m.group(1), m.group(2)


def to_lines(words, y_tol=3):
    """
    Raggruppa le words di pdfplumber in righe.
    """
    rows = []
    for w in sorted(words, key=lambda x: (round(x["top"], 1), x["x0"])):
        placed = False
        for row in rows:
            if abs(row["top"] - w["top"]) <= y_tol:
                row["words"].append(w)
                placed = True
                break
        if not placed:
            rows.append({"top": w["top"], "words": [w]})

    normalized_rows = []
    for row in rows:
        ws = sorted(row["words"], key=lambda x: x["x0"])
        text = " ".join(w["text"] for w in ws)
        normalized_rows.append(
            {
                "top": row["top"],
                "words": ws,
                "text": normalize_text(text),
            }
        )
    return normalized_rows


def is_product_header(line_text: str) -> bool:
    return bool(re.match(r"^I[0-9A-Z]+\s*-\s*[0-9A-Z.]+", line_text))


def parse_header(line_text: str):
    """
    Esempio:
    I030468 - 01.60 Landon Pant 100% Cotton 'Robertson' Denim, 12 oz Blue heavy stone wash
    """
    m = re.match(r"^(I[0-9A-Z]+)\s*-\s*([0-9A-Z.]+)\s+(.*)$", line_text)
    if not m:
        return None, None, None
    codice = m.group(1).strip()
    colore_raw = m.group(2).strip()
    descrizione = normalize_text(m.group(3))
    colore = colore_raw.replace(".", "")
    return codice, colore, descrizione


def extract_size_positions(taglia_line):
    """
    Restituisce lista di tuple (taglia, x0).
    Ignora la parola 'Taglia'.
    """
    sizes = []
    for w in taglia_line["words"]:
        txt = w["text"].strip()
        if txt.lower() == "taglia":
            continue
        sizes.append((txt, w["x0"]))
    return sizes


def extract_qty_positions(qta_line):
    """
    Restituisce lista di tuple (quantità, x0).
    Ignora la parola 'Quantità'.
    """
    qtys = []
    for w in qta_line["words"]:
        txt = w["text"].strip()
        if txt.lower() == "quantità":
            continue
        if re.fullmatch(r"\d+", txt):
            qtys.append((int(txt), w["x0"]))
    return qtys


def map_quantities_to_sizes(sizes, qtys, max_distance=25):
    """
    Abbina ogni quantità alla taglia con x più vicino.
    Questo funziona bene su PDF come il tuo, dove le quantità sono allineate
    sotto le taglie anche se i vuoti non vengono estratti come testo.
    """
    result = {size: 0 for size, _ in sizes}
    if not sizes or not qtys:
        return result

    size_positions = [(size, x) for size, x in sizes]

    for qty, qx in qtys:
        nearest_size = None
        nearest_dist = None

        for size, sx in size_positions:
            dist = abs(sx - qx)
            if nearest_dist is None or dist < nearest_dist:
                nearest_dist = dist
                nearest_size = size

        if nearest_size is not None and nearest_dist is not None and nearest_dist <= max_distance:
            result[nearest_size] += qty

    return result


def parse_product_block(lines):
    """
    lines = righe del singolo blocco prodotto
    """
    header_line = lines[0]["text"]
    codice, colore, descrizione = parse_header(header_line)
    if not codice:
        return None

    prezzo_whs = None
    prezzo_rtl = None
    taglia_line = None
    qta_line = None

    for line in lines:
        text = line["text"]

        if "Prezzo" in text and "EUR" in text:
            prezzo_whs, prezzo_rtl = parse_price_line(text)

        if text.startswith("Taglia"):
            taglia_line = line

        if text.startswith("Quantità"):
            qta_line = line

    sizes = extract_size_positions(taglia_line) if taglia_line else []
    qtys = extract_qty_positions(qta_line) if qta_line else []
    mapped = map_quantities_to_sizes(sizes, qtys)

    record = {
        "CODICE": codice,
        "COLORE": colore,
        "DESCRIZIONE": descrizione,
        "PREZZO WHS": prezzo_whs,
        "PREZZO RTL": prezzo_rtl,
    }

    for size, qty in mapped.items():
        if qty:
            record[str(size)] = qty

    return record


def parse_pdf(file_obj):
    records = []

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=True,
            )
            lines = to_lines(words)

            product_start_indexes = [
                idx for idx, line in enumerate(lines)
                if is_product_header(line["text"])
            ]

            for i, start_idx in enumerate(product_start_indexes):
                end_idx = product_start_indexes[i + 1] if i + 1 < len(product_start_indexes) else len(lines)
                block_lines = lines[start_idx:end_idx]

                product = parse_product_block(block_lines)
                if product:
                    records.append(product)

    return records


def build_dataframe(all_records):
    if not all_records:
        return pd.DataFrame()

    static_cols = ["CODICE", "COLORE", "DESCRIZIONE", "PREZZO WHS", "PREZZO RTL"]

    size_cols = set()
    for r in all_records:
        for key in r.keys():
            if key not in static_cols:
                size_cols.add(key)

    def size_sort_key(val):
        if re.fullmatch(r"\d+", val):
            return (0, int(val))
        return (1, val)

    ordered_size_cols = sorted(size_cols, key=size_sort_key)

    rows = []
    for r in all_records:
        row = {col: r.get(col, "") for col in static_cols}
        for size in ordered_size_cols:
            row[size] = r.get(size, "")
        rows.append(row)

    return pd.DataFrame(rows)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ORDINE")
        ws = writer.book["ORDINE"]

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                cell_value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(cell_value))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    output.seek(0)
    return output.getvalue()


uploaded_files = st.file_uploader(
    "Carica PDF",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_records = []

    with st.spinner("Sto leggendo il PDF e creando l'Excel..."):
        for uploaded_file in uploaded_files:
            try:
                records = parse_pdf(uploaded_file)
                all_records.extend(records)
            except Exception as e:
                st.error(f"Errore su {uploaded_file.name}: {e}")

    df = build_dataframe(all_records)

    if df.empty:
        st.warning("Non sono riuscito a trovare prodotti nel PDF.")
    else:
        st.success(f"Prodotti estratti: {len(df)}")
        st.dataframe(df, use_container_width=True)

        excel_bytes = dataframe_to_excel_bytes(df)

        st.download_button(
            label="Scarica Excel",
            data=excel_bytes,
            file_name="ordine_trasposto.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
