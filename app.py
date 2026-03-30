import io
import os
import re

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(page_title="PDF to Excel Transposer", layout="wide")
st.title("PDF to Excel - Trasposizione ordini")
st.write(
    "Carica uno o più PDF ordine. "
    "Lo script estrae CODICE, COLORE, DESCRIZIONE, PREZZO WHS, PREZZO RTL "
    "e le quantità per taglia, poi genera un file Excel."
)


STATIC_COLS = ["CODICE", "COLORE", "DESCRIZIONE", "PREZZO WHS", "PREZZO RTL"]

ALPHA_SIZE_ORDER = {
    "XXS": 0,
    "XS": 1,
    "S": 2,
    "M": 3,
    "L": 4,
    "XL": 5,
    "XXL": 6,
    "XS-S": 7,
    "S-M": 8,
    "M-L": 9,
    "L-XL": 10,
    "XL-XXL": 11,
}


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_size(size: str) -> str:
    size = normalize_text(size).upper()

    if size == "0":
        return "UNICA"

    # Normalizza trattini diversi e spazi
    size = size.replace("–", "-").replace("—", "-")
    size = re.sub(r"\s*-\s*", "-", size)

    # Normalizza slash in dash: S/M -> S-M
    size = size.replace("/", "-")
    size = re.sub(r"\s*-\s*", "-", size)

    return size


def parse_price_line(text: str):
    match = re.search(r"Prezzo\s+EUR\s+([\d.,]+)\s*/\s*([\d.,]+)", text, re.IGNORECASE)
    if not match:
        return "", ""
    return match.group(1), match.group(2)


def parse_header(line_text: str):
    match = re.match(r"^(I[0-9A-Z]+)\s*-\s*([0-9A-Z.]+)\s+(.*)$", line_text)
    if not match:
        return None, None, None

    codice = match.group(1).strip()
    colore_raw = match.group(2).strip()
    descrizione = normalize_text(match.group(3))
    colore = colore_raw.replace(".", "")

    return codice, colore, descrizione


def extract_order_confirmation_number(file_obj):
    """
    Estrae il numero 'Conferma dell'ordine' dal PDF, es. 26VA028707
    """
    file_obj.seek(0)

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            text = normalize_text(text)

            match = re.search(r"Conferma dell['’]ordine\s+([A-Z0-9]+)", text, re.IGNORECASE)
            if match:
                file_obj.seek(0)
                return match.group(1).strip()

    file_obj.seek(0)
    return ""


def build_output_filename(uploaded_files, order_confirmation_number):
    """
    Se c'è un solo PDF:
    nome originale + _numero conferma + .xlsx

    Se ci sono più PDF:
    export_trasposto + .xlsx
    """
    if len(uploaded_files) == 1:
        original_name = uploaded_files[0].name
        base_name = os.path.splitext(original_name)[0]

        if order_confirmation_number:
            return f"{base_name}_{order_confirmation_number}.xlsx"
        return f"{base_name}.xlsx"

    if order_confirmation_number:
        return f"export_trasposto_{order_confirmation_number}.xlsx"

    return "export_trasposto.xlsx"


def is_product_header(line_text: str) -> bool:
    return bool(re.match(r"^I[0-9A-Z]+\s*-\s*[0-9A-Z.]+", line_text))


def to_lines(words, y_tol=3):
    """
    Raggruppa le words di pdfplumber in righe usando la coordinata verticale.
    """
    rows = []

    for word in sorted(words, key=lambda x: (round(x["top"], 1), x["x0"])):
        placed = False
        for row in rows:
            if abs(row["top"] - word["top"]) <= y_tol:
                row["words"].append(word)
                placed = True
                break

        if not placed:
            rows.append({"top": word["top"], "words": [word]})

    normalized_rows = []
    for row in rows:
        sorted_words = sorted(row["words"], key=lambda x: x["x0"])
        text = " ".join(w["text"] for w in sorted_words)
        normalized_rows.append(
            {
                "top": row["top"],
                "words": sorted_words,
                "text": normalize_text(text),
            }
        )

    return normalized_rows


def extract_size_positions(taglia_line):
    """
    Restituisce lista di tuple (taglia, x0), ignorando la parola 'Taglia'.
    """
    sizes = []
    if not taglia_line:
        return sizes

    for word in taglia_line["words"]:
        txt = normalize_text(word["text"])
        if txt.lower() == "taglia":
            continue
        sizes.append((normalize_size(txt), word["x0"]))

    return sizes


def extract_qty_positions(qta_line):
    """
    Restituisce lista di tuple (quantità, x0), ignorando la parola 'Quantità'.
    Esclude numeri troppo grandi per evitare di prendere riepiloghi tipo 103.
    """
    qtys = []
    if not qta_line:
        return qtys

    for word in qta_line["words"]:
        txt = normalize_text(word["text"])

        if txt.lower() == "quantità":
            continue

        if re.fullmatch(r"\d+", txt):
            val = int(txt)
            if val < 50:
                qtys.append((val, word["x0"]))

    return qtys


def map_quantities_to_sizes(sizes, qtys, max_distance=35):
    """
    Abbina le quantità alle taglie per vicinanza orizzontale.
    Caso speciale:
    se c'è una sola taglia nel blocco, assegna tutta la quantità a quella taglia.
    """
    result = {size: 0 for size, _ in sizes}

    if not sizes or not qtys:
        return result

    if len(sizes) == 1:
        only_size = sizes[0][0]
        result[only_size] = sum(qty for qty, _ in qtys)
        return result

    for qty, qx in qtys:
        nearest_size = None
        nearest_dist = None

        for size, sx in sizes:
            dist = abs(sx - qx)
            if nearest_dist is None or dist < nearest_dist:
                nearest_dist = dist
                nearest_size = size

        if nearest_size is not None and nearest_dist is not None and nearest_dist <= max_distance:
            result[nearest_size] += qty

    return result


def parse_product_block(lines):
    """
    Estrae un singolo prodotto dal blocco di righe.
    """
    if not lines:
        return None

    header_line = lines[0]["text"]
    codice, colore, descrizione = parse_header(header_line)
    if not codice:
        return None

    prezzo_whs = ""
    prezzo_rtl = ""
    taglia_line = None
    qta_line = None

    for line in lines:
        text = line["text"]

        if "Prezzo" in text and "EUR" in text:
            prezzo_whs, prezzo_rtl = parse_price_line(text)

        if text.startswith("Taglia"):
            taglia_line = line

        if text.startswith("Quantità") and "totale" not in text.lower():
            qta_line = line

        # Bonus: ferma il parsing prima del riepilogo del blocco
        if text.startswith("Totale"):
            break

    sizes = extract_size_positions(taglia_line)
    qtys = extract_qty_positions(qta_line)
    mapped = map_quantities_to_sizes(sizes, qtys, max_distance=35)

    record = {
        "CODICE": codice,
        "COLORE": colore,
        "DESCRIZIONE": descrizione,
        "PREZZO WHS": prezzo_whs,
        "PREZZO RTL": prezzo_rtl,
    }

    for size, qty in mapped.items():
        if qty:
            record[size] = qty

    return record


def parse_pdf(file_obj):
    records = []

    file_obj.seek(0)
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
                end_idx = (
                    product_start_indexes[i + 1]
                    if i + 1 < len(product_start_indexes)
                    else len(lines)
                )
                block_lines = lines[start_idx:end_idx]

                product = parse_product_block(block_lines)
                if product:
                    records.append(product)

    file_obj.seek(0)
    return records


def size_sort_key(value: str):
    value = normalize_size(value)

    # 1. UNICA
    if value == "UNICA":
        return (0, 0)

    # 2. numeriche
    if re.fullmatch(r"\d+", value):
        return (1, int(value))

    # 3. alfabetiche/custom
    if value in ALPHA_SIZE_ORDER:
        return (2, ALPHA_SIZE_ORDER[value])

    # 4. fallback
    return (3, value)


def build_dataframe(all_records):
    if not all_records:
        return pd.DataFrame()

    size_cols = set()
    for record in all_records:
        for key in record.keys():
            if key not in STATIC_COLS:
                size_cols.add(key)

    ordered_size_cols = sorted(size_cols, key=size_sort_key)

    rows = []
    for record in all_records:
        row = {col: record.get(col, "") for col in STATIC_COLS}
        for size_col in ordered_size_cols:
            row[size_col] = record.get(size_col, "")
        rows.append(row)

    return pd.DataFrame(rows)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ORDINE")
        ws = writer.book["ORDINE"]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                cell_value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(cell_value))

            ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    output.seek(0)
    return output.getvalue()


def calculate_total_qty(df: pd.DataFrame) -> int:
    total_qty = 0
    for col in df.columns:
        if col not in STATIC_COLS:
            total_qty += pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
    return int(total_qty)


uploaded_files = st.file_uploader(
    "Carica PDF",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_records = []
    order_confirmation_number = ""

    with st.spinner("Sto leggendo il PDF e creando l'Excel..."):
        for uploaded_file in uploaded_files:
            try:
                if not order_confirmation_number:
                    order_confirmation_number = extract_order_confirmation_number(uploaded_file)

                uploaded_file.seek(0)
                records = parse_pdf(uploaded_file)
                all_records.extend(records)

            except Exception as exc:
                st.error(f"Errore su {uploaded_file.name}: {exc}")

    df = build_dataframe(all_records)

    if df.empty:
        st.warning("Non sono riuscito a trovare prodotti nel PDF.")
    else:
        total_qty = calculate_total_qty(df)
        output_filename = build_output_filename(uploaded_files, order_confirmation_number)

        st.success(f"Prodotti estratti: {len(df)}")
        st.info(f"Totale quantità estratte: {total_qty}")

        if order_confirmation_number:
            st.info(f"Conferma ordine rilevata: {order_confirmation_number}")

        st.dataframe(df, use_container_width=True)

        excel_bytes = dataframe_to_excel_bytes(df)

        st.download_button(
            label="Scarica Excel",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
