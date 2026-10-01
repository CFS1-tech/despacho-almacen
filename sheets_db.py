"""Backend Google Sheets (se usa automáticamente si existe st.secrets["gcp_oauth"]).

Autenticación OAuth con tu propia cuenta (client_id + client_secret + refresh_token),
igual que en Picking Subcedis. Cada tabla de schema.TABLES es una pestaña del Sheet.

Cuidado con los códigos: se escribe con value_input_option="RAW" y las columnas se
formatean como TEXTO, así Google Sheets no convierte "000123" en 123. Al leer se usa
get_all_values(), que siempre devuelve texto.
"""
import gspread
import pandas as pd
import streamlit as st
from google.oauth2.credentials import Credentials
from gspread.utils import rowcol_to_a1

from schema import TABLES

BACKEND = "Google Sheets"


@st.cache_resource(show_spinner=False)
def _book():
    cfg = st.secrets["gcp_oauth"]
    creds = Credentials(
        token=None,
        refresh_token=cfg["refresh_token"],
        client_id=cfg["client_id"],
        client_secret=cfg["client_secret"],
        token_uri="https://oauth2.googleapis.com/token",
    )
    gc = gspread.authorize(creds)
    sheet_id = (st.secrets.get("sheet_id") or st.secrets.get("spreadsheet_id")
                or cfg.get("sheet_id") or cfg.get("spreadsheet_id"))
    if not sheet_id:
        raise RuntimeError("Falta sheet_id en los Secrets.")
    return gc.open_by_key(sheet_id)


_ws_cache = {}


def _ws(tabla: str):
    if tabla in _ws_cache:
        return _ws_cache[tabla]
    book = _book()
    cols = TABLES[tabla]
    try:
        ws = book.worksheet(tabla)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=tabla, rows=1000, cols=max(len(cols), 10))
        _formatear_texto(ws, len(cols))
        ws.update(range_name="A1", values=[cols], value_input_option="RAW")
        ws.freeze(rows=1)
    _ws_cache[tabla] = ws
    return ws


def _formatear_texto(ws, ncols: int):
    ultima = rowcol_to_a1(1, ncols).rstrip("0123456789")
    ws.format(f"A:{ultima}", {"numberFormat": {"type": "TEXT"}})


@st.cache_resource(show_spinner="Preparando el Google Sheet…")
def init():
    """Crea las pestañas que falten y agrega columnas nuevas al final del encabezado."""
    for tabla, cols in TABLES.items():
        ws = _ws(tabla)
        encabezado = ws.row_values(1)
        faltan = [c for c in cols if c not in encabezado]
        if faltan:
            nuevo = encabezado + faltan
            if ws.col_count < len(nuevo):
                ws.add_cols(len(nuevo) - ws.col_count)
            ws.update(range_name="A1", values=[nuevo], value_input_option="RAW")
            _formatear_texto(ws, len(nuevo))
    return True


@st.cache_data(ttl=30, show_spinner=False)
def _valores(tabla: str):
    return _ws(tabla).get_all_values()


def clear_cache():
    _valores.clear()


def _a_df(tabla, valores):
    cols = TABLES[tabla]
    if not valores:
        return pd.DataFrame(columns=cols)
    enc = valores[0]
    filas = [f + [""] * (len(enc) - len(f)) for f in valores[1:] if any(str(x).strip() for x in f)]
    df = pd.DataFrame(filas, columns=enc) if filas else pd.DataFrame(columns=enc)
    return df.reindex(columns=cols).fillna("").astype(str)


def read(tabla: str) -> pd.DataFrame:
    return _a_df(tabla, _valores(tabla))


def _encabezado(ws):
    return ws.row_values(1)


def append(tabla: str, filas: list):
    if not filas:
        return
    ws = _ws(tabla)
    enc = _encabezado(ws)
    datos = [[("" if f.get(k) is None else str(f.get(k))) for k in enc] for f in filas]
    ws.append_rows(datos, value_input_option="RAW", insert_data_option="INSERT_ROWS", table_range="A1")
    clear_cache()


def update_many(tabla: str, cambios: dict):
    """Actualiza varias filas (por id) en una sola llamada a la API."""
    if not cambios:
        return
    ws = _ws(tabla)
    valores = ws.get_all_values()  # lectura fresca, sin caché
    enc = valores[0]
    idx_id = enc.index("id")
    fila_de = {f[idx_id]: i + 1 for i, f in enumerate(valores) if i > 0 and len(f) > idx_id}
    data = []
    for id_, campos in cambios.items():
        n = fila_de.get(str(id_))
        if not n:
            continue
        actual = valores[n - 1] + [""] * (len(enc) - len(valores[n - 1]))
        for k, v in campos.items():
            if k in enc and k != "id":
                actual[enc.index(k)] = "" if v is None else str(v)
        data.append({"range": f"A{n}:{rowcol_to_a1(n, len(enc))}", "values": [actual[: len(enc)]]})
    if data:
        ws.batch_update(data, value_input_option="RAW")
    clear_cache()


def update(tabla: str, id_: str, campos: dict):
    update_many(tabla, {id_: campos})


def delete(tabla: str, id_: str):
    ws = _ws(tabla)
    valores = ws.get_all_values()
    idx_id = valores[0].index("id")
    for i, f in enumerate(valores[1:], start=2):
        if len(f) > idx_id and f[idx_id] == str(id_):
            ws.delete_rows(i)
            break
    clear_cache()


def replace_all(tabla: str, df: pd.DataFrame):
    """Reescribe toda la pestaña (se usa en maestros y para cargar el histórico)."""
    ws = _ws(tabla)
    cols = TABLES[tabla]
    df = df.reindex(columns=cols).fillna("").astype(str)
    filas = [cols] + df.values.tolist()
    if ws.row_count < len(filas) + 10:
        ws.add_rows(len(filas) + 10 - ws.row_count)
    ws.clear()
    _formatear_texto(ws, len(cols))
    ws.update(range_name="A1", values=filas, value_input_option="RAW")
    clear_cache()
