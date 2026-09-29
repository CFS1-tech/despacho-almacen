"""Excel descargables (pandas + openpyxl)."""
from io import BytesIO

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from logic import paradas
from schema import BASE_COLS, PEDIDO
from utils import a_fecha, num, telefono_de

NUMERICAS = set(PEDIDO) - {"pedido_gr", "n_orden"} | {"viaje"}


def _formatear(ws, anchos):
    azul = PatternFill("solid", fgColor="1F4FB8")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = azul
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _a_excel(hojas: list) -> bytes:
    """hojas = [(nombre, DataFrame, anchos, columnas_fecha)]"""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nombre, df, anchos, fechas in hojas:
            df.to_excel(xw, sheet_name=nombre, index=False)
            ws = xw.sheets[nombre]
            for col in fechas:
                j = list(df.columns).index(col) + 1
                for fila in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    fila[0].number_format = "DD/MM/YYYY"
            _formatear(ws, anchos)
    return buf.getvalue()


def excel_base(df: pd.DataFrame) -> bytes:
    out = pd.DataFrame()
    for k, h in BASE_COLS:
        if k == "fecha_entrega":
            out[h] = df[k].map(a_fecha)
        elif k in NUMERICAS:
            out[h] = df[k].map(lambda v: num(v) if num(v) is not None else "")
        else:
            out[h] = df[k]
    anchos = [12, 22, 28, 9, 9, 8, 16, 18, 16, 16, 36, 26, 14, 22, 14, 12, 8, 8, 8, 8, 8, 8, 9, 40, 14, 13, 14, 18,
              16, 6, 14, 12]
    return _a_excel([("BASE DE ENTREGAS", out, anchos, ["FECHA ENTREGA"])])


def excel_hoja_choferes(rutas: pd.DataFrame, sols: pd.DataFrame) -> bytes:
    s = sols.set_index("id")
    filas = []
    for _, r in rutas.iterrows():
        for n, i in enumerate(paradas(r), 1):
            if i not in s.index:
                continue
            x = s.loc[i]
            filas.append({
                "CÓDIGO RUTA": r["codigo"], "CHOFER": r["chofer"], "VIAJE": num(r["viaje"]) or 1,
                "T  TRANSPORTE": r["t_transporte"], "PLACA": r["placa"], "ORDEN": n,
                "FECHA ENTREGA": a_fecha(x["fecha_entrega"]), "PROCESO": x["proceso"], "CUENTA": x["cuenta"],
                "CLIENTE": x["cliente"], "N°ORDEN": x["n_orden"], "PEDIDO - GR": x["pedido_gr"],
                "HORA DE CITA": x["hora_cita"], "DIRECCION": x["direccion"],
                "CONTACTO": x["contacto"].replace("\n", " / "), "TELEFONO": telefono_de(x["contacto"]),
                "REFERENCIA": x["referencia"], "CANT DE CAJA": num(x["cant_caja"]) or "",
                "PALET'S": num(x["palets"]) or "", "OBSERVACION": x["observacion"], "CÓDIGO DESPACHO": x["codigo"],
            })
    df = pd.DataFrame(filas)
    anchos = [16, 16, 6, 14, 9, 6, 12, 8, 20, 26, 10, 12, 9, 36, 26, 13, 22, 8, 7, 30, 16]
    return _a_excel([("CONDUCTORES", df, anchos, ["FECHA ENTREGA"] if not df.empty else [])])
