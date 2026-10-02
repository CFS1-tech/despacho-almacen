"""Distribución CargoFlex Supply — Solicitud → Ruteo → Despacho → Base de entregas.

Streamlit ejecuta este archivo de arriba a abajo en cada interacción. La sección
activa se guarda en st.session_state["seccion"]; cada sección es una función.
"""
from datetime import date, timedelta

import re

import pandas as pd
import streamlit as st

import logic as L
from parser_plantilla import df_choferes_defecto, df_tipos_defecto, leer_plantilla
from report import excel_base, excel_hoja_choferes, excel_plantilla
from schema import CATEGORIAS, ESTADOS, PEDIDO, PEDIDO_LABELS, PEDIDO_TEXTO, ROLES, TIPOS_CHOFER
from utils import (ALMACEN, _hora_actual, ahora_str, es_activo, fmt_fecha, hora_str, hoy, norm, nuevo_id, num, num0,
                   slug, txt)

st.set_page_config(page_title="Distribución CargoFlex Supply", page_icon="🚚", layout="wide")


# ================================================================ backend
def _usar_sheets() -> bool:
    try:
        return "gcp_oauth" in st.secrets
    except Exception:
        return False


if _usar_sheets():
    import sheets_db as store
else:
    import db as store


@st.cache_resource(show_spinner=False)
def _init_sqlite():
    store.init()
    return True


if store.BACKEND == "SQLite local":
    _init_sqlite()
else:
    try:
        store.init()
    except Exception as e:  # muestra el motivo real (Streamlit Cloud oculta el mensaje original)
        st.error(f"No se pudo conectar con el Google Sheet: {type(e).__name__}: {e}")
        st.info("Revisa en Settings → Secrets: sheet_id y los valores de [gcp_oauth] "
                "(client_id, client_secret, refresh_token).")
        st.stop()

ss = st.session_state


def flash(msg, tipo="success"):
    ss["_flash"] = (tipo, msg)


def mostrar_flash():
    if "_flash" in ss:
        tipo, msg = ss.pop("_flash")
        getattr(st, tipo)(msg)


# ================================================================ estilos
CSS = """
<style>
:root{--navy:#0e3b2c;--azul:#1f6b45;--azul2:#2e8b57;--lima:#a4cb4c;--rojo:#e0322f;--tinta:#17231d;--gris:#66756c;--linea:#e1e8e3;--fondo:#f3f6f4}
.stApp{background:var(--fondo)}
h1,h2,h3,h4{color:var(--navy)!important;letter-spacing:-.01em}
div[data-testid="stVerticalBlockBorderWrapper"]{background:#fff;border-radius:12px!important;border-color:var(--linea)!important;
  box-shadow:0 1px 2px rgba(16,42,74,.04),0 4px 14px rgba(16,42,74,.05)}
.stButton>button[kind="primary"],.stDownloadButton>button[kind="primary"],.stFormSubmitButton>button[kind="primary"]{
  background:linear-gradient(135deg,var(--azul),var(--azul2));border:0}
.wms-hero{max-width:440px;margin:48px auto 22px;border-radius:16px;padding:34px 32px 30px;text-align:center;color:#fff;
  background:linear-gradient(140deg,#071f17 0%,#0e3b2c 50%,#1f6b45 100%);box-shadow:0 18px 40px rgba(7,31,23,.28);position:relative;overflow:hidden}
.wms-hero:after{content:"";position:absolute;left:0;right:0;bottom:0;height:5px;background:linear-gradient(90deg,var(--lima) 0 70%,var(--rojo) 70% 100%)}
.wms-logo{background:#fff;border-radius:12px;padding:14px 18px;display:inline-block;box-shadow:0 6px 18px rgba(0,0,0,.18)}
.wms-logo img{width:230px;max-width:100%;display:block}
.wms-hero h1{color:#fff!important;font-size:22px;margin:22px 0 4px;font-weight:700}
.wms-hero small{display:block;color:var(--lima);font-size:12px;font-weight:700;letter-spacing:.14em;text-transform:uppercase}
.wms-hero p{color:#cfe3d6;font-size:13px;margin:12px 0 0}
.app-head{display:flex;align-items:center;gap:16px}
.app-head img{height:44px}
.app-head .t{font-size:13px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--azul);border-left:3px solid var(--lima);padding-left:12px;line-height:1.25}
.app-head .t b{display:block;font-size:20px;letter-spacing:-.01em;text-transform:none;color:var(--navy)}
section[data-testid="stSidebar"]{background:#0e3b2c}
section[data-testid="stSidebar"] *{color:#e8f1ec}
section[data-testid="stSidebar"] .stButton>button{background:transparent;border-color:#3d6b57}
section[data-testid="stSidebar"] hr{border-color:#2c5a46}
.pill{display:inline-block;font-size:11px;font-weight:700;border-radius:20px;padding:2px 10px;letter-spacing:.03em;white-space:nowrap}
.st-pendiente{background:#fff3dc;color:#9a5b00}.st-ruteado{background:#e3ebfb;color:#1f4fb8}.st-en_ruta{background:#ece6fb;color:#5b3fb0}
.st-entregado{background:#def3e7;color:#1d7a4c}.st-no_entregado{background:#fbe1de;color:#b3261e}
.proc{display:inline-block;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px;font-weight:700;border-radius:4px;padding:1px 6px}
.proc-OUT{background:#e3ebfb;color:#1f4fb8}.proc-IN{background:#def3e7;color:#1d7a4c}
.sol-cod{font-family:ui-monospace,Menlo,Consolas,monospace;font-weight:700;font-size:12px;color:var(--navy);white-space:nowrap}
.sol-grp{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:10.5px;color:var(--gris);white-space:nowrap}
.sol-sub{color:var(--gris);font-size:12px;line-height:1.35}
.sol-tit{font-weight:600;color:var(--tinta);font-size:14px;line-height:1.3}
.sol-fecha{font-weight:600;color:var(--tinta);font-size:13px}
.kpi-mini{display:flex;gap:10px;flex-wrap:wrap;margin:2px 0 10px}
.kpi-mini span{background:#fff;border:1px solid var(--linea);border-radius:20px;padding:3px 12px;font-size:12px;color:#4a5a66}
.kpi-mini b{color:var(--navy)}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ================================================================ marca
import base64
import os

MARCA = "Distribución CargoFlex Supply"


@st.cache_resource
def _logo_b64():
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    if not os.path.exists(ruta):
        return ""
    with open(ruta, "rb") as f:
        return base64.b64encode(f.read()).decode()


def logo_html(clase=""):
    b = _logo_b64()
    return f'<img class="{clase}" src="data:image/png;base64,{b}" alt="CargoFlex Supply">' if b else "<b>CargoFlex Supply</b>"


# ================================================================ login
def _dni_ok(v: str) -> bool:
    return v.isdigit() and 8 <= len(v) <= 12


def pantalla_login():
    st.markdown("<style>section[data-testid='stSidebar']{display:none}"
                "div[data-testid='stTextInput'] label p{font-size:11px!important;font-weight:700!important;"
                "letter-spacing:.08em;text-transform:uppercase;color:#4a5a50!important}</style>",
                unsafe_allow_html=True)
    usuarios = store.read("USUARIOS")
    primera = usuarios.empty
    st.markdown(f"""<div class="wms-hero"><div class="wms-logo">{logo_html()}</div>
        <h1>Sistema de Distribución</h1><small>CargoFlex Supply</small>
        <p>{"Primera vez: registra al administrador" if primera else "Ingresa tu DNI para continuar"}</p></div>""",
                unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 1.4, 1])
    with centro, st.container(border=True):
        if primera:
            with st.form("f_admin", border=False):
                u = st.text_input("DNI", placeholder="Número de DNI", max_chars=12).strip()
                n = st.text_input("Nombre completo")
                if st.form_submit_button("Crear administrador", type="primary", width="stretch"):
                    if not _dni_ok(u):
                        st.error("Escribe un DNI válido (solo números, 8 dígitos).")
                    else:
                        store.append("USUARIOS", [{"id": u, "nombre": n or u, "clave_hash": "", "roles": "admin",
                                                   "activo": "TRUE", "creado": ahora_str()}])
                        ss["usuario"] = u
                        st.rerun()
        else:
            with st.form("f_login", border=False):
                u = st.text_input("DNI", placeholder="Ingresa tu número de DNI", max_chars=12).strip()
                if st.form_submit_button("Ingresar al sistema", type="primary", width="stretch"):
                    fila = usuarios[usuarios["id"] == u]
                    if fila.empty or not es_activo(fila.iloc[0]["activo"]):
                        st.error("Ese DNI no está registrado o está inactivo. Pide al administrador que te dé acceso.")
                    else:
                        ss["usuario"] = u
                        st.rerun()


if "usuario" not in ss:
    pantalla_login()
    st.stop()

_u = store.read("USUARIOS")
_u = _u[_u["id"] == ss["usuario"]]
if _u.empty or not es_activo(_u.iloc[0]["activo"]):
    ss.clear()
    st.rerun()
USUARIO = ss["usuario"]
NOMBRE = _u.iloc[0]["nombre"] or USUARIO
MIS_ROLES = {r.strip() for r in _u.iloc[0]["roles"].split(",") if r.strip()}


def tiene(*roles) -> bool:
    return "admin" in MIS_ROLES or any(r in MIS_ROLES for r in roles)


# ================================================================ navegación
SECCIONES = [
    ("solicitudes", "📝 Solicitud de transporte", ("solicitante", "ruteador", "supervisor")),
    ("regularizacion", "🧾 Regularización de despacho", ("solicitante", "ruteador", "almacen", "supervisor")),
    ("movcliente", "🏷️ Recepción / despacho del cliente", ("solicitante", "ruteador", "almacen", "supervisor")),
    ("ruteo", "🗺️ Ruteo", ("ruteador",)),
    ("despacho", "🚚 Despacho", ("almacen", "ruteador")),
    ("reportes", "📊 Base y reportes", ("supervisor", "ruteador")),
    ("maestros", "🗂️ Maestros", ("supervisor",)),
    ("usuarios", "👥 Usuarios", ()),
]
visibles = [s for s in SECCIONES if tiene(*s[2])]
if not visibles:
    st.warning("Tu usuario no tiene roles asignados. Pide a un administrador que te asigne uno.")
    if st.button("Salir"):
        ss.clear()
        st.rerun()
    st.stop()
claves = [s[0] for s in visibles]
if ss.get("seccion") not in claves:
    ss["seccion"] = claves[0]

with st.sidebar:
    st.markdown(f"<div style='background:#fff;border-radius:10px;padding:10px 12px;margin-bottom:14px'>{logo_html()}</div>",
                unsafe_allow_html=True)
    st.markdown(f"**{NOMBRE}**  \n"
                + " · ".join(ROLES.get(r, r) for r in sorted(MIS_ROLES)))
    st.radio("Ir a", claves, format_func=lambda k: dict((s[0], s[1]) for s in SECCIONES)[k], key="seccion",
             label_visibility="collapsed")
    st.divider()
    if st.button("🔄 Actualizar datos"):
        store.clear_cache()
        st.rerun()
    if st.button("Cerrar sesión"):
        ss.clear()
        st.rerun()
    st.caption(f"Datos: {store.BACKEND}")


def ir_a(sec):
    ss["seccion"] = sec


c1, c2, c3 = st.columns([5, 2, 2])
c1.markdown(f"<div class='app-head'>{logo_html()}<div class='t'>Sistema de<b>Distribución</b></div></div>",
            unsafe_allow_html=True)
if "regularizacion" in claves:
    c2.button("🧾 Regularización de despacho", width="stretch", on_click=ir_a, args=("regularizacion",),
              type="primary" if ss["seccion"] == "regularizacion" else "secondary")
if "movcliente" in claves:
    c3.button("🏷️ Registro de recepción / despacho", width="stretch", on_click=ir_a, args=("movcliente",),
              type="primary" if ss["seccion"] == "movcliente" else "secondary")
mostrar_flash()


# ================================================================ datos comunes
def cuentas_act():
    return L.activos(store.read("CUENTAS")).sort_values("nombre")


def tipos_act():
    return L.activos(store.read("TIPOS_TRANSPORTE")).sort_values("nombre")


def choferes_act():
    return L.activos(store.read("CHOFERES")).sort_values("nombre")


def clientes_de(cuenta_id):
    cl = L.activos(store.read("CLIENTES"))
    cl = cl[cl["cuenta_id"] == cuenta_id].copy()
    cl["_u"] = cl["usos"].map(num0)
    return cl.sort_values(["_u", "cliente"], ascending=[False, True])


def pill_estado(e):
    return ESTADOS.get(e, e)


# ================================================================ formulario compartido
NUEVO = "__nuevo__"


def _auto_od(p):
    """Origen/destino automáticos según proceso y cliente elegido."""
    lugar = ss.get(f"{p}_lugar", "")
    if ss.get(f"{p}_proc", "OUT") == "OUT":
        ss[f"{p}_ori"], ss[f"{p}_des"] = ALMACEN, lugar
    else:
        ss[f"{p}_ori"], ss[f"{p}_des"] = lugar or ("CLIENTE" if p == "m" else ""), ALMACEN


def _on_cuenta(p):
    ss[f"{p}_cli"] = None
    ss[f"{p}_lugar"] = ""
    _auto_od(p)


def _on_cliente(p):
    cid = ss.get(f"{p}_cli")
    ss[f"{p}_lugar"] = ""
    if cid and cid != NUEVO:
        cl = store.read("CLIENTES").set_index("id")
        if cid in cl.index:
            r = cl.loc[cid]
            ss[f"{p}_lugar"] = r["destino"]
            for campo, clave in (("direccion", "dir"), ("referencia", "ref"), ("contacto", "contacto")):
                if f"{p}_{clave}" in ss and not ss.get(f"{p}_{clave}") and r[campo]:
                    ss[f"{p}_{clave}"] = r[campo]
    _auto_od(p)


def _defaults(p):
    base = {"proc": "IN" if p == "m" else "OUT", "cuenta": None, "cli": None, "cli_nuevo": "", "dir": "", "ref": "",
            "ori": "", "des": "", "fecha": hoy(), "hora": None, "lleg": None, "tipo": None, "chofer": None, "placa": "",
            "contacto": "", "emp": "", "cond": "", "motivo": "", "obs": "", "lugar": ""}
    base.update({f"ped_{k}": ("" if k in PEDIDO_TEXTO else None) for k in PEDIDO})
    return base


def _init_form(p, limpiar=False):
    for k, v in _defaults(p).items():
        if limpiar or f"{p}_{k}" not in ss:
            if limpiar and k in ("cuenta", "fecha", "proc"):  # se mantienen para el siguiente registro
                continue
            ss[f"{p}_{k}"] = v
    if limpiar or not ss.get(f"{p}_ori"):
        _auto_od(p)


def form_movimiento(p):
    """p = 's' solicitud, 'g' regularización, 'm' recepción/despacho del cliente."""
    _init_form(p)
    es_m, es_g = p == "m", p == "g"
    opciones = ["IN", "OUT"] if es_m else ["OUT", "IN"]
    etiquetas = {"OUT": "OUT · Despacho", "IN": "IN · Recepción"}
    st.radio("Proceso", opciones, format_func=etiquetas.get, horizontal=True, key=f"{p}_proc", on_change=_auto_od,
             args=(p,))
    cu = cuentas_act()
    st.selectbox("Cuenta", cu["id"].tolist(), format_func=dict(zip(cu["id"], cu["nombre"])).get, index=None,
                 placeholder="Selecciona la cuenta…", key=f"{p}_cuenta", on_change=_on_cuenta, args=(p,))
    cuenta_id = ss.get(f"{p}_cuenta")
    cl = clientes_de(cuenta_id) if cuenta_id else pd.DataFrame(columns=["id", "cliente", "destino"])
    etiqueta_cli = {r["id"]: r["cliente"] + (f"  ·  {r['destino']}" if r["destino"] else "") for _, r in cl.iterrows()}
    etiqueta_cli[NUEVO] = "➕ Cliente nuevo (escribir nombre)"
    st.selectbox("Cliente", cl["id"].tolist() + [NUEVO], format_func=etiqueta_cli.get, index=None,
                 placeholder="Escribe para buscar…" if cuenta_id else "Primero elige la cuenta",
                 key=f"{p}_cli", on_change=_on_cliente, args=(p,), disabled=not cuenta_id)
    if ss.get(f"{p}_cli") == NUEVO:
        st.text_input("Nombre del cliente nuevo", key=f"{p}_cli_nuevo",
                      help="Se guardará en la lista de clientes de esta cuenta.")
    if not es_m:
        st.text_input("Dirección", key=f"{p}_dir")
        st.text_input("Referencia (opcional)", key=f"{p}_ref")
    a, b = st.columns(2)
    a.text_input("Origen", key=f"{p}_ori")
    b.text_input("Destino", key=f"{p}_des")
    a, b = st.columns(2)
    a.date_input("Fecha" if es_m else "Fecha de entrega", key=f"{p}_fecha", format="DD/MM/YYYY")
    b.time_input("Hora" if es_m else "Hora de cita", key=f"{p}_hora", step=timedelta(minutes=15))
    if es_g:
        a, b = st.columns(2)
        a.time_input("Hora de llegada", key=f"{p}_lleg", step=timedelta(minutes=5))
    if es_m:
        a, b = st.columns(2)
        a.text_input("Empresa de transporte", key=f"{p}_emp", placeholder="Transporte del cliente")
        b.text_input("Placa", key=f"{p}_placa")
        st.text_area("Conductor / persona que entrega o recoge", key=f"{p}_cond", placeholder="Nombre, DNI…",
                     height=80)
    else:
        ti = tipos_act()
        if not es_g:
            ti = ti[ti["categoria"] != "CLIENTE"]
        st.selectbox("T. transporte" + ("" if es_g else " (opcional)"), ti["nombre"].tolist(), index=None,
                     placeholder="Sin preferencia" if not es_g else "Selecciona…", key=f"{p}_tipo")
        if es_g:
            ch = choferes_act()
            a, b = st.columns(2)
            a.selectbox("Chofer / agencia", ch["nombre"].tolist(), index=None, key=f"{p}_chofer")
            b.text_input("Placa (opcional)", key=f"{p}_placa")
        st.text_area("Contacto", key=f"{p}_contacto", placeholder="Nombre, DNI, teléfono de quien recibe…",
                     height=80)
    st.number_input("Cant. de cajas / bultos (opcional)", min_value=0.0, step=1.0, key=f"{p}_ped_cant_caja",
                    format="%g")
    with st.expander("Más detalle del pedido: GR, orden, inner, palets, unidades… (opcional)"):
        cols = st.columns(3)
        for n, k in enumerate(k for k in PEDIDO if k != "cant_caja"):
            if k in PEDIDO_TEXTO:
                cols[n % 3].text_input(PEDIDO_LABELS[k], key=f"{p}_ped_{k}")
            else:
                cols[n % 3].number_input(PEDIDO_LABELS[k], min_value=0.0, step=1.0, key=f"{p}_ped_{k}", format="%g")
    if es_g:
        st.text_area("Motivo de la regularización", key=f"{p}_motivo", height=80)
    st.text_area("Observaciones", key=f"{p}_obs", height=80)
    boton = {"s": "Registrar solicitud", "g": "Registrar regularización", "m": "Registrar"}[p]
    st.button(boton, type="primary", on_click=_guardar_movimiento, args=(p,), key=f"{p}_btn")


def _guardar_movimiento(p):
    cu = store.read("CUENTAS").set_index("id")
    cuenta_id = ss.get(f"{p}_cuenta")
    cid = ss.get(f"{p}_cli")
    if not cuenta_id:
        return flash("Elige la cuenta.", "error")
    if cid == NUEVO:
        cliente = txt(ss.get(f"{p}_cli_nuevo"))
    elif cid:
        cliente = store.read("CLIENTES").set_index("id").loc[cid, "cliente"]
    else:
        cliente = ""
    if not cliente:
        return flash("Elige o escribe el cliente.", "error")
    datos = {"proceso": ss[f"{p}_proc"], "cuenta_id": cuenta_id, "cuenta": cu.loc[cuenta_id, "nombre"],
             "cliente": cliente, "origen": txt(ss[f"{p}_ori"]).upper(), "destino": txt(ss[f"{p}_des"]).upper(),
             "observacion": txt(ss[f"{p}_obs"])}
    for k in PEDIDO:
        v = ss.get(f"{p}_ped_{k}")
        datos[k] = txt(v) if k in PEDIDO_TEXTO else ("" if v is None else str(num(v)))
    fecha = ss[f"{p}_fecha"].isoformat()
    hora = hora_str(ss.get(f"{p}_hora"))
    if p == "s":
        if not txt(ss["s_dir"]) or not hora or not txt(ss["s_contacto"]):
            return flash("Completa dirección, hora de cita y contacto.", "error")
        datos.update({"cliente_id": "" if cid == NUEVO else cid, "direccion": txt(ss["s_dir"]),
                      "referencia": txt(ss["s_ref"]), "fecha_entrega": fecha, "hora_cita": hora,
                      "t_transporte_sol": ss.get("s_tipo") or "", "contacto": txt(ss["s_contacto"])})
        tabla = "SOLICITUDES"
    elif p == "g":
        if not ss.get("g_tipo") or not txt(ss["g_motivo"]):
            return flash("Completa el tipo de transporte y el motivo.", "error")
        datos.update({"direccion": txt(ss["g_dir"]), "referencia": txt(ss["g_ref"]), "fecha_entrega": fecha,
                      "hora_cita": hora, "hora_llegada": hora_str(ss.get("g_lleg")), "t_transporte": ss["g_tipo"],
                      "chofer": ss.get("g_chofer") or "", "placa": txt(ss["g_placa"]).upper(),
                      "contacto": txt(ss["g_contacto"]), "motivo": txt(ss["g_motivo"])})
        tabla = "REGULARIZACIONES"
    else:
        if not hora:
            return flash("Indica la hora.", "error")
        datos.update({"fecha": fecha, "hora": hora, "empresa": txt(ss["m_emp"]), "placa": txt(ss["m_placa"]).upper(),
                      "conductor": txt(ss["m_cond"])})
        tabla = "MOV_CLIENTE"
    fila = L.crear_movimiento(store, tabla, datos, USUARIO)
    _init_form(p, limpiar=True)
    flash(f"Registrado con código **{fila['codigo']}**.")


def tabla_registros(df: pd.DataFrame, columnas: dict, alto=None):
    if df.empty:
        st.info("Aún no hay registros.")
        return
    st.dataframe(df[list(columnas)].rename(columns=columnas), hide_index=True, width="stretch",
                 height=alto)


# ================================================================ secciones
# ================================================================ solicitud de transporte (varios puntos)
PUNTO_COLS = {"cliente": "Cliente / punto de entrega", "direccion": "Dirección", "destino": "Destino (distrito)",
              "contacto": "Contacto"}
SIN_CLIENTE = "(Sin cliente)"
TIPO_FECHA = {"exacta": "En fecha exacta", "hasta": "Hasta una fecha límite"}
EDITABLES = ("pendiente", "ruteado")


def _txt_fecha(s) -> str:
    return ("Hasta " if s.get("tipo_fecha") == "hasta" else "") + fmt_fecha(s["fecha_entrega"])


def _sol_reset():
    for k in [k for k in ss.keys() if str(k).startswith("sol_")]:
        if k not in ("sol_cuenta", "sol_fecha", "sol_proc", "sol_tipo_fecha", "sol_n"):
            del ss[k]
    ss["sol_n"] = ss.get("sol_n", 0) + 1


def form_solicitud():
    if ss.pop("_sol_reset", False):
        _sol_reset()
    ss.setdefault("sol_proc", "OUT")
    ss.setdefault("sol_fecha", hoy())
    ss.setdefault("sol_tipo_fecha", "exacta")
    n = ss.get("sol_n", 0)

    # 1) fecha y hora de cita
    st.markdown("**Fecha y hora**")
    st.radio("Fecha de entrega", list(TIPO_FECHA), format_func=TIPO_FECHA.get, key="sol_tipo_fecha", horizontal=True,
             label_visibility="collapsed")
    a, b = st.columns(2)
    a.date_input("Fecha límite" if ss["sol_tipo_fecha"] == "hasta" else "Fecha de entrega", key="sol_fecha",
                 format="DD/MM/YYYY")
    b.time_input("Hora de cita (opcional)", value=None, key=f"sol_hora_{n}", step=timedelta(minutes=15))
    st.divider()

    # 2) proceso, cuenta y clientes
    st.radio("Proceso", ["OUT", "IN"], format_func={"OUT": "OUT · Despacho", "IN": "IN · Recepción"}.get,
             horizontal=True, key="sol_proc")
    cu = cuentas_act()
    cuenta_id = st.selectbox("Cuenta", cu["id"].tolist(), format_func=dict(zip(cu["id"], cu["nombre"])).get,
                             index=None, placeholder="Selecciona la cuenta…", key="sol_cuenta")
    cl = clientes_de(cuenta_id) if cuenta_id else pd.DataFrame(columns=["id", "cliente", "destino", "direccion"])
    etiqueta = {r["id"]: r["cliente"] + (f"  ·  {r['destino']}" if r["destino"] else "") for _, r in cl.iterrows()}
    sel = st.multiselect("Clientes / puntos de entrega (opcional)", cl["id"].tolist(),
                         format_func=lambda v: etiqueta.get(v, f"➕ {v} (nuevo)"),
                         accept_new_options=True, key=f"sol_sel_{cuenta_id}_{n}", disabled=not cuenta_id,
                         placeholder="Busca o escribe un cliente nuevo · puedes elegir varios" if cuenta_id
                         else "Primero elige la cuenta",
                         help="Cada cliente es un punto de entrega distinto: se registra una solicitud por punto. "
                              "Si no está en la lista, escribe el nombre y presiona Enter.")
    cli_idx = cl.set_index("id") if not cl.empty else pd.DataFrame(columns=["cliente", "direccion", "destino"])
    filas = []
    for v in sel:
        if v in cli_idx.index:
            r = cli_idx.loc[v]
            filas.append({"id": v, "cliente": r["cliente"], "direccion": r["direccion"], "destino": r["destino"],
                          "contacto": ""})
        else:
            nombre = re.sub(r"^(➕\s*)+|(\s*\(nuevo\))+$", "", txt(v), flags=re.I).strip().upper()
            filas.append({"id": "", "cliente": nombre, "direccion": "", "destino": "", "contacto": ""})
    if not filas:
        filas = [{"id": "", "cliente": SIN_CLIENTE, "direccion": "", "destino": "", "contacto": ""}]
    df = pd.DataFrame(filas, columns=["id", *PUNTO_COLS])
    st.markdown("**Datos de entrega**")
    ed = st.data_editor(
        df, num_rows="fixed", hide_index=True, width="stretch",
        key=f"sol_ed_{n}_{cuenta_id}_{'|'.join(sel)}",
        column_config={"id": None,
                       "cliente": st.column_config.TextColumn("Cliente", disabled=True, width="small"),
                       "direccion": st.column_config.TextColumn("Dirección", width="medium"),
                       "destino": st.column_config.TextColumn("Destino", width="small"),
                       "contacto": st.column_config.TextColumn("Contacto", width="small")})
    st.caption("La dirección y el destino quedan guardados para la próxima vez. El contacto no se guarda.")
    st.divider()

    # 3) transporte y observaciones
    ti = tipos_act()
    ti = ti[ti["categoria"] != "CLIENTE"]
    st.selectbox("T. transporte (opcional)", ti["nombre"].tolist(), index=None, placeholder="Sin preferencia",
                 key=f"sol_tipo_{n}")
    st.text_area("Observaciones", key=f"sol_obs_{n}", height=80, placeholder="Qué se envía, indicaciones de acceso…")
    if st.button("Registrar solicitud", type="primary", width="stretch"):
        _guardar_solicitud(cuenta_id, ed, cli_idx, n)


def _guardar_solicitud(cuenta_id, ed, cli_idx, n):
    if not cuenta_id:
        st.error("Elige la cuenta.")
        return
    ed = ed.copy()
    for k in PUNTO_COLS:
        ed[k] = ed[k].map(txt)
    sin_dir = ed[ed["direccion"] == ""]
    if not sin_dir.empty:
        st.error("Falta la dirección de: " + ", ".join(sin_dir["cliente"]))
        return
    cuenta = store.read("CUENTAS").set_index("id").loc[cuenta_id, "nombre"]
    proceso = ss["sol_proc"]
    grupo = L.nuevo_grupo() if len(ed) > 1 else ""
    codigos = []
    for _, p in ed.iterrows():
        lugar = p["destino"].upper()
        cliente = "" if p["cliente"] == SIN_CLIENTE else p["cliente"]
        datos = {"proceso": proceso, "cuenta_id": cuenta_id, "cuenta": cuenta, "cliente": cliente,
                 "cliente_id": p["id"] if p["id"] in cli_idx.index else "",
                 "direccion": p["direccion"], "referencia": "", "contacto": p["contacto"],
                 "origen": ALMACEN if proceso == "OUT" else (lugar or "CLIENTE"),
                 "destino": lugar if proceso == "OUT" else ALMACEN,
                 "fecha_entrega": ss["sol_fecha"].isoformat(), "tipo_fecha": ss["sol_tipo_fecha"],
                 "hora_cita": hora_str(ss.get(f"sol_hora_{n}")), "t_transporte_sol": ss.get(f"sol_tipo_{n}") or "",
                 "observacion": txt(ss.get(f"sol_obs_{n}")), "grupo": grupo}
        codigos.append(L.crear_movimiento(store, "SOLICITUDES", datos, USUARIO)["codigo"])
    ss["_sol_reset"] = True
    flash(f"Registrado: **{', '.join(codigos)}**" + (f" (pedido {grupo})" if grupo else ""))
    st.rerun()


@st.dialog("Editar solicitud", width="large")
def dlg_editar_solicitud(sol_id):
    sols = store.read("SOLICITUDES").set_index("id")
    if sol_id not in sols.index:
        st.error("No se encontró la solicitud.")
        return
    s = sols.loc[sol_id]
    st.markdown(f"<span class='sol-cod'>{s['codigo']}</span> &nbsp; <span class='proc proc-{s['proceso']}'>{s['proceso']}</span>"
                f" &nbsp; <span class='pill st-{s['estado']}'>{ESTADOS.get(s['estado'], s['estado'])}</span>"
                f"<div class='sol-sub'>{s['cuenta']}</div>", unsafe_allow_html=True)
    if s["estado"] == "ruteado":
        st.info("Esta solicitud ya está en una ruta. Los cambios se verán en la hoja de ruta.")
    k = f"ed_{sol_id}_"
    a, b = st.columns(2)
    cliente = a.text_input("Cliente / punto de entrega", s["cliente"], key=k + "cli")
    lugar_actual = s["destino"] if s["proceso"] == "OUT" else s["origen"]
    destino = b.text_input("Destino (distrito)", "" if lugar_actual in (ALMACEN, "CLIENTE") else lugar_actual, key=k + "des")
    direccion = st.text_input("Dirección", s["direccion"], key=k + "dir")
    a, b = st.columns(2)
    referencia = a.text_input("Referencia", s["referencia"], key=k + "ref")
    contacto = b.text_input("Contacto", s["contacto"].replace("\n", " / "), key=k + "con")
    a, b, c = st.columns(3)
    tf = a.radio("Fecha de entrega", list(TIPO_FECHA), format_func=TIPO_FECHA.get, key=k + "tf",
                 index=1 if s.get("tipo_fecha") == "hasta" else 0)
    fecha = b.date_input("Fecha", date.fromisoformat(s["fecha_entrega"]) if s["fecha_entrega"] else hoy(),
                         key=k + "fe", format="DD/MM/YYYY")
    hora_txt = c.text_input("Hora de cita (HH:MM, opcional)", s["hora_cita"], key=k + "ho")
    ti = tipos_act()
    ti = ti[ti["categoria"] != "CLIENTE"]["nombre"].tolist()
    tipo = st.selectbox("T. transporte (opcional)", ti, index=ti.index(s["t_transporte_sol"]) if s["t_transporte_sol"] in ti else None,
                        placeholder="Sin preferencia", key=k + "ti")
    cols = st.columns(5)
    valores = {}
    for n, campo in enumerate(["cant_caja", "pedido_gr", "n_orden", "palets", "unidades", "cant_inner", "accesorio",
                               "maquina", "x_und_acc"]):
        valores[campo] = cols[n % 5].text_input(PEDIDO_LABELS[campo], s[campo], key=k + campo)
    obs = st.text_area("Observaciones", s["observacion"], key=k + "obs", height=70)
    a, b = st.columns([1, 1])
    if a.button("💾 Guardar cambios", type="primary", width="stretch", key=k + "save"):
        h = txt(hora_txt)
        if h and not (len(h) == 5 and h[2] == ":" and h.replace(":", "").isdigit()):
            st.error("La hora debe tener el formato HH:MM, por ejemplo 10:30.")
            return
        if not txt(direccion):
            st.error("La dirección no puede quedar vacía.")
            return
        lugar = txt(destino).upper()
        cambios = {"cliente": txt(cliente), "direccion": txt(direccion), "referencia": txt(referencia),
                   "contacto": txt(contacto), "tipo_fecha": tf, "fecha_entrega": fecha.isoformat(), "hora_cita": h,
                   "t_transporte_sol": tipo or "", "observacion": txt(obs), "actualizado": ahora_str(),
                   "origen": ALMACEN if s["proceso"] == "OUT" else (lugar or "CLIENTE"),
                   "destino": lugar if s["proceso"] == "OUT" else ALMACEN}
        for campo, v in valores.items():
            cambios[campo] = txt(v) if campo in PEDIDO_TEXTO else ("" if num(v) is None else str(num(v)))
        store.update("SOLICITUDES", sol_id, cambios)
        L.aprender_cliente(store, {**s.to_dict(), **cambios}, contar=False)
        flash(f"Solicitud {s['codigo']} actualizada.")
        st.rerun()
    if s["estado"] == "pendiente":
        conf = b.checkbox("Confirmo que quiero anularla", key=k + "conf")
        if b.button("🗑️ Anular solicitud", width="stretch", disabled=not conf, key=k + "anu"):
            L.anular_solicitud(store, sol_id, USUARIO)
            flash(f"Solicitud {s['codigo']} anulada.")
            st.rerun()


def lista_mis_solicitudes():
    sols = store.read("SOLICITUDES")
    mias = sols[(sols["solicitante"] == USUARIO) & (sols["estado"] != "anulado")]
    cuenta = mias["estado"].value_counts()
    st.markdown("#### Mis solicitudes")
    st.markdown("<div class='kpi-mini'>" + "".join(
        f"<span>{ESTADOS[e]}&nbsp; <b>{cuenta.get(e, 0)}</b></span>" for e in ("pendiente", "ruteado"))
        + "</div>", unsafe_allow_html=True)
    f1, f2 = st.columns([1, 2])
    est = f1.selectbox("Estado", ["", "pendiente", "ruteado"],
                       format_func=lambda e: ESTADOS.get(e, "Todos"), key="ms_est")
    q = f2.text_input("Buscar", placeholder="Cuenta, cliente, código, GR, destino", key="ms_q")
    if est:
        mias = mias[mias["estado"] == est]
    if q:
        mias = mias[(mias["cuenta"] + " " + mias["cliente"] + " " + mias["codigo"] + " " + mias["pedido_gr"] + " "
                     + mias["destino"] + " " + mias["grupo"]).map(norm).str.contains(norm(q), regex=False)]
    mias = mias.sort_values("creado", ascending=False)
    if mias.empty:
        st.info("No hay solicitudes con estos filtros." if est or q else "Aún no registraste solicitudes.")
        return
    rutas = store.read("RUTAS").set_index("id")
    limite = 40
    for _, s in mias.head(limite).iterrows():
        r = rutas.loc[s["ruta_id"]] if s["ruta_id"] in rutas.index else None
        lugar = s["destino"] if s["proceso"] == "OUT" else s["origen"]
        with st.container(border=True):
            c0, c1, c2, c3 = st.columns([1.45, 3, 1.75, 0.6], vertical_alignment="center")
            c0.markdown(f"<div class='sol-cod'>{s['codigo']}</div><span class='proc proc-{s['proceso']}'>{s['proceso']}</span>"
                        + (f" <div class='sol-grp'>⛓ {s['grupo']}</div>" if s["grupo"] else ""), unsafe_allow_html=True)
            c1.markdown(f"<div class='sol-tit'>{s['cuenta']}</div>"
                        f"<div class='sol-sub'><b>{s['cliente'] or '—'}</b>{' · ' + lugar if lugar else ''}</div>"
                        f"<div class='sol-sub'>{s['direccion']}</div>", unsafe_allow_html=True)
            c2.markdown(f"<span class='pill st-{s['estado']}'>{ESTADOS.get(s['estado'], s['estado'])}</span>"
                        f"<div class='sol-fecha' style='margin-top:4px'>{_txt_fecha(s)}</div>"
                        f"<div class='sol-sub'>{'Cita ' + s['hora_cita'] if s['hora_cita'] else 'Sin hora de cita'}"
                        f"{' · ' + s['cant_caja'] + ' cj' if s['cant_caja'] else ''}</div>"
                        + (f"<div class='sol-sub'>🚚 {r['codigo']} · {r['chofer']}</div>" if r is not None else "")
                        + (f"<div class='sol-sub'>{s['nota_entrega']}</div>" if s["nota_entrega"] else ""),
                        unsafe_allow_html=True)
            if c3.button("✏️", key=f"edit_{s['id']}", help="Editar" if s["estado"] in EDITABLES
                         else "Ya salió a ruta: no se puede editar", disabled=s["estado"] not in EDITABLES):
                dlg_editar_solicitud(s["id"])
    if len(mias) > limite:
        st.caption(f"Mostrando las {limite} más recientes de {len(mias)}. Usa el buscador para encontrar otras.")


def sec_solicitudes():
    st.subheader("Solicitud de unidad de transporte")
    st.caption("Queda **pendiente** hasta que el área de ruteo la asigne a una ruta.")
    izq, der = st.columns([6, 5], gap="large")
    with izq, st.container(border=True):
        form_solicitud()
    with der:
        lista_mis_solicitudes()


def _lista_propia(tabla, titulo, columnas):
    df = store.read(tabla).sort_values("creado", ascending=False)
    todos = tiene("supervisor", "ruteador")
    if not todos:
        df = df[df["registrado_por"] == USUARIO]
    st.markdown(f"**{'Todos los registros' if todos else titulo}**")
    df = df.assign(Fecha=df[[c for c in ("fecha_entrega", "fecha") if c in df][0]].map(fmt_fecha))
    tabla_registros(df, columnas, alto=520)


def sec_regularizacion():
    st.subheader("Regularización de despacho")
    st.caption("Despachos o recepciones que ya se hicieron sin pasar por una solicitud, para que queden en la base.")
    izq, der = st.columns([2, 3], gap="large")
    with izq:
        with st.container(border=True):
            form_movimiento("g")
    with der:
        _lista_propia("REGULARIZACIONES", "Mis regularizaciones",
                      {"codigo": "Código", "proceso": "Proc.", "cuenta": "Cuenta", "cliente": "Cliente",
                       "Fecha": "Fecha", "t_transporte": "Transporte", "chofer": "Chofer", "motivo": "Motivo",
                       "registrado_por": "Registró"})


def sec_movcliente():
    st.subheader("Recepción / despacho gestionado por el cliente")
    st.caption("Cuando el cliente trae o recoge la mercadería con su propio transporte.")
    izq, der = st.columns([2, 3], gap="large")
    with izq:
        with st.container(border=True):
            form_movimiento("m")
    with der:
        _lista_propia("MOV_CLIENTE", "Mis registros",
                      {"codigo": "Código", "proceso": "Proc.", "cuenta": "Cuenta", "cliente": "Cliente",
                       "Fecha": "Fecha", "hora": "Hora", "empresa": "Empresa", "placa": "Placa",
                       "pedido_gr": "GR", "registrado_por": "Registró"})


def _descripcion_ruta(r, sols_idx):
    n = len(L.paradas(r))
    return (f"**{r['codigo']}** · {ESTADOS.get(r['estado'], r['estado'])} · {r['t_transporte']}"
            + (f" · {r['placa']}" if r["placa"] else "") + f" · {r['chofer']} (viaje {r['viaje'] or 1})"
            + f" · {n} parada(s)" + (f" · {r['cant_caja']} cajas" if num0(r["cant_caja"]) else ""))


def _tabla_paradas(r, sols_idx):
    filas = []
    for n, i in enumerate(L.paradas(r), 1):
        if i in sols_idx.index:
            x = sols_idx.loc[i]
            filas.append({"#": n, "Proc.": x["proceso"], "Cuenta": x["cuenta"], "Cliente": x["cliente"],
                          "Destino": x["destino"] if x["proceso"] == "OUT" else x["origen"], "Cita": x["hora_cita"],
                          "Dirección": x["direccion"], "Cajas": x["cant_caja"], "GR": x["pedido_gr"],
                          "Estado": ESTADOS.get(x["estado"], x["estado"]), "Llegada": x["hora_llegada"],
                          "Nota": x["nota_entrega"], "Código": x["codigo"]})
    return pd.DataFrame(filas)


def sec_ruteo():
    st.subheader("Ruteo")
    st.caption("Marca solicitudes pendientes, define el orden y asígnales transporte y chofer. Cada ruta recibe su código.")
    sols = store.read("SOLICITUDES")
    rutas = store.read("RUTAS")
    f1, f2, f3 = st.columns([1, 1, 1])
    fecha = f1.date_input("Fecha de entrega", hoy(), format="DD/MM/YYYY", key="rt_fecha")
    agrupar = f2.selectbox("Ordenar por", ["Destino", "Cuenta", "Transporte solicitado", "Hora de cita"])
    atras = f3.checkbox("Incluir pendientes atrasados", key="rt_atras")
    fs = fecha.isoformat()
    hasta = sols["tipo_fecha"] == "hasta"
    pend = sols[(sols["estado"] == "pendiente") & ((sols["fecha_entrega"] == fs) | (hasta & (sols["fecha_entrega"] > fs))
                                                   | (atras & (sols["fecha_entrega"] < fs)))]
    atrasados = sols[(sols["estado"] == "pendiente") & (sols["fecha_entrega"] < fs)]
    if len(atrasados) and not atras:
        st.warning(f"Hay {len(atrasados)} solicitud(es) pendiente(s) de fechas anteriores. "
                   "Marca «Incluir pendientes atrasados» para verlas.")
    izq, der = st.columns([3, 2], gap="large")
    with izq:
        if pend.empty:
            st.info(f"No hay solicitudes pendientes para el {fecha:%d/%m/%Y}.")
            elegidas = pd.DataFrame()
        else:
            v = pend.copy()
            v["Lugar"] = [d if p == "OUT" else o for p, d, o in zip(v["proceso"], v["destino"], v["origen"])]
            clave = {"Destino": "Lugar", "Cuenta": "cuenta", "Transporte solicitado": "t_transporte_sol",
                     "Hora de cita": "hora_cita"}[agrupar]
            v = v.sort_values([clave, "hora_cita"])
            vista = pd.DataFrame({
                "Sel": False, "Orden": None, "Código": v["codigo"], "Proc.": v["proceso"], "Cuenta": v["cuenta"],
                "Cliente": v["cliente"], "Destino": v["Lugar"], "Cita": v["hora_cita"],
                "Fecha": [_txt_fecha(x) for _, x in v.iterrows()], "Cajas": v["cant_caja"],
                "Pide": v["t_transporte_sol"], "GR": v["pedido_gr"], "Dirección": v["direccion"], "id": v["id"],
            })
            ed = st.data_editor(
                vista, hide_index=True, width="stretch", key=f"rt_ed_{fs}_{ss.get('rt_n', 0)}",
                column_config={
                    "Sel": st.column_config.CheckboxColumn("✔", width="small"),
                    "Orden": st.column_config.NumberColumn("Orden", min_value=1, step=1, width="small",
                                                           help="Opcional: orden de la parada en la ruta"),
                    "id": None,
                },
                disabled=[c for c in vista.columns if c not in ("Sel", "Orden")],
            )
            elegidas = ed[ed["Sel"]].copy()
            if not elegidas.empty:
                elegidas["_o"] = elegidas["Orden"].fillna(9999)
                elegidas = elegidas.reset_index(drop=True).reset_index().sort_values(["_o", "index"])
    with der:
        with st.container(border=True):
            st.markdown("**Nueva ruta**")
            ti = tipos_act()
            ti = ti[ti["categoria"] != "CLIENTE"]
            ch = choferes_act()
            a, b = st.columns(2)
            tipo = a.selectbox("T. transporte", ti["nombre"].tolist(), index=None, placeholder="Selecciona…")
            chofer_id = b.selectbox("Chofer / agencia", ch["id"].tolist(), index=None, placeholder="Selecciona…",
                                    format_func=dict(zip(ch["id"], ch["nombre"] + " · " + ch["tipo"])).get)
            a, b = st.columns(2)
            placa = a.text_input("Placa (opcional)")
            salida = b.time_input("Hora de salida", value=None, step=timedelta(minutes=15))
            fecha_ruta = st.date_input("Fecha de la ruta", fecha, format="DD/MM/YYYY")
            notas = st.text_area("Notas para el chofer", height=70)
            if not elegidas.empty:
                cajas = sum(num0(x) for x in elegidas["Cajas"])
                st.markdown(f"**{len(elegidas)} parada(s)**" + (f" · {cajas:g} cajas" if cajas else ""))
                st.dataframe(elegidas[["Cliente", "Destino", "Cita"]].reset_index(drop=True).rename(lambda i: i + 1),
                             width="stretch")
                if tipo:
                    otros = [x for x in elegidas["Pide"] if x and x != tipo]
                    if otros:
                        st.warning(f"{len(otros)} pedido(s) pidieron otro transporte ({', '.join(sorted(set(otros)))}).")
                if chofer_id:
                    n = L.viajes_del_chofer(rutas, chofer_id, fecha_ruta.isoformat())
                    if n:
                        st.caption(f"Será el viaje {n + 1} de este chofer ese día.")
            else:
                st.caption("Marca solicitudes en la tabla para agregarlas como paradas.")
            if st.button("Crear ruta", type="primary", disabled=elegidas.empty):
                if not tipo or not chofer_id:
                    st.error("Elige el tipo de transporte y el chofer.")
                else:
                    try:
                        r = L.crear_ruta(store, elegidas["id"].tolist(), fecha_ruta.isoformat(), tipo, chofer_id,
                                         ch.set_index("id").loc[chofer_id, "nombre"], placa, hora_str(salida), notas,
                                         USUARIO)
                        ss["rt_n"] = ss.get("rt_n", 0) + 1
                        flash(f"Ruta **{r['codigo']}** creada con {len(L.paradas(r))} parada(s).")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

    st.divider()
    rutas = store.read("RUTAS")
    del_dia = rutas[(rutas["fecha"] == fs) & (rutas["estado"] != "anulada")].sort_values(["chofer", "viaje"])
    a, b = st.columns([3, 1])
    a.markdown(f"#### Rutas del {fecha:%d/%m/%Y}")
    sols = store.read("SOLICITUDES")
    if not del_dia.empty:
        b.download_button("⬇️ Hoja de choferes (Excel)", excel_hoja_choferes(del_dia, sols),
                          file_name=f"HOJA_CHOFERES_{fs}.xlsx", width="stretch")
    if del_dia.empty:
        st.info("Aún no hay rutas para este día.")
    sidx = sols.set_index("id")
    for _, r in del_dia.iterrows():
        with st.expander(_descripcion_ruta(r, sidx)):
            st.dataframe(_tabla_paradas(r, sidx), hide_index=True, width="stretch")
            a, b, c = st.columns(3)
            a.download_button("Hoja de ruta (texto)", L.texto_hoja_ruta(r, sols), file_name=f"{r['codigo']}.txt",
                              key=f"txt_{r['id']}")
            if r["estado"] == "programada":
                pars = [p for p in L.paradas(r) if p in sidx.index]
                quitar = b.selectbox("Quitar parada", pars, index=None, key=f"q_{r['id']}",
                                     format_func=lambda i: sidx.loc[i, "cliente"])
                if quitar and b.button("Quitar", key=f"qb_{r['id']}"):
                    L.quitar_parada(store, r, quitar)
                    flash("Parada quitada; volvió a pendientes.")
                    st.rerun()
                conf = c.checkbox("Confirmo anular", key=f"ca_{r['id']}")
                if c.button("Anular ruta", key=f"an_{r['id']}", disabled=not conf):
                    L.anular_ruta(store, r)
                    flash("Ruta anulada; sus pedidos volvieron a pendientes.")
                    st.rerun()


def sec_despacho():
    st.subheader("Despacho")
    st.caption("Registra la salida de cada ruta y el resultado de cada parada. Quien registra la salida queda como encargado.")
    f1, f2 = st.columns(2)
    fecha = f1.date_input("Fecha", hoy(), format="DD/MM/YYYY", key="dp_fecha")
    est = f2.selectbox("Estado", ["", "programada", "en_ruta", "completada"],
                       format_func=lambda e: ESTADOS.get(e, "Activas y cerradas"))
    rutas = store.read("RUTAS")
    sols = store.read("SOLICITUDES")
    sidx = sols.set_index("id")
    lista = rutas[(rutas["fecha"] == fecha.isoformat()) & (rutas["estado"] != "anulada")]
    if est:
        lista = lista[lista["estado"] == est]
    orden = {"en_ruta": 0, "programada": 1, "completada": 2}
    lista = lista.assign(_o=lista["estado"].map(orden)).sort_values(["_o", "hora_salida_plan", "chofer"])
    if lista.empty:
        st.info(f"No hay rutas para el {fecha:%d/%m/%Y}.")
    for _, r in lista.iterrows():
        with st.expander(_descripcion_ruta(r, sidx), expanded=r["estado"] == "en_ruta"):
            if r["salida_real"]:
                st.caption(f"Salió {r['salida_real']}" + (f" · cerró {r['cierre']}" if r["cierre"] else ""))
            if r["estado"] == "programada":
                st.dataframe(_tabla_paradas(r, sidx), hide_index=True, width="stretch")
                if st.button("🚚 Registrar salida", key=f"sal_{r['id']}", type="primary"):
                    h = L.registrar_salida(store, r, USUARIO)
                    flash(f"Salida registrada a las {h}.")
                    st.rerun()
            elif r["estado"] == "en_ruta":
                ids = [i for i in L.paradas(r) if i in sidx.index]
                ahora = _hora_actual()
                vista = pd.DataFrame({
                    "id": ids,
                    "Cliente": [sidx.loc[i, "cliente"] for i in ids],
                    "Proc.": [sidx.loc[i, "proceso"] for i in ids],
                    "Destino": [sidx.loc[i, "destino"] for i in ids],
                    "Cita": [sidx.loc[i, "hora_cita"] for i in ids],
                    "Resultado": [{"entregado": "Entregado", "no_entregado": "No entregado"}.get(sidx.loc[i, "estado"], "En ruta")
                                  for i in ids],
                    "Hora": [sidx.loc[i, "hora_llegada"] or ahora for i in ids],
                    "Motivo / nota": [sidx.loc[i, "nota_entrega"] for i in ids],
                })
                ed = st.data_editor(vista, hide_index=True, width="stretch", key=f"ed_{r['id']}",
                                    column_config={"id": None,
                                                   "Resultado": st.column_config.SelectboxColumn(
                                                       options=["En ruta", "Entregado", "No entregado"], required=True),
                                                   "Hora": st.column_config.TextColumn(help="HH:MM")},
                                    disabled=["Cliente", "Proc.", "Destino", "Cita"])
                a, b = st.columns(2)
                if a.button("💾 Guardar resultados", key=f"gr_{r['id']}", type="primary"):
                    faltan = ed[(ed["Resultado"] == "No entregado") & (ed["Motivo / nota"].fillna("").str.strip() == "")]
                    if not faltan.empty:
                        st.error("Escribe el motivo de cada parada no entregada.")
                    else:
                        est_map = {"Entregado": "entregado", "No entregado": "no_entregado", "En ruta": "en_ruta"}
                        L.registrar_resultados(store, {
                            x["id"]: {"estado": est_map[x["Resultado"]],
                                      "hora_llegada": txt(x["Hora"])[:5] if x["Resultado"] != "En ruta" else "",
                                      "nota_entrega": txt(x["Motivo / nota"])}
                            for _, x in ed.iterrows()})
                        flash("Resultados guardados.")
                        st.rerun()
                todas = all(sidx.loc[i, "estado"] in ("entregado", "no_entregado") for i in ids)
                if b.button("✅ Cerrar ruta", key=f"cr_{r['id']}", disabled=not todas,
                            help=None if todas else "Guarda el resultado de todas las paradas primero"):
                    L.cerrar_ruta(store, r)
                    flash("Ruta cerrada.")
                    st.rerun()
            else:
                st.dataframe(_tabla_paradas(r, sidx), hide_index=True, width="stretch")


def _rango(tipo):
    t = hoy()
    if tipo == "7":
        ss["q_desde"], ss["q_hasta"] = t - timedelta(days=6), t
    elif tipo == "mes":
        ss["q_desde"], ss["q_hasta"] = t.replace(day=1), t
    elif tipo == "ant":
        fin = t.replace(day=1) - timedelta(days=1)
        ss["q_desde"], ss["q_hasta"] = fin.replace(day=1), fin
    else:
        fechas = [x for x in store.read("HISTORICO")["fecha_entrega"].tolist() + store.read("SOLICITUDES")["fecha_entrega"].tolist() if x]
        ss["q_desde"] = date.fromisoformat(min(fechas)) if fechas else t
        ss["q_hasta"] = t


def sec_reportes():
    st.subheader("Base de entregas y reportes")
    ss.setdefault("q_desde", hoy() - timedelta(days=6))
    ss.setdefault("q_hasta", hoy())
    b = st.columns(5)
    for col, (k, lab) in zip(b, [("7", "7 días"), ("mes", "Este mes"), ("ant", "Mes anterior"), ("todo", "Todo")]):
        col.button(lab, on_click=_rango, args=(k,), width="stretch")
    f = st.columns([1, 1, 2, 1, 1])
    desde = f[0].date_input("Desde", key="q_desde", format="DD/MM/YYYY")
    hasta = f[1].date_input("Hasta", key="q_hasta", format="DD/MM/YYYY")
    cu = cuentas_act()
    cuenta = f[2].selectbox("Cuenta", [""] + cu["nombre"].tolist(), format_func=lambda x: x or "Todas")
    proceso = f[3].selectbox("Proceso", ["", "OUT", "IN"], format_func=lambda x: x or "Todos")
    hist = f[4].checkbox("Incluir histórico", value=True)
    R = L.base_entregas(store, desde.isoformat(), hasta.isoformat(), cuenta, proceso, hist)
    if R.empty:
        st.info("Sin registros en el periodo.")
        return
    R["_viaje"] = R.apply(L.clave_viaje, axis=1)
    sol = R[R["registro"] == "Solicitud"]
    ok = R[R["revision"].str.upper().str.startswith("OK")]
    ent = sol[sol["estado"] == "entregado"]
    a_tiempo = ent[(ent["hora_llegada"] == "") | (ent["hora_cita"] == "") | (ent["hora_llegada"] <= ent["hora_cita"])]
    m = st.columns(6)
    m[0].metric("Registros", len(R))
    m[1].metric("OUT / IN", f"{(R['proceso'] == 'OUT').sum()} / {(R['proceso'] == 'IN').sum()}")
    m[2].metric("Rutas / viajes", R["_viaje"].replace("", pd.NA).nunique())
    m[3].metric("Confirmadas (OK)", len(ok))
    m[4].metric("Llegó a la cita", f"{round(len(a_tiempo) / len(ent) * 100)}%" if len(ent) else "—")
    m[5].metric("Pendientes", (sol["estado"] == "pendiente").sum())
    m = st.columns(6)
    m[0].metric("No entregadas", (sol["estado"] == "no_entregado").sum())
    m[1].metric("Regularizaciones", (R["registro"] == "Regularización").sum())
    m[2].metric("Gestión cliente", (R["registro"] == "Cliente").sum())
    m[3].metric("Histórico", (R["registro"] == "Histórico").sum())
    m[4].metric("Cajas", f"{R['cant_caja'].map(num0).sum():,.0f}")
    m[5].metric("Unidades", f"{R['unidades'].map(num0).sum():,.0f}")

    x = R.assign(_n=1)
    g1, g2 = st.columns(2)
    dias = (hasta - desde).days + 1
    if dias > 45:
        g1.markdown("**Registros por mes**")
        g1.bar_chart(x.groupby(x["fecha_entrega"].str[:7])["_n"].sum(), color="#1f6b45")
    else:
        g1.markdown("**Registros por día**")
        g1.bar_chart(x.groupby("fecha_entrega")["_n"].sum(), color="#1f6b45")
    g2.markdown("**Top cuentas**")
    g2.bar_chart(x.groupby("cuenta")["_n"].sum().sort_values(ascending=False).head(10), horizontal=True,
                 color="#1f6b45")

    t1, t2, t3, t4, t5 = st.tabs(["Base de entregas", "Diario", "Por cuenta", "Por transporte", "Por chofer"])
    from schema import BASE_COLS
    with t1:
        st.dataframe(R[[k for k, _ in BASE_COLS]].rename(columns=dict(BASE_COLS)), hide_index=True,
                     width="stretch", height=460)
    agg = dict(Registros=("_n", "sum"), OUT=("proceso", lambda s: (s == "OUT").sum()),
               IN=("proceso", lambda s: (s == "IN").sum()), Viajes=("_viaje", lambda s: s[s != ""].nunique()),
               Cajas=("cant_caja", lambda s: s.map(num0).sum()), Palets=("palets", lambda s: s.map(num0).sum()),
               Unidades=("unidades", lambda s: s.map(num0).sum()))
    with t2:
        d = x.groupby("fecha_entrega").agg(**agg).sort_index(ascending=False)
        d.index = d.index.map(fmt_fecha)
        st.dataframe(d, width="stretch")
    with t3:
        st.dataframe(x.groupby("cuenta").agg(**agg).sort_values("Registros", ascending=False), width="stretch")
    with t4:
        st.dataframe(x.groupby(["t_transporte", "transporte"]).agg(**agg).sort_values("Registros", ascending=False),
                     width="stretch")
    with t5:
        c = x[x["chofer"] != ""]
        st.dataframe(c.groupby("chofer").agg(Paradas=("_n", "sum"), Viajes=("_viaje", lambda s: s[s != ""].nunique()),
                                             Dias=("fecha_entrega", "nunique"),
                                             Cajas=("cant_caja", lambda s: s.map(num0).sum()))
                     .sort_values("Paradas", ascending=False), width="stretch")

    a, b = st.columns(2)
    a.download_button("⬇️ Descargar BASE DE ENTREGAS (Excel)", excel_base(R),
                      file_name=f"BASE_DE_ENTREGAS_{desde}_a_{hasta}.xlsx", type="primary", width="stretch")
    b.download_button("⬇️ Descargar en CSV", R[[k for k, _ in BASE_COLS]].rename(columns=dict(BASE_COLS))
                      .to_csv(index=False, sep=";").encode("utf-8-sig"), file_name=f"BASE_DE_ENTREGAS_{desde}_a_{hasta}.csv",
                      width="stretch")


def _editor_maestro(tabla, df, config, key, orden):
    df = df.sort_values(orden).reset_index(drop=True)
    df["activo"] = df["activo"].map(es_activo)
    ed = st.data_editor(df, num_rows="dynamic", hide_index=True, width="stretch", key=key,
                        column_config={"id": None, "activo": st.column_config.CheckboxColumn("Activo", default=True),
                                       **config})
    return ed


def _guardar_maestro(tabla, ed, resto=None, id_de=None):
    ed = ed.copy().fillna("")
    ed = ed[ed.drop(columns=["id", "activo"], errors="ignore").astype(str).apply(lambda r: "".join(r).strip(), axis=1) != ""]
    ed["id"] = [i if txt(i) else (id_de(r) if id_de else nuevo_id()) for i, (_, r) in zip(ed["id"], ed.iterrows())]
    ed["activo"] = ed["activo"].map(lambda v: "TRUE" if v in (True, "TRUE", "True", "") else "FALSE")
    final = pd.concat([resto, ed]) if resto is not None else ed
    store.replace_all(tabla, final.drop_duplicates("id", keep="last"))
    flash("Cambios guardados.")
    st.rerun()


def sec_maestros():
    st.subheader("Maestros")
    t1, t2, t3, t4, t5 = st.tabs(["Cuentas", "Clientes por cuenta", "Tipos de transporte", "Choferes y agencias",
                                  "Importar Excel"])
    with t1:
        st.caption("Puedes agregar filas al final o pegar desde Excel. Desmarca «Activo» en vez de borrar.")
        ed = _editor_maestro("CUENTAS", store.read("CUENTAS"), {"nombre": "Cuenta"}, "m_cu", "nombre")
        if st.button("Guardar cuentas", type="primary"):
            ed["nombre"] = ed["nombre"].map(lambda s: txt(s).upper())
            _guardar_maestro("CUENTAS", ed, id_de=lambda r: slug(r["nombre"]))
    with t2:
        cu = store.read("CUENTAS").sort_values("nombre")
        cuenta_id = st.selectbox("Cuenta", cu["id"].tolist(), format_func=dict(zip(cu["id"], cu["nombre"])).get,
                                 key="m_cl_cuenta")
        todos = store.read("CLIENTES")
        sub = todos[todos["cuenta_id"] == cuenta_id]
        st.caption(f"{len(sub)} cliente(s). Pega filas desde Excel: Cliente, Dirección, Destino, Contacto, Referencia.")
        ed = _editor_maestro("CLIENTES", sub.drop(columns=["cuenta_id", "cuenta"]),
                             {"cliente": "Cliente", "direccion": "Dirección", "destino": "Destino",
                              "contacto": "Contacto", "referencia": "Referencia",
                              "usos": st.column_config.TextColumn("Usos", disabled=True)}, f"m_cl_{cuenta_id}", "cliente")
        if st.button("Guardar clientes", type="primary"):
            ed["cuenta_id"] = cuenta_id
            ed["cuenta"] = cu.set_index("id").loc[cuenta_id, "nombre"]
            ed["destino"] = ed["destino"].map(lambda s: txt(s).upper())
            _guardar_maestro("CLIENTES", ed, resto=todos[todos["cuenta_id"] != cuenta_id])
    with t3:
        ed = _editor_maestro("TIPOS_TRANSPORTE", store.read("TIPOS_TRANSPORTE"),
                             {"nombre": "Tipo de transporte",
                              "categoria": st.column_config.SelectboxColumn("Categoría", options=CATEGORIAS, required=True)},
                             "m_ti", "nombre")
        if st.button("Guardar tipos", type="primary"):
            ed["nombre"] = ed["nombre"].map(lambda s: txt(s).upper())
            _guardar_maestro("TIPOS_TRANSPORTE", ed, id_de=lambda r: slug(r["nombre"]))
    with t4:
        ed = _editor_maestro("CHOFERES", store.read("CHOFERES"),
                             {"nombre": "Nombre", "tipo": st.column_config.SelectboxColumn("Tipo", options=TIPOS_CHOFER),
                              "dni": "DNI", "telefono": "Teléfono"}, "m_ch", "nombre")
        if st.button("Guardar choferes", type="primary"):
            ed["nombre"] = ed["nombre"].map(lambda s: txt(s).upper())
            _guardar_maestro("CHOFERES", ed, id_de=lambda r: "ch-" + slug(r["nombre"]))
    with t5:
        sec_importar()


def sec_importar():
    st.markdown("Sube tu Excel de trabajo (con las hojas **MATRICES** y **RUTEO**) para cargar cuentas, "
                "clientes e histórico.")
    st.caption("Puedes subir tu Excel de siempre (PROPUESTA_PLANTILLA) tal cual, o descargar esta plantilla "
               "con el formato exacto y las instrucciones.")
    cu = store.read("CUENTAS")
    st.download_button(
        "⬇️ Descargar plantilla de importación (Excel)",
        excel_plantilla(sorted(cu["nombre"].tolist()) if not cu.empty else [],
                        sorted(store.read("TIPOS_TRANSPORTE")["nombre"].tolist()),
                        sorted(store.read("CHOFERES")["nombre"].tolist())),
        file_name="PLANTILLA_IMPORTACION_DESPACHO.xlsx")
    st.divider()
    if store.read("TIPOS_TRANSPORTE").empty or store.read("CHOFERES").empty:
        if st.button("Cargar tipos de transporte y choferes por defecto"):
            if store.read("TIPOS_TRANSPORTE").empty:
                store.replace_all("TIPOS_TRANSPORTE", df_tipos_defecto())
            if store.read("CHOFERES").empty:
                store.replace_all("CHOFERES", df_choferes_defecto())
            flash("Tipos de transporte y choferes cargados.")
            st.rerun()
    archivo = st.file_uploader("Excel (.xlsx)", type=["xlsx", "xlsm"])
    if not archivo:
        return
    with st.spinner("Leyendo el Excel…"):
        datos = leer_plantilla(archivo)
    a, b, c = st.columns(3)
    a.metric("Cuentas (MATRICES)", len(datos["cuentas"]))
    b.metric("Clientes deducidos", len(datos["clientes"]))
    c.metric("Filas de histórico (RUTEO)", len(datos["historico"]))
    if st.button("Agregar cuentas y clientes nuevos (no duplica)"):
        cu = store.read("CUENTAS")
        nuevas = datos["cuentas"][~datos["cuentas"]["id"].isin(cu["id"])]
        if not nuevas.empty:
            store.append("CUENTAS", nuevas.to_dict("records"))
        cl = store.read("CLIENTES")
        existentes = set(zip(cl["cuenta_id"], cl["cliente"].map(norm))) | set(cl["id"])
        nuevos = datos["clientes"][[(r["cuenta_id"], norm(r["cliente"])) not in existentes and r["id"] not in existentes
                                    for _, r in datos["clientes"].iterrows()]]
        if not nuevos.empty:
            store.append("CLIENTES", nuevos.to_dict("records"))
        flash(f"{len(nuevas)} cuenta(s) y {len(nuevos)} cliente(s) agregados.")
        st.rerun()
    conf = st.checkbox("Entiendo que el histórico actual se reemplaza por el del archivo")
    if st.button("Cargar histórico", disabled=not conf):
        store.replace_all("HISTORICO", datos["historico"])
        flash(f"Histórico cargado: {len(datos['historico'])} filas.")
        st.rerun()


def sec_usuarios():
    st.subheader("Usuarios y roles")
    u = store.read("USUARIOS").sort_values("id")
    roles_sin_admin = list(ROLES)
    vista = pd.DataFrame({"id": u["id"], "Nombre": u["nombre"], "Activo": u["activo"].map(es_activo)})
    for r in roles_sin_admin:
        vista[ROLES[r]] = u["roles"].map(lambda s, r=r: r in s.split(","))
    ed = st.data_editor(vista, hide_index=True, width="stretch", key="u_ed",
                        column_config={"id": st.column_config.TextColumn("DNI / usuario", disabled=True)})
    if st.button("Guardar roles", type="primary"):
        base = u.set_index("id")
        filas = []
        for _, r in ed.iterrows():
            roles = [k for k in roles_sin_admin if r[ROLES[k]]]
            if r["id"] == USUARIO and "admin" not in roles:
                roles.append("admin")  # no te quites el acceso a ti misma
            filas.append({"id": r["id"], "nombre": r["Nombre"], "clave_hash": base.loc[r["id"], "clave_hash"],
                          "roles": ",".join(roles), "activo": "TRUE" if r["Activo"] else "FALSE",
                          "creado": base.loc[r["id"], "creado"]})
        store.replace_all("USUARIOS", pd.DataFrame(filas))
        flash("Roles guardados.")
        st.rerun()
    with st.form("f_nuevo_u", clear_on_submit=True):
        st.markdown("**Nuevo usuario**")
        a, b, c = st.columns([1, 2, 2])
        nu = a.text_input("DNI", max_chars=12).strip()
        nn = b.text_input("Nombre completo")
        nr = c.multiselect("Roles", list(ROLES), format_func=ROLES.get)
        if st.form_submit_button("Crear usuario", type="primary"):
            if not _dni_ok(nu):
                st.error("Escribe un DNI válido (solo números, 8 dígitos).")
            elif nu in set(u["id"]):
                st.error("Ya existe un usuario con ese DNI.")
            else:
                store.append("USUARIOS", [{"id": nu, "nombre": nn or nu, "clave_hash": "",
                                           "roles": ",".join(nr), "activo": "TRUE", "creado": ahora_str()}])
                flash(f"Usuario {nu} creado. Ya puede ingresar con su DNI.")
                st.rerun()


# ================================================================ router
{
    "solicitudes": sec_solicitudes, "regularizacion": sec_regularizacion, "movcliente": sec_movcliente,
    "ruteo": sec_ruteo, "despacho": sec_despacho, "reportes": sec_reportes, "maestros": sec_maestros,
    "usuarios": sec_usuarios,
}[ss["seccion"]]()
