"""Despacho Almacén — Solicitud → Ruteo → Despacho → Base de entregas.

Streamlit ejecuta este archivo de arriba a abajo en cada interacción. La sección
activa se guarda en st.session_state["seccion"]; cada sección es una función.
"""
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import logic as L
from auth import hash_clave, verificar_clave
from parser_plantilla import df_choferes_defecto, df_tipos_defecto, leer_plantilla
from report import excel_base, excel_hoja_choferes
from schema import CATEGORIAS, ESTADOS, PEDIDO, PEDIDO_LABELS, PEDIDO_TEXTO, ROLES, TIPOS_CHOFER
from utils import (ALMACEN, _hora_actual, ahora_str, es_activo, fmt_fecha, hora_str, hoy, norm, nuevo_id, num, num0,
                   slug, txt)

st.set_page_config(page_title="Despacho Almacén", page_icon="🚚", layout="wide")


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
    store.init()

ss = st.session_state


def flash(msg, tipo="success"):
    ss["_flash"] = (tipo, msg)


def mostrar_flash():
    if "_flash" in ss:
        tipo, msg = ss.pop("_flash")
        getattr(st, tipo)(msg)


# ================================================================ login
def pantalla_login():
    st.title("🚚 Despacho Almacén")
    usuarios = store.read("USUARIOS")
    if usuarios.empty:
        st.info("Primera vez: crea el usuario **administrador**.")
        with st.form("f_admin"):
            u = st.text_input("Usuario").strip().lower()
            n = st.text_input("Nombre")
            c1 = st.text_input("Clave", type="password")
            c2 = st.text_input("Repite la clave", type="password")
            if st.form_submit_button("Crear administrador", type="primary"):
                if not u or len(c1) < 6:
                    st.error("Escribe un usuario y una clave de al menos 6 caracteres.")
                elif c1 != c2:
                    st.error("Las claves no coinciden.")
                else:
                    store.append("USUARIOS", [{"id": u, "nombre": n or u, "clave_hash": hash_clave(c1),
                                               "roles": "admin", "activo": "TRUE", "creado": ahora_str()}])
                    ss["usuario"] = u
                    st.rerun()
        return
    with st.form("f_login"):
        u = st.text_input("Usuario").strip().lower()
        c = st.text_input("Clave", type="password")
        if st.form_submit_button("Ingresar", type="primary"):
            fila = usuarios[usuarios["id"] == u]
            if fila.empty or not es_activo(fila.iloc[0]["activo"]) or not verificar_clave(c, fila.iloc[0]["clave_hash"]):
                st.error("Usuario o clave incorrectos.")
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
    st.markdown(f"### 🚚 Despacho Almacén\n**{NOMBRE}**  \n"
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
c1.markdown("## 🚚 Despacho Almacén")
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
def sec_solicitudes():
    st.subheader("Solicitud de unidad de transporte")
    st.caption("Queda **pendiente** hasta que el área de ruteo la asigne a una ruta.")
    izq, der = st.columns([2, 3], gap="large")
    with izq:
        with st.container(border=True):
            form_movimiento("s")
    with der:
        st.markdown("**Mis solicitudes**")
        sols = store.read("SOLICITUDES")
        mias = sols[sols["solicitante"] == USUARIO].sort_values("creado", ascending=False)
        f1, f2 = st.columns(2)
        est = f1.selectbox("Estado", [""] + list(ESTADOS)[:6], format_func=lambda e: ESTADOS.get(e, "Todos"))
        q = f2.text_input("Buscar", placeholder="Cuenta, cliente, código, GR")
        if est:
            mias = mias[mias["estado"] == est]
        if q:
            mias = mias[(mias["cuenta"] + " " + mias["cliente"] + " " + mias["codigo"] + " " + mias["pedido_gr"])
                        .map(norm).str.contains(norm(q), regex=False)]
        rutas = store.read("RUTAS").set_index("id")
        v = mias.copy()
        v["Estado"] = v["estado"].map(pill_estado)
        v["Ruta"] = v["ruta_id"].map(lambda i: f"{rutas.loc[i, 'codigo']} · {rutas.loc[i, 'chofer']}"
                                     if i in rutas.index else "")
        v["Fecha"] = v["fecha_entrega"].map(fmt_fecha)
        tabla_registros(v, {"codigo": "Código", "proceso": "Proc.", "cuenta": "Cuenta", "cliente": "Cliente",
                            "destino": "Destino", "Fecha": "Entrega", "hora_cita": "Cita", "Estado": "Estado",
                            "Ruta": "Ruta", "nota_entrega": "Nota"}, alto=420)
        pend = mias[mias["estado"] == "pendiente"]
        if not pend.empty:
            with st.expander("Anular una solicitud pendiente"):
                sid = st.selectbox("Solicitud", pend["id"].tolist(),
                                   format_func=dict(zip(pend["id"], pend["codigo"] + " · " + pend["cliente"])).get)
                if st.button("Anular", type="secondary"):
                    L.anular_solicitud(store, sid, USUARIO)
                    flash("Solicitud anulada.")
                    st.rerun()


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
    pend = sols[(sols["estado"] == "pendiente") & ((sols["fecha_entrega"] == fs) | (atras & (sols["fecha_entrega"] < fs)))]
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
                "Fecha": v["fecha_entrega"].map(fmt_fecha), "Cajas": v["cant_caja"],
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
        g1.bar_chart(x.groupby(x["fecha_entrega"].str[:7])["_n"].sum(), color="#1f4fb8")
    else:
        g1.markdown("**Registros por día**")
        g1.bar_chart(x.groupby("fecha_entrega")["_n"].sum(), color="#1f4fb8")
    g2.markdown("**Top cuentas**")
    g2.bar_chart(x.groupby("cuenta")["_n"].sum().sort_values(ascending=False).head(10), horizontal=True,
                 color="#1f4fb8")

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
                        column_config={"id": st.column_config.TextColumn("Usuario", disabled=True)})
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
    a, b = st.columns(2, gap="large")
    with a, st.form("f_nuevo_u", clear_on_submit=True):
        st.markdown("**Nuevo usuario**")
        nu = st.text_input("Usuario (sin espacios)").strip().lower()
        nn = st.text_input("Nombre")
        nc = st.text_input("Clave inicial", type="password")
        nr = st.multiselect("Roles", list(ROLES), format_func=ROLES.get)
        if st.form_submit_button("Crear usuario", type="primary"):
            if not nu or " " in nu or len(nc) < 6:
                st.error("Usuario sin espacios y clave de al menos 6 caracteres.")
            elif nu in set(u["id"]):
                st.error("Ese usuario ya existe.")
            else:
                store.append("USUARIOS", [{"id": nu, "nombre": nn or nu, "clave_hash": hash_clave(nc),
                                           "roles": ",".join(nr), "activo": "TRUE", "creado": ahora_str()}])
                flash(f"Usuario {nu} creado.")
                st.rerun()
    with b, st.form("f_clave", clear_on_submit=True):
        st.markdown("**Restablecer clave**")
        cu_ = st.selectbox("Usuario", u["id"].tolist())
        c1 = st.text_input("Nueva clave", type="password")
        if st.form_submit_button("Cambiar clave"):
            if len(c1) < 6:
                st.error("La clave debe tener al menos 6 caracteres.")
            else:
                store.update("USUARIOS", cu_, {"clave_hash": hash_clave(c1)})
                flash("Clave actualizada.")
                st.rerun()


# ================================================================ router
{
    "solicitudes": sec_solicitudes, "regularizacion": sec_regularizacion, "movcliente": sec_movcliente,
    "ruteo": sec_ruteo, "despacho": sec_despacho, "reportes": sec_reportes, "maestros": sec_maestros,
    "usuarios": sec_usuarios,
}[ss["seccion"]]()
