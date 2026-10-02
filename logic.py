"""Lógica de negocio. Recibe el backend (`store` = db o sheets_db) y nunca sabe cuál es."""
import pandas as pd

from schema import BASE_COLS, ESTADOS, PEDIDO, TABLES
from utils import (ALMACEN, _hora_actual, ahora_str, es_activo, fmt_fecha, norm, nuevo_codigo, nuevo_id, num0,
                   telefono_de, txt)


# ---------------------------------------------------------------- maestros
def activos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    return df[df["activo"].map(es_activo)]


def categoria_de(tipos: pd.DataFrame, t_transporte: str) -> str:
    t = txt(t_transporte)
    if not t:
        return ""
    fila = tipos[tipos["nombre"].map(norm) == norm(t)]
    if not fila.empty:
        return fila.iloc[0]["categoria"]
    tu = t.upper()
    if "SUPPLY" in tu:
        return "SUPPLY"
    if "CLIENTE" in tu:
        return "CLIENTE"
    if any(x in tu for x in ("AGENCIA", "OLTURSA", "CRUZ DEL SUR")):
        return "AGENCIA"
    return "TERCERIZADO"


def nombres_usuarios(store) -> dict:
    u = store.read("USUARIOS")
    return {r["id"]: (r["nombre"] or r["id"]) for _, r in u.iterrows()}


def nuevo_grupo() -> str:
    return nuevo_codigo("P")


def aprender_cliente(store, mov: dict, contar: bool = True):
    """Si el cliente no existe en la cuenta lo crea. Si existe, guarda la dirección y el destino
    usados (ubicaciones fijas: la última registrada queda como la habitual) y completa datos vacíos."""
    cuenta_id, cliente = mov.get("cuenta_id"), txt(mov.get("cliente"))
    if not cuenta_id or not cliente:
        return
    lugar = mov.get("destino") if mov.get("proceso") == "OUT" else mov.get("origen")
    lugar = "" if txt(lugar).upper() in ("", ALMACEN, "CLIENTE") else txt(lugar).upper()
    cl = store.read("CLIENTES")
    ex = cl[(cl["cuenta_id"] == cuenta_id) & (cl["cliente"].map(norm) == norm(cliente))]
    if ex.empty:
        store.append("CLIENTES", [{
            "id": nuevo_id(), "cuenta_id": cuenta_id, "cuenta": mov.get("cuenta", ""), "cliente": cliente,
            "direccion": mov.get("direccion", ""), "destino": lugar, "contacto": "",
            "referencia": mov.get("referencia", ""), "usos": "1", "activo": "TRUE",
        }])
        return
    r = ex.iloc[0]
    cambios = {"usos": str(int(num0(r["usos"])) + 1)} if contar else {}
    for k, v in (("direccion", mov.get("direccion")), ("destino", lugar)):
        if txt(v) and txt(v) != txt(r[k]):
            cambios[k] = txt(v)
    for k, v in (("referencia", mov.get("referencia")),):
        if not txt(r[k]) and txt(v):
            cambios[k] = txt(v)
    if cambios:
        store.update("CLIENTES", r["id"], cambios)


# ---------------------------------------------------------------- registros
def crear_movimiento(store, tabla: str, datos: dict, usuario: str) -> dict:
    """Crea una solicitud, regularización o registro de cliente."""
    prefijo = {"SOLICITUDES": "S", "REGULARIZACIONES": "RG"}.get(tabla) or ("RC" if datos.get("proceso") == "IN" else "DC")
    fila = {k: "" for k in TABLES[tabla]}
    fila.update({k: ("" if v is None else v) for k, v in datos.items() if k in fila})
    fila.update({"id": nuevo_id(), "codigo": nuevo_codigo(prefijo), "creado": ahora_str()})
    if tabla == "SOLICITUDES":
        fila.update({"estado": "pendiente", "solicitante": usuario, "actualizado": ahora_str()})
    else:
        fila["registrado_por"] = usuario
    store.append(tabla, [fila])
    try:
        aprender_cliente(store, fila)
    except Exception:
        pass
    return fila


def anular_solicitud(store, sol_id: str, usuario: str):
    store.update("SOLICITUDES", sol_id, {"estado": "anulado", "nota_entrega": f"Anulado por {usuario}",
                                         "actualizado": ahora_str()})


# ---------------------------------------------------------------- ruteo
def viajes_del_chofer(rutas: pd.DataFrame, chofer_id: str, fecha: str) -> int:
    if rutas.empty:
        return 0
    r = rutas[(rutas["chofer_id"] == chofer_id) & (rutas["fecha"] == fecha) & (rutas["estado"] != "anulada")]
    return len(r)


def crear_ruta(store, sol_ids: list, fecha: str, t_transporte: str, chofer_id: str, chofer: str, placa: str,
               salida: str, notas: str, usuario: str) -> dict:
    sols = store.read("SOLICITUDES").set_index("id")
    ids = [i for i in sol_ids if i in sols.index and sols.loc[i, "estado"] == "pendiente"]
    if not ids:
        raise ValueError("Las solicitudes elegidas ya no están pendientes.")
    rutas = store.read("RUTAS")
    ruta = {k: "" for k in TABLES["RUTAS"]}
    ruta.update({
        "id": nuevo_id(), "codigo": nuevo_codigo("R"), "fecha": fecha, "hora_salida_plan": salida,
        "t_transporte": t_transporte, "chofer_id": chofer_id, "chofer": chofer, "placa": placa.upper(),
        "viaje": str(viajes_del_chofer(rutas, chofer_id, fecha) + 1), "paradas": ",".join(ids),
        "estado": "programada", "notas": notas, "cant_caja": str(sum(num0(sols.loc[i, "cant_caja"]) for i in ids)),
        "creado_por": usuario, "creado": ahora_str(),
    })
    store.append("RUTAS", [ruta])
    store.update_many("SOLICITUDES", {i: {"estado": "ruteado", "ruta_id": ruta["id"], "orden_parada": str(n + 1),
                                          "actualizado": ahora_str()} for n, i in enumerate(ids)})
    return ruta


def paradas(ruta) -> list:
    return [p for p in txt(ruta["paradas"]).split(",") if p]


def anular_ruta(store, ruta):
    store.update("RUTAS", ruta["id"], {"estado": "anulada"})
    store.update_many("SOLICITUDES", {i: {"estado": "pendiente", "ruta_id": "", "orden_parada": "",
                                          "actualizado": ahora_str()} for i in paradas(ruta)})


def quitar_parada(store, ruta, sol_id: str):
    quedan = [p for p in paradas(ruta) if p != sol_id]
    sols = store.read("SOLICITUDES").set_index("id")
    cajas = sum(num0(sols.loc[i, "cant_caja"]) for i in quedan if i in sols.index)
    store.update("RUTAS", ruta["id"], {"paradas": ",".join(quedan), "cant_caja": str(cajas)})
    store.update("SOLICITUDES", sol_id, {"estado": "pendiente", "ruta_id": "", "orden_parada": "",
                                         "actualizado": ahora_str()})


def registrar_salida(store, ruta, usuario: str):
    hora = _hora_actual()
    store.update("RUTAS", ruta["id"], {"estado": "en_ruta", "salida_real": hora, "despachado_por": usuario})
    store.update_many("SOLICITUDES", {i: {"estado": "en_ruta", "actualizado": ahora_str()} for i in paradas(ruta)})
    return hora


def registrar_resultados(store, resultados: dict):
    """resultados = {sol_id: {"estado": ..., "hora_llegada": ..., "nota_entrega": ...}}"""
    for v in resultados.values():
        v["actualizado"] = ahora_str()
    store.update_many("SOLICITUDES", resultados)


def cerrar_ruta(store, ruta):
    store.update("RUTAS", ruta["id"], {"estado": "completada", "cierre": _hora_actual()})


def texto_hoja_ruta(ruta, sols: pd.DataFrame) -> str:
    s = sols.set_index("id")
    lineas = [f"HOJA DE RUTA {ruta['codigo']} — {fmt_fecha(ruta['fecha'])}",
              f"Transporte: {ruta['t_transporte']}" + (f" · {ruta['placa']}" if ruta["placa"] else ""),
              f"Chofer: {ruta['chofer']} (viaje {ruta['viaje'] or 1})",
              f"Salida: {ruta['hora_salida_plan']}" if ruta["hora_salida_plan"] else "", ""]
    for n, i in enumerate(paradas(ruta), 1):
        if i not in s.index:
            continue
        x = s.loc[i]
        lineas += [f"{n}. [{x['proceso']}] {x['cuenta']} - {x['cliente']}",
                   f"   {x['direccion']}" + (f" (Ref: {x['referencia']})" if x["referencia"] else ""),
                   f"   Cita: {x['hora_cita'] or '—'}" + (f" · {x['cant_caja']} cajas" if x["cant_caja"] else "")
                   + (f" · GR {x['pedido_gr']}" if x["pedido_gr"] else ""),
                   f"   Contacto: {x['contacto'].replace(chr(10), ' / ')}", ""]
    if ruta["notas"]:
        lineas.append("Notas: " + ruta["notas"])
    return "\n".join(lineas)


# ---------------------------------------------------------------- base de entregas
def base_entregas(store, desde: str, hasta: str, cuenta: str = "", proceso: str = "",
                  incluir_historico: bool = True) -> pd.DataFrame:
    """Une solicitudes, regularizaciones, registros de cliente e histórico con las columnas de tu Excel."""
    tipos = store.read("TIPOS_TRANSPORTE")
    usuarios = nombres_usuarios(store)
    rutas = store.read("RUTAS").set_index("id")
    filas = []

    def ok(fecha, cu, pr):
        return desde <= fecha <= hasta and (not cuenta or norm(cu) == norm(cuenta)) and (not proceso or pr == proceso)

    def ped(x):
        return {k: x[k] for k in PEDIDO}

    for _, s in store.read("SOLICITUDES").iterrows():
        r = rutas.loc[s["ruta_id"]] if s["ruta_id"] in rutas.index else None
        # "hasta una fecha": si ya tiene ruta, la fecha real es la de la ruta
        fecha = r["fecha"] if (r is not None and s["tipo_fecha"] == "hasta") else s["fecha_entrega"]
        if s["estado"] == "anulado" or not ok(fecha, s["cuenta"], s["proceso"]):
            continue
        t = (r["t_transporte"] if r is not None else "") or s["t_transporte_sol"]
        rev = {"entregado": "OK", "no_entregado": f"NO ENTREGADO: {s['nota_entrega']}"}.get(
            s["estado"], ESTADOS.get(s["estado"], s["estado"]).upper())
        filas.append({
            "fecha_entrega": fecha, "cuenta": s["cuenta"], "cliente": s["cliente"],
            "hora_cita": s["hora_cita"], "hora_llegada": s["hora_llegada"], "proceso": s["proceso"],
            "origen": s["origen"], "destino": s["destino"], "t_transporte": t,
            "chofer": r["chofer"] if r is not None else "", "direccion": s["direccion"],
            "contacto": s["contacto"].replace("\n", " / "), "telefono": telefono_de(s["contacto"]),
            "referencia": s["referencia"], **ped(s), "observacion": s["observacion"],
            "encargado": usuarios.get(r["despachado_por"], r["despachado_por"]) if r is not None else "",
            "transporte": categoria_de(tipos, t), "revision": rev, "codigo_despacho": s["codigo"],
            "codigo_ruta": r["codigo"] if r is not None else "", "viaje": r["viaje"] if r is not None else "",
            "registro": "Solicitud", "estado": s["estado"],
        })
    for _, g in store.read("REGULARIZACIONES").iterrows():
        if not ok(g["fecha_entrega"], g["cuenta"], g["proceso"]):
            continue
        filas.append({
            "fecha_entrega": g["fecha_entrega"], "cuenta": g["cuenta"], "cliente": g["cliente"],
            "hora_cita": g["hora_cita"], "hora_llegada": g["hora_llegada"], "proceso": g["proceso"],
            "origen": g["origen"], "destino": g["destino"], "t_transporte": g["t_transporte"], "chofer": g["chofer"],
            "direccion": g["direccion"], "contacto": g["contacto"].replace("\n", " / "),
            "telefono": telefono_de(g["contacto"]), "referencia": g["referencia"], **ped(g),
            "observacion": " · ".join(x for x in (g["observacion"], f"Regularización: {g['motivo']}" if g["motivo"] else "") if x),
            "encargado": usuarios.get(g["registrado_por"], g["registrado_por"]),
            "transporte": categoria_de(tipos, g["t_transporte"]), "revision": "REGULARIZADO",
            "codigo_despacho": g["codigo"], "codigo_ruta": "", "viaje": "", "registro": "Regularización",
            "estado": "regularizado",
        })
    for _, m in store.read("MOV_CLIENTE").iterrows():
        if not ok(m["fecha"], m["cuenta"], m["proceso"]):
            continue
        filas.append({
            "fecha_entrega": m["fecha"], "cuenta": m["cuenta"], "cliente": m["cliente"], "hora_cita": m["hora"],
            "hora_llegada": m["hora"], "proceso": m["proceso"], "origen": m["origen"], "destino": m["destino"],
            "t_transporte": "CLIENTE", "chofer": "Cliente", "direccion": "",
            "contacto": m["conductor"].replace("\n", " / "), "telefono": telefono_de(m["conductor"]),
            "referencia": " · ".join(x for x in (m["empresa"], m["placa"]) if x), **ped(m),
            "observacion": m["observacion"], "encargado": usuarios.get(m["registrado_por"], m["registrado_por"]),
            "transporte": "CLIENTE", "revision": "OK", "codigo_despacho": m["codigo"], "codigo_ruta": "", "viaje": "",
            "registro": "Cliente", "estado": "ok",
        })
    if incluir_historico:
        h = store.read("HISTORICO")
        if not h.empty:
            h = h[(h["fecha_entrega"] >= desde) & (h["fecha_entrega"] <= hasta)]
            if cuenta:
                h = h[h["cuenta"].map(norm) == norm(cuenta)]
            if proceso:
                h = h[h["proceso"] == proceso]
            h = h.assign(codigo_despacho="", codigo_ruta="", registro="Histórico", estado="historico")
            h["transporte"] = [t or categoria_de(tipos, tt) for t, tt in zip(h["transporte"], h["t_transporte"])]
            filas += h.to_dict("records")
    cols = [k for k, _ in BASE_COLS]
    df = pd.DataFrame(filas, columns=cols).fillna("")
    return df.sort_values(["fecha_entrega", "codigo_ruta", "hora_cita"], kind="stable").reset_index(drop=True)


def clave_viaje(fila) -> str:
    """Identifica una ruta/viaje también en el histórico (fecha + chofer + n.° de viaje)."""
    if fila["codigo_ruta"]:
        return fila["codigo_ruta"]
    ch = txt(fila["chofer"]).upper()
    if fila["registro"] == "Histórico" and ch and ch not in ("CLIENTE", "AGENCIA"):
        return f"H|{fila['fecha_entrega']}|{ch}|{fila['viaje'] or 1}"
    return ""
