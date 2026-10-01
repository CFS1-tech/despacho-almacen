"""Definición de tablas (= pestañas del Google Sheet). Todas las columnas se guardan como TEXTO.

La primera columna de cada tabla es `id` (clave única), salvo USUARIOS donde `id` es el nombre de usuario.
"""

# Campos del detalle del pedido, iguales a las columnas de tu BASE DE ENTREGAS
PEDIDO = ["pedido_gr", "n_orden", "cant_caja", "cant_inner", "accesorio", "maquina", "x_und_acc", "palets", "unidades"]
PEDIDO_LABELS = {
    "pedido_gr": "Pedido - GR", "n_orden": "N.° orden", "cant_caja": "Cant. de cajas",
    "cant_inner": "Cant. inner", "accesorio": "Accesorio", "maquina": "Máquina",
    "x_und_acc": "X und. acc.", "palets": "Palets", "unidades": "Unidades",
}
PEDIDO_TEXTO = {"pedido_gr", "n_orden"}

TABLES = {
    "USUARIOS": ["id", "nombre", "clave_hash", "roles", "activo", "creado"],
    "CUENTAS": ["id", "nombre", "activo"],
    "CLIENTES": ["id", "cuenta_id", "cuenta", "cliente", "direccion", "destino", "contacto", "referencia", "usos", "activo"],
    "TIPOS_TRANSPORTE": ["id", "nombre", "categoria", "activo"],
    "CHOFERES": ["id", "nombre", "tipo", "dni", "telefono", "activo"],
    "SOLICITUDES": [
        "id", "codigo", "proceso", "cuenta_id", "cuenta", "cliente_id", "cliente", "direccion", "referencia",
        "origen", "destino", "fecha_entrega", "hora_cita", "t_transporte_sol", "contacto", *PEDIDO,
        "observacion", "estado", "ruta_id", "orden_parada", "hora_llegada", "nota_entrega",
        "solicitante", "creado", "actualizado", "tipo_fecha", "grupo",
    ],
    "RUTAS": [
        "id", "codigo", "fecha", "hora_salida_plan", "t_transporte", "chofer_id", "chofer", "placa", "viaje",
        "paradas", "estado", "notas", "cant_caja", "salida_real", "cierre", "creado_por", "despachado_por", "creado",
    ],
    "REGULARIZACIONES": [
        "id", "codigo", "proceso", "cuenta_id", "cuenta", "cliente", "direccion", "referencia", "origen", "destino",
        "fecha_entrega", "hora_cita", "hora_llegada", "t_transporte", "chofer", "placa", "contacto", *PEDIDO,
        "motivo", "observacion", "registrado_por", "creado",
    ],
    "MOV_CLIENTE": [
        "id", "codigo", "proceso", "cuenta_id", "cuenta", "cliente", "origen", "destino", "fecha", "hora",
        "empresa", "placa", "conductor", *PEDIDO, "observacion", "registrado_por", "creado",
    ],
    # Histórico importado de la hoja RUTEO (mismas columnas que la base)
    "HISTORICO": [
        "fecha_entrega", "cuenta", "cliente", "hora_cita", "hora_llegada", "proceso", "origen", "destino",
        "t_transporte", "chofer", "direccion", "contacto", "telefono", "referencia", *PEDIDO,
        "observacion", "encargado", "transporte", "revision", "viaje",
    ],
}

# Columnas de la BASE DE ENTREGAS exportada (clave interna -> encabezado de tu Excel)
BASE_COLS = [
    ("fecha_entrega", "FECHA ENTREGA"), ("cuenta", "CUENTA"), ("cliente", "CLIENTE"), ("hora_cita", "HORA DE CITA"),
    ("hora_llegada", "HORA DE LLEGADA"), ("proceso", "PROCESO"), ("origen", "ORIGEN"), ("destino", "DESTINO"),
    ("t_transporte", "T  TRANSPORTE"), ("chofer", "CHOFER"), ("direccion", "DIRECCION"), ("contacto", "CONTACTO"),
    ("telefono", "TELEFONO"), ("referencia", "REFERENCIA"), ("pedido_gr", "PEDIDO - GR"), ("n_orden", "N°ORDEN"),
    ("cant_caja", "CANT DE CAJA"), ("cant_inner", "CANT INNER"), ("accesorio", "ACSESORIO"), ("maquina", "MAQUINA"),
    ("x_und_acc", "X UND ACC"), ("palets", "PALET'S"), ("unidades", "UNIDADES"), ("observacion", "OBSERVACION"),
    ("encargado", "ENCARGADO"), ("transporte", "TRANSPORTE"), ("revision", "REVISIÓN"),
    ("codigo_despacho", "CÓDIGO DESPACHO"), ("codigo_ruta", "CÓDIGO RUTA"), ("viaje", "VIAJE"),
    ("registro", "REGISTRO"), ("estado", "ESTADO"),
]

ROLES = {
    "solicitante": "Solicitante",
    "ruteador": "Ruteador",
    "almacen": "Almacén / Despacho",
    "supervisor": "Supervisor",
    "admin": "Administrador",
}

ESTADOS = {
    "pendiente": "Pendiente", "ruteado": "Ruteado", "en_ruta": "En ruta", "entregado": "Entregado",
    "no_entregado": "No entregado", "anulado": "Anulado",
    "programada": "Programada", "completada": "Completada", "anulada": "Anulada",
}

CATEGORIAS = ["SUPPLY", "TERCERIZADO", "AGENCIA", "CLIENTE"]
TIPOS_CHOFER = ["Chofer propio", "Chofer tercerizado", "Agencia"]
