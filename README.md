# Despacho Almacén

Sistema de solicitudes de transporte, ruteo, despacho y base de entregas.
Stack: **Python + Streamlit** (Streamlit Community Cloud), datos en **Google Sheets**, código en **GitHub**.
Misma arquitectura que *Picking Subcedis*.

## Flujo

```
Solicitud (OUT/IN) → Ruteo (código de ruta + n.° de viaje) → Despacho (salida, entregas, cierre) → BASE DE ENTREGAS
          ↑ Regularización de despacho      ↑ Recepción / despacho gestionado por el cliente
```

## Archivos

| Archivo | Qué hace |
|---|---|
| `app.py` | La app. Login, menú por rol y una función por sección. |
| `db.py` | Backend SQLite local (si no hay credenciales de Google). |
| `sheets_db.py` | Backend Google Sheets (se usa si existe `st.secrets["gcp_oauth"]`). Misma interfaz que `db.py`. |
| `schema.py` | Tablas (= pestañas del Sheet) y columnas de la BASE DE ENTREGAS. |
| `logic.py` | Reglas de negocio: crear solicitud, crear/anular ruta, salida, resultados, cierre, armar la base. |
| `report.py` | Excel descargables: BASE DE ENTREGAS y HOJA DE CHOFERES. |
| `parser_plantilla.py` | Lee tu Excel (hojas MATRICES y RUTEO) para cargar cuentas, clientes e histórico. |
| `auth.py` | Claves hasheadas (PBKDF2). |
| `utils.py` | Hora de Perú (`_ahora()`, `_hora_actual()`), códigos, normalización. |

## Pestañas del Google Sheet (se crean solas)

`USUARIOS, CUENTAS, CLIENTES, TIPOS_TRANSPORTE, CHOFERES, SOLICITUDES, RUTAS, REGULARIZACIONES, MOV_CLIENTE, HISTORICO`

Todo se escribe con `value_input_option="RAW"` y columnas en formato TEXTO para que los códigos (GR, N° orden) no pierdan ceros.

## Puesta en marcha

1. **Google Sheet**: crea un Sheet vacío llamado *Despacho Almacén* y copia su ID (lo que va entre `/d/` y `/edit` en la URL).
2. **GitHub**: crea el repo (por ejemplo `CFS1-tech/despacho-almacen`, **privado**) y sube todos los archivos de esta carpeta (incluida la carpeta `.streamlit` con `config.toml`).
3. **Streamlit Cloud**: *New app* → elige el repo → archivo principal `app.py`.
4. **Secrets** (Settings → Secrets): pega el contenido de `.streamlit/secrets.toml.example` con tus datos.
   Puedes reutilizar el `client_id`, `client_secret` y `refresh_token` de Picking Subcedis.
5. Abre la app: la primera vez te pide crear el **usuario administrador**.
6. **Maestros → Importar Excel**:
   - «Cargar tipos de transporte y choferes por defecto».
   - Sube tu `PROPUESTA_PLANTILLA.xlsx` → «Agregar cuentas y clientes nuevos» y «Cargar histórico».
7. **Usuarios**: crea a cada persona con su clave inicial y marca sus roles.

## Roles

| Rol | Ve |
|---|---|
| Solicitante | Solicitud de transporte, Regularización, Recepción/despacho del cliente |
| Ruteador | Lo anterior + Ruteo, Despacho, Base y reportes |
| Almacén / Despacho | Despacho, Regularización, Recepción/despacho del cliente |
| Supervisor | Solicitudes, Base y reportes, Maestros |
| Administrador | Todo + Usuarios |

## Probar en tu computadora (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py        # sin secrets usa SQLite (despacho.db)
```
