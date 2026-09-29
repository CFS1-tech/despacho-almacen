"""Utilidades comunes: hora de Perú, códigos, normalización de texto y números."""
import random
import re
import string
import unicodedata
import uuid
from datetime import date, datetime, timedelta, timezone

TZ_PERU = timezone(timedelta(hours=-5))  # Lima no usa horario de verano
ALMACEN = "VES SERATRA"


def _ahora() -> datetime:
    """Fecha y hora actual en Perú (Streamlit Cloud corre en UTC)."""
    return datetime.now(TZ_PERU)


def hoy() -> date:
    return _ahora().date()


def ahora_str() -> str:
    return _ahora().strftime("%Y-%m-%d %H:%M:%S")


def _hora_actual() -> str:
    return _ahora().strftime("%H:%M")


def nuevo_id() -> str:
    return uuid.uuid4().hex[:12]


def nuevo_codigo(prefijo: str) -> str:
    """Código legible tipo S-260928-A1K3 (prefijo, fecha, 4 caracteres)."""
    sufijo = "".join(random.choices(string.ascii_uppercase + string.digits, k=4))
    return f"{prefijo}-{_ahora():%y%m%d}-{sufijo}"


def norm(s) -> str:
    """Texto sin tildes, en minúsculas y con espacios simples (para comparar)."""
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def slug(s) -> str:
    return re.sub(r"[^a-z0-9]+", "-", norm(s)).strip("-")[:60]


def num(v):
    """Convierte a número; devuelve None si está vacío o no es número."""
    if v is None:
        return None
    s = str(v).strip().replace(",", ".")
    if s == "" or s.lower() == "nan":
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return int(n) if n.is_integer() else n


def num0(v) -> float:
    n = num(v)
    return n if n is not None else 0


def txt(v) -> str:
    """Valor como texto limpio (None/NaN -> '')."""
    if v is None:
        return ""
    s = str(v)
    return "" if s.lower() in ("nan", "none", "nat") else s.strip()


def fmt_fecha(s: str) -> str:
    """'2026-09-28' -> '28/09/2026'."""
    s = txt(s)
    if len(s) >= 10 and s[4] == "-":
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s


def a_fecha(s):
    """Texto 'YYYY-MM-DD' -> date (o None)."""
    try:
        return datetime.strptime(txt(s)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def hora_str(t) -> str:
    """datetime.time / texto -> 'HH:MM'."""
    if t is None:
        return ""
    if hasattr(t, "strftime"):
        return t.strftime("%H:%M")
    return txt(t)[:5]


TEL_RE = re.compile(r"(?:\+?51\s?)?(?:9\d{2}[\s-]?\d{3}[\s-]?\d{3}|0?1\s?\d{3}\s?\d{4})")


def telefono_de(texto) -> str:
    m = TEL_RE.search(txt(texto))
    return m.group(0) if m else ""


def es_activo(v) -> bool:
    return txt(v).upper() not in ("FALSE", "0", "NO", "N")
