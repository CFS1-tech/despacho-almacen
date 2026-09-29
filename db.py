"""Backend SQLite local (se usa cuando no hay credenciales de Google en los Secrets).

Misma interfaz que sheets_db.py:
    init(), read(tabla), append(tabla, filas), update(tabla, id, campos),
    update_many(tabla, {id: campos}), delete(tabla, id), replace_all(tabla, df), clear_cache()
"""
import sqlite3
import threading

import pandas as pd

from schema import TABLES

DB_PATH = "despacho.db"
BACKEND = "SQLite local"
_lock = threading.Lock()


def _conn():
    c = sqlite3.connect(DB_PATH, check_same_thread=False)
    return c


def init():
    with _lock, _conn() as c:
        for t, cols in TABLES.items():
            c.execute(f'CREATE TABLE IF NOT EXISTS "{t}" ({", ".join(f"{k} TEXT" for k in cols)})')
            existentes = {r[1] for r in c.execute(f'PRAGMA table_info("{t}")')}
            for k in cols:
                if k not in existentes:
                    c.execute(f'ALTER TABLE "{t}" ADD COLUMN {k} TEXT')


def clear_cache():
    pass


def read(tabla: str) -> pd.DataFrame:
    cols = TABLES[tabla]
    with _lock, _conn() as c:
        df = pd.read_sql_query(f'SELECT {", ".join(cols)} FROM "{tabla}"', c)
    return df.fillna("").astype(str)


def _fila(tabla, d):
    return [("" if d.get(k) is None else str(d.get(k))) for k in TABLES[tabla]]


def append(tabla: str, filas: list):
    if not filas:
        return
    cols = TABLES[tabla]
    with _lock, _conn() as c:
        c.executemany(
            f'INSERT INTO "{tabla}" ({", ".join(cols)}) VALUES ({", ".join("?" * len(cols))})',
            [_fila(tabla, f) for f in filas],
        )


def update_many(tabla: str, cambios: dict):
    if not cambios:
        return
    with _lock, _conn() as c:
        for id_, campos in cambios.items():
            campos = {k: v for k, v in campos.items() if k in TABLES[tabla] and k != "id"}
            if not campos:
                continue
            sets = ", ".join(f"{k} = ?" for k in campos)
            c.execute(f'UPDATE "{tabla}" SET {sets} WHERE id = ?',
                      [("" if v is None else str(v)) for v in campos.values()] + [str(id_)])


def update(tabla: str, id_: str, campos: dict):
    update_many(tabla, {id_: campos})


def delete(tabla: str, id_: str):
    with _lock, _conn() as c:
        c.execute(f'DELETE FROM "{tabla}" WHERE id = ?', [str(id_)])


def replace_all(tabla: str, df: pd.DataFrame):
    cols = TABLES[tabla]
    df = df.reindex(columns=cols).fillna("").astype(str)
    with _lock, _conn() as c:
        c.execute(f'DELETE FROM "{tabla}"')
        c.executemany(
            f'INSERT INTO "{tabla}" ({", ".join(cols)}) VALUES ({", ".join("?" * len(cols))})',
            df.values.tolist(),
        )
