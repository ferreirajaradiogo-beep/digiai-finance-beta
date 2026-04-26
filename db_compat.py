import os
import re
import sqlite3
from pathlib import Path

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:  # pragma: no cover - optional dependency in local sqlite mode
    psycopg = None
    dict_row = None


DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
DB_BACKEND = "postgres" if DATABASE_URL.startswith(("postgres://", "postgresql://")) else "sqlite"

if psycopg:
    DB_INTEGRITY_ERRORS = (sqlite3.IntegrityError, psycopg.IntegrityError)
else:
    DB_INTEGRITY_ERRORS = (sqlite3.IntegrityError,)


def _normalize_postgres_query(query):
    normalized = query

    normalized = re.sub(
        r"INSERT\s+OR\s+IGNORE\s+INTO\s+categorias",
        "INSERT INTO categorias",
        normalized,
        flags=re.IGNORECASE,
    )
    if "INSERT INTO categorias" in normalized and "ON CONFLICT" not in normalized:
        normalized = normalized.rstrip() + " ON CONFLICT (usuario, nome) DO NOTHING"

    normalized = re.sub(
        r"INSERT\s+OR\s+IGNORE\s+INTO\s+contas",
        "INSERT INTO contas",
        normalized,
        flags=re.IGNORECASE,
    )
    if "INSERT INTO contas" in normalized and "ON CONFLICT" not in normalized:
        normalized = normalized.rstrip() + " ON CONFLICT (usuario, nome) DO NOTHING"

    normalized = normalized.replace("ORDER BY nome COLLATE NOCASE", "ORDER BY LOWER(nome)")
    normalized = normalized.replace("date('now')", "CAST(CURRENT_DATE AS TEXT)")
    normalized = normalized.replace("strftime('%Y', data)", "LEFT(data, 4)")
    normalized = normalized.replace("?", "%s")
    return normalized


class CompatCursor:
    def __init__(self, cursor, backend):
        self._cursor = cursor
        self._backend = backend

    def execute(self, query, params=None):
        final_query = _normalize_postgres_query(query) if self._backend == "postgres" else query
        if params is None:
            self._cursor.execute(final_query)
        else:
            self._cursor.execute(final_query, params)
        return self

    def executemany(self, query, seq_of_params):
        final_query = _normalize_postgres_query(query) if self._backend == "postgres" else query
        self._cursor.executemany(final_query, seq_of_params)
        return self

    def fetchone(self):
        return self._cursor.fetchone()

    def fetchall(self):
        return self._cursor.fetchall()

    @property
    def lastrowid(self):
        return getattr(self._cursor, "lastrowid", None)

    def __getattr__(self, item):
        return getattr(self._cursor, item)


class CompatConnection:
    def __init__(self, conn, backend):
        self._conn = conn
        self._backend = backend

    def cursor(self):
        return CompatCursor(self._conn.cursor(), self._backend)

    def commit(self):
        return self._conn.commit()

    def close(self):
        return self._conn.close()

    def rollback(self):
        return self._conn.rollback()

    def __getattr__(self, item):
        return getattr(self._conn, item)


def connect(db_path):
    if DB_BACKEND == "postgres":
        if not psycopg:
            raise RuntimeError("psycopg is required when DATABASE_URL is configured.")
        conn = psycopg.connect(DATABASE_URL, row_factory=dict_row)
        return CompatConnection(conn, DB_BACKEND)

    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return CompatConnection(conn, DB_BACKEND)


def get_table_columns(cursor, table_name):
    if DB_BACKEND == "postgres":
        cursor.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = %s
            """,
            (table_name,),
        )
        return {row["column_name"] for row in cursor.fetchall()}

    cursor.execute(f"PRAGMA table_info({table_name})")
    return {row["name"] for row in cursor.fetchall()}
