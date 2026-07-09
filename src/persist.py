# src/persist.py
"""Database persistence helpers for the Excel Formula Analyzer.

The module expects PostgreSQL connection parameters via environment variables in
`config/.env` (loaded by `python-dotenv` in the calling scripts).
"""

import json
import os
from typing import Any, Dict

import psycopg2
from psycopg2.extras import Json

# NOTE: Environment variables should already be loaded by the caller (e.g., tracker.py).


def _get_connection():
    """Create a new psycopg2 connection using env vars.

    Returns
    -------
    psycopg2.extensions.connection
    """
    host = os.getenv("DB_HOST", "localhost")
    dbname = os.getenv("DB_NAME")
    user = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")
    port = os.getenv("DB_PORT", "5432")
    if not all([dbname, user, password]):
        raise RuntimeError(
            "Database credentials missing. Ensure DB_HOST, DB_NAME, DB_USER, DB_PASSWORD are set in .env or the environment."
        )
    # Debug output for connection parameters (remove in production)
    print(f"[DB] Connecting to {user}@{host}:{port}/{dbname}")
    conn = psycopg2.connect(host=host, dbname=dbname, user=user, password=password, port=port)
    return conn


def save_workbook_record(file_name: str, version: str, data: Dict[str, Any]):
    """Insert a new record into ``workbook_records``.

    Parameters
    ----------
    file_name: str
        Name of the Excel file (basename).
    version: str
        Arbitrary version identifier – you can derive it from the filename or pass any string.
    data: dict
        JSON‑serialisable payload containing column mappings, formulas, etc.
    """
    conn = _get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO workbook_records (file_name, version, data)
                VALUES (%s, %s, %s)
                """,
                (file_name, version, Json(data))
            )
        conn.commit()
        print(f"[DB] Inserted record for {file_name} version {version}")
    finally:
        conn.close()


def get_record_count() -> int:
    """Return the total number of rows in workbook_records.
    Useful for quick verification after an INSERT.
    """
    conn = _get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM workbook_records")
            count = cur.fetchone()[0]
        return count
    finally:
        conn.close()
