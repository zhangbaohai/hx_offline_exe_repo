
import sqlite3
from pathlib import Path

def ensure_db(db_path: str):
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS ibps (
        code TEXT PRIMARY KEY,
        name TEXT,
        raw_line TEXT,
        source TEXT,
        updated_at TEXT DEFAULT (datetime('now'))
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS cnaps (
        code TEXT PRIMARY KEY,
        name TEXT,
        raw_line TEXT,
        source TEXT,
        updated_at TEXT DEFAULT (datetime('now'))
    )""")
    conn.commit(); conn.close()

def upsert_many_simple(db_path: str, table: str, rows: list):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    sql = f"""INSERT INTO {table}(code, name, raw_line, source)
              VALUES (?,?,?,?)
              ON CONFLICT(code) DO UPDATE SET
                name=excluded.name,
                raw_line=excluded.raw_line,
                source=excluded.source,
                updated_at=datetime('now')"""
    cur.executemany(sql, rows)
    conn.commit(); conn.close()

def replace_all(db_path: str, table: str, rows: list):
    conn = sqlite3.connect(db_path); cur = conn.cursor()
    cur.execute(f"DELETE FROM {table}")
    conn.commit(); conn.close()
    upsert_many_simple(db_path, table, rows)

def query(db_path: str, table: str, keyword: str, limit: int = 1000):
    conn = sqlite3.connect(db_path); conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    kw = f"%{keyword}%"
    cur.execute(f"""SELECT code, name FROM {table}
                    WHERE code LIKE ? OR name LIKE ?
                    ORDER BY name LIMIT ?""", (kw, kw, limit))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def lookup_by_bankname(db_path: str, table: str, bank_name: str):
    conn = sqlite3.connect(db_path); conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(f"""SELECT code, name FROM {table}
                    WHERE name LIKE ? ORDER BY LENGTH(name) DESC LIMIT 1""",
                (f"%{bank_name}%",))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None
