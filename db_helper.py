
import sqlite3
from pathlib import Path
from typing import Iterable, List, Tuple

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

def upsert_many_batched(db_path: str, table: str, rows: Iterable[Tuple[str,str,str,str]], batch_size: int = 20000):
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
    except Exception:
        pass
    cur = conn.cursor()
    sql = f"""INSERT INTO {table}(code, name, raw_line, source)
              VALUES (?,?,?,?)
              ON CONFLICT(code) DO UPDATE SET
                name=excluded.name,
                raw_line=excluded.raw_line,
                source=excluded.source,
                updated_at=datetime('now')"""
    batch = []
    count = 0
    for r in rows:
        batch.append(r)
        if len(batch) >= batch_size:
            cur.executemany(sql, batch)
            conn.commit()
            count += len(batch)
            batch.clear()
    if batch:
        cur.executemany(sql, batch)
        conn.commit()
        count += len(batch)
    conn.close()
    return count

def replace_all(db_path: str, table: str, rows: List[Tuple[str,str,str,str]]):
    conn = sqlite3.connect(db_path); cur = conn.cursor()
    cur.execute(f"DELETE FROM {table}")
    conn.commit(); conn.close()
    upsert_many_batched(db_path, table, rows)

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
