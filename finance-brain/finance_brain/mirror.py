"""DuckDB mirror of captured tables plus DDL: the migration proof."""
from __future__ import annotations

import re
from pathlib import Path

import duckdb
import pandas as pd

OUT = Path(__file__).resolve().parent.parent / "out"
DB_PATH = OUT / "brain.duckdb"


def table_name(cap: dict, t: dict) -> str:
    return f"{cap['workbook']['slug']}__{t['slug']}"


def mirror(cap: dict, db_path: Path = DB_PATH) -> list[dict]:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect(str(db_path))
    loaded = []
    try:
        for t in cap["tables"]:
            if t["kind"] != "table" or not t["columns"] or not t["rows"]:
                continue
            cols = [c["slug"] for c in t["columns"]]
            df = pd.DataFrame(t["rows"], columns=cols)
            for c in t["columns"]:
                if c["type"] == "date":
                    df[c["slug"]] = pd.to_datetime(df[c["slug"]], errors="coerce").dt.date
                elif c["type"] in ("integer", "decimal"):
                    df[c["slug"]] = pd.to_numeric(df[c["slug"]], errors="coerce")
                elif c["type"] == "boolean":
                    df[c["slug"]] = df[c["slug"]].astype("boolean")
                else:
                    df[c["slug"]] = df[c["slug"]].astype("string")
            name = table_name(cap, t)
            con.register("df_tmp", df)
            con.execute(f'CREATE OR REPLACE TABLE "{name}" AS SELECT * FROM df_tmp')
            con.unregister("df_tmp")
            loaded.append({"table": name, "rows": len(df), "columns": cols, "role": t["role"]})
    finally:
        con.close()
    return loaded


SAFE_SQL = re.compile(r"^\s*(select|with|show|describe|explain|pragma)\b", re.I)


def run_sql(sql: str, db_path: Path = DB_PATH, limit: int = 200) -> dict:
    if not SAFE_SQL.match(sql or ""):
        raise ValueError("read-only: statements must start with SELECT, WITH, SHOW, DESCRIBE or EXPLAIN")
    con = duckdb.connect(str(db_path), read_only=True)
    try:
        cur = con.execute(sql)
        cols = [d[0] for d in cur.description]
        rows = cur.fetchmany(limit)
        return {"columns": cols, "rows": [[(str(v) if not isinstance(v, (int, float, type(None))) else v) for v in r] for r in rows], "truncated": len(rows) >= limit}
    finally:
        con.close()


def list_tables(db_path: Path = DB_PATH) -> list[dict]:
    if not db_path.exists():
        return []
    con = duckdb.connect(str(db_path), read_only=True)
    try:
        rows = con.execute("select table_name, estimated_size from duckdb_tables() order by 1").fetchall()
        return [{"table": r[0], "rows": r[1]} for r in rows]
    finally:
        con.close()
