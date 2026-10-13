"""Finance Brain web app: drop a workbook, browse the catalog and lineage, see findings, ask the brain, run SQL."""
from __future__ import annotations

import json
import os
import shutil
import time
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

from .brain import ask  # noqa: E402
from .capture import capture  # noqa: E402
from .cards import table_card  # noqa: E402
from .gsheet import fetch_gsheet  # noqa: E402
from .mirror import list_tables, mirror, run_sql  # noqa: E402
from .store_local import LocalStore  # noqa: E402

OUT = ROOT / "out"
UPLOADS = ROOT / "uploads"
STATIC = Path(__file__).resolve().parent / "static"
OUT.mkdir(exist_ok=True); UPLOADS.mkdir(exist_ok=True)

app = FastAPI(title="finDB")
local = LocalStore(OUT)
local.load_all()
_hydra = {"store": None, "error": None, "tried": False, "syncing": 0, "last": None}


def hydra():
    """The HydraDB store once the background connection is up; None while connecting or in local mode."""
    if os.environ.get("FB_STORE", "local") != "hydradb":
        return None
    return _hydra["store"]


@app.on_event("startup")
def _connect_hydra():
    if os.environ.get("FB_STORE", "local") != "hydradb":
        return
    import threading

    def run():
        try:
            from .store_hydra import HydraStore
            h = HydraStore(database=os.environ.get("FB_HYDRA_DB", "finance_brain"), local=local)
            h.ensure_database(wait=True)
            _hydra["store"] = h
        except Exception as e:
            _hydra["error"] = str(e)[:300]
        finally:
            _hydra["tried"] = True
    threading.Thread(target=run, daemon=True).start()


def store():
    return hydra() or local


def cards_for(cap: dict, sem: dict | None = None) -> dict[str, str]:
    return {t["id"]: table_card(cap, t, sem) for t in cap["tables"] if t["kind"] == "table"}


def process(path: Path, enrich: bool = False) -> dict:
    t0 = time.time()
    cap = capture(path)
    slug = cap["workbook"]["slug"]
    out = OUT / slug
    out.mkdir(parents=True, exist_ok=True)
    sem = None
    sem_error = None
    if not enrich and (out / "semantics.json").exists():     # keep earlier Claude descriptions when a file is dropped again
        try:
            sem = json.loads((out / "semantics.json").read_text())
        except Exception:
            sem = None
    if enrich:
        try:
            from .semantics import enrich as _enrich
            sem = _enrich(cap, out_dir=out)
        except Exception as e:
            sem_error = str(e)[:300]
    cards = cards_for(cap, sem)
    (out / "capture.json").write_text(json.dumps(cap, default=str))
    (out / "ddl.sql").write_text(cap["ddl"])
    (out / "cards").mkdir(exist_ok=True)
    for tid, md in cards.items():
        (out / "cards" / f"{tid}.md").write_text(md)
    loaded = mirror(cap)
    local.add_capture(cap, sem)
    hres = None
    h = hydra()
    if h:   # push to HydraDB in the background so the drop returns at once
        import threading
        _hydra["syncing"] += 1

        def _push():
            try:
                res = h.add_capture(cap, cards)
                _hydra["last"] = {"slug": slug, "documents": len(res["documents"]), "rows": res["rows"], "notes": res["notes"], "errors": res["errors"][:3]}
            except Exception as e:
                _hydra["last"] = {"slug": slug, "errors": [str(e)[:300]]}
            finally:
                _hydra["syncing"] -= 1
        threading.Thread(target=_push, daemon=True).start()
        hres = {"status": "syncing"}
    sc = cap["scorecard"]
    return {"slug": slug, "filename": cap["workbook"]["filename"], "seconds": round(time.time() - t0, 2),
            "sheets": len(cap["workbook"]["sheets"]), "tables": [t["id"] for t in cap["tables"] if t["kind"] == "table"],
            "blocks": len(cap["tables"]), "edges": len(cap["edges"]), "relationships": len(cap["relationships"]), "findings": len(cap["findings"]),
            "formula_cells": sc["formula_cells"], "duckdb_tables": loaded, "hydra": hres, "semantics": bool(sem), "semantics_error": sem_error}


# ----------------------------------------------------------------------------- pages
@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.get("/api/status")
def status():
    h = hydra()
    wants = os.environ.get("FB_STORE", "local") == "hydradb"
    return {"store": "hydradb" if (h or (wants and not _hydra["tried"])) else "local", "hydra_connecting": wants and not _hydra["tried"],
            "hydra_error": _hydra["error"], "hydra_syncing": _hydra["syncing"] > 0, "hydra_last": _hydra["last"], "hydra_configured": bool(os.environ.get("HYDRA_DB_API_KEY")),
            "model": os.environ.get("FB_MODEL", "claude-opus-5-5"), "effort": os.environ.get("FB_EFFORT", "low"),
            "anthropic_key": bool(os.environ.get("ANTHROPIC_API_KEY")), "workbooks": len(local.captures), "chunks": len(local.chunks),
            "graph_nodes": local.graph.number_of_nodes(), "graph_edges": local.graph.number_of_edges(), "duckdb_tables": len(list_tables())}


@app.post("/api/upload")
async def upload(file: UploadFile | None = File(None), gsheet_url: str | None = Form(None), enrich: bool = Form(False)):
    if gsheet_url:
        try:
            path = fetch_gsheet(gsheet_url, UPLOADS)
        except Exception as e:
            raise HTTPException(400, str(e))
    elif file is not None:
        if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(400, "upload an .xlsx or .xlsm file (export Google Sheets or .xls to xlsx first)")
        path = UPLOADS / Path(file.filename).name
        path.write_bytes(await file.read())
    else:
        raise HTTPException(400, "no file or gsheet_url")
    try:
        return process(path, enrich=enrich)
    except Exception as e:
        raise HTTPException(500, f"capture failed: {type(e).__name__}: {e}")


@app.get("/api/workbooks")
def workbooks():
    out = []
    for slug, cap in local.captures.items():
        sc = cap["scorecard"]
        out.append({"slug": slug, "filename": cap["workbook"]["filename"], "captured_at": cap["workbook"]["captured_at"],
                    "sheets": [s["name"] for s in cap["workbook"]["sheets"]],
                    "tables": [{"id": t["id"], "sheet": t["sheet"], "kind": t["kind"], "role": t["role"], "n_rows": t["n_rows"], "range": t["range"], "cols": len(t["columns"])} for t in cap["tables"]],
                    "scorecard": {k: v for k, v in sc.items() if k != "findings"}, "findings": len(cap["findings"]),
                    "findings_high": sum(1 for f in cap["findings"] if f["severity"] == "high"), "enriched": slug in local.semantics})
    return out


@app.get("/api/workbooks/{slug}")
def workbook(slug: str):
    cap = local.captures.get(slug)
    if not cap:
        raise HTTPException(404, "unknown workbook")
    sem = local.semantics.get(slug)
    tables = []
    for t in cap["tables"]:
        tt = dict(t); tt["rows"] = t["rows"][:25]
        if t["kind"] == "table":
            tt["card"] = table_card(cap, t, sem)
        tables.append(tt)
    return {"workbook": cap["workbook"], "tables": tables, "relationships": cap["relationships"], "notes": cap["notes"],
            "findings": cap["findings"], "scorecard": {k: v for k, v in cap["scorecard"].items() if k != "findings"}, "ddl": cap["ddl"],
            "semantics": sem, "cross_relationships": [r for r in local.cross_rels if slug in (r["child_workbook"], r["parent_workbook"])]}


SKIP_PRED = {"CONTAINS", "HAS_COLUMN", "LOOKS_UP", "VLOOKUP_KEY", "HLOOKUP_KEY", "XLOOKUP_KEY"}


def table_graph(cap: dict) -> dict:
    """Lineage rolled up to one node per table: which tables feed which (solid) and which are keyed to which (dashed)."""
    import re
    from openpyxl.utils import column_index_from_string
    blocks = {t["id"]: t for t in cap["tables"]}
    by_sheet: dict[str, list] = {}
    for t in cap["tables"]:
        by_sheet.setdefault(t["sheet"], []).append(t)

    def group(nid: str):
        n = cap["nodes"].get(nid) or {}
        if n.get("table") in blocks:
            return n["table"]
        if n.get("type") == "CELL":
            m = re.match(r"([A-Z]+)(\d+)$", n.get("address") or "")
            if m:
                col, row = column_index_from_string(m.group(1)), int(m.group(2))
                for t in by_sheet.get(n.get("sheet"), []):
                    if t["top_row"] <= row <= t["last_row"] and t["c1"] <= col <= t["c2"]:
                        return t["id"]
            return f"sheet:{n.get('sheet')}"
        if n.get("type") in ("EXTERNAL_WORKBOOK", "NAMED_RANGE"):
            return nid
        return None

    agg: dict[tuple[str, str], dict] = {}
    for e in cap["edges"]:
        if e["predicate"] in SKIP_PRED:
            continue
        s, d = group(e["src"]), group(e["dst"])
        if not s or not d or s == d:
            continue
        kind = "key" if e["predicate"] == "REFERENCES" else "feeds"
        if kind == "key":
            s, d = d, s          # lookup table -> table that uses it
        a = agg.setdefault((s, d), {"feeds": 0, "key": 0})
        a[kind] += e.get("count", 1)
    used = {x for pair in agg for x in pair}
    nodes = []
    for gid in used:
        t = blocks.get(gid)
        if t:
            label, role = t["name"], (t["role"] if t["kind"] == "table" else "inputs")
        elif gid.startswith("sheet:"):
            label, role = gid[6:], "inputs"
        else:
            label, role = cap["nodes"].get(gid, {}).get("name", gid), "external"
        nodes.append({"data": {"id": gid, "label": label if len(label) <= 34 else label[:32] + "…", "full": label, "type": "TABLE", "role": role}})
    edges = [{"data": {"id": f"{s}->{d}", "source": s, "target": d, "kind": "feeds" if a["feeds"] else "key", "count": a["feeds"] or a["key"]}}
             for (s, d), a in agg.items()]
    return {"elements": nodes + edges, "n_nodes": len(nodes), "n_edges": len(edges)}


@app.get("/api/workbooks/{slug}/graph")
def graph(slug: str, structural: bool = False, level: str = "column"):
    cap = local.captures.get(slug)
    if not cap:
        raise HTTPException(404, "unknown workbook")
    if level == "table":
        return table_graph(cap)
    keep = {"CONTAINS", "HAS_COLUMN"} if structural else set()
    edges = [e for e in cap["edges"] if e["predicate"] not in {"CONTAINS", "HAS_COLUMN"} or e["predicate"] in keep]
    used = {e["src"] for e in edges} | {e["dst"] for e in edges}
    nodes = [{"data": {"id": nid, "label": n.get("name", nid)[:60], "type": n.get("type"), "sheet": n.get("sheet"), "table": n.get("table"), "full": n.get("name", nid)}}
             for nid, n in cap["nodes"].items() if nid in used]
    els = nodes + [{"data": {"id": f"{e['src']}->{e['dst']}:{e['predicate']}", "source": e["src"], "target": e["dst"], "predicate": e["predicate"], "context": e.get("context", ""), "count": e.get("count", 1)}} for e in edges]
    return {"elements": els, "n_nodes": len(nodes), "n_edges": len(edges)}


@app.get("/api/workbooks/{slug}/ddl", response_class=PlainTextResponse)
def ddl(slug: str):
    cap = local.captures.get(slug)
    if not cap:
        raise HTTPException(404, "unknown workbook")
    return cap["ddl"]


@app.delete("/api/workbooks/{slug}")
def delete(slug: str):
    if slug not in local.captures:
        raise HTTPException(404, "unknown workbook")
    folder = OUT / slug                                      # remove the capture but keep semantics.json for a later re-drop
    for f in ("capture.json", "ddl.sql"):
        (folder / f).unlink(missing_ok=True)
    shutil.rmtree(folder / "cards", ignore_errors=True)
    local.captures.pop(slug, None); local.semantics.pop(slug, None); local.rebuild()
    return {"deleted": slug}


class AskBody(BaseModel):
    question: str
    persona: str | None = "analyst"
    graph: bool = True
    k: int = 12


@app.post("/api/ask")
def api_ask(body: AskBody):
    return ask(body.question, store(), persona=body.persona or None, graph=body.graph, k=body.k)


class RememberBody(BaseModel):
    persona: str
    text: str


@app.post("/api/remember")
def api_remember(body: RememberBody):
    return store().remember(body.persona, body.text)


@app.get("/api/memories")
def api_memories(persona: str):
    return local.memories_for(persona)


@app.get("/api/sql")
def api_sql(q: str):
    try:
        return run_sql(q)
    except Exception as e:
        raise HTTPException(400, str(e))


@app.get("/api/sql/tables")
def api_tables():
    return list_tables()


@app.post("/api/enrich/{slug}")
def api_enrich(slug: str):
    cap = local.captures.get(slug)
    if not cap:
        raise HTTPException(404, "unknown workbook")
    try:
        from .semantics import enrich as _enrich
        sem = _enrich(cap, out_dir=OUT / slug)
    except Exception as e:
        raise HTTPException(500, f"enrich failed: {type(e).__name__}: {e}")
    local.add_capture(cap, sem)
    for tid, md in cards_for(cap, sem).items():
        (OUT / slug / "cards" / f"{tid}.md").write_text(md)
    return {"slug": slug, "usage": sem.get("_usage"), "tables": len(sem.get("tables", {})), "lineage": len(sem.get("lineage", {}))}


@app.post("/api/hydra/push")
def api_hydra_push():
    h = hydra()
    if not h:
        raise HTTPException(400, f"HydraDB not active: FB_STORE={os.environ.get('FB_STORE')} error={_hydra['error']}")
    results = {}
    for slug, cap in local.captures.items():
        results[slug] = h.add_capture(cap, cards_for(cap, local.semantics.get(slug)))
    return results


@app.post("/api/samples")
def api_samples():
    """Load the three Ardent sample workbooks (demo reset)."""
    samples = ROOT.parent / "samples"
    return [process(p) for p in sorted(samples.glob("Ardent_*.xlsx"))]
