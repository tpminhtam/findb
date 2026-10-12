"""HydraDB store: same interface as LocalStore, backed by the hosted HydraDB context API.

Written against docs.hydradb.com (v2, Oct 12 2026). UNTESTED until a key is available: every call is logged to
out/hydra_log.jsonl so the first on-site run can be debugged quickly. Switch on with FB_STORE=hydradb and HYDRA_DB_API_KEY.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import time
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "out"
LOG = OUT / "hydra_log.jsonl"
ROW_CAP = int(os.environ.get("FB_HYDRA_ROW_CAP", "200"))   # HydraDB limits ingestion to 1,000 items per window


class HydraNotConfigured(RuntimeError):
    pass


def _log(event: str, **data) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    with LOG.open("a") as f:
        f.write(json.dumps({"at": dt.datetime.now().isoformat(timespec="seconds"), "event": event, **data}, default=str) + "\n")


def _get(obj, *path, default=None):
    cur = obj
    for p in path:
        if cur is None:
            return default
        cur = cur.get(p) if isinstance(cur, dict) else getattr(cur, p, None)
    return default if cur is None else cur


class HydraStore:
    def __init__(self, database: str = "finance_brain", token: str | None = None, local=None):
        token = token or os.environ.get("HYDRA_DB_API_KEY")
        if not token:
            raise HydraNotConfigured("HYDRA_DB_API_KEY is not set")
        from hydra_db import HydraDB  # hydradb-sdk
        self.client = HydraDB(token=token)
        self.database = database
        self.local = local            # LocalStore kept alongside for graph paths and the UI
        self.catalog = "catalog"
        self.rows = "rows"

    # ------------------------------------------------------------------ database
    def ensure_database(self, wait: bool = True) -> dict:
        try:   # already there? then skip creation
            if _get(self.client.databases.status(database=self.database), "data", "infra", "ready_for_ingestion"):
                return {"database": self.database, "ready_for_ingestion": True}
        except Exception:
            pass
        schema = [{"name": n, "type": "string", "enable_match": True} for n in ("workbook", "sheet", "table", "record_type", "period", "entity")]
        try:
            r = self.client.databases.create(database=self.database, database_metadata_schema=schema)
            _log("databases.create", response=str(r)[:500])
        except Exception as e:  # already exists or schema shape differs: log and continue
            _log("databases.create.error", error=str(e)[:500])
            try:
                r = self.client.databases.create(database=self.database)
                _log("databases.create.plain", response=str(r)[:500])
            except Exception as e2:
                _log("databases.create.plain.error", error=str(e2)[:500])
        status = None
        for _ in range(60 if wait else 1):
            st = self.client.databases.status(database=self.database)
            status = _get(st, "data", "infra", "ready_for_ingestion")
            _log("databases.status", ready=status)
            if status or not wait:
                break
            time.sleep(5)
        return {"database": self.database, "ready_for_ingestion": bool(status)}

    # ------------------------------------------------------------------ ingestion
    def _ingest(self, **kw):
        """context.ingest with a wait-and-retry on HydraDB's RATE_LIMITED response."""
        for attempt in range(4):
            try:
                return self.client.context.ingest(**kw)
            except Exception as e:
                m = re.search(r"retry in (\d+) second", str(e))
                if "RATE_LIMITED" in str(e) and m and attempt < 3:
                    _log("rate_limited", wait=int(m.group(1)))
                    time.sleep(int(m.group(1)) + 1)
                    continue
                raise

    def graph_payload_for(self, cap: dict, t: dict) -> dict:
        """Entities and relations around one table: its columns, metric rows, direct inputs and FKs."""
        nodes = cap["nodes"]
        ents: dict[str, dict] = {}
        rels: list[dict] = []
        def ent(nid: str) -> str | None:
            n = nodes.get(nid)
            if not n:
                return None
            key = re.sub(r"[^A-Za-z0-9_]", "_", nid)[:80]
            if key not in ents:
                ents[key] = {"name": n.get("name", nid)[:200], "type": n.get("type", "NODE"), "namespace": cap["workbook"]["slug"]}
            return key
        tkey = ent(f"table:{t['id']}")
        col_ids = {c["id"] for c in t["columns"]} | {f"{t['id']}#r{r}" for r in t["derived_rows"]}
        for cid in col_ids:
            ck = ent(cid)
            if tkey and ck:
                rels.append({"source": tkey, "target": ck, "predicate": "HAS_COLUMN"})
        for e in cap["edges"]:
            if e["dst"] in col_ids and e["predicate"] not in ("CONTAINS", "HAS_COLUMN") and len(rels) < 400:
                s, d = ent(e["src"]), ent(e["dst"])
                if s and d:
                    rels.append({"source": d, "target": s, "predicate": "DERIVES_FROM", "context": f"{e['predicate']}: {e.get('context', '')[:180]}"})
        for r in cap["relationships"]:
            if r["child_table"] == t["id"]:
                s, d = ent(r["child_col"]), ent(r["parent_col"])
                if s and d:
                    rels.append({"source": s, "target": d, "predicate": "REFERENCES", "context": r["kind"]})
        return {"entities": ents, "relations": rels}

    def add_capture(self, cap: dict, cards: dict[str, str]) -> dict:
        """cards: table id -> markdown card text. Ingests cards (+graph), rows (capped), notes."""
        wb = cap["workbook"]["slug"]
        results = {"documents": [], "rows": 0, "notes": 0, "errors": []}
        for t in cap["tables"]:
            if t["kind"] != "table":
                continue
            doc_id = f"tbl:{wb}:{t['id']}"
            md = cards.get(t["id"], "")
            payload = {doc_id: self.graph_payload_for(cap, t)}
            try:
                r = self._ingest(
                    type="knowledge", database=self.database, collection=self.catalog,
                    documents=[(f"{t['id']}.md", md.encode("utf-8"), "text/markdown")],
                    document_metadata=json.dumps([{"id": doc_id, "additional_metadata": {"workbook": wb, "sheet": t["sheet"], "table": t["id"], "role": t["role"]}}]),
                    graph_payload=json.dumps(payload))
                rid = _get(r, "data", "results", default=[None])
                rid = rid[0].id if rid and hasattr(rid[0], "id") else (rid[0].get("id") if rid and isinstance(rid[0], dict) else None)
                results["documents"].append({"doc": doc_id, "hydra_id": rid})
                _log("context.ingest.document", doc=doc_id, hydra_id=rid)
            except Exception as e:
                results["errors"].append(f"{doc_id}: {e}")
                _log("context.ingest.document.error", doc=doc_id, error=str(e)[:800])
            if t["role"] in ("data", "dimension", "assumptions") and t["rows"]:
                items = []
                for i, row in enumerate(t["rows"][:ROW_CAP]):
                    rec = {c["slug"]: v for c, v in zip(t["columns"], row)}
                    title = f"{t['id']} row {i + 1}: " + ", ".join(f"{k}={v}" for k, v in list(rec.items())[:6])
                    meta = {"workbook": wb, "sheet": t["sheet"], "table": t["id"], "record_type": t["id"]}
                    for c in t["columns"]:
                        if c["type"] == "date" and rec.get(c["slug"]):
                            meta["period"] = str(rec[c["slug"]])[:7]; break
                    if "entity" in rec:
                        meta["entity"] = str(rec["entity"])
                    items.append({"id": f"row:{wb}:{t['id']}:{i}", "database": self.database, "collection": self.rows, "title": title[:200],
                                  "kind": "custom", "provider": "excel", "external_id": f"{wb}:{t['id']}:{i}", "timestamp": dt.datetime.now().isoformat(),
                                  "fields": {"kind": "custom", "data": {"record_type": t["id"], **{k: (v if isinstance(v, (int, float, str, bool)) or v is None else str(v)) for k, v in rec.items()}}},
                                  "metadata": meta, "additional_metadata": {"source_file": cap["workbook"]["filename"], "range": t["range"]}})
                for start in range(0, len(items), 100):
                    batch = items[start:start + 100]
                    try:
                        self._ingest(type="knowledge", database=self.database, collection=self.rows, app_knowledge=json.dumps(batch))
                        results["rows"] += len(batch)
                        _log("context.ingest.rows", table=t["id"], n=len(batch))
                    except Exception as e:
                        results["errors"].append(f"rows {t['id']}: {e}")
                        _log("context.ingest.rows.error", table=t["id"], error=str(e)[:800])
                        break
        notes = []
        for i, n in enumerate(cap["notes"]):
            notes.append({"id": f"note:{wb}:{i}", "database": self.database, "collection": self.catalog, "title": f"{n['kind']} {n['sheet']}!{n['address']}"[:200],
                          "kind": "comment" if n["kind"] == "comment" else "knowledge_base", "provider": "excel", "external_id": f"{wb}:note:{i}",
                          "timestamp": dt.datetime.now().isoformat(),
                          "fields": ({"kind": "comment", "body": n["text"], "author": n.get("author") or "unknown"} if n["kind"] == "comment" else {"kind": "knowledge_base", "title": f"{n['sheet']} {n['kind']}", "body": n["text"]}),
                          "metadata": {"workbook": wb, "sheet": n["sheet"], "record_type": f"note_{n['kind']}"}})
        for f in cap["findings"]:
            notes.append({"id": f"finding:{wb}:{len(notes)}", "database": self.database, "collection": self.catalog, "title": f"finding {f['kind']} {f.get('sheet')}"[:200],
                          "kind": "knowledge_base", "provider": "finance-brain", "external_id": f"{wb}:finding:{len(notes)}", "timestamp": dt.datetime.now().isoformat(),
                          "fields": {"kind": "knowledge_base", "title": f"Finding {f['kind']} ({f['severity']})", "body": json.dumps(f, default=str)},
                          "metadata": {"workbook": wb, "sheet": f.get("sheet") or "", "record_type": "finding"}})
        for start in range(0, len(notes), 100):
            batch = notes[start:start + 100]
            try:
                self._ingest(type="knowledge", database=self.database, collection=self.catalog, app_knowledge=json.dumps(batch))
                results["notes"] += len(batch)
            except Exception as e:
                results["errors"].append(f"notes: {e}")
                _log("context.ingest.notes.error", error=str(e)[:800])
                break
        return results

    def wait_indexed(self, hydra_ids: list[str], collection: str | None = None, timeout: int = 300) -> dict:
        """Poll until every id is completed or errored; returns {id: last indexing_status} (intermediate states such as
        'graph_creation' are returned as-is when the timeout hits; the data is already queryable during graph creation)."""
        t0 = time.time(); last: dict[str, str] = {}
        ids = [i for i in hydra_ids if i]
        while ids and time.time() - t0 < timeout:
            st = self.client.context.status(database=self.database, collection=collection or self.catalog, ids=ids)
            for s in _get(st, "data", "statuses", default=[]) or []:
                last[_get(s, "id")] = _get(s, "indexing_status")
            ids = [i for i in ids if last.get(i) not in ("completed", "errored")]
            if ids:
                time.sleep(3)
        return last

    # ------------------------------------------------------------------ query
    def query(self, q: str, persona: str | None = None, k: int = 10, graph: bool = True) -> dict:
        from hydra_db.helpers import build_string
        collections = [self.catalog, self.rows] + ([f"analyst_{persona}"] if persona else [])
        kwargs = dict(database=self.database, query=q, type="all" if persona else "knowledge", collections=collections,
                      query_by="hybrid", graph_context=graph, mode="thinking", query_apps=True, max_results=k)
        try:
            r = self.client.query(**kwargs)
        except Exception as e:
            _log("query.error", q=q, error=str(e)[:600])
            if persona:   # the persona collection may not exist yet
                kwargs["collections"] = [self.catalog, self.rows]; kwargs["type"] = "knowledge"
                r = self.client.query(**kwargs)
            else:
                raise
        d = _get(r, "data", default=r)
        chunks = []
        for c in (_get(d, "chunks", default=[]) or []):
            chunks.append({"id": str(_get(c, "id") or _get(c, "chunk_uuid") or ""), "kind": f"hydra:{_get(c, 'collection') or ''}",
                           "text": _get(c, "chunk_content") or "", "score": None, "nodes": [], "meta": _get(c, "additional_metadata") or {}})
        paths = []
        g = _get(d, "graph")
        for p in (_get(g, "paths", default=[]) or []):
            paths.append(f"HYDRADB FACT (score {float(_get(p, 'relevancy_score', default=0) or 0):.2f}): {_get(p, 'combined_context')}")
        gc = _get(d, "graph_context")
        for cr in (_get(gc, "chunk_relations", default=[]) or []):
            paths.append(f"HYDRADB RELATION: {_get(cr, 'combined_context')}")
        try:
            context_md = build_string(r)
        except Exception as e:
            _log("build_string.error", error=str(e)[:300]); context_md = ""
        _log("query", q=q, n_chunks=len(chunks), n_paths=len(paths), collections=collections)
        mems = [c for c in chunks if c["kind"].startswith("hydra:analyst_")]
        if self.local and graph:
            paths.extend(self.local.query(q, persona=None, k=4, graph=True)["paths"])
        local_mems = self.local.memories_for(persona) if self.local else []
        return {"chunks": chunks, "paths": paths[:60], "nodes": [], "memories": [{"text": m["text"]} for m in mems] or local_mems,
                "context_md": context_md[:14000], "store": "hydradb"}

    def remember(self, persona: str, text: str) -> dict:
        r = self.client.context.ingest(type="memory", database=self.database, collection=f"analyst_{persona}",
                                       memories=json.dumps([{"text": text, "infer": True}]))
        _log("memory.ingest", persona=persona, response=str(r)[:300])
        if self.local:
            self.local.remember(persona, text)
        return {"text": text, "hydra": str(r)[:200]}
