"""Offline knowledge store: BM25 retrieval over catalog/notes/findings/lineage chunks + a lineage graph.

Same interface as HydraStore so the app and brain work before the HydraDB key arrives.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import networkx as nx
from rank_bm25 import BM25Okapi

from .capture import cross_workbook_relationships

OUT = Path(__file__).resolve().parent.parent / "out"
STRUCTURAL = {"CONTAINS", "HAS_COLUMN"}
IMPACT_RE = re.compile(r"\b(break|breaks|change|changes|changed|impact|affect|depend|downstream|if we|what happens|rename|retire|delete|grow|append)\b", re.I)


def tokenize(text: str) -> list[str]:
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    toks = re.findall(r"[a-z0-9]+(?:[.%][a-z0-9]+)?", text.lower())
    out = []
    for t in toks:
        out.append(t)
        if "_" in t:
            out.extend(t.split("_"))
    return [t for t in out if len(t) >= 2]


class LocalStore:
    def __init__(self, root: Path = OUT):
        self.root = Path(root)
        self.captures: dict[str, dict] = {}
        self.semantics: dict[str, dict] = {}
        self.chunks: list[dict] = []
        self.bm25: BM25Okapi | None = None
        self.graph = nx.DiGraph()
        self.cross_rels: list[dict] = []
        self.mem_path = self.root / "memories.json"
        self.memories: dict[str, list[dict]] = json.loads(self.mem_path.read_text()) if self.mem_path.exists() else {}

    # ------------------------------------------------------------------ loading
    def load_all(self) -> None:
        for p in sorted(self.root.glob("*/capture.json")):
            cap = json.loads(p.read_text())
            self.captures[cap["workbook"]["slug"]] = cap
            sp = p.parent / "semantics.json"
            if sp.exists():
                self.semantics[cap["workbook"]["slug"]] = json.loads(sp.read_text())
        self.rebuild()

    def add_capture(self, cap: dict, semantics: dict | None = None) -> None:
        self.captures[cap["workbook"]["slug"]] = cap
        if semantics:
            self.semantics[cap["workbook"]["slug"]] = semantics
        self.rebuild()

    def rebuild(self) -> None:
        self.cross_rels = cross_workbook_relationships(list(self.captures.values()))
        self.graph = nx.DiGraph()
        self.chunks = []
        for cap in self.captures.values():
            for nid, n in cap["nodes"].items():
                self.graph.add_node(nid, **n, workbook=cap["workbook"]["slug"])
            for e in cap["edges"]:
                self.graph.add_edge(e["src"], e["dst"], predicate=e["predicate"], context=e.get("context", ""), count=e.get("count", 1), via=e.get("via_name"))
            self.chunks.extend(self.chunks_for(cap))
        for r in self.cross_rels:
            self.graph.add_edge(r["child_col"], r["parent_col"], predicate="REFERENCES", context=f"cross-workbook FK coverage {r['coverage']}", count=1)
            self.chunks.append({"id": f"xrel:{r['child_col']}->{r['parent_col']}", "kind": "relationship", "workbook": r["child_workbook"],
                                "text": f"Cross-workbook relationship: {r['child_workbook']} {r['child_col']} references {r['parent_workbook']} {r['parent_col']} (value coverage {r['coverage']}).",
                                "nodes": [r["child_col"], r["parent_col"]]})
        self.bm25 = BM25Okapi([tokenize(c["text"]) for c in self.chunks]) if self.chunks else None

    # ------------------------------------------------------------------ chunking
    def node_name(self, nid: str) -> str:
        n = self.graph.nodes.get(nid) or {}
        return n.get("name", nid)

    def chunks_for(self, cap: dict) -> list[dict]:
        wb = cap["workbook"]["slug"]; fn = cap["workbook"]["filename"]
        sem = self.semantics.get(wb, {})
        sc = cap["scorecard"]
        out = []
        tables = [t for t in cap["tables"] if t["kind"] == "table"]
        brief = (f"Workbook {fn} (id {wb}). Sheets: {', '.join(s['name'] for s in cap['workbook']['sheets'])}. "
                 f"Tables: {', '.join(f'{t['id']} ({t['role']}, {t['n_rows']} rows)' for t in tables)}. "
                 f"Cells {sc['cells']}, formula cells {sc['formula_cells']} ({sc['formula_share']:.0%}), named ranges {sc['named_ranges']}, manual inputs {sc['manual_inputs']}. "
                 f"Findings by severity: {sc['findings_by_severity']}. Migration plan: load as tables {sc['migration_plan']['load_as_tables']}; rebuild as views or metrics {sc['migration_plan']['rebuild_as_views_or_metrics']}; "
                 f"keep as assumptions {sc['migration_plan']['keep_as_assumptions']}; ingest as knowledge {sc['migration_plan']['ingest_as_knowledge']}.")
        if sem.get("workbook_brief"):
            brief += " Summary: " + sem["workbook_brief"]
        out.append({"id": f"wb:{wb}", "kind": "workbook", "workbook": wb, "text": brief, "nodes": [f"wb:{wb}"]})
        edges_by_dst: dict[str, list] = defaultdict(list)
        for e in cap["edges"]:
            if e["predicate"] not in STRUCTURAL:
                edges_by_dst[e["dst"]].append(e)
        names = cap["nodes"]
        def nm(nid): return names.get(nid, {}).get("name", nid)
        for t in cap["tables"]:
            tid = t["id"]
            tsem = sem.get("tables", {}).get(tid, {})
            if t["kind"] == "table":
                cols = "; ".join(f"{c['name']} [{c['slug']}] {c['semantic']}" + (" key" if c["is_key"] else "") + (f" ({tsem.get('columns', {}).get(c['slug'])})" if tsem.get("columns", {}).get(c["slug"]) else "")
                                 + (f" e.g. {c['samples'][:2]}" if c["samples"] else "") for c in t["columns"])
                fks = [r for r in cap["relationships"] if r["child_table"] == tid]
                fk_txt = "; ".join(f"{r['child_col'].split('.')[-1]} -> {r['parent_table']}.{r['parent_col'].split('.')[-1]} ({r['kind']})" for r in fks)
                derived = "; ".join(t["derived_rows"].values())
                text = (f"Table {tid} on sheet '{t['sheet']}' of {fn}, range {t['range']}, header row {t['header_row']}, {t['n_rows']} data rows, role {t['role']}, formula share {t['formula_share']:.0%}."
                        + (f" Title: {t['title']}." if t.get("title") else "") + (f" {tsem['description']}" if tsem.get("description") else "")
                        + f" Columns: {cols}." + (f" Derived rows (totals/metrics): {derived}." if derived else "") + (f" Relationships: {fk_txt}." if fk_txt else "")
                        + (f" Helper rows (hidden keys) {t['helper_rows']}." if t["helper_rows"] else "") + (f" Migration note: {tsem['migration_note']}" if tsem.get("migration_note") else ""))
                out.append({"id": f"tbl:{wb}:{tid}", "kind": "table", "workbook": wb, "sheet": t["sheet"], "table": tid, "text": text,
                            "nodes": [f"table:{tid}"] + [c["id"] for c in t["columns"]] + [f"{tid}#r{r}" for r in t["derived_rows"]]})
                if t["narrative"] or t["role"] == "narrative":
                    for i, row in enumerate(t["rows"]):
                        pieces = [f"{c['name']}: {v}" for c, v in zip(t["columns"], row) if v not in (None, "")]
                        out.append({"id": f"row:{wb}:{tid}:{i}", "kind": "narrative_row", "workbook": wb, "sheet": t["sheet"], "table": tid,
                                    "text": f"{t['sheet']} entry: " + " | ".join(pieces), "nodes": [f"table:{tid}"]})
            elif t["kind"] == "kv":
                for p in t["pairs"]:
                    text = f"Input on sheet '{t['sheet']}' {p['address']}: {p['key']} = {p['value']}" + (f" (formula {p['formula']})" if p.get("formula") else "") + (f". Note by {p.get('author')}: {p['comment']}" if p.get("comment") else "")
                    out.append({"id": f"kv:{wb}:{t['sheet']}:{p['address']}", "kind": "input", "workbook": wb, "sheet": t["sheet"], "text": text, "nodes": [f"{re.sub(r'[^a-z0-9]+', '_', t['sheet'].lower()).strip('_')}!{p['address']}"]})
        for n in cap["notes"]:
            if n["kind"] == "fact":
                continue
            who = f" (by {n['author']})" if n.get("author") else ""
            out.append({"id": f"note:{wb}:{n['sheet']}:{n['address']}:{n['kind']}", "kind": f"note_{n['kind']}", "workbook": wb, "sheet": n["sheet"],
                        "text": f"{n['kind'].title()} on sheet '{n['sheet']}' at {n['address']}{who}: {n['text']}", "nodes": [f"{re.sub(r'[^a-z0-9]+', '_', n['sheet'].lower()).strip('_')}!{n['address']}"] if n["kind"] == "comment" else []})
        for i, f in enumerate(cap["findings"]):
            if f["kind"] == "stale_range":
                text = (f"Finding stale_range (high): formulas reference '{f['sheet']}' table {f['table']} only through row {f['range_end']} but the table runs to row {f['table_end']}; "
                        f"{f['excluded_data_rows']} data rows are silently excluded from {f['formulas']} formulas (e.g. {', '.join(f['examples'])}). Appending rows to {f['table']} will not flow into these formulas until the ranges are extended.")
                nodes = [f"table:{f['table']}"]
            else:
                text = f"Finding {f['kind']} ({f['severity']}) on sheet '{f['sheet']}' at {f.get('location')}: {f.get('detail', '')}" + (f" ({f['count']} cells)" if f.get("count") else "")
                nodes = []
            out.append({"id": f"finding:{wb}:{i}", "kind": "finding", "workbook": wb, "sheet": f.get("sheet"), "text": text, "nodes": nodes})
        for dst, es in edges_by_dst.items():
            srcs = sorted({(nm(e["src"]), e["predicate"], e.get("via_name")) for e in es})
            src_txt = "; ".join(f"{s} via {p}" + (f" (named range {v})" if v else "") for s, p, v in srcs[:14])
            ex = next((e["context"] for e in es if e.get("context")), "")
            cells = sorted({c for e in es for c in e.get("cells", [])})[:4]
            expl = sem.get("lineage", {}).get(dst)
            text = f"Lineage: {nm(dst)} (node {dst}) derives from {src_txt}. Example formula at {', '.join(cells)}: {ex}" + (f" In words: {expl}" if expl else "")
            out.append({"id": f"lin:{wb}:{dst}", "kind": "lineage", "workbook": wb, "text": text, "nodes": [dst] + [e["src"] for e in es]})
        return out

    # ------------------------------------------------------------------ query
    def upstream(self, nid: str, depth: int = 4, width: int = 10) -> list[str]:
        lines, frontier, seen = [], [nid], {nid}
        for d in range(depth):
            nxt = []
            for n in frontier:
                preds = [(u, self.graph.edges[u, n]) for u in self.graph.predecessors(n) if self.graph.edges[u, n]["predicate"] not in STRUCTURAL]
                if not preds:
                    continue
                parts = []
                for u, e in preds[:width]:
                    parts.append(f"{self.node_name(u)} [{e['predicate']}" + (f", named range {e['via']}" if e.get("via") else "") + "]")
                    if u not in seen:
                        seen.add(u); nxt.append(u)
                lines.append(f"{self.node_name(n)} <- " + "; ".join(parts))
            frontier = nxt
            if not frontier:
                break
        return lines

    def upstream_nodes(self, nid: str, depth: int = 3) -> list[str]:
        out, frontier, seen = [], [nid], {nid}
        for _ in range(depth):
            nxt = []
            for n in frontier:
                for u in self.graph.predecessors(n):
                    if self.graph.edges[u, n]["predicate"] not in STRUCTURAL and u not in seen:
                        seen.add(u); nxt.append(u); out.append(u)
            frontier = nxt
        return out

    def downstream_nodes(self, nid: str, depth: int = 2) -> list[str]:
        out, frontier, seen = [], [nid], {nid}
        for _ in range(depth):
            nxt = []
            for n in frontier:
                for v in self.graph.successors(n):
                    if self.graph.edges[n, v]["predicate"] not in STRUCTURAL and v not in seen:
                        seen.add(v); nxt.append(v); out.append(v)
            frontier = nxt
        return out

    def downstream(self, nid: str, depth: int = 3, width: int = 12) -> list[str]:
        lines, frontier, seen = [], [nid], {nid}
        for d in range(depth):
            nxt = []
            for n in frontier:
                succs = [(v, self.graph.edges[n, v]) for v in self.graph.successors(n) if self.graph.edges[n, v]["predicate"] not in STRUCTURAL]
                if not succs:
                    continue
                parts = []
                for v, e in succs[:width]:
                    parts.append(f"{self.node_name(v)} [{e['predicate']}]")
                    if v not in seen:
                        seen.add(v); nxt.append(v)
                lines.append(f"{self.node_name(n)} -> " + "; ".join(parts))
            frontier = nxt
            if not frontier:
                break
        return lines

    def query(self, q: str, persona: str | None = None, k: int = 12, graph: bool = True) -> dict:
        if not self.bm25:
            return {"chunks": [], "paths": [], "nodes": [], "memories": self.memories_for(persona)}
        toks = tokenize(q)
        scores = self.bm25.get_scores(toks)
        ql = q.lower()
        boosted = []
        for i, c in enumerate(self.chunks):
            s = float(scores[i])
            for nid in c.get("nodes", []):
                name = self.node_name(nid).lower()
                if name and (name in ql or (nid.split("!")[-1].lower() in ql if "!" in nid else False)):
                    s += 4.0
            if c.get("table") and c["table"].replace("_", " ") in ql:
                s += 2.0
            if c["kind"] == "finding" and IMPACT_RE.search(q):
                s += 1.5
            boosted.append((s, i))
        boosted.sort(reverse=True)
        top = [dict(self.chunks[i], score=round(s, 2)) for s, i in boosted[:k] if s > 0]
        paths, seen_nodes = [], []
        if graph:
            impact = bool(IMPACT_RE.search(q))
            qtok = set(toks)
            cands: dict[str, tuple[int, int]] = {}
            for rank, c in enumerate(top[:8]):
                for nid in c.get("nodes", []):
                    if nid in cands or nid not in self.graph or nid.startswith(("wb:", "sheet:")):
                        continue
                    overlap = len(qtok & set(tokenize(self.node_name(nid))))
                    cands[nid] = (overlap, rank)
            ordered = sorted(cands, key=lambda n: (-cands[n][0], cands[n][1]))
            picked = [n for n in ordered if cands[n][0] > 0][:4] or ordered[:2]
            chain: list[str] = []
            for nid in picked:
                seen_nodes.append(nid)
                up = self.upstream(nid)
                down = self.downstream(nid) if impact or not up else []
                paths.extend(f"UPSTREAM {l}" for l in up[:8])
                paths.extend(f"DOWNSTREAM {l}" for l in down[:8])
                chain.extend(self.upstream_nodes(nid, depth=3))
                if impact:
                    chain.extend(self.downstream_nodes(nid, depth=2))
            # pull the lineage records of the nodes on the chain into the context (they explain each hop)
            have = {c["id"] for c in top}
            by_id = {c["id"]: c for c in self.chunks if c["kind"] == "lineage"}
            extra = []
            chain = list(dict.fromkeys(chain))
            chain.sort(key=lambda n: -len(qtok & set(tokenize(self.node_name(n)))))
            for nid in chain:
                for cid in (f"lin:{wb}:{nid}" for wb in self.captures):
                    if cid in by_id and cid not in have:
                        have.add(cid); extra.append(dict(by_id[cid], score=0.0))
                if len(extra) >= 12:
                    break
            top.extend(extra)
        return {"chunks": top, "paths": paths[:40], "nodes": seen_nodes, "memories": self.memories_for(persona)}

    # ------------------------------------------------------------------ memory
    def remember(self, persona: str, text: str) -> dict:
        import datetime as dt
        m = {"text": text.strip(), "at": dt.datetime.now().isoformat(timespec="seconds")}
        self.memories.setdefault(persona, []).append(m)
        self.root.mkdir(parents=True, exist_ok=True)
        self.mem_path.write_text(json.dumps(self.memories, indent=1))
        return m

    def memories_for(self, persona: str | None) -> list[dict]:
        if not persona:
            return []
        return self.memories.get(persona, [])[-12:]
