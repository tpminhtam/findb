"""Optional Claude enrichment: table descriptions, column meanings, lineage in plain English, migration notes."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

import anthropic

MODEL = os.environ.get("FB_MODEL", "claude-opus-5-5")

PROMPT = """You are documenting a finance team's spreadsheet for migration into a database and a knowledge graph.
Below is a JSON capture of one workbook: tables (with columns, types, samples, derived rows), lineage edges (dst derives from src via a function; context is a sample formula), relationships, notes and findings.

Return ONLY a JSON object with this shape, no prose, no code fences:
{
  "workbook_brief": "2-3 sentences: what this workbook is for and who uses it",
  "tables": {"<table id>": {"description": "1-2 sentences", "columns": {"<column slug>": "short meaning"}, "migration_note": "1 sentence: load as table / rebuild as view or metric / keep as assumptions / treat as knowledge, and why"}},
  "lineage": {"<dst node id>": "one plain-English sentence explaining how this value is computed"}
}
Cover every table. For lineage, cover the 12 most important destination nodes (report columns and metric rows). Be concrete: name sheets, functions, inputs.

CAPTURE:
"""


def compact(cap: dict) -> dict:
    tables = []
    for t in cap["tables"]:
        if t["kind"] != "table":
            continue
        tables.append({"id": t["id"], "sheet": t["sheet"], "title": t.get("title"), "role": t["role"], "n_rows": t["n_rows"], "range": t["range"],
                       "columns": [{"slug": c["slug"], "name": c["name"], "semantic": c["semantic"], "key": c["is_key"], "samples": c["samples"][:3], "formula_share": c["formula_share"]} for c in t["columns"]],
                       "derived_rows": t["derived_rows"], "sample_rows": t["rows"][:3]})
    edges = [{"dst": e["dst"], "dst_name": cap["nodes"].get(e["dst"], {}).get("name"), "src": e["src"], "src_name": cap["nodes"].get(e["src"], {}).get("name"),
              "via": e["predicate"], "formula": e.get("context", "")[:160]} for e in cap["edges"] if e["predicate"] not in ("CONTAINS", "HAS_COLUMN")][:250]
    return {"workbook": cap["workbook"]["filename"], "sheets": [s["name"] for s in cap["workbook"]["sheets"]], "named_ranges": cap["workbook"]["named_ranges"],
            "tables": tables, "kv_inputs": [p for t in cap["tables"] if t["kind"] == "kv" for p in t["pairs"]],
            "edges": edges, "relationships": cap["relationships"], "notes": cap["notes"][:60], "findings": [{k: v for k, v in f.items() if k != "examples"} for f in cap["findings"][:40]]}


def parse_json(text: str) -> dict:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.S)
    start, end = text.find("{"), text.rfind("}")
    return json.loads(text[start:end + 1])


def enrich(cap: dict, out_dir: Path | None = None, model: str | None = None) -> dict:
    client = anthropic.Anthropic()   # ANTHROPIC_API_KEY or the `ant auth login` profile
    body = PROMPT + json.dumps(compact(cap), default=str)
    with client.beta.messages.stream(model=model or MODEL, max_tokens=12000, output_config={"effort": "medium"},
                                     betas=["server-side-fallback-2026-07-01"], fallbacks="default",
                                     messages=[{"role": "user", "content": body}]) as stream:
        resp = stream.get_final_message()
    if resp.stop_reason == "refusal":
        raise RuntimeError("model refused")
    text = "".join(b.text for b in resp.content if b.type == "text")
    sem = parse_json(text)
    sem["_usage"] = {"input": resp.usage.input_tokens, "output": resp.usage.output_tokens, "model": resp.model}
    if out_dir:
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        (Path(out_dir) / "semantics.json").write_text(json.dumps(sem, indent=1))
    return sem
