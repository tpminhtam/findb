"""Deterministic markdown table cards (HydraDB documents and UI detail), optionally enriched by semantics.json."""
from __future__ import annotations


def table_card(cap: dict, t: dict, sem: dict | None = None) -> str:
    wb = cap["workbook"]
    tsem = (sem or {}).get("tables", {}).get(t["id"], {})
    nodes = cap["nodes"]
    lines = [f"# {t['name']} ({t['id']})", "",
             f"Workbook: {wb['filename']} · Sheet: {t['sheet']} · Range: {t['range']} · Header row: {t['header_row']} · Data rows: {t['n_rows']} · Role: {t['role']} · Formula share: {t['formula_share']:.0%}"]
    if t.get("title"):
        lines.append(f"Title: {t['title']}")
    if tsem.get("description"):
        lines += ["", tsem["description"]]
    lines += ["", "## Columns", ""]
    for c in t["columns"]:
        desc = tsem.get("columns", {}).get(c["slug"], "")
        lines.append(f"- **{c['name']}** (`{c['slug']}`, {c['letter']}): {c['semantic']}" + (", candidate key" if c["is_key"] else "") +
                     (f", {c['formula_share']:.0%} formulas" if c["formula_share"] else "") + (f", nulls {c['null_rate']:.0%}" if c["null_rate"] else "") +
                     (f". {desc}" if desc else "") + (f" Samples: {c['samples'][:3]}" if c["samples"] else ""))
    if t["derived_rows"]:
        lines += ["", "## Derived rows (totals and metrics)", ""] + [f"- row {r}: {lab}" for r, lab in t["derived_rows"].items()]
    rels = [r for r in cap["relationships"] if r["child_table"] == t["id"] or r["parent_table"] == t["id"]]
    if rels:
        lines += ["", "## Relationships", ""] + [f"- {r['child_col']} -> {r['parent_col']} ({r['kind']}, coverage {r['coverage']})" for r in rels]
    col_ids = {c["id"] for c in t["columns"]} | {f"{t['id']}#r{r}" for r in t["derived_rows"]}
    ins = [e for e in cap["edges"] if e["dst"] in col_ids and e["predicate"] not in ("CONTAINS", "HAS_COLUMN")]
    if ins:
        lines += ["", "## Inputs (lineage)", ""]
        seen = set()
        for e in ins:
            key = (e["src"], e["dst"], e["predicate"])
            if key in seen:
                continue
            seen.add(key)
            lines.append(f"- {nodes.get(e['dst'], {}).get('name', e['dst'])} <- {nodes.get(e['src'], {}).get('name', e['src'])} via {e['predicate']}" +
                         (f" (named range {e['via_name']})" if e.get("via_name") else "") + (f"; e.g. {e['context'][:110]}" if e.get("context") else ""))
            if len(seen) >= 40:
                lines.append("- ..."); break
    outs = [e for e in cap["edges"] if e["src"] in col_ids and e["predicate"] not in ("CONTAINS", "HAS_COLUMN")]
    if outs:
        feeds = sorted({nodes.get(e["dst"], {}).get("name", e["dst"]) for e in outs})
        lines += ["", "## Feeds", "", "- " + "; ".join(feeds[:30])]
    if tsem.get("migration_note"):
        lines += ["", "## Migration note", "", tsem["migration_note"]]
    lines += ["", "## DDL", "", "```sql", next((blk for blk in cap["ddl"].split("\n\n") if f"__{t['slug']} (" in blk), "").strip(), "```"]
    return "\n".join(lines)
