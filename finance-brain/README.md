# finDB

**Your Finance Org Brain.** Drop an Excel workbook or a link-shared Google Sheet; finDB finds every table,
types every column, infers keys and relationships, follows every formula into a lineage graph, scores what is fragile, and loads the
result into a knowledge store (HydraDB when the key is present, a local BM25 + graph store otherwise) plus a DuckDB mirror with DDL.
Then people and agents can ask where a number comes from, what breaks if it changes, who owns an input, and whether two files agree.

Built for the Company Brain Hackathon (TechEquity Ai @ Snowflake SVAI, Oct 12 2026).

## Run

```bash
cd "~/Desktop/Company Brain Hackathon/finance-brain"
uv run uvicorn finance_brain.app:app --host 127.0.0.1 --port 8765
open http://127.0.0.1:8765
```

Drop any .xlsx on the sidebar (the Ardent samples are in `../samples/`). Tabs: Tables, Lineage, Issues, Ask, SQL.
On load, an intro types "finDB" in the centre, then the word flies into the sidebar corner while the dark panel collapses into the sidebar (about 3.5 s; click or press any key to skip; links with `?…` parameters skip it).
Deep links (handy for the video): `/#lineage`, `/#issues`, `/#ask`, `/#sql`, `/?q=<question>#ask` asks a question, `/?trace=gross margin %#lineage` traces a number, `/?table=actuals` opens a table card.
Lineage tab: opens on a table-level map (solid = feeds, dashed = shares a key). Click a table, pick a number, and the view becomes a two-hop trace; click any box to keep tracing, or switch between Comes from and Feeds.

Tests: `uv run pytest -q` (8 tests: detection, lineage, findings, relationships, local store, DDL executes in DuckDB, API round trip).

## Keys (both optional; the app degrades gracefully)

`.env` in this folder (never committed):

```
ANTHROPIC_API_KEY=sk-ant-...     # generated answers and "enrich with Claude"; without it, Ask returns the retrieved context only
HYDRA_DB_API_KEY=...             # from dashboard.hydradb.com (the hackathon key)
FB_STORE=local                   # set to hydradb once the key is in
FB_MODEL=claude-opus-5-5
FB_EFFORT=low                    # low = fast demo answers; medium for richer ones
FB_HYDRA_ROW_CAP=500             # rows per table pushed to HydraDB as app-source records (DuckDB always holds all rows)
```

Both keys were supplied on Oct 12 and live only in `.env`. Restart the server after editing `.env`.

## On-site: switching on HydraDB (10 minutes)

1. Put `HYDRA_DB_API_KEY=...` and `FB_STORE=hydradb` in `.env`, restart the server. On the first request the app creates database
   `finance_brain` (with a metadata schema: workbook, sheet, table, record_type, period, entity) and waits for `ready_for_ingestion`.
   The header pill shows `store: hydradb`; if it shows an error, read `out/hydra_log.jsonl`.
2. Push what is already captured: `curl -X POST localhost:8765/api/hydra/push` (or just re-upload a file; uploads go to both stores).
   Per table it ingests a markdown card as a knowledge document with a `graph_payload` (TABLE, COLUMN, METRIC_ROW, CELL entities;
   DERIVES_FROM / REFERENCES / HAS_COLUMN relations), up to 500 rows as `custom` app-source records with filterable metadata, and all
   notes, cell comments and findings as app sources. Memories go to collection `analyst_<persona>` with `infer: true`.
3. Ask something. Queries use `type="all"` over `catalog`, `rows` and the persona collection, `query_by="hybrid"`, `graph_context=True`,
   `mode="thinking"`, `query_apps=True`. Graph paths from HydraDB (query_paths / chunk_relations) are passed to Claude as text; the
   local lineage paths are appended too, so the lineage story works even if the HydraDB graph is still indexing.

`finance_brain/store_hydra.py` was tested live on Oct 12 2026 (14:05 to 14:15 PDT) with the hackathon key: database creation with the
metadata schema (ready in about 10 s), 19 table cards with `graph_payload`, 1,417 row records and 76 notes/findings ingested with zero
errors in about 27 s total, memories accepted, queries returning `data.chunks[].chunk_content`, `data.graph.paths[].combined_context`
and `data.graph_context.chunk_relations`. The prompt uses HydraDB's own `build_string(result)` markdown plus the local lineage paths.
Indexing passes through `graph_creation` for several minutes after ingest (entity extraction); data is queryable before that finishes.
Every call is logged to `out/hydra_log.jsonl`. The header pill shows `store: hydradb`.

## What the capture does (finance_brain/capture.py)

1. Blocks: rectangular regions split by blank rows then blank columns, per sheet (hidden rows included).
2. Tables: header row = first row of mostly strings (bold strings win), titles above it, helper rows (hidden keys) kept, totals and margin
   lines marked as derived rows; two-column label/value blocks become key-value inputs; text-only blocks become notes.
3. Schema: type and semantic per column (money, percent, date, identifier, narrative...) from values and number formats; null rate,
   distinct ratio, candidate keys; table role (data, dimension, report, assumptions, narrative).
4. Relationships: foreign keys by value containment into keyed dimension tables (one parent per child column), plus VLOOKUP/XLOOKUP joins.
   Cross-workbook keys are inferred in the store.
5. Lineage: every formula is tokenised (openpyxl Tokenizer); each range maps to column nodes, metric-row nodes (derived rows), cell nodes
   (inputs, assumptions, helper keys) or external workbooks; named ranges are resolved; lookups are split into key and returned column.
6. Findings: stale ranges (a formula stops before the referenced table ends), typed constants inside formulas, lookups without exact match,
   external links, pasted/copied values (from comments), hidden rows and columns, manual inputs that formulas depend on.
7. Scorecard and migration plan: load as tables / rebuild as views or metrics / keep as assumptions / ingest as knowledge; DDL per table.

Outputs per workbook in `out/<slug>/`: `capture.json`, `ddl.sql`, `cards/*.md`, optional `semantics.json` (Claude enrichment).
DuckDB mirror: `out/brain.duckdb` (read-only SQL from the UI). Memories: `out/memories.json`.

## Test, then demo

**Two-minute test**

1. Open http://127.0.0.1:8765. Top right should read `HydraDB` with a green dot (`HydraDB · connecting` for a few seconds after a restart).
2. Hover a workbook in the sidebar, click ×, then drag the same file back from `../samples/`. Tables appear in under a second;
   the chip shows `HydraDB · syncing` for about 10 seconds while the copy goes to HydraDB in the background.
3. Tables: click **Actuals**. You should see eight column chips (Amount tagged money) and five sample rows.
4. Lineage: click **BvA Summary**, then **Gross margin %**. You should see the trace and a one-sentence explanation under it.
5. Issues: the top card reads "46 rows are missing from 516 formulas".
6. Ask: click the first suggestion. A bold one-line answer and up to four bullets appear in about 5 seconds.
7. SQL: press Run. Six departments with Q3 amounts.
8. Optional: `uv run pytest -q` (8 tests, never touches HydraDB).

**Demo (about 2.5 minutes)**

Before recording: remove the three workbooks with × so the screen starts empty. Keep `../samples/` open in Finder.

1. Problem, 15 s, on the empty screen: "Finance runs on spreadsheets nobody fully understands. finDB reads them like a senior analyst."
2. Drop `Ardent_FY26_Budget_vs_Actuals.xlsx`. Say the three numbers: 9 tables, 315 formulas, 8 issues. Open **Actuals**: typed columns, real rows.
3. Lineage. "This is how the workbook actually flows." Click **BvA Summary**, then **Gross margin %**. Read the sentence under the graph.
   Click **Total Revenue** or a YTD box to walk one hop further back to Actuals and the FX rate.
4. Issues. Read the red card: "46 rows are missing from 516 formulas." Click **Ask what breaks**. The answer lists the affected totals with sources.
5. Drop `Ardent_Q3_FY26_Board_KPIs.xlsx` and `Ardent_FY26_Headcount_Plan.xlsx`. In Ask: "Do our files agree on gross margin?" (No, and why.)
6. Memory. Click **+ Teach the brain**, save "Maya Chen owns the BvA file". Ask "Who do I ask before changing the gross margin definition?"
7. SQL. Press Run, then click **Export SQL** on Tables. "The data is out of the spreadsheet: tables in SQL, knowledge and lineage in HydraDB."
8. Close, 10 s: "HydraDB is the brain: catalog, lineage graph, rows and memory. The warehouse is the muscle."

Deep links if you prefer to jump instead of click: `/?trace=gross margin %#lineage`, `/?q=<question>#ask`, `/?table=actuals`.

## Known limits

- Heuristic table detection: tuned on finance layouts (titles, hidden helper rows, totals). Wild sheets may need the Claude fallback
  (not built). Cell-level lineage is aggregated to columns to stay under HydraDB's 10,000-relation limit.
- External workbook links are detected but not followed; cross-file dependencies in the samples are pasted values with comments.
- Google Sheets import needs "anyone with the link"; it uses the xlsx export URL and keeps formulas.
