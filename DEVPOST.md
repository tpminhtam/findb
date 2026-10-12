## Inspiration

Most of a finance team's knowledge lives in spreadsheets. Modern teams are moving that work into databases, and relational tables are the right home for the numbers. But the move drops everything around the numbers: the analyst's notes, the topside adjustments, the one-time placeholders.

I have spent years automating finance processes. As much as we want finance data to be structured, it is always the unstructured part that becomes the bottleneck and stalls the decision makers. I wanted a brain that keeps both.

## What it does

Drop an Excel workbook, or paste a link to a shared Google Sheet, and within a second finDB:

- **Finds the tables.** Including several tables on one sheet, title rows, hidden helper rows, totals and blocks of manual inputs.
- **Builds the schema.** Column types, keys, and relationships between tables, within a workbook and across workbooks.
- **Traces every formula.** Click any number to see what it comes from and what it feeds, with the formula or a plain-English explanation.
- **Flags what is fragile.** In the demo it reports that 46 rows are silently missing from 516 formulas, because the SUMIFS ranges stop short of the data. It also flags pasted values, numbers typed inside formulas and manual inputs.
- **Keeps the context.** Cell comments, commentary tabs, change logs and metric definitions go into the brain alongside the data.
- **Answers questions.** "Where does gross margin come from?" "What breaks if rows are added?" "Do our files agree on gross margin?" Answers are short and cited, and the brain remembers what analysts teach it.
- **Migrates.** Every table gets SQL DDL and is queryable in SQL, with a plan for what to load as tables, rebuild as views, or keep as knowledge.

## How we built it

- **Capture engine (Python, openpyxl).** It detects blocks, headers and totals, types columns from values and number formats, then tokenizes every formula and maps each reference to a table column, a total row or an input cell. Named ranges and lookups are resolved.
- **HydraDB as the brain.** Each table becomes a knowledge document with its own graph: tables, columns, totals and input cells, linked by "derives from" and "references" edges. Rows go in as app-source records with filterable metadata. Notes, comments and findings go in as app sources. What analysts teach it goes into per-analyst memory. Questions use hybrid search with graph context.
- **SQL as the muscle.** HydraDB is a graph and context store, not a relational one, so every table is also mirrored into DuckDB with generated DDL. Numbers run in SQL. Knowledge, lineage and memory live in HydraDB.
- **Claude** writes each answer from the retrieved context only, with numbered citations, and writes the plain-English descriptions.
- **Interface.** FastAPI and a single HTML page, with Cytoscape for the lineage map.

I built it solo in one afternoon, pairing with Claude Code, with eight automated tests and three sample workbooks for a fictional company (19 tables, 458 formulas, planted errors).

## Challenges we ran into

- **Finance data is relational and HydraDB is not.** Instead of forcing it, I split the job: SQL for the numbers, HydraDB for everything a warehouse drops.
- **Spreadsheets are not tables.** One sheet can hold a report, a pasted snapshot and a notes block, with a hidden row the formulas depend on.
- **Lineage explodes at the cell level.** Hundreds of formulas became an unreadable graph. Rolling them up to columns and total rows made it readable and kept it inside HydraDB's graph limits.
- **Rate limits and background indexing.** I capped the rows sent per table, moved the sync to the background with retries, and kept a local copy of the lineage so answers work the moment a file lands.
- **Too much interface.** I rebuilt the UI three times, each time showing less.

## Accomplishments that we're proud of

- It catches a real class of error. Rows appended below a formula range are invisible in Excel. finDB states it in one sentence and lists every total it breaks.
- It answers a cross-file question correctly: two workbooks define gross margin differently, and finDB says no, shows both definitions and names the owners.
- Drop to tables in under a second, with the brain synced to HydraDB about ten seconds later.
- Every answer cites its sources.

## What we learned

- The unstructured part is the product. Tables are easy. The notes, overrides and formula logic are where the knowledge is.
- A graph and memory store complements a warehouse. It does not compete with it.
- Column-level lineage is the right grain for people. Cell-level is for machines.
- Findings need plain sentences. "46 rows are missing" lands. "Stale range" does not.

## What's next for finDB

- Finish the migration: generate warehouse views and metric definitions from the formulas.
- Follow live links between workbooks and whole Google Drive folders.
- Watch a workbook over time: compare each month's version and flag what changed.
- Permissions, so answers respect who is allowed to see what.
- An MCP tool so finance agents can ask the brain directly.
