import json
from pathlib import Path

import pytest

from finance_brain.capture import capture, cross_workbook_relationships
from finance_brain.store_local import LocalStore

SAMPLES = Path(__file__).resolve().parent.parent.parent / "samples"


@pytest.fixture(scope="module")
def caps():
    return {f: capture(SAMPLES / f"{f}.xlsx") for f in ["Ardent_FY26_Budget_vs_Actuals", "Ardent_FY26_Headcount_Plan", "Ardent_Q3_FY26_Board_KPIs"]}


def test_bva_tables(caps):
    cap = caps["Ardent_FY26_Budget_vs_Actuals"]
    ids = {t["id"]: t for t in cap["tables"]}
    assert ids["actuals"]["n_rows"] == 936 and len(ids["actuals"]["columns"]) == 8
    assert ids["budget"]["n_rows"] == 1248
    assert ids["bva_summary"]["header_row"] == 4 and 3 in ids["bva_summary"]["helper_rows"]
    assert set(ids["bva_summary"]["derived_rows"].values()) >= {"Total Revenue", "Total COGS", "Gross profit", "Gross margin %", "EBITDA"}
    assert sum(1 for t in cap["tables"] if t["sheet"] == "BvA Summary" and t["kind"] == "table") == 2
    assert any(t["kind"] == "kv" and t["sheet"] == "Cover" for t in cap["tables"])
    assert cap["workbook"]["named_ranges"]["FX_EUR_USD"] == "Assumptions!$B$2"


def test_bva_relationships_and_lineage(caps):
    cap = caps["Ardent_FY26_Budget_vs_Actuals"]
    rels = {(r["child_col"], r["parent_col"]) for r in cap["relationships"]}
    assert ("actuals.account", "coa.account") in rels and ("actuals.department", "departments.dept_code") in rels
    assert ("bva_summary.account", "coa.account") in rels
    names = {e["src"] for e in cap["edges"] if e["dst"] == "bva_summary.ytd_actual" and e["predicate"] == "SUMIFS"}
    assert {"actuals.amount", "actuals.account", "actuals.entity"} <= names
    assert any(e.get("via_name") == "FX_EUR_USD" for e in cap["edges"])
    gm = {e["src"] for e in cap["edges"] if e["dst"] == "bva_summary#r12"}
    assert gm == {"bva_summary#r11", "bva_summary#r7"}


def test_bva_findings(caps):
    cap = caps["Ardent_FY26_Budget_vs_Actuals"]
    stale = [f for f in cap["findings"] if f["kind"] == "stale_range"]
    assert len(stale) == 1 and stale[0]["excluded_data_rows"] == 46 and stale[0]["table"] == "actuals" and stale[0]["formulas"] > 400
    kinds = {f["kind"] for f in cap["findings"]}
    assert {"constant_in_formula", "pasted_value", "hidden_rows"} <= kinds
    assert sum(1 for n in cap["notes"] if n["kind"] == "comment") == 5


def test_other_workbooks(caps):
    hc = caps["Ardent_FY26_Headcount_Plan"]
    assert {t["id"] for t in hc["tables"] if t["kind"] == "table"} == {"roster", "comp_bands", "hiring_plan", "hc_summary"}
    assert any(e["src"] == "hc_summary.ending_hc" and e["dst"] == "hc_summary.beginning_hc" for e in hc["edges"])
    assert any(p["key"].startswith("Benefits load") and p["value"] == 0.22 for t in hc["tables"] if t["kind"] == "kv" for p in t["pairs"])
    kpi = caps["Ardent_Q3_FY26_Board_KPIs"]
    assert {t["id"] for t in kpi["tables"] if t["kind"] == "table"} >= {"customers", "arr_bridge", "kpi_summary", "cohorts", "definitions"}
    assert any(e["src"] == "arr_bridge.ending_arr" and e["dst"] == "arr_bridge.beginning_arr" for e in kpi["edges"])


def test_cross_workbook(caps):
    rels = cross_workbook_relationships(list(caps.values()))
    pairs = {(r["child_col"], r["parent_col"]) for r in rels}
    assert ("hiring_plan.department", "departments.dept_code") in pairs


def test_local_store_query(caps, tmp_path):
    store = LocalStore(root=tmp_path)
    for cap in caps.values():
        store.add_capture(cap)
    r = store.query("What breaks if more rows are appended to the Actuals tab?", persona="tam", graph=True)
    assert any(c["kind"] == "finding" and "stale_range" in c["text"] for c in r["chunks"])
    r2 = store.query("Where does YTD gross margin on the BvA Summary come from?", graph=True)
    assert any("Gross margin" in p for p in r2["paths"])
    store.remember("tam", "The BvA is owned by Maya Chen.")
    assert store.memories_for("tam")[0]["text"].startswith("The BvA")


def test_ddl_executes_in_duckdb(caps):
    import duckdb
    con = duckdb.connect()
    for cap in caps.values():
        for stmt in cap["ddl"].split(";\n"):
            if "CREATE TABLE" in stmt:
                con.execute(stmt)
    assert con.execute("select count(*) from information_schema.tables").fetchone()[0] >= 15


def test_api_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setenv("FB_STORE", "local")   # never push test data to HydraDB
    from fastapi.testclient import TestClient
    import finance_brain.app as appmod
    monkeypatch.setattr(appmod, "OUT", tmp_path / "out")
    (tmp_path / "out").mkdir()
    appmod.local = appmod.LocalStore(tmp_path / "out")
    client = TestClient(appmod.app)
    with open(SAMPLES / "Ardent_FY26_Budget_vs_Actuals.xlsx", "rb") as f:
        r = client.post("/api/upload", files={"file": ("Ardent_FY26_Budget_vs_Actuals.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert r.status_code == 200 and "bva_summary" in r.json()["tables"]
    slug = r.json()["slug"]
    assert client.get(f"/api/workbooks/{slug}/graph").json()["n_edges"] > 50
    a = client.post("/api/ask", json={"question": "What breaks if we append rows to Actuals?", "persona": "t", "graph": True}).json()
    assert a["mode"] in ("generated", "retrieval-only") and any("stale_range" in c["text"] for c in a["chunks"])
    assert client.get("/api/sql", params={"q": "drop table x"}).status_code == 400
    assert client.get("/api/workbooks/{}/ddl".format(slug)).text.count("CREATE TABLE") >= 8
