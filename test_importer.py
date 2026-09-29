"""Offline tests: synthetic lookups + the sample workbook. No Ajera connection needed.

    python -m pytest -q
"""
from pathlib import Path

import pytest
from openpyxl import load_workbook

import importer
from importer import Lookups, read_workbook, validate

SAMPLE = Path(__file__).parent / "examples" / "Hogwarts sample projects.xlsx"


def fake_lookups(existing=()):
    wb = load_workbook(SAMPLE, data_only=True)
    ws = wb["Lookups"]
    heads = [c.value for c in ws[1]]
    rows = {}
    key = 100
    for i, h in enumerate(heads):
        if h in ("Client", "Department", "Project Type", "Rate Table", "Invoice Format"):
            names = [r[i] for r in ws.iter_rows(min_row=2, values_only=True) if r[i]]
            rows[h] = [(n, key + j) for j, n in enumerate(names)]
            key += 100
    rows["Employee"] = [("Albus Dumbledore", 1), ("Minerva McGonagall", 2), ("Argus Filch", 3)]
    return Lookups(company_key=1, rows=rows, managers={1, 2}, principals={1},
                   existing_ids={e.lower() for e in existing})


def test_sample_validates_against_its_own_lookups():
    projects, problems = read_workbook(SAMPLE)
    assert not problems
    assert len(projects) == 10
    assert validate(projects, fake_lookups()) == []
    p = projects[0].args
    assert p["Project"]["ID"] == "HP001"
    assert p["Project"]["Description"].startswith("Test Project:")
    assert p["Project"]["RateTableKey"] and p["InvoiceGroups"][0]["ClientKey"]
    assert "ProjectManager" not in p["Project"]  # blank = left out, not sent as null
    assert all(ph["Description"] and ph["DepartmentKey"] for ph in p["Phases"])


def test_existing_and_duplicate_ids_are_rejected():
    projects, _ = read_workbook(SAMPLE)
    projects[1].values["Job ID"] = "HP001"
    probs = [str(p) for p in validate(projects, fake_lookups(existing=["HP003"]))]
    assert any("row 3, Job ID: is also used on row 2" in p for p in probs)
    assert any("row 4, Job ID: already exists in Ajera" in p for p in probs)


def test_problem_messages_never_echo_cell_values():
    projects, _ = read_workbook(SAMPLE)
    projects[0].values["Client"] = "Secret Client LLC"
    projects[0].values["Project Manager"] = "Argus Filch"  # exists but is not a PM
    probs = [str(p) for p in validate(projects, fake_lookups())]
    assert probs and not any("Secret" in p or "Filch" in p for p in probs)
    assert any("not flagged as a Project Manager" in p for p in probs)


@pytest.mark.parametrize("value,expected", [("Fixed Fee", "FixedFee"), ("fixedfee", "FixedFee"),
                                            ("Time and Expense", "TimeAndExpense"), ("bogus", None)])
def test_billing_type_labels(value, expected):
    assert importer._billing(value) == expected


def test_project_without_phases_gets_a_default_phase():
    projects, _ = read_workbook(SAMPLE)
    projects[0].phases = []
    validate(projects, fake_lookups())
    assert projects[0].args["Phases"][0]["Description"] == importer.DEFAULT_PHASE


def mock_firm_lookups():
    """A pretend firm whose names share nothing with the Hogwarts sample."""
    return Lookups(
        company_key=1,
        rows={"Client": [("Acme Corp", 11), ("ZZ Test Client", 12)],
              "Department": [("Austin", 21), ("Denver", 22)],
              "Project Type": [("Overhead", 31), ("Environmental", 32), ("Engineering", 33)],
              "Rate Table": [("Acme 2026 Standard Rates", 41), ("2025 Standard", 42), ("2026 Standard", 43)],
              "Invoice Format": [], "Employee": []},
        managers=set(), principals=set(), existing_ids={"hp001", "hp007"})


def test_adapt_sample_passes_check_in_another_firm(tmp_path, monkeypatch):
    import adapt
    lk = mock_firm_lookups()
    monkeypatch.setattr(importer, "load_lookups", lambda aj, want_existing_ids=True: lk)
    out, summary, problems = adapt.adapt_sample(None, dest=tmp_path / "adapted.xlsx")
    assert problems == []
    assert "key 12" in summary[0]                    # the one test-looking client
    assert "key 43" in summary[1]                    # 2026 Standard, not the client-specific one
    assert "HP008 to HP017" in summary[4]            # continues after existing HP007
    projects, _ = read_workbook(out)
    assert {p.values["Project Type"] for p in projects} <= {"Environmental", "Engineering"}  # no Overhead


def test_adapt_sample_needs_a_client_when_ambiguous(monkeypatch):
    import adapt
    lk = mock_firm_lookups()
    lk.rows["Client"].append(("Sample Client", 13))
    monkeypatch.setattr(importer, "load_lookups", lambda aj, want_existing_ids=True: lk)
    with pytest.raises(ValueError, match="Several clients"):
        adapt.adapt_sample(None)


def test_number_job_ids_match_phases():
    assert importer.job_id_text(11001.0) == "11001"
    assert importer.job_id_text(" HP001 ") == "HP001"
