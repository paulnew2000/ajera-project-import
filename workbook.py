"""
Builds nicely formatted import workbooks:

  * a blank template (with a Lookups sheet of YOUR Ajera's names and dropdowns),
  * an export of existing Ajera projects in the same format (handy as a starting point),
  * the sample workbook in examples/.

Every workbook has the same four sheets: Instructions, Projects, Phases, Lookups.
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from importer import BILLING_LABEL, BILLING_TYPES, PHASE_COLUMNS, PROJECT_COLUMNS, STATUSES, Lookups, job_id_text

NAVY = "1F3864"
REQ_FILL = PatternFill("solid", fgColor=NAVY)
OPT_FILL = PatternFill("solid", fgColor="5B7DB1")
BAND_FILL = PatternFill("solid", fgColor="EEF2F8")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="C9D3E3")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WIDTHS = {"Job ID": 12, "Description": 42, "Client": 32, "Department": 20, "Project Type": 22,
          "Billing Type": 18, "Rate Table": 24, "Project Manager": 24, "Principal in Charge": 24,
          "Status": 12, "Start Date": 13, "End Date": 13, "Invoice Group": 18, "Invoice Format": 24,
          "Notes": 60, "Phase Description": 34}
DATE_COLS = {"Start Date", "End Date"}

# Lookups sheet layout: (heading, source) - each gets a Name column and a Key column
LOOKUP_BLOCKS = [("Client", "Client"), ("Department", "Department"), ("Project Type", "Project Type"),
                 ("Rate Table", "Rate Table"), ("Invoice Format", "Invoice Format"),
                 ("Project Manager", "managers"), ("Principal in Charge", "principals")]


def _data_sheet(ws, columns, rows: list[dict]):
    for c, (name, required, help_text) in enumerate(columns, 1):
        cell = ws.cell(1, c, name)
        cell.fill, cell.font, cell.border = (REQ_FILL if required else OPT_FILL), WHITE_BOLD, GRID
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.comment = Comment(("REQUIRED. " if required else "") + help_text, "Ajera Project Import", width=260, height=90)
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(name, 18)
    ws.row_dimensions[1].height = 30
    for r, row in enumerate(rows, 2):
        for c, (name, _, _) in enumerate(columns, 1):
            cell = ws.cell(r, c, row.get(name))
            cell.border = GRID
            cell.alignment = Alignment(vertical="top", wrap_text=name == "Notes")
            if name in DATE_COLS:
                cell.number_format = "yyyy-mm-dd"
            if r % 2 == 1:
                cell.fill = BAND_FILL
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(rows) + 1, 2)}"


def _lookups_sheet(wb, lk_rows: dict[str, list[tuple[str, object]]]) -> dict[str, str]:
    """Write the Lookups sheet; return {column heading: range formula} for dropdowns."""
    ws = wb.create_sheet("Lookups")
    blocks = [(h, lk_rows.get(src, [])) for h, src in LOOKUP_BLOCKS]
    blocks += [("Billing Type", [(b, "") for b in BILLING_TYPES]), ("Status", [(s, "") for s in STATUSES])]
    ranges = {}
    col = 1
    for heading, items in blocks:
        for i, h in enumerate((heading, "Key")):
            cell = ws.cell(1, col + i, h)
            cell.fill, cell.font, cell.border = REQ_FILL, WHITE_BOLD, GRID
        for r, (name, key) in enumerate(items, 2):
            ws.cell(r, col, name).border = GRID
            ws.cell(r, col + 1, key).border = GRID
        ws.column_dimensions[get_column_letter(col)].width = 34
        ws.column_dimensions[get_column_letter(col + 1)].width = 8
        ws.column_dimensions[get_column_letter(col + 2)].width = 3
        letter = get_column_letter(col)
        if items:
            ranges[heading] = f"Lookups!${letter}$2:${letter}${len(items) + 1}"
        col += 3
    ws.freeze_panes = "A2"
    return ranges


def _dropdowns(ws, columns, ranges: dict[str, str], extra: dict[str, str] | None = None):
    for c, (name, _, _) in enumerate(columns, 1):
        src = ranges.get(name)
        if not src:
            continue
        dv = DataValidation(type="list", formula1=f"={src}", allow_blank=True, showErrorMessage=True,
                            errorStyle="warning", errorTitle="Not in the Lookups list",
                            error="This value isn't on the Lookups sheet. Keep it only if you are sure it matches Ajera "
                                  "(or it is an Ajera key number).")
        dv.add(f"{get_column_letter(c)}2:{get_column_letter(c)}1000")
        ws.add_data_validation(dv)


def _instructions(wb, title: str, note: str):
    ws = wb.active
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 90
    ws["B2"] = title
    ws["B2"].font = Font(bold=True, size=18, color=NAVY)
    ws["B3"] = note
    ws["B3"].font = Font(italic=True, color="555555")
    steps = [
        "1.  Fill in one row per project on the Projects sheet.",
        "2.  Optional: list phases on the Phases sheet (one row per phase, same Job ID). No phases = one phase called 'General'.",
        "3.  Names must match Ajera exactly. The Lookups sheet lists what your Ajera contains; dropdowns use it.",
        "4.  Run CHECK first. It reads Ajera but changes nothing, and lists every problem by sheet and row.",
        "5.  Run CREATE only when CHECK is clean. It makes the projects, then saves a *_results.xlsx copy.",
        "6.  IMPORTANT: the Ajera API cannot delete projects. Every project created is logged in created_projects.csv",
        "     so you can find and remove test projects in Ajera by hand.",
    ]
    for i, s in enumerate(steps, 5):
        ws.cell(i, 2, s).font = Font(size=11)
    r = 5 + len(steps) + 1
    for sheet, cols in (("Projects", PROJECT_COLUMNS), ("Phases", PHASE_COLUMNS)):
        ws.cell(r, 2, f"{sheet} sheet columns").font = Font(bold=True, size=13, color=NAVY)
        r += 1
        for c, h in enumerate(("Column", "Required?", "What to enter"), 2):
            cell = ws.cell(r, c, h)
            cell.fill, cell.font, cell.border = REQ_FILL, WHITE_BOLD, GRID
        for name, required, help_text in cols:
            r += 1
            for c, val in enumerate((name, "Yes" if required else "No", help_text), 2):
                cell = ws.cell(r, c, val)
                cell.border = GRID
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if required:
                    cell.font = Font(bold=True)
        r += 2


def build_workbook(path: Path, projects: list[dict], phases: list[dict],
                   lookup_rows: dict[str, list[tuple[str, object]]], title: str, note: str) -> Path:
    wb = Workbook()
    _instructions(wb, title, note)
    ws_p = wb.create_sheet("Projects")
    _data_sheet(ws_p, PROJECT_COLUMNS, projects)
    ws_ph = wb.create_sheet("Phases")
    _data_sheet(ws_ph, PHASE_COLUMNS, phases)
    ranges = _lookups_sheet(wb, lookup_rows)
    _dropdowns(ws_p, PROJECT_COLUMNS, ranges)
    _dropdowns(ws_ph, PHASE_COLUMNS, ranges)
    wb.active = 1  # open on the Projects sheet
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def lookup_rows_from(lk: Lookups) -> dict[str, list[tuple[str, object]]]:
    rows = dict(lk.rows)
    rows["managers"] = [(n, k) for n, k in lk.rows["Employee"] if k in lk.managers]
    rows["principals"] = [(n, k) for n, k in lk.rows["Employee"] if k in lk.principals]
    return rows


# -- export existing projects -------------------------------------------------------------
def _key_of(v):
    return v.get("EmployeeKey") if isinstance(v, dict) else v


def _date(v):
    if not v:
        return None
    try:
        return datetime.fromisoformat(str(v)[:10]).date()
    except ValueError:
        return None


def export_projects(aj, lk: Lookups, keys: list[int]) -> tuple[list[dict], list[dict], dict]:
    """Read projects from Ajera and turn them into Projects/Phases rows (names, not keys).
    Returns (projects, phases, stats). Only top-level phases of the first invoice group's
    client are kept - the import format has one client per project."""
    projects, phases = [], []
    stats = {"projects": 0, "phases": 0, "extra_invoice_groups": 0, "unresolved_names": 0}
    for p in aj.get_projects(keys):
        groups = p.get("InvoiceGroups") or [{}]
        first = groups[0]
        stats["extra_invoice_groups"] += len(groups) - 1
        row = {
            "Job ID": job_id_text(p.get("ID")),
            "Description": p.get("Description"),
            # GetProjects nests the client ({"Client": {"ClientKey": ...}}); CreateProjects takes a flat ClientKey
            "Client": lk.name_of("Client", (first.get("Client") or {}).get("ClientKey") or first.get("ClientKey")),
            "Department": lk.name_of("Department", p.get("DepartmentKey")),
            "Project Type": lk.name_of("Project Type", p.get("ProjectTypeKey")),
            "Billing Type": BILLING_LABEL.get(p.get("BillingType"), p.get("BillingType")),
            "Rate Table": lk.name_of("Rate Table", p.get("RateTableKey")),
            "Project Manager": lk.name_of("Employee", _key_of(p.get("ProjectManager"))),
            "Principal in Charge": lk.name_of("Employee", _key_of(p.get("PrincipalInCharge"))),
            "Status": p.get("Status"),
            "Start Date": _date(p.get("EstimatedStartDate")),
            "End Date": _date(p.get("EstimatedCompletionDate")),
            "Invoice Group": first.get("Description"),
            "Invoice Format": lk.name_of("Invoice Format", first.get("InvoiceFormatKey")) if first.get("InvoiceFormatKey") else None,
            "Notes": p.get("Notes"),
        }
        # a bare number in a name column means the key didn't resolve to an active record
        stats["unresolved_names"] += sum(1 for c in ("Client", "Department", "Project Type", "Rate Table")
                                         if str(row[c]).isdigit())
        projects.append(row)
        for g in groups:
            for ph in g.get("Phases") or []:
                phases.append({
                    "Job ID": row["Job ID"],
                    "Phase Description": ph.get("Description"),
                    "Department": lk.name_of("Department", ph.get("DepartmentKey")) if ph.get("DepartmentKey") else None,
                    "Billing Type": BILLING_LABEL.get(ph.get("BillingType"), ph.get("BillingType")),
                })
        stats["projects"] += 1
    stats["phases"] = len(phases)
    return projects, phases, stats


def today_note() -> str:
    return f"Generated {date.today():%B %d, %Y}"
