"""
Ajera Project Import - command line.

    python cli.py connect                      test the connection, show which database you're on
    python cli.py template [out.xlsx]          blank workbook with YOUR Ajera's names + dropdowns
    python cli.py export --latest 10 [out]     copy existing projects into the import format
    python cli.py export --ids 1001,1002 [out]
    python cli.py check  file.xlsx             validate only - changes nothing in Ajera
    python cli.py create file.xlsx             validate, then create (asks you to type YES)
    python cli.py log                          show projects this tool has created

Windows users can double-click "Ajera Project Import.bat" instead.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

from dotenv import load_dotenv

HERE = Path(__file__).resolve().parent
load_dotenv(HERE / ".env")

from ajera_client import Ajera  # noqa: E402
import importer  # noqa: E402
import workbook  # noqa: E402


def cmd_connect(aj: Ajera, _):
    aj.call("ListCompanies")
    print(f"Connected. DatabaseID = {aj.database_id}  (sample data: {aj.is_sample_data})")
    try:
        aj.check_write_target()
        print("Writes are ENABLED for this database (AJERA_EXPECTED_DB_ID matches).")
    except RuntimeError as e:
        print(f"Writes are BLOCKED: {e}")


def cmd_template(aj: Ajera, a):
    out = Path(a.out or "Ajera Project Import - template.xlsx")
    lk = importer.load_lookups(aj, want_existing_ids=False)
    workbook.build_workbook(out, [], [], workbook.lookup_rows_from(lk),
                            "Ajera Project Import", workbook.today_note() + " from your Ajera database")
    print(f"Template written: {out.resolve()}")
    print("Lookups: " + ", ".join(f"{k} {len(v)}" for k, v in lk.rows.items()))


def cmd_export(aj: Ajera, a):
    out = Path(a.out or "Ajera Project Import - export.xlsx")
    if a.ids:
        wanted = {s.strip().lower() for s in a.ids.split(",") if s.strip()}
        keys = [p["ProjectKey"] for p in aj.list_all_projects() if str(p.get("ID", "")).strip().lower() in wanted]
    else:
        pat = re.compile(a.id_pattern) if a.id_pattern else None
        active = aj.call("ListProjects", {"FilterByStatus": ["Active"]}).get("Projects", [])
        active = [p for p in active if not pat or pat.fullmatch(str(p.get("ID", "")))]
        keys = sorted((p["ProjectKey"] for p in active), reverse=True)[a.skip:a.skip + a.latest]
    if not keys:
        sys.exit("No matching projects found.")
    lk = importer.load_lookups(aj, want_existing_ids=False)
    projects, phases, stats = workbook.export_projects(aj, lk, keys)
    workbook.build_workbook(out, projects, phases, workbook.lookup_rows_from(lk),
                            "Ajera Project Import - export", workbook.today_note() + " (copied from existing Ajera projects)")
    print(f"Export written: {out.resolve()}")
    print("  " + ", ".join(f"{k}={v}" for k, v in stats.items()))


def _print_problems(problems):
    print(f"\n{len(problems)} problem(s) - nothing was created:")
    for p in problems:
        print(f"  - {p}")


def cmd_check(aj: Ajera, a):
    projects, problems, _ = importer.check_file(aj, a.file)
    if problems:
        _print_problems(problems)
        return 1
    n_ph = sum(len(p.args["Phases"]) for p in projects)
    print(f"All good: {len(projects)} project(s), {n_ph} phase(s) ready to create. (Nothing was changed in Ajera.)")
    print(f"Target DatabaseID {aj.database_id}.")
    return 0


def cmd_create(aj: Ajera, a):
    aj.check_write_target()
    if cmd_check(aj, a):
        return 1
    if not a.yes:
        print("\nThe Ajera API cannot delete projects. Anything created stays until someone removes it in Ajera.")
        if input("Type YES to create these projects: ").strip() != "YES":
            print("Cancelled - nothing was created.")
            return 1
    outcomes, problems, results = importer.import_file(aj, a.file)
    if problems:
        _print_problems(problems)
        return 1
    for o in outcomes:
        print(f"  row {o.row:>3}  {o.job_id:<12} {'OK ' if o.ok else 'ERR'}  key={o.project_key or '-':<8} {o.message}")
    ok = sum(o.ok for o in outcomes)
    print(f"\nCreated {ok} of {len(outcomes)}. Results workbook: {results}")
    print(f"Created-projects log: {importer.created_log_path()}")
    return 0 if ok == len(outcomes) else 1


def cmd_log(_aj, _a):
    path = importer.created_log_path()
    if not path.exists():
        print("No projects have been created by this tool yet.")
        return
    with path.open(encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    print(f"{len(rows)} project(s) created by this tool (full list: {path}):")
    for r in rows:
        print(f"  {r['created_at']}  db={r['database_id']}  key={r['project_key']:<8} {r['job_id']}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Create Ajera projects from an Excel workbook.")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("connect")
    t = sub.add_parser("template")
    t.add_argument("out", nargs="?")
    e = sub.add_parser("export")
    e.add_argument("out", nargs="?")
    e.add_argument("--ids", help="comma-separated project IDs")
    e.add_argument("--latest", type=int, default=10, help="newest N active projects (default 10)")
    e.add_argument("--skip", type=int, default=0, help="skip the newest N first")
    e.add_argument("--id-pattern", help="regular expression the project ID must match, e.g. \\d{5}")
    for name in ("check", "create"):
        s = sub.add_parser(name)
        s.add_argument("file")
        if name == "create":
            s.add_argument("--yes", action="store_true", help="don't ask for confirmation")
    sub.add_parser("log")
    a = ap.parse_args(argv)
    if a.cmd == "log":
        return cmd_log(None, a)
    try:
        aj = Ajera()
        return {"connect": cmd_connect, "template": cmd_template, "export": cmd_export,
                "check": cmd_check, "create": cmd_create}[a.cmd](aj, a) or 0
    except RuntimeError as e:
        print(f"Stopped: {e}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
