"""
Ajera Project Import - MCP server.

Lets an AI assistant (Claude Desktop, Claude Code, ...) run the same steps as cli.py:
check a workbook, create the projects, build a template, look at the created-projects log.
Tool results report sheet/row/column and counts - never client or employee names.

Add to your MCP client config (see README):
    "ajera-project-import": {"command": "python", "args": ["<folder>/server.py"]}
"""
from __future__ import annotations

import csv
from pathlib import Path

from dotenv import load_dotenv

HERE = Path(__file__).resolve().parent
load_dotenv(HERE / ".env")

from mcp.server.fastmcp import FastMCP  # noqa: E402

import importer  # noqa: E402
import workbook  # noqa: E402
from ajera_client import Ajera  # noqa: E402

mcp = FastMCP("ajera-project-import")
_aj: Ajera | None = None


def aj() -> Ajera:
    global _aj
    if _aj is None:
        _aj = Ajera()
    return _aj


@mcp.tool()
def check_connection() -> str:
    """Test the Ajera connection and report which database it points at and whether writes are allowed."""
    a = aj()
    a.call("ListCompanies")
    try:
        a.check_write_target()
        writes = "Writes ENABLED (AJERA_EXPECTED_DB_ID matches)."
    except RuntimeError as e:
        writes = f"Writes BLOCKED: {e}"
    return f"Connected to DatabaseID {a.database_id} (sample data: {a.is_sample_data}). {writes}"


@mcp.tool()
def make_template(output_path: str = "Ajera Project Import - template.xlsx") -> str:
    """Create a blank import workbook whose Lookups sheet and dropdowns list this Ajera's clients,
    departments, project types, rate tables, invoice formats, managers and principals."""
    lk = importer.load_lookups(aj(), want_existing_ids=False)
    out = workbook.build_workbook(Path(output_path), [], [], workbook.lookup_rows_from(lk),
                                  "Ajera Project Import", workbook.today_note() + " from your Ajera database")
    return f"Template written to {out.resolve()} (" + ", ".join(f"{k}: {len(v)}" for k, v in lk.rows.items()) + ")"


@mcp.tool()
def check_workbook(path: str) -> str:
    """Validate an import workbook against Ajera WITHOUT creating anything. Always run this first."""
    projects, problems, _ = importer.check_file(aj(), path)
    if problems:
        return f"{len(problems)} problem(s); nothing would be created:\n" + "\n".join(f"- {p}" for p in problems)
    n_ph = sum(len(p.args["Phases"]) for p in projects)
    return (f"All good: {len(projects)} project(s) and {n_ph} phase(s) are ready to create "
            f"in DatabaseID {aj().database_id}. Nothing was changed.")


@mcp.tool()
def create_projects(path: str, confirm: bool = False) -> str:
    """Create the workbook's projects in Ajera. The Ajera API CANNOT delete projects, so only pass
    confirm=True after the user has seen check_workbook's result and explicitly said to go ahead.
    Nothing is created unless every row passes validation."""
    a = aj()
    a.check_write_target()
    if not confirm:
        return ("Not created. Show the user check_workbook's result, remind them that Ajera's API cannot "
                "delete projects, and call again with confirm=True only if they say yes.")
    outcomes, problems, results = importer.import_file(a, path)
    if problems:
        return f"{len(problems)} problem(s); nothing was created:\n" + "\n".join(f"- {p}" for p in problems)
    lines = [f"- row {o.row} {o.job_id}: {'created' if o.ok else 'NOT created'}"
             + (f", ProjectKey {o.project_key}" if o.project_key else "") + (f" ({o.message})" if not o.ok else "")
             for o in outcomes]
    ok = sum(o.ok for o in outcomes)
    return (f"Created {ok} of {len(outcomes)} project(s).\n" + "\n".join(lines)
            + f"\nResults workbook: {results}\nCreated-projects log: {importer.created_log_path()}")


@mcp.tool()
def list_created_projects() -> str:
    """List projects this tool has created (from the created-projects log), so test projects can be
    found and removed by hand in Ajera - the API cannot delete them."""
    path = importer.created_log_path()
    if not path.exists():
        return "No projects have been created by this tool yet."
    with path.open(encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    return f"{len(rows)} project(s) in {path}:\n" + "\n".join(
        f"- {r['created_at']} DatabaseID {r['database_id']} ProjectKey {r['project_key']} Job {r['job_id']}" for r in rows)


if __name__ == "__main__":
    mcp.run()
