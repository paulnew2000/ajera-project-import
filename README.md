# Ajera Project Import

**Create projects in Deltek Ajera from an Excel workbook** using Ajera's API. You type one row per project and one row per phase, then run a **check** (which changes nothing) and a **create**.

It comes three ways, all running the same engine:

| How | Who it's for |
|---|---|
| **`Ajera Project Import.bat`** (double-click menu) | Anyone on Windows. No typing commands. |
| **`cli.py`** (command line) | Mac/Linux users, scripting, scheduled jobs |
| **`server.py`** (MCP server) | Use it from Claude Desktop / Claude Code or another AI assistant: *"check projects.xlsx and create them"* |

**Tested end to end** against an Ajera SaaS sandbox (Ajera 10.30, API v1/v2), September 2026. It created 10 projects with 32 phases, Time & Expense and Fixed Fee, with and without Project Manager/Principal. Every phase was open for time entry immediately.

A non-technical walkthrough is in **[docs/Ajera Project Import - User Guide.pdf](docs/)**.

```
 Excel workbook ──► CHECK (reads Ajera, changes nothing) ──► CREATE ──► read back ──► log + results workbook
   Projects            every name looked up in Ajera           v2 CreateProjects    verify it     created_projects.csv
   Phases              every row checked; any problem =        one project at a      exists       <file>_results_<time>.xlsx
   Lookups             NOTHING is created                      time
```

---

## ⚠️ Read this first: the Ajera API cannot delete projects

Once a project is created, only a person in Ajera can remove it. The API has no delete method. So this tool:

1. **Checks every row before creating anything.** If a single row has a problem, nothing is created.
2. **Logs every project the moment Ajera creates it** to `created_projects.csv` (date, database, ProjectKey, Job ID, description, source file). Menu option 5 / `python cli.py log` shows the list.
3. **Refuses to write** unless `AJERA_EXPECTED_DB_ID` in your settings matches the database your URL points at. A production URL pasted by mistake can't receive test data.

**Practise in your Ajera sandbox / test database first.** Give test projects an obviously fake Job ID (the sample uses `HP001`…) so they're easy to find. To clean up, open each logged project in Ajera and delete it. Ajera only allows that while the project has no time, expenses or invoices; otherwise mark it Inactive.

---

## Quick start (Windows)

1. **Install Python 3.10 or newer** from [python.org](https://www.python.org/downloads/) and tick **"Add python.exe to PATH"**.
2. **Download this repository** (green *Code* button → *Download ZIP*) and unzip it somewhere, e.g. `Documents\Ajera Project Import`.
3. **Double-click `Ajera Project Import.bat`.** The first run installs the packages it needs and opens the settings file (`.env`) in Notepad. Fill in:
   - `AJERA_API_URL`: your Ajera API URL (from your Ajera administrator)
   - `AJERA_USERNAME` / `AJERA_PASSWORD`: the Ajera API user
4. Choose **1 - Check connection**. It tells you which **DatabaseID** you're connected to. If that's the database you want to write to, put the number in `.env` as `AJERA_EXPECTED_DB_ID` (menu option 6 opens it).
5. Choose **2 - Make a blank template**. You get a workbook whose **Lookups** sheet and dropdowns list *your* clients, departments, project types, rate tables, invoice formats, project managers and principals.
6. Fill in the **Projects** sheet (and **Phases** if you want more than one phase).
7. Choose **3 - CHECK**, pick the workbook, fix anything it lists, and repeat until it says **All good**.
8. Choose **4 - CREATE**. Type `YES` to confirm. You'll get a `…_results_<date-time>.xlsx` copy showing each row's outcome and the new Ajera ProjectKey.

### What your Ajera administrator needs to set up
- API access turned on, and an **API user** for this tool.
- These API methods allowed for that user: `CreateProjects` (the only write), `ListCompanies`, `ListDepartments`, `ListProjectTypes`, `ListClients`, `ListRateTables`, `ListEmployees`, `GetEmployees`, `ListProjects`, `GetProjects`, plus optionally `ListInvoiceFormats` (needed only if you fill in the *Invoice Format* column).
- Ideally, a **sandbox/test database** to practise in. It has its own URL, and therefore its own DatabaseID.

---

## The workbook

Every workbook has four sheets: **Instructions**, **Projects**, **Phases**, **Lookups**. Dark-blue headings are required and lighter-blue ones are optional. Hover over a heading to see what goes in it.

### Projects (one row per project)
| Column | Required | Notes |
|---|---|---|
| Job ID | ✔ | Your project number. Must not already exist in Ajera. Max 30 characters. |
| Description | ✔ | Project name. |
| Client | ✔ | Exactly as in Ajera. The project gets one invoice group billed to this client. |
| Department | ✔ | |
| Project Type | ✔ | |
| Billing Type | ✔ | Time and Expense, Fixed Fee, Percent Complete, Unit Price, Percent of Construction Cost, Nonbillable, Marketing, Overhead |
| Rate Table | ✔ | Ajera requires a billing rate table on every project. |
| Project Manager | | Must be flagged *Project Manager* on the employee record. |
| Principal in Charge | | Must be flagged *Principal* on the employee record. |
| Status | | Active (default) or Preliminary |
| Start Date / End Date | | Estimated dates. Use Excel dates or YYYY-MM-DD. |
| Invoice Group | | Name for the invoice group (default *Invoice Group 1*). |
| Invoice Format | | |
| Notes | | Copied into the project's Notes. |

### Phases (optional, one row per phase)
| Column | Required | Notes |
|---|---|---|
| Job ID | ✔ | Must match a Job ID on the Projects sheet. |
| Phase Description | ✔ | e.g. *Task 1 - Field Work* |
| Department | | Defaults to the project's department. |
| Billing Type | | Defaults to the project's billing type. |

A project with no phase rows gets one phase called **General**, because Ajera needs at least one phase to charge time to.

**Names are matched without regard to capitals or extra spaces.** If two Ajera records share a name, type the **Key** number from the Lookups sheet instead.

---

## Try the sample

`examples/Hogwarts sample projects.xlsx` holds 10 made-up projects ("Test Project: Dumbledore", …) with 32 phases. Its clients, departments, project types and rate tables are made up too (Gringotts, Gryffindor…), so they don't exist in your Ajera.

**Make it runnable in your Ajera in one step.** Use menu option **7**, or:

```bash
python cli.py adapt-sample --client "Your Test Client"
```

This writes `Hogwarts sample - ready for my Ajera.xlsx` and runs CHECK on it. It reads your Ajera but doesn't change anything. It keeps the Harry Potter descriptions, notes, dates and phases, and fills the rest from your database:

| Column | Filled with |
|---|---|
| Client | The client you name (by name or Key). Use a test/placeholder client. Leave `--client` off and it picks the one client whose name looks like *test/sample/placeholder/new client*, if there's exactly one. |
| Department, Project Type | Your departments and types, spread across the sample's rows. Overhead/marketing-style types are avoided. |
| Rate Table | `--rate-table` if given, otherwise the newest, plainest one with "standard" in its name (e.g. *2026 Standard*) |
| Invoice Format, Project Manager, Principal | Left blank (optional) |
| Job ID | HP001… continuing after any HP numbers already in your Ajera, so it can be run again |

Then CREATE the file (menu option 4), ideally in your sandbox.

**To make realistic test data from your own projects instead:**

```bash
python cli.py export --latest 10 my_projects.xlsx
```

```bash
python examples/hogwartsify.py my_projects.xlsx test_projects.xlsx
```

`export` copies 10 recent projects into the import format. `hogwartsify` replaces their IDs (HP001…), descriptions, notes, dates, managers and phase names with Harry Potter ones. It keeps your real clients, departments and rate tables, so the file passes CHECK. Add `--scrub-references` to replace those too; that's how the shareable sample was made.

---

## Use it from Claude (MCP)

Add this to your MCP client configuration (Claude Desktop: *Settings → Developer → Edit Config*, file `claude_desktop_config.json`) and restart the client:

```json
{
  "mcpServers": {
    "ajera-project-import": {
      "command": "python",
      "args": ["C:\\path\\to\\ajera-project-import\\server.py"]
    }
  }
}
```

The server reads the same `.env` file. Its tools:

| Tool | What it does |
|---|---|
| `check_connection` | Which database am I on, and are writes allowed? |
| `make_template` | Blank workbook with your Lookups + dropdowns |
| `adapt_sample` | Make the Hogwarts sample runnable in your Ajera, then CHECK it |
| `check_workbook` | Validate a workbook. Changes nothing. |
| `create_projects` | Create the projects. Does nothing unless called with `confirm=true`, which the assistant should only pass after you say yes. |
| `list_created_projects` | The created-projects log, for clean-up |

Then just ask: *"Check C:\Projects\new jobs.xlsx against Ajera"* → *"OK, create them."*

---

## Command line

```bash
python cli.py connect
```

```bash
python cli.py template "my template.xlsx"
```

```bash
python cli.py export --ids 1001,1002 out.xlsx
```

```bash
python cli.py adapt-sample --client "Your Test Client"
```

```bash
python cli.py check projects.xlsx
```

```bash
python cli.py create projects.xlsx
```

```bash
python cli.py log
```

`adapt-sample` also takes `--rate-table` and `--prefix`. `export` also takes `--latest N`, `--skip N` and `--id-pattern REGEX`. `create --yes` skips the typed confirmation (for scripts).

Tests (offline, no Ajera needed):

```bash
python -m pytest -q
```

---

## Privacy by design

Problem messages point at **sheet, row and column** and never repeat what's in the cell (e.g. *"Projects row 4, Client: not found in Ajera"*). The command line and MCP tools report counts, Job IDs and ProjectKeys, not client or employee names. So you can paste output into a help ticket, or let an AI assistant run the tool, without exposing client data. Your workbook, `.env`, the created-projects log and results workbooks are all excluded from git by `.gitignore`.

---

## Ajera API behaviour worth knowing

Learned the hard way. The code handles all of these:

- **`CreateProjects` exists only in API version 2.** Most reads (e.g. `GetEmployees`) are used on version 1, so the client keeps a session for each version.
- **A wrongly shaped `CreateProjects` request returns "success" and creates nothing.** Ajera echoes back an example skeleton. The tool only counts a create when Ajera returns a `ProjectKey`, then reads the project back to confirm it.
- **The request shape that works:** `{"CreateType": "Project", "Project": {...}, "InvoiceGroups": [{"Description", "ClientKey", "InvoiceFormatKey"}], "Phases": [{...}]}`. The client belongs to the invoice group, not the project. `RateTableKey` is required.
- **`GetProjects` returns the client nested** (`InvoiceGroups[].Client.ClientKey`), but `CreateProjects` takes a flat `ClientKey`.
- **Project Manager / Principal** must be employees flagged as such in Ajera, otherwise the create is rejected.
- **`ListProjects` with no status filter returns every project.** A status list silently narrows the results to Active, so duplicate-ID checks must use the unfiltered list.
- **Custom fields are ignored by `CreateProjects`.** Setting them takes a follow-up `UpdateProjects` call (full record plus the unchanged original). This tool doesn't set custom fields.
- The API URL's query string is base64 JSON: `{"ClientID", "DatabaseID", "IsSampleData"}`. That's how the tool knows which database it's pointed at.

## Limitations
- One invoice group (one client) per project. Phases are one level deep (no sub-phases).
- Uses the first company in multi-company Ajera setups.
- Doesn't set custom fields, contract amounts, budgets, contacts or team members. Add those in Ajera afterwards.
- Doesn't create clients or employees. They must already exist.

## Files
| File | What |
|---|---|
| `Ajera Project Import.bat` | Windows double-click menu |
| `cli.py` | Command line |
| `server.py` | MCP server |
| `importer.py` | The engine: read workbook → look up → validate → create → verify → log |
| `workbook.py` | Builds the formatted template/export workbooks |
| `adapt.py` | Fits the Hogwarts sample to your Ajera (`adapt-sample`) |
| `ajera_client.py` | Small Ajera API client (sessions v1/v2, write-target check) |
| `examples/` | Sample workbook and `hogwartsify.py` |
| `docs/` | The non-technical user guide (PDF) and the script that builds it |
| `.env.example` | Settings template: copy to `.env` |

## License
MIT. Not affiliated with or endorsed by Deltek. *Ajera* is a trademark of Deltek, Inc. Harry Potter names are used only as obviously fictional test data.
