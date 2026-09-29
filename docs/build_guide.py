"""Builds docs/Ajera Project Import - User Guide.pdf (plain-language guide for non-technical readers).

    python docs/build_guide.py
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.graphics.shapes import Drawing, Line, Polygon, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (KeepTogether, ListFlowable, ListItem, PageBreak, Paragraph, Preformatted,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

OUT = Path(__file__).resolve().parent / "Ajera Project Import - User Guide.pdf"
NAVY = colors.HexColor("#1F3864")
BLUE = colors.HexColor("#5B7DB1")
PALE = colors.HexColor("#EEF2F8")
AMBER = colors.HexColor("#FFF4D6")
AMBER_EDGE = colors.HexColor("#C98A00")
GREEN = colors.HexColor("#E6F4EA")
GREEN_EDGE = colors.HexColor("#2E7D32")
GREY = colors.HexColor("#555555")

ss = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss["Heading1"], textColor=NAVY, fontSize=18, spaceBefore=6, spaceAfter=8)
H2 = ParagraphStyle("H2", parent=ss["Heading2"], textColor=NAVY, fontSize=13.5, spaceBefore=12, spaceAfter=5)
BODY = ParagraphStyle("Body", parent=ss["BodyText"], fontSize=10.5, leading=14.5, spaceAfter=6)
SMALL = ParagraphStyle("Small", parent=BODY, fontSize=9, leading=12, textColor=GREY)
CELL = ParagraphStyle("Cell", parent=BODY, fontSize=9, leading=11.5, spaceAfter=0)
CELL_B = ParagraphStyle("CellB", parent=CELL, fontName="Helvetica-Bold")
HEAD = ParagraphStyle("Head", parent=CELL, textColor=colors.white, fontName="Helvetica-Bold")
TITLE = ParagraphStyle("Title", parent=H1, fontSize=30, leading=36, alignment=TA_CENTER, spaceAfter=10)
SUB = ParagraphStyle("Sub", parent=BODY, fontSize=13, leading=18, alignment=TA_CENTER, textColor=GREY)
MONO = ParagraphStyle("Mono", fontName="Courier", fontSize=9, leading=11.5, textColor=colors.HexColor("#E8E8E8"))


def P(text, style=BODY):
    return Paragraph(text, style)


def bullets(items, style=BODY):
    return ListFlowable([ListItem(P(t, style), leftIndent=12) for t in items], bulletType="bullet",
                        start="•", leftIndent=14, bulletFontSize=9)


def steps(items):
    return ListFlowable([ListItem(P(t), leftIndent=16) for t in items], bulletType="1", leftIndent=18,
                        bulletFontName="Helvetica-Bold", bulletColor=NAVY)


def table(rows, widths, header=True):
    data = [[P(c, HEAD if (header and r == 0) else (CELL_B if c_i == 0 else CELL)) for c_i, c in enumerate(row)]
            for r, row in enumerate(rows)]
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C9D3E3")),
             ("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
    for r in range(1 if header else 0, len(rows)):
        if r % 2 == 0:
            style.append(("BACKGROUND", (0, r), (-1, r), PALE))
    t.setStyle(TableStyle(style))
    return t


def callout(title, body, fill=AMBER, edge=AMBER_EDGE):
    inner = [P(f"<b>{title}</b>", ParagraphStyle("ct", parent=BODY, textColor=edge, fontSize=11.5))]
    inner += [P(b) if isinstance(b, str) else b for b in body]
    t = Table([[inner]], colWidths=[6.5 * inch])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), fill), ("BOX", (0, 0), (-1, -1), 1.2, edge),
                           ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                           ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return KeepTogether([Spacer(1, 4), t, Spacer(1, 8)])


def flow_diagram():
    """Excel -> CHECK -> CREATE -> Ajera, with the log underneath."""
    d = Drawing(468, 150)
    boxes = [("Your Excel", "workbook", PALE, NAVY), ("CHECK", "safe: changes nothing", GREEN, GREEN_EDGE),
             ("CREATE", "you type YES", AMBER, AMBER_EDGE), ("Ajera", "projects appear", PALE, NAVY)]
    w, h, gap, y = 96, 58, 28, 78
    for i, (t1, t2, fill, edge) in enumerate(boxes):
        x = i * (w + gap)
        d.add(Rect(x, y, w, h, rx=8, ry=8, fillColor=fill, strokeColor=edge, strokeWidth=1.4))
        d.add(String(x + w / 2, y + 34, t1, fontName="Helvetica-Bold", fontSize=12, fillColor=edge, textAnchor="middle"))
        d.add(String(x + w / 2, y + 17, t2, fontName="Helvetica", fontSize=8.5, fillColor=GREY, textAnchor="middle"))
        if i < len(boxes) - 1:
            ax = x + w + 3
            d.add(Line(ax, y + h / 2, ax + gap - 8, y + h / 2, strokeColor=NAVY, strokeWidth=1.6))
            d.add(Polygon([ax + gap - 8, y + h / 2 + 4, ax + gap - 2, y + h / 2, ax + gap - 8, y + h / 2 - 4],
                          fillColor=NAVY, strokeColor=NAVY))
    # problems loop back from CHECK to Excel
    cx = 1 * (w + gap) + w / 2
    d.add(Line(cx, y + h, cx, y + h + 10, strokeColor=AMBER_EDGE, strokeWidth=1.2))
    d.add(Line(cx, y + h + 10, w / 2, y + h + 10, strokeColor=AMBER_EDGE, strokeWidth=1.2))
    d.add(Line(w / 2, y + h + 10, w / 2, y + h + 2, strokeColor=AMBER_EDGE, strokeWidth=1.2))
    d.add(String((cx + w / 2) / 2, y + h + 13, "problems? fix the row and check again", fontSize=8,
                 fillColor=AMBER_EDGE, textAnchor="middle"))
    # outputs under CREATE
    x = 2 * (w + gap)
    d.add(Line(x + w / 2, y, x + w / 2, 44, strokeColor=GREY, strokeWidth=1, strokeDashArray=[3, 2]))
    d.add(Rect(x - 50, 4, w + 100, 40, rx=6, ry=6, fillColor=colors.white, strokeColor=GREY, strokeWidth=0.8))
    d.add(String(x + w / 2, 28, "results workbook  +  created-projects log", fontSize=8.5, fillColor=GREY, textAnchor="middle"))
    d.add(String(x + w / 2, 15, "(so you always know what was made)", fontSize=8, fillColor=GREY, textAnchor="middle"))
    return d


def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(GREY)
    if doc.page > 1:
        canvas.drawString(inch, 0.55 * inch, "Ajera Project Import - User Guide")
        canvas.drawRightString(7.5 * inch, 0.55 * inch, f"Page {doc.page}")
    canvas.restoreState()


def build():
    s = []
    # -- cover ------------------------------------------------------------------------
    s += [Spacer(1, 1.6 * inch), P("Ajera Project Import", TITLE),
          P("Create projects in Deltek Ajera from an Excel spreadsheet", SUB), Spacer(1, 0.35 * inch),
          flow_diagram(), Spacer(1, 0.4 * inch),
          P("A plain-language guide for the people who set up projects. No programming needed.", SUB),
          Spacer(1, 1.3 * inch),
          P(f"Version 1.0 &nbsp;·&nbsp; {date.today():%B %Y} &nbsp;·&nbsp; github.com/paulnew2000/ajera-project-import",
            ParagraphStyle("c", parent=SMALL, alignment=TA_CENTER)),
          P("Not affiliated with or endorsed by Deltek. Ajera is a trademark of Deltek, Inc.",
            ParagraphStyle("c2", parent=SMALL, alignment=TA_CENTER)),
          PageBreak()]

    # -- 1 what it does -------------------------------------------------------------------
    s += [P("1. What this tool does", H1),
          P("Setting up a new project in Ajera normally means clicking through several screens for each job. "
            "With this tool you <b>list your new projects in a spreadsheet</b>, one row per project, and it creates "
            "them in Ajera for you, correctly and all in one go."),
          P("It always works in two steps:"),
          table([["Step", "What happens", "Does it change Ajera?"],
                 ["CHECK", "Reads your spreadsheet and compares every name (client, department, rate table, "
                           "people...) with what is in Ajera. Lists anything that doesn't match, by row.", "No, never"],
                 ["CREATE", "Runs the check again, then makes the projects. Only runs if <b>every</b> row passed, "
                            "and only after you type YES.", "Yes"]],
                [1.0 * inch, 4.1 * inch, 1.4 * inch]),
          Spacer(1, 6),
          callout("The one thing to remember",
                  ["Ajera's API <b>cannot delete projects</b>. Once CREATE makes a project, only a person working in "
                   "Ajera can remove it. That's why the tool checks everything first, asks you to type YES, and "
                   "keeps a log of every project it makes. See section 7."]),
          P("2. Before you start", H1),
          P("You need three things. The first two are a one-time job for whoever looks after Ajera and your computer."),
          table([["You need", "Who sets it up", "Details"],
                 ["Ajera API access", "Your Ajera administrator",
                  "An <b>API user</b> (name + password) and the <b>API web address</b>. The API user must be allowed "
                  "to create projects. The README lists the exact permissions."],
                 ["Python", "You or IT (5 minutes)",
                  "Free, from python.org. During install, tick <b>Add python.exe to PATH</b>."],
                 ["A practice database", "Your Ajera administrator (recommended)",
                  "An Ajera <b>sandbox</b> or test copy, where you can try things without affecting real projects."]],
                [1.45 * inch, 1.6 * inch, 3.45 * inch]),
          PageBreak()]

    # -- 3 setup ------------------------------------------------------------------------
    menu = ("  ==============================================\n"
            "            AJERA PROJECT IMPORT\n"
            "  ==============================================\n"
            "   1  Check connection  (which database am I on?)\n"
            "   2  Make a blank template from my Ajera\n"
            "   3  CHECK a workbook   (safe - changes nothing)\n"
            "   4  CREATE projects from a workbook\n"
            "   5  Show projects this tool has created\n"
            "   6  Open the settings file (.env)\n"
            "   Q  Quit")
    menu_box = Table([[Preformatted(menu, MONO)]], colWidths=[4.6 * inch])
    menu_box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0C0C0C")),
                                  ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 8),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    s += [P("3. One-time setup", H1),
          steps([
              "Download the tool from GitHub (green <b>Code</b> button, then <b>Download ZIP</b>) and unzip it to a "
              "folder you'll remember, such as <i>Documents\\Ajera Project Import</i>.",
              "Double-click <b>Ajera Project Import.bat</b>. The first time, it installs what it needs (about a minute) "
              "and opens a settings file in Notepad.",
              "In Notepad, paste the <b>API web address</b>, <b>user name</b> and <b>password</b> from your "
              "administrator after the = signs. Save and close Notepad.",
              "Choose <b>1 - Check connection</b>. It shows a <b>DatabaseID</b> number. Confirm with your administrator "
              "that this is the database you want (ideally the sandbox), then choose <b>6</b> and type that number "
              "after <b>AJERA_EXPECTED_DB_ID=</b>. Until this is filled in, the tool refuses to create anything.",
          ]),
          Spacer(1, 4), P("This is the menu you'll see each time:"), menu_box, Spacer(1, 6),
          callout("Keep the settings file private",
                  ["The settings file (<b>.env</b>) holds your Ajera password. Don't email it or put it on a shared "
                   "drive. Everyone who uses the tool should have their own."], fill=PALE, edge=NAVY),
          PageBreak()]

    # -- 4 everyday use --------------------------------------------------------------------
    s += [P("4. Creating projects: the everyday routine", H1),
          steps([
              "<b>Get a template.</b> Menu option <b>2</b> makes <i>Ajera Project Import - template.xlsx</i> in the "
              "tool's folder. Its <b>Lookups</b> sheet lists your Ajera's clients, departments, project types, rate "
              "tables, project managers and principals, and the spreadsheet's drop-down lists use them. Make a fresh "
              "template now and then so new clients appear.",
              "<b>Fill it in.</b> One row per project on the <b>Projects</b> sheet. Add rows on the <b>Phases</b> "
              "sheet if a project needs more than one phase (section 5).",
              "<b>CHECK it.</b> Option <b>3</b>, then pick your file. You'll either see <b>All good</b> or a list "
              "such as <i>Projects row 4, Client: not found in Ajera</i>. Fix those rows, save, and check again.",
              "<b>CREATE.</b> Option <b>4</b>, pick the same file, and type <b>YES</b> when asked. Each project is "
              "created and then read back from Ajera to confirm it's really there.",
              "<b>Look at the results.</b> A copy of your spreadsheet is saved next to it, named "
              "<i>...results_&lt;date-time&gt;.xlsx</i>. It has three new columns: <b>Result</b>, "
              "<b>Ajera Project Key</b> and <b>Message</b>.",
          ]),
          callout("Good habits",
                  [bullets(["Always CHECK before CREATE. It costs nothing and changes nothing.",
                            "Keep each spreadsheet small (a day's or a week's new jobs), so problems are easy to find.",
                            "Once a file has been created, don't run CREATE on it again. The Job IDs now exist, "
                            "so the check will stop you anyway."])],
                  fill=GREEN, edge=GREEN_EDGE),
          PageBreak()]

    # -- 5 columns ---------------------------------------------------------------------------
    s += [P("5. Filling in the spreadsheet", H1),
          P("Dark-blue column headings must be filled in; light-blue ones are optional. Hover over a heading in Excel "
            "for a reminder. Spelling must match Ajera, but capital letters and extra spaces don't matter."),
          P("Projects sheet: one row per project", H2),
          table([["Column", "Required?", "What to enter"],
                 ["Job ID", "Yes", "Your project number. It must not already exist in Ajera."],
                 ["Description", "Yes", "The project name."],
                 ["Client", "Yes", "Pick from the list. The client must already exist in Ajera."],
                 ["Department", "Yes", "Pick from the list."],
                 ["Project Type", "Yes", "Pick from the list."],
                 ["Billing Type", "Yes", "Time and Expense, Fixed Fee, Percent Complete, Unit Price, Nonbillable..."],
                 ["Rate Table", "Yes", "The billing rate table. Ajera insists on one."],
                 ["Project Manager", "No", "Pick from the list (only people set up as project managers)."],
                 ["Principal in Charge", "No", "Pick from the list (only people set up as principals)."],
                 ["Status", "No", "Active (the default) or Preliminary."],
                 ["Start / End Date", "No", "Estimated dates."],
                 ["Invoice Group", "No", "A name for the billing group. Leave blank for <i>Invoice Group 1</i>."],
                 ["Invoice Format", "No", "Pick from the list."],
                 ["Notes", "No", "Anything you like; it's copied into the project's Notes."]],
                [1.45 * inch, 0.85 * inch, 4.2 * inch]),
          P("Phases sheet: optional", H2),
          P("Phases are the pieces of a project that staff charge time to. Add one row per phase, using the same "
            "<b>Job ID</b> as the project. <b>If you add no phases, the project gets one phase called "
            "<i>General</i></b>, so time can be charged straight away."),
          table([["Column", "Required?", "What to enter"],
                 ["Job ID", "Yes", "The project this phase belongs to."],
                 ["Phase Description", "Yes", "e.g. <i>Task 1 - Field Work</i>."],
                 ["Department", "No", "Leave blank to use the project's department."],
                 ["Billing Type", "No", "Leave blank to use the project's billing type."]],
                [1.45 * inch, 0.85 * inch, 4.2 * inch]),
          Spacer(1, 6),
          P("<b>Two people or clients with the same name?</b> The Lookups sheet shows a <b>Key</b> number next to each "
            "name. Type that number instead of the name, and the tool will know exactly which one you mean."),
          PageBreak()]

    # -- 6 messages ---------------------------------------------------------------------------
    s += [P("6. What the messages mean", H1),
          P("Every message tells you the <b>sheet</b>, the <b>row number</b> and the <b>column</b>. It deliberately "
            "doesn't repeat what you typed, so you can safely forward a message to a colleague or IT."),
          table([["Message", "What to do"],
                 ["...not found in Ajera", "Spelling doesn't match Ajera. Choose from the drop-down, or check the "
                                           "Lookups sheet. If it's brand new in Ajera, make a fresh template."],
                 ["...already exists in Ajera", "That Job ID is taken. Use a different number."],
                 ["...is also used on row N", "The same Job ID appears twice in your spreadsheet."],
                 ["...is required", "A dark-blue column was left empty."],
                 ["...matches more than one Ajera record", "Two records share that name. Type the Key number from the "
                                                           "Lookups sheet instead."],
                 ["...not flagged as a Project Manager / Principal", "Ajera only accepts people set up that way. Pick "
                                                                     "someone else, or ask your administrator."],
                 ["...does not match any Job ID on the Projects sheet", "A Phases row points at a project that isn't "
                                                                        "on the Projects sheet (often a typo)."],
                 ["...is not a date / is before the Start Date", "Fix the date."],
                 ["Writes are BLOCKED / AJERA_EXPECTED_DB_ID...", "The safety switch is working: the settings don't "
                                                                  "say you meant to write to this database. See section 3, step 4."],
                 ["Could not sign in to Ajera", "The user name, password or web address in the settings file is wrong, "
                                                "or the API user is disabled."],
                 ["Ajera refused it: ...", "Ajera itself said no during CREATE. The rest of the message says why. "
                                           "Other rows still go ahead, and the results file shows which were made."]],
                [2.55 * inch, 3.95 * inch]),
          PageBreak()]

    # -- 7 cleanup ------------------------------------------------------------------------------
    s += [P("7. Projects can't be deleted by this tool", H1),
          callout("Why this matters",
                  ["Ajera's API, the doorway this tool uses, has no delete button. A project created by mistake "
                   "has to be removed by a person, inside Ajera."]),
          P("To make that easy, the tool <b>writes down every project the moment Ajera creates it</b>, in a file "
            "called <b>created_projects.csv</b> in the tool's folder. Menu option <b>5</b> shows the list. Each "
            "line records:"),
          bullets(["when it was created and in which database (DatabaseID)",
                   "the <b>Ajera Project Key</b> (Ajera's internal number) and the <b>Job ID</b>",
                   "the description and which spreadsheet it came from"]),
          P("How to clean up test projects", H2),
          steps(["Open the log (option 5, or open created_projects.csv in Excel).",
                 "In Ajera, find each project by its Job ID.",
                 "Delete it. Ajera only allows this while nothing has been charged or billed to the project. "
                 "If anything has been, set its status to <b>Inactive</b> instead."]),
          P("<b>Tip:</b> give practice projects an obviously fake Job ID, such as <b>HP001</b>. They're easy to "
            "spot, and they can't clash with a real job number."),
          P("8. Built-in safety features", H1),
          table([["Feature", "What it protects against"],
                 ["CHECK changes nothing", "You can check as often as you like."],
                 ["All-or-nothing check", "If any row has a problem, nothing at all is created, so you never end up "
                                          "with half a batch."],
                 ["Type YES to create", "Creating by accident."],
                 ["Database safety switch", "Writing test data into your live Ajera because the wrong web address "
                                            "was pasted. The tool only writes to the database number in the settings."],
                 ["Read-back check", "Ajera sometimes says \"success\" without actually creating anything. The tool "
                                     "looks for each new project before reporting success."],
                 ["Created-projects log", "Losing track of what was made (section 7)."],
                 ["Messages don't repeat your data", "Client and staff names leaking when you share an error message."]],
                [1.9 * inch, 4.6 * inch]),
          PageBreak()]

    # -- 9 sample ------------------------------------------------------------------------------
    s += [P("9. Try it with the Hogwarts sample", H1),
          P("The <b>examples</b> folder has <b>Hogwarts sample projects.xlsx</b>: ten obviously fictional projects "
            "that show how a finished spreadsheet looks. For example:"),
          table([["Job ID", "Description", "Client", "Billing Type"],
                 ["HP001", "Test Project: Dumbledore", "Gringotts Wizarding Bank", "Time and Expense"],
                 ["HP002", "Test Project: McGonagall", "Ministry of Magic", "Time and Expense"],
                 ["HP003", "Test Project: Hagrid", "Hogwarts School of Witchcraft and Wizardry", "Fixed Fee"]],
                [0.7 * inch, 1.85 * inch, 2.5 * inch, 1.45 * inch]),
          Spacer(1, 6),
          P("Gringotts and Gryffindor aren't in your Ajera, so <b>CHECK will list those names as not found</b>. "
            "That's a good way to see the messages. To create the sample projects in your <b>sandbox</b>, change the "
            "Client, Department, Project Type, Rate Table and Invoice Format columns (and the Department column on "
            "the Phases sheet) to names from your own template's Lookups sheet, then CHECK and CREATE."),
          P("10. Optional: use it by asking an AI assistant", H1),
          P("If you use Claude Desktop (or another assistant that supports <i>MCP</i> tools), IT can connect this tool "
            "to it using the instructions in the README. You can then just ask:"),
          bullets(["<i>\"Which Ajera database is the project import tool connected to?\"</i>",
                   "<i>\"Check the spreadsheet New jobs.xlsx in my Documents folder.\"</i>",
                   "<i>\"Those look right. Create them.\"</i>"]),
          P("The assistant uses the same CHECK and CREATE steps and safety features. It must ask for your go-ahead "
            "before creating anything."),
          P("11. Words you'll see", H1),
          table([["Term", "Meaning"],
                 ["API", "The \"side door\" that lets programs talk to Ajera without clicking through screens."],
                 ["API user", "A special Ajera login, used only by programs like this one."],
                 ["Sandbox", "A practice copy of Ajera. Changes there don't affect your real data."],
                 ["DatabaseID", "The number that identifies which Ajera database (live or sandbox) you're connected to."],
                 ["Project Key", "Ajera's own internal number for a project (different from your Job ID)."],
                 ["Phase", "A part of a project that people charge time to."],
                 ["Invoice group", "How Ajera groups phases for billing. This tool makes one per project, billed to the client."],
                 ["Rate table", "The list of billing rates Ajera uses to price time on the project."]],
                [1.3 * inch, 5.2 * inch])]

    doc = SimpleDocTemplate(str(OUT), pagesize=LETTER, leftMargin=inch, rightMargin=inch,
                            topMargin=0.85 * inch, bottomMargin=0.85 * inch,
                            title="Ajera Project Import - User Guide", author="paulnew2000",
                            subject="Create Ajera projects from Excel")
    doc.build(s, onFirstPage=on_page, onLaterPages=on_page)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
