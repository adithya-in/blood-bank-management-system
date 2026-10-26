from pathlib import Path
import re
import subprocess

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

BASE = Path(__file__).resolve().parent
ASSETS = BASE / "assets"
DOCX = BASE / "Blood_Bank_PBCMT404_Report.docx"
PDF = BASE / "Blood_Bank_PBCMT404_Report.pdf"

BODY_FONT = "Times New Roman"
CODE_FONT = "Courier New"
BLUE = "366091"
DEEP_BLUE = "17365D"
BLACK = "000000"

doc = Document()
doc.core_properties.title = "Online Blood Bank and Donor Management System - PBCMT404 Report"
doc.core_properties.subject = "Database Systems Mini Project Report"
doc.core_properties.author = "ADITHYA"
doc.core_properties.keywords = "blood bank, database systems, Flask, SQLite, PBCMT404"


def page_setup(section):
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.25)
    section.footer_distance = Inches(0.25)


page_setup(doc.sections[0])


def font(run, size=12, bold=False, italic=False, color=BLACK, name=BODY_FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for key in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(key), name)


def para(text="", size=12, bold=False, italic=False, color=BLACK, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         before=0, after=0, line=1.15, name=BODY_FONT):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.widow_control = True
    if text:
        font(p.add_run(text), size=size, bold=bold, italic=italic, color=color, name=name)
    return p


def spacer(pts=18):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(pts)
    p.paragraph_format.line_spacing = 1.0
    return p


def chapter(number, title):
    spacer(10)
    para(f"CHAPTER - {number}", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         before=0, after=0, line=1.0)
    para(title.upper(), size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         before=0, after=16, line=1.0)


def section_heading(text, page_break=False):
    p = para(text, size=14, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
             before=10, after=6, line=1.0, name=BODY_FONT)
    p.paragraph_format.page_break_before = page_break
    return p


def body(text):
    return para(text, after=0)


def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.left_indent = Inches(0.3)
    pf.first_line_indent = Inches(-0.18)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15
    font(p.add_run(text), size=12)
    return p


def labeled(label, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_before = Pt(4)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.15
    font(p.add_run(label), size=12, bold=True)
    font(p.add_run(text), size=12)
    return p


def code(lines):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.left_indent = Inches(0.22)
    pf.space_before = Pt(4)
    pf.space_after = Pt(9)
    pf.line_spacing = 1.05
    run = p.add_run("\n".join(lines))
    font(run, size=9.2, name=CODE_FONT)
    return p


def code_text(text):
    code(text.splitlines())


def figure(filename, caption, width=6.2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(ASSETS / filename), width=Inches(width))
    para(caption, size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         before=0, after=8, line=1.0)


def no_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "nil")
        borders.append(e)
    tbl_pr.append(borders)


def cell_margins(cell, top=0, left=0, bottom=0, right=0):
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(value))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tc_pr.append(mar)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)
    font(run, size=11)


def restart_numbering(section):
    sect_pr = section._sectPr
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:start"), "1")


def page_break():
    doc.add_page_break()


# ---------------------------------------------------------------------------
# Cover page
# ---------------------------------------------------------------------------

para("PBCMT404 : DATABASE SYSTEMS", size=20, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0, line=1.0)
para("MINI PROJECT REPORT", size=20, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0, line=1.0)
para("“Online Blood Bank and Donor Management System”", size=20, bold=True, color=DEEP_BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=14, after=0, line=1.0)
para("Using Python and SQLite", size=20, bold=True, color=DEEP_BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=14, after=0, line=1.0)
para("(2025-2026 ACADEMIC YEAR)", size=20, bold=True, color=DEEP_BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=14, after=0, line=1.0)
para("B.Tech Semester IV", size=20, bold=True, color=DEEP_BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=14, after=0, line=1.0)
spacer(50)
para("SREE  BUDDHA  COLLEGE  OF  ENGINEERING", size=18, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0, line=1.0)
para("(AUTONOMOUS)", size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     before=0, after=10, line=1.0)
para("PATTOOR P.O., NOORANAD , ALAPPUZHA , KERALA", size=14, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=18, line=1.0)
para("DEPARTMENT  OF", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     before=0, after=1, line=1.0)
para("COMPUTER SCIENCE AND ENGINEERING (ARTIFICIAL", size=16.5, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0, line=1.0)
para("INTELLIGENCE AND MACHINE LEARNING)", size=16.5, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=24, line=1.0)
para("PREPARED BY", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     before=0, after=10, line=1.0)
names_table = doc.add_table(rows=1, cols=2)
names_table.alignment = WD_TABLE_ALIGNMENT.CENTER
names_table.autofit = False
no_borders(names_table)
row = names_table.rows[0]
row.height = Inches(0.44)
row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
for cell in row.cells:
    cell_margins(cell)
row.cells[0].width = Inches(3.8)
row.cells[1].width = Inches(2.7)
p = row.cells[0].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_after = Pt(0)
font(p.add_run("ADITHYA"), size=17)
p = row.cells[1].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_after = Pt(0)
font(p.add_run("[Roll No.]"), size=17)

# ---------------------------------------------------------------------------
# Contents page
# ---------------------------------------------------------------------------

contents_section = doc.add_section(WD_SECTION.NEW_PAGE)
page_setup(contents_section)
contents_section.footer.is_linked_to_previous = False

para("CONTENTS", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     before=0, after=30, line=1.0)

toc = doc.add_table(rows=6, cols=3)
toc.alignment = WD_TABLE_ALIGNMENT.CENTER
toc.autofit = False
no_borders(toc)
headers = ("CHAPTERS", "CONTENTS", "PAGE No")
toc_rows = [
    ("Chapter 1", "Introduction", "[[C1]]"),
    ("Chapter 2", "Methodology", "[[C2]]"),
    ("Chapter 3", "Implementation", "[[C3]]"),
    ("Chapter 4", "Results", "[[C4]]"),
    ("Chapter 5", "Conclusion", "[[C5]]"),
]
for i, row in enumerate(toc.rows):
    row.height = Inches(0.62)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    values = headers if i == 0 else toc_rows[i - 1]
    for j, cell in enumerate(row.cells):
        cell.width = [Inches(2.15), Inches(2.85), Inches(1.3)][j]
        cell_margins(cell, left=0, right=0)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j == 2 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        font(p.add_run(values[j]), size=14 if i == 0 else 12, bold=True)

# ---------------------------------------------------------------------------
# Body section
# ---------------------------------------------------------------------------

body_section = doc.add_section(WD_SECTION.NEW_PAGE)
page_setup(body_section)
body_section.footer.is_linked_to_previous = False
restart_numbering(body_section)
add_page_number(body_section.footer.paragraphs[0])

# =========================== CHAPTER 1 ====================================
chapter(1, "Introduction")
body("The Online Blood Bank and Donor Management System (OBBDMS) is a database-driven web application designed to manage donor records, hospital blood stock, blood requests, donation records and donation appointments. In a traditional blood bank, these records are often maintained in paper registers, spreadsheets and disconnected hospital systems. Staff must reconcile donor details, available units, pending requests and donation history before they can answer a simple question: does a hospital currently hold usable units of a given blood group?")
body("That manual approach creates duplicated records, stale stock counts and requests whose status is known only to the person who received the phone call. Donors likewise have no reliable way to check when they last donated or when they become eligible to donate again. The project addresses these problems by centralising the core data in one relational database and giving each role a focused interface for the work it needs to perform.")
body("The system is implemented using Python 3.10, Flask 3.1 and SQLite. Flask provides the application layer and route-level access control, Jinja templates and a custom CSS interface provide the presentation layer, and SQLite stores the relational data. Two roles are supported. Donors register, sign in, check their 90-day eligibility, book a donation slot, search for blood and raise requests. Administrators manage hospitals, inventory, requests, donations and donor records from a dispatch-board interface.")
body("From a database-systems perspective, the project applies relational schema design, primary and foreign keys, uniqueness constraints, check constraints, indexes and parameterized SQL. The schema contains six related tables: users, hospitals, inventory, requests, donations and appointments. Inventory is unique per hospital and blood group, requests and donations reference both a user and a hospital, and appointments reference a donor and a hospital. These constraints keep the data consistent when records are inserted or updated through the application.")
body("The primary objectives of the project are to centralise donor and stock data, provide public availability search, compute donor eligibility from donation history, enforce a controlled request lifecycle, protect state-changing operations with authentication and CSRF tokens, and render the same relational data through a usable desktop and mobile interface. The system is deliberately scoped as an academic prototype with synthetic data; it does not verify clinical compatibility, screen blood or replace hospital procedures.")
body("The report follows the PBCMT404 mini-project structure. Chapter 2 explains the requirement analysis, system design, architecture and tools. Chapter 3 presents the database schema and the implementation modules. Chapter 4 shows the working interface and the results of the exercised workflows. Chapter 5 concludes the report and identifies the limitations of the prototype.")

# =========================== CHAPTER 2 ====================================
page_break()
chapter(2, "Methodology")
section_heading("2.1 Requirement Analysis")
body("The requirements were identified from the daily work of a small blood bank. The following functional requirements shaped the design:")
bullet("Public availability search by blood group and city without requiring a login.")
bullet("Donor registration and secure login with email and password.")
bullet("Donor dashboard with last donation, 90-day eligibility, donation history, appointment booking and personal request history.")
bullet("Administrator dashboard with stock strips, critical groups, pending requests, recent donations and hospital totals.")
bullet("Hospital and inventory management with one stock line per hospital and blood group.")
bullet("Request lifecycle: Pending, Approved, Completed and Cancelled, with stock deducted only on valid completion.")
bullet("Donation logging, hospital management and donor record search.")
bullet("Role-based access control, CSRF protection and server-side validation on every write.")

section_heading("2.2 System Design")
body("The design follows an entity-relationship model. The entities and relationships are:")
bullet("users: donors and administrators, with a unique email and a role constrained to donor or admin.")
bullet("hospitals: blood-bank sites with name, city, phone and address.")
bullet("inventory: one row per hospital and blood group, with a non-negative unit count and a UNIQUE pair (hospital_id, blood_group).")
bullet("requests: raised by a user for a hospital, blood group, unit count, urgency and status.")
bullet("donations: completed donations by a donor at a hospital, used for history and eligibility.")
bullet("appointments: booked donation slots for a donor at a hospital, with date, time and status.")
body("A user can raise many requests, make many donations and book many appointments. A hospital can hold many inventory lines, receive many requests and record many donations. Each of these one-to-many relationships is enforced with a foreign key, and SQLite foreign-key enforcement is enabled on every connection.")

figure("er.png", "fig(2.1)  Entity–Relationship Diagram", 6.05)

section_heading("2.3 Architecture")
body("The application uses a three-tier architecture:")
bullet("Presentation tier: server-rendered Jinja templates, a custom CSS design system and a small amount of JavaScript for status updates, counters and motion.")
bullet("Application tier: Flask routes, authentication decorators, CSRF validation, request validation and parameterized database queries.")
bullet("Database tier: SQLite with six related tables, foreign-key constraints, check constraints, unique constraints and search indexes.")
body("The browser sends an HTTP request to Flask. The application verifies the session, role and CSRF token where required, runs a parameterized query and renders a template with the result. The same route structure powers the desktop and mobile layouts; the responsive behaviour is implemented in CSS rather than a separate mobile application.")

figure("architecture.png", "fig(2.2)  Three-Tier Architecture", 6.25)

section_heading("2.4 Tools and Technologies Used")
bullet("Programming language: Python 3.10.")
bullet("Web framework: Flask 3.1.3.")
bullet("Database: SQLite 3, with a MySQL-portable relational schema.")
bullet("Templates: Jinja2.")
bullet("Frontend: HTML5, custom CSS and vanilla JavaScript.")
bullet("Security: Werkzeug password hashing, signed sessions and CSRF tokens.")
bullet("Testing: Flask test client and Google Chrome through Playwright.")
bullet("Diagrams: Matplotlib; screenshots: Playwright at 1440×900 and 390×844.")

# =========================== CHAPTER 3 ====================================
page_break()
chapter(3, "Implementation")
section_heading("3.1 Database Implementation")
body("The database is created automatically on first run. The schema enforces primary keys, foreign keys, check constraints and a unique constraint on inventory. The statements below define the six tables and their indexes.")
code_text("""CREATE TABLE users (
  user_id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  email TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('donor','admin')),
  phone TEXT, city TEXT, blood_group TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE hospitals (
  hospital_id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  city TEXT NOT NULL,
  phone TEXT, address TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE inventory (
  inventory_id INTEGER PRIMARY KEY AUTOINCREMENT,
  hospital_id INTEGER NOT NULL,
  blood_group TEXT NOT NULL,
  units INTEGER NOT NULL DEFAULT 0 CHECK (units >= 0),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE (hospital_id, blood_group),
  FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id)
);

CREATE TABLE requests (
  request_id INTEGER PRIMARY KEY AUTOINCREMENT,
  requester_id INTEGER NOT NULL,
  hospital_id INTEGER NOT NULL,
  blood_group TEXT NOT NULL,
  units INTEGER NOT NULL CHECK (units > 0),
  urgency TEXT NOT NULL DEFAULT 'Routine',
  note TEXT,
  status TEXT NOT NULL DEFAULT 'Pending'
    CHECK (status IN ('Pending','Approved','Completed','Cancelled')),
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (requester_id) REFERENCES users(user_id),
  FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id)
);""")
code_text("""CREATE TABLE donations (
  donation_id INTEGER PRIMARY KEY AUTOINCREMENT,
  donor_id INTEGER NOT NULL,
  hospital_id INTEGER NOT NULL,
  donation_date TEXT NOT NULL,
  units INTEGER NOT NULL DEFAULT 1 CHECK (units > 0),
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (donor_id) REFERENCES users(user_id),
  FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id)
);

CREATE TABLE appointments (
  appointment_id INTEGER PRIMARY KEY AUTOINCREMENT,
  donor_id INTEGER NOT NULL,
  hospital_id INTEGER NOT NULL,
  slot_date TEXT NOT NULL,
  slot_time TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'Booked'
    CHECK (status IN ('Booked','Kept','Missed','Cancelled')),
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (donor_id) REFERENCES users(user_id),
  FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id)
);

CREATE INDEX idx_inventory_group ON inventory(blood_group);
CREATE INDEX idx_requests_status ON requests(status);
CREATE INDEX idx_donations_donor ON donations(donor_id);""")
body("The schema holds one fact per table: a hospital holds stock, a user raises a request, a donor makes a donation and a donor books an appointment. Hospital names are never duplicated inside inventory, and donor details are never duplicated inside requests or donations.")

section_heading("3.2 Database Connection and Query Helpers")
body("A SQLite connection is opened per Flask application context. Foreign-key enforcement is switched on explicitly because SQLite does not enable it by default. A teardown hook closes the connection at the end of the request.")
code_text("""def get_db():
    if "db" not in g:
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db

@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()

def query(sql, args=(), one=False):
    cur = get_db().execute(sql, args)
    rows = cur.fetchall()
    cur.close()
    return (rows[0] if rows else None) if one else rows

def execute(sql, args=()):
    db = get_db()
    cur = db.execute(sql, args)
    db.commit()
    last_id = cur.lastrowid
    cur.close()
    return last_id""")

section_heading("3.3 Authentication, Roles and CSRF")
body("Registration validates the email, password length, blood group and required fields, hashes the password with Werkzeug and inserts the user with the donor role. Login verifies the hash and stores the user id and role in the signed session. The admin_required decorator returns HTTP 403 when a signed-in donor attempts to open an administrator route.")
body("Every POST is guarded by a before_request hook. The token is generated per session, embedded in forms and in a meta tag for fetch calls, and compared with secrets.compare_digest. A missing or stale token is rejected before the route can touch data.")
code_text("""@app.before_request
def csrf_protect():
    if request.method == "POST":
        sent = (request.form.get("csrf_token")
                or request.headers.get("X-CSRF-Token"))
        if not sent or not secrets.compare_digest(sent, session.get("csrf", "")):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Invalid CSRF token."}), 400
            abort(400, description="Invalid CSRF token.")

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please sign in to continue.", "error")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped

def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped""")

section_heading("3.4 Request Lifecycle and Stock Consistency")
body("Status changes are not accepted as free text. The API defines the permitted next state for each current state and returns HTTP 409 for an invalid transition. Completion reads the inventory row for the request's hospital and blood group, refuses the transition if there are not enough units, and only then decrements the inventory and marks the request Completed.")
code_text("""allowed = {
    "Pending":   {"Approved", "Cancelled"},
    "Approved":  {"Completed", "Cancelled"},
    "Completed": set(),
    "Cancelled": set(),
}

if new_status not in allowed.get(row["status"], set()):
    return jsonify({"error": "Invalid transition."}), 409

if new_status == "Completed":
    inv = query(
        "SELECT * FROM inventory WHERE hospital_id = ? AND blood_group = ?",
        (row["hospital_id"], row["blood_group"]), one=True)
    if not inv or inv["units"] < row["units"]:
        return jsonify({"error": "Not enough units in stock."}), 409
    execute("UPDATE inventory SET units = units - ? "
            "WHERE inventory_id = ?",
            (row["units"], inv["inventory_id"]))

execute("UPDATE requests SET status = ? WHERE request_id = ?",
        (new_status, request_id))""")

section_heading("3.5 Availability Search Implementation")
body("Search is built from optional predicates rather than separate queries for every combination. A blood group uses exact equality, while a city uses a parameterized LIKE. The query joins inventory to hospitals so the result can show the site name, city and phone number without duplicating those fields inside the inventory table.")
code_text("""sql = \"\"\"
    SELECT i.blood_group, i.units, i.updated_at,
           h.name, h.city, h.phone
    FROM inventory i
    JOIN hospitals h ON h.hospital_id = i.hospital_id
    WHERE 1=1
\"\"\"
if blood:
    sql += " AND i.blood_group = ?"
    args.append(blood)
if city:
    sql += " AND h.city LIKE ?"
    args.append(f"%{city}%")
sql += " AND i.units > 0 ORDER BY i.units DESC, h.name"
rows = query(sql, args)""")
body("Values are always bound as parameters. An injection-style city string is treated as literal data and returns zero rows rather than altering the query structure.")

section_heading("3.6 Donor Module")
body("The donor dashboard combines eligibility, history, appointments and requests on one page. Eligibility is computed from the most recent donation plus a 90-day cooldown, and the same page exposes appointment booking and personal request submission.")
code_text("""def donor_stats(user_id):
    last = query(
        "SELECT donation_date FROM donations WHERE donor_id = ? "
        "ORDER BY donation_date DESC LIMIT 1",
        (user_id,), one=True)
    total = query(
        "SELECT COUNT(*) AS c, COALESCE(SUM(units),0) AS u "
        "FROM donations WHERE donor_id = ?",
        (user_id,), one=True)
    last_date = next_date = None
    days_left = 0
    if last:
        last_date = datetime.strptime(last["donation_date"], "%Y-%m-%d").date()
        next_date = last_date + timedelta(days=90)
        days_left = max(0, (next_date - date.today()).days)
    return {"last_date": last_date, "next_date": next_date,
            "days_left": days_left, "eligible": days_left == 0,
            "count": total["c"], "units": total["u"]}")

@app.route("/book", methods=["POST"])
@login_required
def book_appointment():
    hospital_id = request.form.get("hospital_id", type=int)
    slot_date = request.form.get("slot_date", "")
    slot_time = request.form.get("slot_time", "")
    chosen = datetime.strptime(slot_date, "%Y-%m-%d").date()
    if chosen < date.today():
        flash("Pick today or a future date.", "error")
        return redirect(url_for("donor_dashboard"))
    execute("INSERT INTO appointments (donor_id,hospital_id,slot_date,slot_time) "
            "VALUES (?,?,?,?)",
            (session["user_id"], hospital_id, slot_date, slot_time))""")

section_heading("3.7 Administrator Module")
body("The administrator board summarises stock by blood group, marks critical groups, and lists pending requests, recent donations and hospital totals. Each pending-request slip joins users and hospitals so it can show the requester, reference, blood group, unit count, urgency and hospital without storing duplicate names.")
code_text("""slips = query(\"\"\"
    SELECT r.*, u.name AS requester, h.name AS hospital, h.city
    FROM requests r
    JOIN users u ON u.user_id = r.requester_id
    JOIN hospitals h ON h.hospital_id = r.hospital_id
    WHERE r.status = 'Pending'
    ORDER BY CASE r.urgency WHEN 'Critical' THEN 0 ELSE 1 END,
             r.created_at
    LIMIT 4
\"\"\")""")
body("Donation logging is a separate administrator action. The route validates the donor, hospital, date and unit count, inserts the donation, reads the donor's blood group, and updates the matching inventory row. If the hospital does not yet hold that group, a new inventory line is created.")
code_text("""@app.route("/admin/donations", methods=["GET", "POST"])
@admin_required
def admin_donations():
    if request.method == "POST":
        donor_id = request.form.get("donor_id", type=int)
        hospital_id = request.form.get("hospital_id", type=int)
        units = request.form.get("units", type=int) or 1
        execute("INSERT INTO donations (donor_id,hospital_id,donation_date,units,note) "
                "VALUES (?,?,?,?,?)",
                (donor_id, hospital_id, donation_date, units, note))
        group = query("SELECT blood_group FROM users WHERE user_id = ?",
                      (donor_id,), one=True)["blood_group"]
        execute("UPDATE inventory SET units = units + ? "
                "WHERE hospital_id = ? AND blood_group = ?",
                (units, hospital_id, group))
    return render_template("admin_donations.html", ...)""")

section_heading("3.8 Interface Implementation")
body("The interface uses a hospital dispatch-desk metaphor rather than a generic admin template. Cream paper, an oxidized-red rail, ink typography, amber signals and teal approvals form a clear operational surface. Stock is rendered as paper strips with split-flap counters; pending requests appear as paper slips; approval presses a rubber stamp before the slip is filed.")
body("The layout is responsive. At mobile width the fixed rail becomes a compact header with a horizontally scrollable navigation strip, the board collapses to one column, stock strips become two-up and action buttons grow to a 40-pixel minimum height. Wide tables scroll inside their own container. A reduced-motion media query collapses animation and reveals all content immediately, so no state depends on motion.")

# =========================== CHAPTER 4 ====================================
page_break()
chapter(4, "Results")
body("The Online Blood Bank and Donor Management System was implemented using Python, Flask, SQLite and a responsive HTML interface. The following screens were captured from the running application with the seeded synthetic demo database. Each screen is followed by the implemented features and the observed result.")

section_heading("4.1 Login Interface")
figure("login.jpg", "Fig. 4.1  Login Page", 5.35)
body("This is the entry screen for donors and administrators. The server verifies the password hash, creates a signed session, and redirects the user according to their role.")
labeled("Features: ", "email and password inputs, CSRF-protected form, password hashing, role-based redirect and clear error messages.")
labeled("Result: ", "A valid administrator reaches /admin, a valid donor reaches /dashboard, and a missing CSRF token returns HTTP 400.")

section_heading("4.2 Public Landing Page", page_break=True)
figure("landing.jpg", "Fig. 4.2  Public Landing Page", 6.2)
body("The landing page states the problem and gives visitors two immediate actions: search for availability or register as a donor. Group-level stock totals are visible without signing in.")
labeled("Features: ", "public availability search, donor registration, live stock summary, low-stock signalling and a sign-in route.")
labeled("Result: ", "A visitor can check whether blood is available in a city before travelling, while donor and administrator actions remain protected.")

section_heading("4.3 Availability Search", page_break=True)
figure("search.jpg", "Fig. 4.3  Availability Search for O+ in Pune", 6.2)
body("The search page returns hospital rows for the selected blood group and city. Each row includes the hospital name, city, phone number and available units.")
labeled("Features: ", "blood-group filter, city filter, joined hospital data, positive-stock filter and parameterized SQL.")
labeled("Result: ", "The O+ search in Pune returned City Hospital with its contact number. An injection-style city string returned zero rows and did not expose or alter data.")

section_heading("4.4 Donor Dashboard", page_break=True)
figure("donor.jpg", "Fig. 4.4  Donor Dashboard", 5.8)
body("The donor dashboard answers the questions a donor actually asks: when did I last donate, when can I donate again, what have I given and what is happening to my request?")
labeled("Features: ", "90-day eligibility strip, donation history, appointment booking, personal request form, request history and appointment cancellation.")
labeled("Result: ", "A donor can see the next eligible date, book a future slot and review only their own records. Ownership checks prevent cancellation of another donor's appointment.")

section_heading("4.5 Administrator Dashboard", page_break=True)
figure("board.jpg", "Fig. 4.5  Administrator Dispatch Board", 6.2)
body("The administrator board is the first viewport after login. Group stock is arranged as split-flap strips, critical groups are marked in red, and pending requests appear as paper slips.")
labeled("Features: ", "stock strips, critical-group signalling, pending request rail, recent donations, hospital stock totals and top-level statistics.")
labeled("Result: ", "An administrator can understand stock and pending work at a glance and move directly to the request queue or inventory page.")

section_heading("4.6 Request Management and Status Actions", page_break=True)
figure("requests.jpg", "Fig. 4.6  Request Queue with Status Actions", 6.15)
body("The request page exposes the full queue with status filters. Pending requests offer Approve and Decline; approved requests offer Complete and Cancel. Completing a request is the only action that changes stock.")
labeled("Features: ", "status filters, permitted actions by state, server-side transition checks, stock validation, confirmation toast and retained terminal rows.")
labeled("Result: ", "Approval returned HTTP 200, repeat completion returned HTTP 409, and completing a valid request reduced A+ stock from 12 to 11 units.")

section_heading("4.7 Mobile Responsive Interface", page_break=True)
figure("mobile-pair.jpg", "Fig. 4.7  Mobile Layout at 390×844", 4.9)
body("The mobile layout keeps the same information architecture rather than hiding core actions. The rail becomes a compact header with a scrollable tab strip, the board collapses to a single column and stock strips become two-up.")
labeled("Features: ", "responsive navigation, one-column board, internally scrolling tables, 40-pixel action targets and a reduced-motion path.")
labeled("Result: ", "Browser checks found no page-level horizontal overflow at 390 pixels, sampled text met WCAG AA contrast, and reduced motion revealed content immediately.")

section_heading("4.8 Overall System Result", page_break=True)
body("All planned modules were implemented and exercised: registration, login, public search, donor dashboard, administrator board, hospital and inventory management, donation logging, request processing and appointments.")
bullet("The six-table relational schema is created automatically and enforces primary keys, foreign keys, uniqueness and check constraints.")
bullet("Fourteen functional and security checks were executed, including missing CSRF, role denial, invalid state transitions and injection-style search.")
bullet("Browser checks covered desktop and mobile overflow, sampled contrast, reduced motion and loaded assets.")
bullet("Stock changed only after a valid completion, and database row counts changed only after a successful request or appointment insert.")
body("The prototype is not an exhaustive penetration test or a clinical validation. It is a working demonstration that database constraints, server-side validation and interface state can be kept consistent through one application.")

# =========================== CHAPTER 5 ====================================
page_break()
chapter(5, "Conclusion")
body("The Online Blood Bank and Donor Management System demonstrates how a relational database can support a time-sensitive healthcare workflow. The six-table schema captures users, hospitals, inventory, requests, donations and appointments through explicit primary keys, foreign keys, uniqueness rules and check constraints. Each role receives a focused surface: donors find blood, check eligibility and manage their contribution history; administrators see the shelf, clear the request rail and record donations.")
body("From a DBMS perspective, the project applies relational schema design, normalisation, foreign-key integrity, unique and check constraints, indexes and parameterized SQL. The application commits each write, performs stock validation before deducting units, and enforces every write through server-side validation. Authentication, role guards and CSRF protection show that database quality must be paired with application-level security; a correct schema alone does not prevent an unauthorised action.")
body("The interface contributes to the result because it exposes the same relational facts that the schema encodes. Stock appears as ordered units rather than free text, requests move through an explicit state machine rather than undefined statuses, and terminal rows are retained as filed records. The mobile layout keeps the same actions available on a phone, and the reduced-motion path keeps the content usable without animation.")
body("The prototype has clear limits. It uses SQLite, which is adequate for a single-node demonstration but not for concurrent multi-branch operation. It does not verify clinical compatibility, screen donated blood, confirm donor identity or replace hospital procedures. It also lacks email notifications, audit logging, rate limiting and production backup. Future work can migrate the schema to MySQL or PostgreSQL, add versioned migrations, audit trails, notifications, pagination and an automated CI test pipeline.")
body("Overall, the project meets its academic objective: it turns a database exercise into a coherent working application and shows how schema constraints, server-side validation, role-based access and interface craft combine to produce trustworthy behaviour. With the migration and operational work described above, the same structure could serve as the foundation for a larger blood-bank information system.")

doc.save(DOCX)

subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(BASE), str(DOCX)],
    check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
)

from pypdf import PdfReader

reader = PdfReader(PDF)
body_pages = len(reader.pages) - 2
starts = {}
for i, page in enumerate(reader.pages):
    if i < 2:
        continue
    text = page.extract_text() or ""
    for number in range(1, 6):
        if number not in starts and re.search(rf"CHAPTER\s*[-–—]\s*{number}\b", text):
            starts[number] = i - 1
ranges = {}
for number in range(1, 6):
    if number not in starts:
        ranges[number] = "—"
        continue
    start = starts[number]
    end = starts.get(number + 1, body_pages + 1) - 1
    ranges[number] = str(start) if start == end else f"{start}-{end}"

d = Document(DOCX)
for table in d.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for number in range(1, 6):
                    token = f"[[C{number}]]"
                    if token in p.text:
                        for run in p.runs:
                            if token in run.text:
                                run.text = run.text.replace(token, ranges[number])
d.save(DOCX)

subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(BASE), str(DOCX)],
    check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
)

print("docx", DOCX, DOCX.stat().st_size)
print("pdf", PDF, PDF.stat().st_size)
print("pages", len(PdfReader(PDF).pages))
print("chapter starts", starts)
print("toc ranges", ranges)
