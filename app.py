"""
Online Blood Bank & Donor Management System
Flask + SQLite. Two roles: donor and hospital admin.

Run:
    python3 app.py
Then open http://127.0.0.1:5000

Demo accounts (synthetic data):
    Admin  admin@bloodbank.test / Admin@123
    Donor  donor@bloodbank.test / Donor@123
"""

import os
import re
import sqlite3
import secrets
from datetime import datetime, date, timedelta
from functools import wraps

from flask import (
    Flask, g, render_template, request, redirect, url_for,
    session, flash, jsonify, abort
)
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "instance", "bloodbank.db")
BLOOD_GROUPS = ["O+", "O-", "A+", "A-", "B+", "B-", "AB+", "AB-"]
CRITICAL_UNITS = 5
DONOR_COOLDOWN_DAYS = 90

app = Flask(__name__)


def _load_secret_key():
    env_key = os.environ.get("BLOODBANK_SECRET")
    if env_key:
        return env_key
    key_file = os.path.join(BASE_DIR, "instance", ".secret_key")
    os.makedirs(os.path.dirname(key_file), exist_ok=True)
    if os.path.exists(key_file):
        return open(key_file, "r", encoding="utf-8").read().strip()
    key = secrets.token_hex(32)
    with open(key_file, "w", encoding="utf-8") as fh:
        fh.write(key)
    return key


app.config.update(
    SECRET_KEY=_load_secret_key(),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    MAX_CONTENT_LENGTH=1 * 1024 * 1024,
)


# --------------------------------------------------------------------------
# Database helpers
# --------------------------------------------------------------------------

def get_db():
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
    return last_id


# --------------------------------------------------------------------------
# Schema + seed
# --------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('donor','admin')),
    phone TEXT,
    city TEXT,
    blood_group TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS hospitals (
    hospital_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    city TEXT NOT NULL,
    phone TEXT,
    address TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS inventory (
    inventory_id INTEGER PRIMARY KEY AUTOINCREMENT,
    hospital_id INTEGER NOT NULL,
    blood_group TEXT NOT NULL,
    units INTEGER NOT NULL DEFAULT 0 CHECK (units >= 0),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (hospital_id, blood_group),
    FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS requests (
    request_id INTEGER PRIMARY KEY AUTOINCREMENT,
    requester_id INTEGER NOT NULL,
    hospital_id INTEGER NOT NULL,
    blood_group TEXT NOT NULL,
    units INTEGER NOT NULL CHECK (units > 0),
    urgency TEXT NOT NULL DEFAULT 'Routine',
    note TEXT,
    status TEXT NOT NULL DEFAULT 'Pending' CHECK (status IN ('Pending','Approved','Completed','Cancelled')),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (requester_id) REFERENCES users(user_id),
    FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id)
);

CREATE TABLE IF NOT EXISTS donations (
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

CREATE TABLE IF NOT EXISTS appointments (
    appointment_id INTEGER PRIMARY KEY AUTOINCREMENT,
    donor_id INTEGER NOT NULL,
    hospital_id INTEGER NOT NULL,
    slot_date TEXT NOT NULL,
    slot_time TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Booked' CHECK (status IN ('Booked','Kept','Missed','Cancelled')),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (donor_id) REFERENCES users(user_id),
    FOREIGN KEY (hospital_id) REFERENCES hospitals(hospital_id)
);

CREATE INDEX IF NOT EXISTS idx_inventory_group ON inventory(blood_group);
CREATE INDEX IF NOT EXISTS idx_requests_status ON requests(status);
CREATE INDEX IF NOT EXISTS idx_donations_donor ON donations(donor_id);
"""


def seed(db):
    """Synthetic demo data. Clearly labelled in the UI."""
    admin_id = db.execute(
        "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
        ("Dr. Meera Joshi", "admin@bloodbank.test",
         generate_password_hash("Admin@123"), "admin", "020-4000-1200", "Pune", None),
    ).lastrowid
    donor1 = db.execute(
        "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
        ("Rahul Patil", "donor@bloodbank.test",
         generate_password_hash("Donor@123"), "donor", "99999 00001", "Pune", "O+"),
    ).lastrowid
    donor2 = db.execute(
        "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
        ("Sneha Sharma", "sneha@bloodbank.test",
         generate_password_hash("Donor@123"), "donor", "99999 00002", "Pune", "A+"),
    ).lastrowid
    donor3 = db.execute(
        "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
        ("Amit Kumar", "amit@bloodbank.test",
         generate_password_hash("Donor@123"), "donor", "99999 00003", "Mumbai", "B+"),
    ).lastrowid
    donor4 = db.execute(
        "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
        ("Priya Nair", "priya@bloodbank.test",
         generate_password_hash("Donor@123"), "donor", "99999 00004", "Pune", "O-"),
    ).lastrowid
    donor5 = db.execute(
        "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
        ("Imran Shaikh", "imran@bloodbank.test",
         generate_password_hash("Donor@123"), "donor", "99999 00005", "Mumbai", "AB+"),
    ).lastrowid

    hospitals = [
        ("City Hospital", "Pune", "020-4000-1200", "12 Station Road, Shivajinagar"),
        ("Life Care Hospital", "Mumbai", "022-6000-3400", "88 Marine Drive, Churchgate"),
        ("Sunrise Blood Centre", "Nashik", "0253-220-1180", "4 Gangapur Road, Nashik"),
    ]
    hospital_ids = []
    for h in hospitals:
        hospital_ids.append(db.execute(
            "INSERT INTO hospitals (name,city,phone,address) VALUES (?,?,?,?)", h
        ).lastrowid)

    stock = {
        0: {"O+": 18, "O-": 4, "A+": 12, "A-": 3, "B+": 9, "B-": 2, "AB+": 6, "AB-": 1},
        1: {"O+": 7, "O-": 8, "A+": 5, "A-": 6, "B+": 14, "B-": 4, "AB+": 3, "AB-": 2},
        2: {"O+": 11, "O-": 2, "A+": 7, "A-": 1, "B+": 4, "B-": 5, "AB+": 2, "AB-": 0},
    }
    for idx, hid in enumerate(hospital_ids):
        for group, units in stock[idx].items():
            db.execute(
                "INSERT INTO inventory (hospital_id,blood_group,units) VALUES (?,?,?)",
                (hid, group, units),
            )

    requests = [
        (donor1, hospital_ids[0], "O+", 2, "Critical", "Post-surgery transfusion, ward 4", "Approved"),
        (donor2, hospital_ids[0], "A+", 1, "Routine", "Day-care top-up", "Pending"),
        (donor3, hospital_ids[1], "B+", 3, "Critical", "Trauma bay 2", "Pending"),
        (donor1, hospital_ids[2], "O-", 1, "Routine", "Scheduled procedure", "Completed"),
        (donor2, hospital_ids[1], "AB-", 2, "Critical", "ICU request", "Pending"),
        (donor4, hospital_ids[2], "O-", 1, "Routine", "Planned orthopaedic surgery", "Pending"),
    ]
    for r in requests:
        db.execute(
            "INSERT INTO requests (requester_id,hospital_id,blood_group,units,urgency,note,status) VALUES (?,?,?,?,?,?,?)",
            r,
        )

    donations = [
        (donor1, hospital_ids[0], "2026-09-14", 1, "Whole blood"),
        (donor2, hospital_ids[0], "2026-08-30", 1, "Whole blood"),
        (donor3, hospital_ids[1], "2026-09-28", 1, "Whole blood"),
        (donor1, hospital_ids[0], "2026-06-02", 1, "Whole blood"),
        (donor2, hospital_ids[2], "2026-05-18", 1, "Whole blood"),
        (donor4, hospital_ids[0], "2026-07-11", 1, "Whole blood"),
        (donor5, hospital_ids[1], "2026-08-02", 1, "Whole blood"),
    ]
    for d in donations:
        db.execute(
            "INSERT INTO donations (donor_id,hospital_id,donation_date,units,note) VALUES (?,?,?,?,?)",
            d,
        )

    tomorrow = (date.today() + timedelta(days=1)).isoformat()
    db.execute(
        "INSERT INTO appointments (donor_id,hospital_id,slot_date,slot_time) VALUES (?,?,?,?)",
        (donor1, hospital_ids[0], tomorrow, "10:30"),
    )


def init_db():
    db = get_db()
    db.executescript(SCHEMA)
    count = db.execute("SELECT COUNT(*) AS c FROM users").fetchone()["c"]
    if count == 0:
        seed(db)
        db.commit()


# --------------------------------------------------------------------------
# Auth + CSRF
# --------------------------------------------------------------------------

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please sign in to continue.", "error")
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please sign in to continue.", "error")
            return redirect(url_for("login", next=request.path))
        if session.get("role") != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped


@app.before_request
def ensure_session():
    if "user_id" not in session:
        session["csrf"] = session.get("csrf") or secrets.token_hex(16)


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return session["csrf"]


app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def csrf_protect():
    if request.method == "POST":
        sent = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token")
        if not sent or not secrets.compare_digest(sent, session.get("csrf", "")):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Invalid CSRF token. Refresh the page."}), 400
            abort(400, description="Invalid CSRF token. Refresh the page and try again.")


@app.context_processor
def inject_globals():
    user = None
    if session.get("user_id"):
        user = query("SELECT * FROM users WHERE user_id = ?", (session["user_id"],), one=True)
    return {
        "current_user": user,
        "blood_groups": BLOOD_GROUPS,
        "critical_units": CRITICAL_UNITS,
        "today": date.today(),
    }


@app.errorhandler(403)
def forbidden(_e):
    return render_template("error.html", code=403,
                           title="Not your desk",
                           message="This area is limited to hospital administrators."), 403


@app.errorhandler(404)
def not_found(_e):
    return render_template("error.html", code=404,
                           title="Page not found",
                           message="That board or record does not exist."), 404


# --------------------------------------------------------------------------
# Public + auth
# --------------------------------------------------------------------------

@app.route("/")
def index():
    if session.get("user_id"):
        return redirect(url_for("admin_board" if session.get("role") == "admin" else "donor_dashboard"))
    totals = query("""
        SELECT blood_group, SUM(units) AS units
        FROM inventory GROUP BY blood_group ORDER BY blood_group
    """)
    total_units = sum(r["units"] for r in totals) if totals else 0
    hospitals = query("SELECT COUNT(*) AS c FROM hospitals", one=True)["c"]
    donors = query("SELECT COUNT(*) AS c FROM users WHERE role='donor'", one=True)["c"]
    return render_template("landing.html", totals=totals, total_units=total_units,
                           hospitals=hospitals, donors=donors)


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect(url_for("admin_board" if session.get("role") == "admin" else "donor_dashboard"))
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        user = query("SELECT * FROM users WHERE email = ?", (email,), one=True)
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["user_id"]
            session["role"] = user["role"]
            session["name"] = user["name"]
            session["csrf"] = secrets.token_hex(16)
            flash(f"Signed in as {user['name']}.", "success")
            target = request.args.get("next")
            if target and target.startswith("/") and not target.startswith("//"):
                return redirect(target)
            return redirect(url_for("admin_board" if user["role"] == "admin" else "donor_dashboard"))
        flash("Email or password is incorrect.", "error")
    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if session.get("user_id"):
        return redirect(url_for("index"))
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        phone = request.form.get("phone", "").strip()
        city = request.form.get("city", "").strip()
        blood = request.form.get("blood_group", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm", "")

        errors = []
        if len(name) < 2:
            errors.append("Enter your full name.")
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            errors.append("Enter a valid email address.")
        if len(password) < 8:
            errors.append("Use at least 8 characters for the password.")
        if password != confirm:
            errors.append("The two passwords do not match.")
        if blood not in BLOOD_GROUPS:
            errors.append("Choose a valid blood group.")
        if query("SELECT user_id FROM users WHERE email = ?", (email,), one=True):
            errors.append("An account with that email already exists.")

        if errors:
            for e in errors:
                flash(e, "error")
        else:
            uid = execute(
                "INSERT INTO users (name,email,password_hash,role,phone,city,blood_group) VALUES (?,?,?,?,?,?,?)",
                (name, email, generate_password_hash(password), "donor", phone, city, blood),
            )
            session.clear()
            session["user_id"] = uid
            session["role"] = "donor"
            session["name"] = name
            session["csrf"] = secrets.token_hex(16)
            flash("Your donor card is ready. Welcome aboard.", "success")
            return redirect(url_for("donor_dashboard"))
    return render_template("register.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    flash("Signed out. The board is still here when you need it.", "success")
    return redirect(url_for("index"))


# --------------------------------------------------------------------------
# Shared search
# --------------------------------------------------------------------------

@app.route("/search")
def search():
    blood = request.args.get("blood_group", "").strip()
    city = request.args.get("city", "").strip()
    rows = []
    if blood or city:
        sql = """
            SELECT i.blood_group, i.units, i.updated_at, h.name, h.city, h.phone
            FROM inventory i
            JOIN hospitals h ON h.hospital_id = i.hospital_id
            WHERE 1=1
        """
        args = []
        if blood:
            sql += " AND i.blood_group = ?"
            args.append(blood)
        if city:
            sql += " AND h.city LIKE ?"
            args.append(f"%{city}%")
        sql += " AND i.units > 0 ORDER BY i.units DESC, h.name"
        rows = query(sql, args)
        if session.get("user_id") and session.get("role") == "donor" and blood:
            session["last_search"] = blood
    totals = query("""
        SELECT blood_group, COALESCE(SUM(units),0) AS units
        FROM inventory GROUP BY blood_group
    """)
    counts = {r["blood_group"]: r["units"] for r in totals}
    totals = [{"blood_group": g, "units": counts.get(g, 0)} for g in BLOOD_GROUPS]
    template = "search.html" if session.get("user_id") else "search_public.html"
    return render_template(template, rows=rows, blood=blood, city=city, totals=totals)


# --------------------------------------------------------------------------
# Donor portal
# --------------------------------------------------------------------------

def donor_stats(user_id):
    last = query(
        "SELECT donation_date FROM donations WHERE donor_id = ? ORDER BY donation_date DESC LIMIT 1",
        (user_id,), one=True
    )
    total = query(
        "SELECT COUNT(*) AS c, COALESCE(SUM(units),0) AS u FROM donations WHERE donor_id = ?",
        (user_id,), one=True
    )
    last_date = None
    next_date = None
    days_left = 0
    if last:
        last_date = datetime.strptime(last["donation_date"], "%Y-%m-%d").date()
        next_date = last_date + timedelta(days=DONOR_COOLDOWN_DAYS)
        days_left = max(0, (next_date - date.today()).days)
    return {
        "last_date": last_date,
        "next_date": next_date,
        "days_left": days_left,
        "eligible": days_left == 0,
        "count": total["c"],
        "units": total["u"],
    }


@app.route("/dashboard")
@login_required
def donor_dashboard():
    if session.get("role") == "admin":
        return redirect(url_for("admin_board"))
    user = query("SELECT * FROM users WHERE user_id = ?", (session["user_id"],), one=True)
    stats = donor_stats(user["user_id"])
    history = query("""
        SELECT d.*, h.name AS hospital
        FROM donations d JOIN hospitals h ON h.hospital_id = d.hospital_id
        WHERE d.donor_id = ? ORDER BY d.donation_date DESC
    """, (user["user_id"],))
    appointments = query("""
        SELECT a.*, h.name AS hospital, h.city
        FROM appointments a JOIN hospitals h ON h.hospital_id = a.hospital_id
        WHERE a.donor_id = ? AND a.status = 'Booked'
        ORDER BY a.slot_date, a.slot_time
    """, (user["user_id"],))
    my_requests = query("""
        SELECT r.*, h.name AS hospital
        FROM requests r JOIN hospitals h ON h.hospital_id = r.hospital_id
        WHERE r.requester_id = ? ORDER BY r.created_at DESC
    """, (user["user_id"],))
    hospitals = query("SELECT * FROM hospitals ORDER BY name")
    return render_template("donor_dashboard.html", user=user, stats=stats,
                           history=history, appointments=appointments,
                           my_requests=my_requests, hospitals=hospitals)


@app.route("/book", methods=["POST"])
@login_required
def book_appointment():
    if session.get("role") != "donor":
        abort(403)
    hospital_id = request.form.get("hospital_id", type=int)
    slot_date = request.form.get("slot_date", "")
    slot_time = request.form.get("slot_time", "")
    if not hospital_id or not slot_date or not slot_time:
        flash("Choose a hospital, date and time.", "error")
        return redirect(url_for("donor_dashboard"))
    try:
        chosen = datetime.strptime(slot_date, "%Y-%m-%d").date()
    except ValueError:
        flash("That date is not valid.", "error")
        return redirect(url_for("donor_dashboard"))
    if chosen < date.today():
        flash("Pick today or a future date.", "error")
        return redirect(url_for("donor_dashboard"))
    execute(
        "INSERT INTO appointments (donor_id,hospital_id,slot_date,slot_time) VALUES (?,?,?,?)",
        (session["user_id"], hospital_id, slot_date, slot_time),
    )
    flash("Donation slot booked. Please arrive hydrated.", "success")
    return redirect(url_for("donor_dashboard"))


@app.route("/appointment/<int:appointment_id>/cancel", methods=["POST"])
@login_required
def cancel_appointment(appointment_id):
    row = query("SELECT * FROM appointments WHERE appointment_id = ?", (appointment_id,), one=True)
    if not row or (row["donor_id"] != session["user_id"] and session.get("role") != "admin"):
        abort(404)
    execute("UPDATE appointments SET status = 'Cancelled' WHERE appointment_id = ?", (appointment_id,))
    flash("Appointment cancelled.", "success")
    return redirect(url_for("donor_dashboard"))


@app.route("/request", methods=["POST"])
@login_required
def create_request():
    hospital_id = request.form.get("hospital_id", type=int)
    blood = request.form.get("blood_group", "")
    units = request.form.get("units", type=int)
    urgency = request.form.get("urgency", "Routine")
    note = request.form.get("note", "").strip()[:200]
    if not hospital_id or blood not in BLOOD_GROUPS or not units or units < 1 or units > 20:
        flash("Choose a hospital, blood group and 1–20 units.", "error")
        return redirect(url_for("donor_dashboard"))
    if urgency not in ("Routine", "Critical"):
        urgency = "Routine"
    execute(
        "INSERT INTO requests (requester_id,hospital_id,blood_group,units,urgency,note) VALUES (?,?,?,?,?,?)",
        (session["user_id"], hospital_id, blood, units, urgency, note),
    )
    flash("Request raised. The hospital desk will review it.", "success")
    return redirect(url_for("donor_dashboard"))


# --------------------------------------------------------------------------
# Admin board
# --------------------------------------------------------------------------

def stock_by_group():
    rows = query("""
        SELECT blood_group, SUM(units) AS units
        FROM inventory GROUP BY blood_group
    """)
    counts = {r["blood_group"]: r["units"] for r in rows}
    return [{"group": g, "units": counts.get(g, 0),
             "critical": counts.get(g, 0) <= CRITICAL_UNITS} for g in BLOOD_GROUPS]


def board_stats():
    req = query("""
        SELECT status, COUNT(*) AS c FROM requests GROUP BY status
    """)
    by_status = {r["status"]: r["c"] for r in req}
    donors = query("SELECT COUNT(*) AS c FROM users WHERE role='donor'", one=True)["c"]
    units = query("SELECT COALESCE(SUM(units),0) AS u FROM inventory", one=True)["u"]
    pending = by_status.get("Pending", 0)
    return {
        "donors": donors,
        "units": units,
        "pending": pending,
        "approved": by_status.get("Approved", 0),
        "completed": by_status.get("Completed", 0),
        "critical_groups": sum(1 for s in stock_by_group() if s["critical"]),
    }


@app.route("/admin")
@admin_required
def admin_board():
    stock = stock_by_group()
    stats = board_stats()
    slips = query("""
        SELECT r.*, u.name AS requester, h.name AS hospital, h.city
        FROM requests r
        JOIN users u ON u.user_id = r.requester_id
        JOIN hospitals h ON h.hospital_id = r.hospital_id
        WHERE r.status = 'Pending'
        ORDER BY CASE r.urgency WHEN 'Critical' THEN 0 ELSE 1 END, r.created_at
        LIMIT 4
    """)
    recent = query("""
        SELECT d.*, u.name AS donor, u.blood_group, h.name AS hospital
        FROM donations d
        JOIN users u ON u.user_id = d.donor_id
        JOIN hospitals h ON h.hospital_id = d.hospital_id
        ORDER BY d.donation_date DESC, d.donation_id DESC LIMIT 8
    """)
    hospitals = query("""
        SELECT h.*, COALESCE(SUM(i.units),0) AS units,
               COUNT(DISTINCT i.inventory_id) AS lines
        FROM hospitals h LEFT JOIN inventory i ON i.hospital_id = h.hospital_id
        GROUP BY h.hospital_id ORDER BY h.name
    """)
    donations_30 = query("""
        SELECT COUNT(*) AS c, COALESCE(SUM(units),0) AS u
        FROM donations WHERE donation_date >= date('now','-30 day')
    """, one=True)
    return render_template("admin_board.html", stock=stock, stats=stats,
                           slips=slips, recent=recent, hospitals=hospitals,
                           donations_30=donations_30)


@app.route("/admin/hospitals", methods=["GET", "POST"])
@admin_required
def admin_hospitals():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        city = request.form.get("city", "").strip()
        phone = request.form.get("phone", "").strip()
        address = request.form.get("address", "").strip()
        if not name or not city:
            flash("Hospital name and city are required.", "error")
        else:
            hid = execute(
                "INSERT INTO hospitals (name,city,phone,address) VALUES (?,?,?,?)",
                (name, city, phone, address),
            )
            for group in BLOOD_GROUPS:
                execute("INSERT INTO inventory (hospital_id,blood_group,units) VALUES (?,?,0)",
                        (hid, group))
            flash(f"{name} added with a fresh stock sheet.", "success")
        return redirect(url_for("admin_hospitals"))
    hospitals = query("""
        SELECT h.*, COALESCE(SUM(i.units),0) AS units,
               COUNT(DISTINCT i.inventory_id) AS lines
        FROM hospitals h LEFT JOIN inventory i ON i.hospital_id = h.hospital_id
        GROUP BY h.hospital_id ORDER BY h.name
    """)
    return render_template("admin_hospitals.html", hospitals=hospitals)


@app.route("/admin/hospitals/<int:hospital_id>/delete", methods=["POST"])
@admin_required
def delete_hospital(hospital_id):
    row = query("SELECT * FROM hospitals WHERE hospital_id = ?", (hospital_id,), one=True)
    if not row:
        abort(404)
    used = query("SELECT COUNT(*) AS c FROM requests WHERE hospital_id = ?", (hospital_id,), one=True)["c"]
    if used:
        flash("Cannot remove a hospital that has requests. Keep it for the record.", "error")
        return redirect(url_for("admin_hospitals"))
    execute("DELETE FROM hospitals WHERE hospital_id = ?", (hospital_id,))
    flash("Hospital removed.", "success")
    return redirect(url_for("admin_hospitals"))


@app.route("/admin/inventory", methods=["GET", "POST"])
@admin_required
def admin_inventory():
    hospitals = query("SELECT * FROM hospitals ORDER BY name")
    if request.method == "POST":
        hospital_id = request.form.get("hospital_id", type=int)
        blood = request.form.get("blood_group", "")
        try:
            units = int(request.form.get("units", ""))
        except ValueError:
            units = -1
        if not hospital_id or blood not in BLOOD_GROUPS or units < 0 or units > 999:
            flash("Pick a hospital and blood group, and enter 0–999 units.", "error")
            return redirect(url_for("admin_inventory"))
        existing = query(
            "SELECT * FROM inventory WHERE hospital_id = ? AND blood_group = ?",
            (hospital_id, blood), one=True
        )
        hospital = query("SELECT name FROM hospitals WHERE hospital_id = ?", (hospital_id,), one=True)
        if existing:
            execute(
                "UPDATE inventory SET units = ?, updated_at = datetime('now') WHERE inventory_id = ?",
                (units, existing["inventory_id"]),
            )
            flash(f"{blood} at {hospital['name']} set to {units} units.", "success")
        else:
            execute("INSERT INTO inventory (hospital_id,blood_group,units) VALUES (?,?,?)",
                    (hospital_id, blood, units))
            flash("Stock line created.", "success")
        return redirect(url_for("admin_inventory"))
    rows = query("""
        SELECT i.*, h.name AS hospital, h.city
        FROM inventory i JOIN hospitals h ON h.hospital_id = i.hospital_id
        ORDER BY CASE WHEN i.units <= ? THEN 0 ELSE 1 END, i.units ASC, h.name
    """, (CRITICAL_UNITS,))
    return render_template("admin_inventory.html", hospitals=hospitals, rows=rows)


@app.route("/admin/requests")
@admin_required
def admin_requests():
    status = request.args.get("status", "All")
    if status not in ("All", "Pending", "Approved", "Completed", "Cancelled"):
        status = "All"
    sql = """
        SELECT r.*, u.name AS requester, h.name AS hospital, h.city
        FROM requests r
        JOIN users u ON u.user_id = r.requester_id
        JOIN hospitals h ON h.hospital_id = r.hospital_id
    """
    args = []
    if status != "All":
        sql += " WHERE r.status = ?"
        args.append(status)
    sql += " ORDER BY CASE r.urgency WHEN 'Critical' THEN 0 ELSE 1 END, r.created_at DESC"
    rows = query(sql, args)
    hospitals = query("SELECT * FROM hospitals ORDER BY name")
    return render_template("admin_requests.html", rows=rows, status=status,
                           hospitals=hospitals)


@app.route("/admin/requests/new", methods=["POST"])
@admin_required
def admin_new_request():
    hospital_id = request.form.get("hospital_id", type=int)
    blood = request.form.get("blood_group", "")
    units = request.form.get("units", type=int)
    urgency = request.form.get("urgency", "Routine")
    note = request.form.get("note", "").strip()[:200]
    if not hospital_id or blood not in BLOOD_GROUPS or not units or units < 1 or units > 20:
        flash("Choose a hospital, blood group and 1–20 units.", "error")
    else:
        execute(
            "INSERT INTO requests (requester_id,hospital_id,blood_group,units,urgency,note) VALUES (?,?,?,?,?,?)",
            (session["user_id"], hospital_id, blood, units, urgency, note),
        )
        flash("Request added to the rail.", "success")
    return redirect(url_for("admin_requests"))


@app.route("/api/requests/<int:request_id>/status", methods=["POST"])
@admin_required
def api_request_status(request_id):
    payload = request.get_json(silent=True) or {}
    new_status = payload.get("status", "")
    row = query("SELECT * FROM requests WHERE request_id = ?", (request_id,), one=True)
    if not row:
        return jsonify({"error": "Request not found."}), 404
    allowed = {
        "Pending": {"Approved", "Cancelled"},
        "Approved": {"Completed", "Cancelled"},
        "Completed": set(),
        "Cancelled": set(),
    }
    if new_status not in allowed.get(row["status"], set()):
        return jsonify({"error": f"Cannot move {row['status']} to {new_status}."}), 409

    if new_status == "Completed":
        inv = query(
            "SELECT * FROM inventory WHERE hospital_id = ? AND blood_group = ?",
            (row["hospital_id"], row["blood_group"]), one=True
        )
        if not inv or inv["units"] < row["units"]:
            return jsonify({"error": "Not enough units in stock at that hospital to complete this request."}), 409
        execute("UPDATE inventory SET units = units - ?, updated_at = datetime('now') WHERE inventory_id = ?",
                (row["units"], inv["inventory_id"]))
    execute("UPDATE requests SET status = ?, updated_at = datetime('now') WHERE request_id = ?",
            (new_status, request_id))
    return jsonify({"ok": True, "status": new_status,
                    "units": query("SELECT SUM(units) AS u FROM inventory WHERE blood_group = ?",
                                   (row["blood_group"],), one=True)["u"]})


@app.route("/admin/donations", methods=["GET", "POST"])
@admin_required
def admin_donations():
    donors = query("SELECT * FROM users WHERE role='donor' ORDER BY name")
    hospitals = query("SELECT * FROM hospitals ORDER BY name")
    if request.method == "POST":
        donor_id = request.form.get("donor_id", type=int)
        hospital_id = request.form.get("hospital_id", type=int)
        donation_date = request.form.get("donation_date", "")
        units = request.form.get("units", type=int) or 1
        note = request.form.get("note", "").strip()[:200]
        try:
            datetime.strptime(donation_date, "%Y-%m-%d")
        except ValueError:
            flash("Enter a valid donation date.", "error")
            return redirect(url_for("admin_donations"))
        if not donor_id or not hospital_id or units < 1 or units > 2:
            flash("Choose a donor, hospital and 1–2 units.", "error")
            return redirect(url_for("admin_donations"))
        execute(
            "INSERT INTO donations (donor_id,hospital_id,donation_date,units,note) VALUES (?,?,?,?,?)",
            (donor_id, hospital_id, donation_date, units, note),
        )
        group = query("SELECT blood_group FROM users WHERE user_id = ?", (donor_id,), one=True)["blood_group"]
        if group:
            inv = query("SELECT * FROM inventory WHERE hospital_id = ? AND blood_group = ?",
                        (hospital_id, group), one=True)
            if inv:
                execute("UPDATE inventory SET units = units + ?, updated_at = datetime('now') WHERE inventory_id = ?",
                        (units, inv["inventory_id"]))
            else:
                execute("INSERT INTO inventory (hospital_id,blood_group,units) VALUES (?,?,?)",
                        (hospital_id, group, units))
        flash("Donation logged and stock updated.", "success")
        return redirect(url_for("admin_donations"))
    rows = query("""
        SELECT d.*, u.name AS donor, u.blood_group, h.name AS hospital
        FROM donations d
        JOIN users u ON u.user_id = d.donor_id
        JOIN hospitals h ON h.hospital_id = d.hospital_id
        ORDER BY d.donation_date DESC, d.donation_id DESC
    """)
    return render_template("admin_donations.html", donors=donors, hospitals=hospitals, rows=rows)


@app.route("/admin/donors")
@admin_required
def admin_donors():
    search = request.args.get("q", "").strip()
    sql = """
        SELECT u.*,
          (SELECT MAX(donation_date) FROM donations WHERE donor_id = u.user_id) AS last_donation,
          (SELECT COUNT(*) FROM donations WHERE donor_id = u.user_id) AS donations
        FROM users u WHERE u.role = 'donor'
    """
    args = []
    if search:
        sql += " AND (u.name LIKE ? OR u.email LIKE ? OR u.city LIKE ? OR u.blood_group = ?)"
        like = f"%{search}%"
        args = [like, like, like, search]
    sql += " ORDER BY u.name"
    donors = query(sql, args)
    return render_template("admin_donors.html", donors=donors, search=search)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

@app.cli.command("init-db")
def init_db_command():
    init_db()
    print("Database ready at", DB_PATH)


with app.app_context():
    init_db()


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
