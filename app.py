# Alien Tees x Bodyfitters — Stock Room (একটাই ফাইল: সার্ভার + ওয়েবসাইট)
# চালানো:  pip install flask requests pillow openpyxl reportlab   →   python app.py   →   http://localhost:5000
import base64, csv, importlib, io, json, os, pickle, re, sqlite3, subprocess, sys, tempfile, threading, time, shutil, zipfile, secrets
from datetime import datetime, time as datetime_time, timedelta
from zoneinfo import ZoneInfo
from colorsys import rgb_to_hsv
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import wraps
from urllib.parse import urlparse
import requests
from openpyxl import load_workbook, Workbook
from flask import Flask, Response, jsonify, request, session, send_file
from PIL import Image
from werkzeug.security import check_password_hash, generate_password_hash

SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret-key")
ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASS = os.environ.get("ADMIN_PASS", "alien@123")   # প্রথমবার চালানোর আগে বদলান
CACHE_TTL, STOCK_TTL = 600, 60

APP_BUILD = "ERP-final-2026-10-19-stock-recon-profit-alerts-v3"
app = Flask(__name__)
app.secret_key = SECRET_KEY
app.permanent_session_lifetime = timedelta(hours=9)
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax")
DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.db")
INSTANCE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "instance")
EXCEL_STOCK_FILE = os.path.join(INSTANCE, "stock_master.xlsx")
EXCEL_BACKUP_DIR = os.path.join(INSTANCE, "stock_backups")
FULL_BACKUP_DIR = os.path.join(INSTANCE, "full_backups")
EXCEL_LOCK = threading.RLock()
DHAKA_TZ = ZoneInfo("Asia/Dhaka")

PERMS = {
    "order_sheet_upload": "Can upload Order Sheet Excel",
    "order_view": "Can view Order Dashboard",
    "return_sheet_upload": "Can upload Return Sheet Excel",
    "return_view": "Can view Return Dashboard",
    "courier_view": "Can view Courier Tracking",
    "courier_update": "Can update Courier Status",
    "pos_use": "Can make POS sales (deducts stock)",
    "pos_orders": "Can view POS orders, reports, and receipts",
    "pos_delete": "Can delete POS sales (restores stock and records an audit trail)",
    "purchase_view": "Can view Purchase Management",
    "purchase_manage": "Can create and manage purchases",
    "supplier_manage": "Can manage suppliers and supplier payments",
    "expense_view": "Can view Expense Management",
    "expense_manage": "Can create, approve, and manage expenses",
    "hr_view": "Can view HR Management",
    "hr_manage": "Can manage employees and attendance",
    "hr_leave_approve": "Can review and approve leave requests",
    "hr_payroll": "Can manage payroll and payslips",
    "price_edit": "Can edit prices",
    "stock_manage": "Full stock management (all changes)",
    "stock_dashboard": "Can open Stock Dashboard",
    "stock_view": "Can view Stock Dashboard",
    "stock_value_view": "Can view prices and stock value",
    "stock_history_view": "Can view Stock History",
    "stock_add": "Can add stock quantity",
    "stock_reduce": "Can reduce stock quantity",
    "stock_set": "Can set stock quantity directly",
    "stock_comment": "Can add stock sale comments",
    "stock_comment_delete": "Can delete stock sale comments",
    "view_stock": "Can view stock",
    "search": "Can search by text",
    "image_search": "Can search by image",
    "view_product": "Can view product details",
    "size_filter": "Can filter by size",
    "download": "Can download HD images",
    "copy": "Can copy HD images",
    "link_copy": "Can copy links",
    "report": "Can report products",
    "id_copy": "Can copy Product ID",
    "name_copy": "Can copy Product Name",
    "full_info_copy": "Can copy full information",
    "share_product": "Can share products",
    "website_view": "Can open the website",
}
DEFAULT_PERMS = {k: False for k in PERMS}

# Staff/User: আগের কাজগুলো চালু থাকবে; price_edit আলাদা permission।
STAFF_DEFAULTS = {
    "search": True, "image_search": True, "view_product": True,
    "view_stock": False, "size_filter": True,
    "download": True, "copy": True, "link_copy": True,
    "favorite": True, "report": True,
    "id_copy": True, "name_copy": True, "full_info_copy": True,
    "share_product": True, "website_view": True,
    "price_edit": False,
    "stock_manage": False,
    "stock_dashboard": False, "stock_view": False, "stock_value_view": False, "stock_history_view": False,
    "stock_comment": True, "stock_comment_delete": False,
}

# CS: শুধু search → product view → real website link।
CS_DEFAULTS = {
    "search": True, "image_search": True, "view_product": True, "website_view": True,
    "price_edit": False, "view_stock": False, "size_filter": False,
    "stock_manage": False,
    "stock_dashboard": False, "stock_view": False, "stock_value_view": False, "stock_history_view": False,
    "stock_comment": True, "stock_comment_delete": False,
    "download": False, "copy": False, "link_copy": False,
    "favorite": False, "report": False,
    "id_copy": False, "name_copy": False, "full_info_copy": False, "share_product": False,
}

# Guest/Customer permissions Admin Management থেকে আলাদা control করা যাবে।
GUEST_PERM_LABELS = {
    "search": "🔎 টেক্সট সার্চ",
    "image_search": "🖼️ ইমেজ সার্চ",
    "view_product": "👁️ প্রোডাক্ট দেখা",
    "view_stock": "📦 স্টক দেখা",
    "size_filter": "📏 সাইজ ফিল্টার",
    "download": "⬇️ HD ডাউনলোড",
    "copy": "⧉ HD ছবি কপি",
    "link_copy": "🔗 লিংক কপি",
    "id_copy": "📋 ID কপি",
    "name_copy": "📝 নাম কপি",
    "full_info_copy": "📄 Full Info",
    "share_product": "📤 Share",
    "website_view": "🌐 ওয়েবসাইটে দেখুন",
}
GUEST_DEFAULTS = {
    "search": True, "image_search": True, "view_product": True,
    "view_stock": False, "size_filter": False,
    "download": False, "copy": False, "link_copy": False,
    "id_copy": True, "name_copy": True, "full_info_copy": True,
    "share_product": True, "website_view": True,
}

def db():
    # A bounded wait avoids transient "database is locked" errors during simultaneous POS/ERP saves.
    c = sqlite3.connect(DB, timeout=20)
    c.execute("PRAGMA busy_timeout=20000")
    try: c.execute("PRAGMA synchronous=NORMAL")
    except sqlite3.DatabaseError: pass
    c.row_factory = sqlite3.Row
    return c

def init():
    with db() as c:
        # WAL improves read/write concurrency for the many small API requests from POS/mobile.
        try:
            c.execute("PRAGMA journal_mode=WAL")
            c.execute("PRAGMA synchronous=NORMAL")
        except sqlite3.DatabaseError:
            pass
        c.execute("CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE, pw TEXT, role TEXT, perms TEXT NOT NULL DEFAULT '{}')")
        try: c.execute("ALTER TABLE users ADD COLUMN perms TEXT NOT NULL DEFAULT '{}'" )
        except sqlite3.OperationalError: pass
        c.execute("CREATE TABLE IF NOT EXISTS activity_logs(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, action TEXT, query TEXT, details TEXT, created_at REAL)")
        c.execute("CREATE TABLE IF NOT EXISTS favorites(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, product_key TEXT, created_at REAL, UNIQUE(username, product_key))")
        c.execute("CREATE TABLE IF NOT EXISTS recent_views(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, product_key TEXT, viewed_at REAL)")
        c.execute("CREATE TABLE IF NOT EXISTS product_reports(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, product_key TEXT, reason TEXT, created_at REAL, status TEXT DEFAULT 'open')")
        c.execute("CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS sources(id INTEGER PRIMARY KEY, kind TEXT, name TEXT, url TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS prices(k TEXT PRIMARY KEY, price REAL)")
        c.execute("CREATE TABLE IF NOT EXISTS excel_stock_history(id INTEGER PRIMARY KEY AUTOINCREMENT, sheet TEXT, row_number INTEGER, product_code TEXT, size TEXT, old_qty INTEGER, new_qty INTEGER, username TEXT, note TEXT, created_at REAL)")
        c.execute("CREATE TABLE IF NOT EXISTS excel_stock_batches(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, change_count INTEGER, plus_units INTEGER, minus_units INTEGER, created_at REAL)")
        c.execute("CREATE TABLE IF NOT EXISTS excel_stock_comments(id INTEGER PRIMARY KEY AUTOINCREMENT, sheet TEXT, product_code TEXT, size TEXT, sold_qty INTEGER NOT NULL, message TEXT, username TEXT NOT NULL, created_at REAL NOT NULL)")
        c.execute("CREATE TABLE IF NOT EXISTS excel_stock_imports(id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, sha256 TEXT, bytes INTEGER, username TEXT, created_at REAL)")
        c.execute("CREATE TABLE IF NOT EXISTS order_history(id INTEGER PRIMARY KEY AUTOINCREMENT, batch_id TEXT NOT NULL, source_file TEXT NOT NULL, row_number INTEGER NOT NULL, order_date TEXT, source TEXT, assigned_to TEXT, payment_status TEXT, order_status TEXT, delivery_status TEXT, cons_id TEXT, address TEXT, phone TEXT, customer_name TEXT, product_code TEXT, size TEXT, product_price TEXT, delivery_charge TEXT, advance_paid TEXT, collectible_amount TEXT, rating TEXT, note TEXT, image_url TEXT, row_json TEXT NOT NULL, uploaded_by TEXT NOT NULL, uploaded_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_order_history_date ON order_history(order_date)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_order_history_phone ON order_history(phone)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_order_history_code ON order_history(product_code)")
        c.execute("CREATE TABLE IF NOT EXISTS return_history(id INTEGER PRIMARY KEY AUTOINCREMENT, batch_id TEXT NOT NULL, source_file TEXT NOT NULL, row_number INTEGER NOT NULL, assigned_to TEXT, cons_id TEXT, address TEXT, phone TEXT, customer_name TEXT, product_code TEXT, size TEXT, collected_amount TEXT, note TEXT, pathao_bill_status TEXT, return_reason TEXT, row_json TEXT NOT NULL, uploaded_by TEXT NOT NULL, uploaded_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_return_history_cons ON return_history(cons_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_return_history_code ON return_history(product_code)")
        c.execute("CREATE TABLE IF NOT EXISTS order_history_comments(id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL, message TEXT NOT NULL, username TEXT NOT NULL, created_at REAL NOT NULL)")
        # Additive ERP tracking tables; existing order/return and stock records remain untouched.
        c.execute("CREATE TABLE IF NOT EXISTS courier_tracking(id INTEGER PRIMARY KEY AUTOINCREMENT, cons_id TEXT NOT NULL, courier TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'Pending', rating TEXT NOT NULL DEFAULT '', details TEXT NOT NULL DEFAULT '', order_id INTEGER, updated_by TEXT NOT NULL DEFAULT '', updated_at REAL NOT NULL, UNIQUE(cons_id,courier))")
        for col, definition in (("api_amount", "REAL"), ("api_amount_field", "TEXT NOT NULL DEFAULT ''")):
            try: c.execute(f"ALTER TABLE courier_tracking ADD COLUMN {col} {definition}")
            except sqlite3.OperationalError: pass
        c.execute("CREATE INDEX IF NOT EXISTS idx_courier_tracking_status ON courier_tracking(status)")
        c.execute("CREATE TABLE IF NOT EXISTS courier_events(id INTEGER PRIMARY KEY AUTOINCREMENT, cons_id TEXT NOT NULL, courier TEXT NOT NULL DEFAULT '', status TEXT NOT NULL, details TEXT NOT NULL DEFAULT '', source TEXT NOT NULL DEFAULT 'manual', created_at REAL NOT NULL, UNIQUE(cons_id,courier,status,source,details))")
        c.execute("CREATE TABLE IF NOT EXISTS courier_fraud_checks(id INTEGER PRIMARY KEY AUTOINCREMENT, phone TEXT NOT NULL, courier TEXT NOT NULL DEFAULT '', orders INTEGER NOT NULL DEFAULT 0, delivered INTEGER NOT NULL DEFAULT 0, returned INTEGER NOT NULL DEFAULT 0, provider_result TEXT NOT NULL DEFAULT '', checked_by TEXT NOT NULL DEFAULT '', checked_at REAL NOT NULL)")
        c.execute("CREATE TABLE IF NOT EXISTS courier_order_mappings(courier TEXT PRIMARY KEY, status_path TEXT NOT NULL DEFAULT '/status_by_cid/{cons_id}', updated_at REAL NOT NULL DEFAULT 0)")
        for col, definition in (("status_path", "TEXT NOT NULL DEFAULT '/status_by_cid/{cons_id}'"),):
            try: c.execute(f"ALTER TABLE courier_order_mappings ADD COLUMN {col} {definition}")
            except sqlite3.OperationalError: pass
        c.execute("CREATE TABLE IF NOT EXISTS pos_sales(id INTEGER PRIMARY KEY AUTOINCREMENT, invoice TEXT UNIQUE, username TEXT, customer_name TEXT, phone TEXT, address TEXT, subtotal REAL, discount REAL, total REAL, paid REAL, method TEXT, items TEXT, created_at REAL)")
        # Indexed normalized phone fields keep POS customer autocomplete fast on older, larger datasets.
        for table in ("pos_sales", "order_history"):
            try: c.execute(f"ALTER TABLE {table} ADD COLUMN phone_digits TEXT NOT NULL DEFAULT ''")
            except sqlite3.OperationalError: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_pos_sales_phone_digits ON pos_sales(phone_digits COLLATE NOCASE)")
        except sqlite3.OperationalError: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_order_history_phone_digits ON order_history(phone_digits COLLATE NOCASE)")
        except sqlite3.OperationalError: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_pos_sales_created_at ON pos_sales(created_at)")
        except sqlite3.OperationalError: pass
        # One-time backfill. Rows with no phone stay blank and do not affect customer lookup.
        for table in ("pos_sales", "order_history"):
            try:
                missing = c.execute(f"SELECT id, phone FROM {table} WHERE COALESCE(phone,'')<>'' AND COALESCE(phone_digits,'')='' LIMIT 50000").fetchall()
                if missing:
                    c.executemany(f"UPDATE {table} SET phone_digits=? WHERE id=?", [(re.sub(r"\D", "", str(r["phone"] or "")), int(r["id"])) for r in missing])
            except sqlite3.OperationalError:
                pass
        try: c.execute("ALTER TABLE pos_sales ADD COLUMN company_key TEXT NOT NULL DEFAULT 'alien_tees'")
        except sqlite3.OperationalError: pass
        c.execute("CREATE TABLE IF NOT EXISTS pos_memo_tokens(sale_id INTEGER PRIMARY KEY, token TEXT UNIQUE NOT NULL, created_at REAL NOT NULL)")
        c.execute("CREATE TABLE IF NOT EXISTS pos_deleted_sales(id INTEGER PRIMARY KEY AUTOINCREMENT, original_sale_id INTEGER NOT NULL, invoice TEXT NOT NULL, deleted_by TEXT NOT NULL, deleted_at REAL NOT NULL, reason TEXT NOT NULL DEFAULT '', sale_json TEXT NOT NULL, restored_units INTEGER NOT NULL DEFAULT 0)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_pos_deleted_sales_invoice ON pos_deleted_sales(invoice)")
        c.execute("CREATE TABLE IF NOT EXISTS order_history_memo_tokens(order_id INTEGER PRIMARY KEY, token TEXT UNIQUE NOT NULL, created_at REAL NOT NULL)")
        c.execute("CREATE TABLE IF NOT EXISTS erp_purchase_memo_tokens(purchase_id INTEGER PRIMARY KEY, token TEXT UNIQUE NOT NULL, created_at REAL NOT NULL)")
        # Additive Purchase / Expense Management tables. No existing table or row is reset.
        c.execute("CREATE TABLE IF NOT EXISTS erp_suppliers(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT NOT NULL DEFAULT '', email TEXT NOT NULL DEFAULT '', address TEXT NOT NULL DEFAULT '', opening_balance REAL NOT NULL DEFAULT 0, notes TEXT NOT NULL DEFAULT '', active INTEGER NOT NULL DEFAULT 1, created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_suppliers_name ON erp_suppliers(name)")
        c.execute("CREATE TABLE IF NOT EXISTS erp_purchases(id INTEGER PRIMARY KEY AUTOINCREMENT, purchase_no TEXT UNIQUE NOT NULL, supplier_id INTEGER NOT NULL, supplier_name TEXT NOT NULL, invoice_no TEXT NOT NULL DEFAULT '', purchase_date TEXT NOT NULL, subtotal REAL NOT NULL DEFAULT 0, discount REAL NOT NULL DEFAULT 0, total REAL NOT NULL DEFAULT 0, notes TEXT NOT NULL DEFAULT '', stock_received INTEGER NOT NULL DEFAULT 0, created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_purchases_date ON erp_purchases(purchase_date)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_purchases_supplier ON erp_purchases(supplier_id)")
        c.execute("CREATE TABLE IF NOT EXISTS erp_purchase_items(id INTEGER PRIMARY KEY AUTOINCREMENT, purchase_id INTEGER NOT NULL, stock_sheet TEXT NOT NULL DEFAULT '', item_code TEXT NOT NULL DEFAULT '', item_name TEXT NOT NULL DEFAULT '', size TEXT NOT NULL DEFAULT '', quantity INTEGER NOT NULL, unit_cost REAL NOT NULL, line_total REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_purchase_items_purchase ON erp_purchase_items(purchase_id)")
        c.execute("CREATE TABLE IF NOT EXISTS erp_supplier_payments(id INTEGER PRIMARY KEY AUTOINCREMENT, supplier_id INTEGER NOT NULL, purchase_id INTEGER, amount REAL NOT NULL, payment_date TEXT NOT NULL, payment_method TEXT NOT NULL DEFAULT 'Cash', reference TEXT NOT NULL DEFAULT '', notes TEXT NOT NULL DEFAULT '', created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_supplier_payments_supplier ON erp_supplier_payments(supplier_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_supplier_payments_purchase ON erp_supplier_payments(purchase_id)")
        c.execute("CREATE TABLE IF NOT EXISTS erp_expenses(id INTEGER PRIMARY KEY AUTOINCREMENT, expense_no TEXT UNIQUE NOT NULL, title TEXT NOT NULL, category TEXT NOT NULL, amount REAL NOT NULL, expense_date TEXT NOT NULL, payment_method TEXT NOT NULL DEFAULT 'Cash', paid_to TEXT NOT NULL DEFAULT '', reference TEXT NOT NULL DEFAULT '', notes TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'Pending', created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL, updated_by TEXT NOT NULL DEFAULT '', updated_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_expenses_date ON erp_expenses(expense_date)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_erp_expenses_category ON erp_expenses(category)")
        # Daily cash close is additive and does not alter existing sales/expense records.
        c.execute("""CREATE TABLE IF NOT EXISTS erp_cash_closings(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            closing_date TEXT NOT NULL,
            opening_cash REAL NOT NULL DEFAULT 0,
            expected_cash REAL NOT NULL DEFAULT 0,
            actual_cash REAL NOT NULL DEFAULT 0,
            difference REAL NOT NULL DEFAULT 0,
            notes TEXT NOT NULL DEFAULT '',
            closed_by TEXT NOT NULL DEFAULT '',
            created_at REAL NOT NULL,
            updated_at REAL NOT NULL,
            UNIQUE(closing_date)
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_cash_closings_date ON erp_cash_closings(closing_date)")
        # Physical stock count / reconciliation trail; original inventory tables stay untouched.
        c.execute("""CREATE TABLE IF NOT EXISTS erp_stock_reconciliations(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sheet TEXT NOT NULL,
            row_number INTEGER NOT NULL,
            product_code TEXT NOT NULL,
            size TEXT NOT NULL,
            system_qty INTEGER NOT NULL DEFAULT 0,
            counted_qty INTEGER NOT NULL DEFAULT 0,
            variance INTEGER NOT NULL DEFAULT 0,
            action TEXT NOT NULL DEFAULT 'counted',
            note TEXT NOT NULL DEFAULT '',
            username TEXT NOT NULL DEFAULT '',
            created_at REAL NOT NULL
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_stock_recon_item ON erp_stock_reconciliations(sheet,row_number,product_code,size,created_at DESC)")
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_expenses_date_status_method ON erp_expenses(expense_date,status,payment_method)")
        except sqlite3.OperationalError: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_supplier_payments_date_method ON erp_supplier_payments(payment_date,payment_method)")
        except sqlite3.OperationalError: pass
        # Additive HR module. These tables are separate from existing ERP data.
        c.execute("CREATE TABLE IF NOT EXISTS hr_employees(id INTEGER PRIMARY KEY AUTOINCREMENT, employee_code TEXT UNIQUE NOT NULL, full_name TEXT NOT NULL, phone TEXT NOT NULL DEFAULT '', email TEXT NOT NULL DEFAULT '', department TEXT NOT NULL DEFAULT '', designation TEXT NOT NULL DEFAULT '', employment_type TEXT NOT NULL DEFAULT 'Full-time', join_date TEXT NOT NULL DEFAULT '', salary REAL NOT NULL DEFAULT 0, shift_minutes INTEGER NOT NULL DEFAULT 0, status TEXT NOT NULL DEFAULT 'Active', address TEXT NOT NULL DEFAULT '', emergency_contact TEXT NOT NULL DEFAULT '', notes TEXT NOT NULL DEFAULT '', photo_file TEXT NOT NULL DEFAULT '', created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL, updated_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hr_employees_name ON hr_employees(full_name)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hr_employees_status ON hr_employees(status)")
        try: c.execute("ALTER TABLE hr_employees ADD COLUMN shift_minutes INTEGER NOT NULL DEFAULT 0")
        except sqlite3.OperationalError: pass
        # Per-employee overtime rate is added without changing existing salary records.
        try: c.execute("ALTER TABLE hr_employees ADD COLUMN overtime_rate REAL NOT NULL DEFAULT 0")
        except sqlite3.OperationalError: pass
        c.execute("CREATE TABLE IF NOT EXISTS hr_attendance(id INTEGER PRIMARY KEY AUTOINCREMENT, employee_id INTEGER NOT NULL, attendance_date TEXT NOT NULL, check_in TEXT NOT NULL DEFAULT '', check_out TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'Present', work_hours REAL NOT NULL DEFAULT 0, notes TEXT NOT NULL DEFAULT '', photo_file TEXT NOT NULL DEFAULT '', created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL, updated_at REAL NOT NULL, UNIQUE(employee_id,attendance_date))")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hr_attendance_date ON hr_attendance(attendance_date)")
        # Overtime tracking is additive: older attendance, payroll and ERP data remain intact.
        for col, definition in (("worked_minutes", "INTEGER NOT NULL DEFAULT 0"), ("shift_minutes", "INTEGER NOT NULL DEFAULT 480"), ("overtime_minutes", "INTEGER NOT NULL DEFAULT 0")):
            try: c.execute(f"ALTER TABLE hr_attendance ADD COLUMN {col} {definition}")
            except sqlite3.OperationalError: pass
        c.execute("INSERT OR IGNORE INTO settings(k,v) VALUES('hr_shift_minutes','480')")
        # Global monthly grace for late arrivals, and deduction applied to each late day beyond it.
        c.execute("INSERT OR IGNORE INTO settings(k,v) VALUES('hr_late_allowed_days','3')")
        c.execute("INSERT OR IGNORE INTO settings(k,v) VALUES('hr_late_deduction_per_day','100')")
        c.execute("UPDATE hr_attendance SET worked_minutes=CAST(ROUND(work_hours*60) AS INTEGER) WHERE worked_minutes=0 AND work_hours>0")
        c.execute("UPDATE hr_attendance SET overtime_minutes=MAX(0,worked_minutes-shift_minutes) WHERE overtime_minutes=0 AND worked_minutes>shift_minutes")
        c.execute("CREATE TABLE IF NOT EXISTS hr_leaves(id INTEGER PRIMARY KEY AUTOINCREMENT, employee_id INTEGER NOT NULL, leave_type TEXT NOT NULL DEFAULT 'Casual', start_date TEXT NOT NULL, end_date TEXT NOT NULL, days REAL NOT NULL DEFAULT 1, reason TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'Pending', reviewed_by TEXT NOT NULL DEFAULT '', reviewed_at REAL, created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hr_leaves_status ON hr_leaves(status)")
        c.execute("CREATE TABLE IF NOT EXISTS hr_payroll(id INTEGER PRIMARY KEY AUTOINCREMENT, employee_id INTEGER NOT NULL, pay_month TEXT NOT NULL, basic_salary REAL NOT NULL DEFAULT 0, allowances REAL NOT NULL DEFAULT 0, bonus REAL NOT NULL DEFAULT 0, deductions REAL NOT NULL DEFAULT 0, net_pay REAL NOT NULL DEFAULT 0, status TEXT NOT NULL DEFAULT 'Draft', paid_date TEXT NOT NULL DEFAULT '', payment_method TEXT NOT NULL DEFAULT 'Bank Transfer', reference TEXT NOT NULL DEFAULT '', notes TEXT NOT NULL DEFAULT '', created_by TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL, updated_at REAL NOT NULL, UNIQUE(employee_id,pay_month))")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hr_payroll_month ON hr_payroll(pay_month)")
        # Snapshot OT hours/rate/pay in each payslip so printed historical payroll stays consistent.
        for col, definition in (("overtime_hours", "REAL NOT NULL DEFAULT 0"),
                                 ("overtime_rate", "REAL NOT NULL DEFAULT 0"),
                                 ("overtime_pay", "REAL NOT NULL DEFAULT 0"),
                                 # -1 marks an older payroll row that predates late-deduction snapshots.
                                 ("late_days", "INTEGER NOT NULL DEFAULT -1"),
                                 ("late_allowed_days", "INTEGER NOT NULL DEFAULT -1"),
                                 ("late_deduction", "REAL NOT NULL DEFAULT -1")):
            try: c.execute(f"ALTER TABLE hr_payroll ADD COLUMN {col} {definition}")
            except sqlite3.OperationalError: pass
        # Additive-only POS memo fields; existing sales remain untouched.
        for col, definition in (("delivery_charge", "REAL NOT NULL DEFAULT 0"), ("delivery_zone", "TEXT NOT NULL DEFAULT ''")):
            try: c.execute(f"ALTER TABLE pos_sales ADD COLUMN {col} {definition}")
            except sqlite3.OperationalError: pass
        c.execute("CREATE TABLE IF NOT EXISTS notifications(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, message TEXT NOT NULL, kind TEXT NOT NULL DEFAULT 'info', ref TEXT NOT NULL DEFAULT '', created_at REAL NOT NULL, read_at REAL)")
        # Additive-only table: existing prices/sources and Google Sheet data remain untouched.
        c.execute("CREATE TABLE IF NOT EXISTS price_history(id INTEGER PRIMARY KEY AUTOINCREMENT, product_key TEXT, display_id TEXT, title TEXT, store TEXT, img TEXT, old_price REAL, new_price REAL, username TEXT, action TEXT, created_at REAL)")
        if not c.execute("SELECT 1 FROM users").fetchone():
            c.execute("INSERT INTO users(username,pw,role) VALUES(?,?,?)", (ADMIN_USER, generate_password_hash(ADMIN_PASS), "super_admin"))
        # Create indexes only after each relevant module table exists (additive, no rows deleted).
        for idx_sql in (
            "CREATE INDEX IF NOT EXISTS idx_erp_purchases_date_supplier ON erp_purchases(purchase_date,supplier_id)",
            "CREATE INDEX IF NOT EXISTS idx_erp_expenses_status_date ON erp_expenses(status,expense_date)",
            "CREATE INDEX IF NOT EXISTS idx_erp_purchase_items_cost_lookup ON erp_purchase_items(item_code,size)",
            "CREATE INDEX IF NOT EXISTS idx_hr_attendance_employee_date ON hr_attendance(employee_id,attendance_date)",
            "CREATE INDEX IF NOT EXISTS idx_erp_supplier_payments_supplier_date ON erp_supplier_payments(supplier_id,payment_date)",
            "CREATE INDEX IF NOT EXISTS idx_pos_sales_phone_created ON pos_sales(phone_digits COLLATE NOCASE,created_at DESC)",
            "CREATE INDEX IF NOT EXISTS idx_order_history_phone_id ON order_history(phone_digits COLLATE NOCASE,id DESC)",
        ):
            try: c.execute(idx_sql)
            except sqlite3.OperationalError: pass
        # বর্তমান মূল Admin account-কে Super Admin করা হলো; পুরোনো users/data delete বা reset হচ্ছে না।
        c.execute("UPDATE users SET role='super_admin' WHERE username=? AND role IN ('admin','user','staff')", (ADMIN_USER,))
        # Guest/Customer permissions additive setting; আগে থেকে থাকলে untouched।
        if not c.execute("SELECT 1 FROM settings WHERE k='guest_perms'").fetchone():
            c.execute("INSERT INTO settings(k,v) VALUES('guest_perms',?)", (json.dumps(GUEST_DEFAULTS, ensure_ascii=False),))
        old = c.execute("SELECT v FROM settings WHERE k='sheet_url'").fetchone()   # পুরনো ডাটাবেস থেকে শিট নিয়ে আসা
        if old and old["v"] and not c.execute("SELECT 1 FROM sources WHERE kind='sheet'").fetchone():
            c.execute("INSERT INTO sources(kind,name,url) VALUES('sheet','Sheet 1',?)", (old["v"],))

def get_setting(k):
    with db() as c: r = c.execute("SELECT v FROM settings WHERE k=?", (k,)).fetchone()
    return r["v"] if r else ""

def set_setting(k, v):
    with db() as c: c.execute("INSERT OR REPLACE INTO settings(k,v) VALUES(?,?)", (k, v))

def sources(kind):
    with db() as c: return [dict(r) for r in c.execute("SELECT * FROM sources WHERE kind=? ORDER BY id", (kind,))]

def current_user():
    u = session.get("user")
    if not u: return None
    with db() as c: return c.execute("SELECT * FROM users WHERE username=?", (u,)).fetchone()

def user_role(u=None):
    u = u or current_user()
    if not u: return ""
    return u["role"]

def user_perms(u=None):
    u = u or current_user()
    if not u: return dict(DEFAULT_PERMS)
    role = u["role"]
    if role in ("admin", "super_admin"): return {k: True for k in PERMS}
    try: p = json.loads(u["perms"] or "{}")
    except Exception: p = {}
    defaults = CS_DEFAULTS if role == "cs" else STAFF_DEFAULTS
    effective = {k: bool(p.get(k, defaults.get(k, False))) for k in PERMS}
    effective["view_stock"] = effective.get("stock_view", False) or effective.get("stock_manage", False)
    return effective

def guest_perms():
    try: p = json.loads(get_setting("guest_perms") or "{}")
    except Exception: p = {}
    return {k: bool(p.get(k, GUEST_DEFAULTS.get(k, False))) for k in GUEST_PERM_LABELS}

def is_admin_level():
    return bool(session.get("user") and session.get("role") in ("admin", "super_admin"))

def can_action(name):
    if session.get("user"):
        return bool(user_perms().get(name))
    return bool(guest_perms().get(name))

def has_perm(name):
    return bool(name in PERMS and user_perms().get(name))

def log_activity(action, query="", details=""):
    u = session.get("user") or "guest"
    with db() as c:
        c.execute("INSERT INTO activity_logs(username,action,query,details,created_at) VALUES(?,?,?,?,?)",
                  (u, str(action)[:60], str(query)[:500], str(details)[:1000], time.time()))

def admin_only(f):
    @wraps(f)
    def w(*a, **k):
        if session.get("user") and session.get("role") in ("admin", "super_admin"): return f(*a, **k)
        return jsonify(error="admin_login"), 401
    return w

def perm_auth(name):
    def deco(f):
        @wraps(f)
        def w(*a, **k):
            if session.get("user") and (session.get("role") in ("admin", "super_admin") or has_perm(name)): return f(*a, **k)
            return jsonify(error="permission_denied"), 403
        return w
    return deco

def stock_dashboard_auth(f):
    @wraps(f)
    def w(*a, **k):
        if session.get("user") and (session.get("role") in ("admin", "super_admin") or
                                     user_perms().get("stock_dashboard") or user_perms().get("stock_manage") or
                                     user_perms().get("stock_history_view")):
            return f(*a, **k)
        return jsonify(error="permission_denied"), 403
    return w

def super_admin_only(f):
    @wraps(f)
    def w(*a, **k):
        if session.get("user") and session.get("role") == "super_admin": return f(*a, **k)
        return jsonify(error="super_admin_required"), 403
    return w

def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if session.get("user"): return f(*a, **k)
        return jsonify(error="login_required"), 401
    return w

# ---------- ডাটা ----------
C = {"products": [], "stock": {}, "bytab": {}, "stk": {}, "errors": [], "tabs": [], "cat_t": 0, "stock_t": 0, "force": False, "sheet_order": {}, "order_version": 0}
P = {"stage": "শুরু", "cat_done": 0, "cat_total": 0, "sig_done": 0, "sig_total": 0, "ready": False, "catalog_ready": False, "ai_ready": False}
_BG = {"catalog": False, "stock": False, "ai": False}
SKIPH = {"total", "মোট", "price", "rate", "name", "product", "title", "sum", "grand total", "amount", "cost"}
STOP = {"full", "half", "sleeve", "shirt", "shart", "cargo", "cargos", "tee", "tees", "the", "and", "for"}
UA = {"User-Agent": "Mozilla/5.0"}

def num(s):
    try: return float(re.sub(r"[^\d.]", "", s or "") or 0)
    except ValueError: return 0.0

def _split_media_urls(value):
    """Catalog media cells can contain comma/semicolon/pipe/newline separated URLs."""
    out=[]
    for x in re.split(r"[,;|\n\r]+", str(value or "")):
        x=x.strip()
        if x.startswith(("http://", "https://")): out.append(x)
    return out

def fetch_catalog(s):
    r = requests.get(s["url"], timeout=40, headers=UA); r.raise_for_status()
    out = []
    for feed_idx, row in enumerate(csv.DictReader(io.StringIO(r.content.decode("utf-8-sig", "ignore")))):
        r_ = {(k or "").strip().lower(): (v or "") for k, v in row.items()}
        price, sale = num(r_.get("price")), num(r_.get("sale_price"))
        pid = r_.get("id") or r_.get("sku", "")
        primary = next(iter(_split_media_urls(r_.get("image_link"))), r_.get("image_link", "").strip())
        media=[]; size_chart_image=""; size_chart_text=""
        for key, value in r_.items():
            nk=re.sub(r"[^a-z0-9]+", "_", key.lower()).strip("_")
            vals=_split_media_urls(value)
            is_media = ("image" in nk) and ("logo" not in nk) and (bool(vals) or "link" in nk or "url" in nk)
            if is_media:
                for u in vals:
                    if u not in media: media.append(u)
            if "size" in nk and ("chart" in nk or "guide" in nk or "measurement" in nk):
                if vals and not size_chart_image:
                    size_chart_image=vals[0]
                elif value.strip() and not vals and not size_chart_text:
                    size_chart_text=value.strip()
        if primary and primary not in media: media.insert(0, primary)
        if not primary and media: primary=media[0]
        extra=[u for u in media if u != primary and u != size_chart_image]
        out.append(dict(
            id=pid, title=r_.get("title") or r_.get("name", ""), store=s["name"],
            price=sale if sale and sale < price else price, old=price if sale and sale < price else 0,
            img=primary, link=r_.get("link") or r_.get("url") or r_.get("product_link", ""),
            desc=r_.get("description", ""), extra=extra,
            size_chart_image=size_chart_image, size_chart_text=size_chart_text, feed_idx=feed_idx
        ))
    return out

def discover_tabs(sid):
    c = C.setdefault("tabc", {})
    if sid in c and time.time() - c[sid][0] < 300: return c[sid][1]
    names = []
    try:
        r = requests.get("https://docs.google.com/spreadsheets/d/%s/htmlview" % sid, timeout=30, headers=UA)
        for n, g in re.findall(r'items\.push\(\{name:\s*"((?:[^"\\]|\\.)*)".*?gid:\s*"(\d+)"', r.text):
            names.append((json.loads('"%s"' % n), g))
    except Exception as e: print("tab discover failed:", e)
    if names or sid not in c: c[sid] = (time.time(), names or [("Sheet1", "0")])
    return c[sid][1]

def read_tab(sid, gid):
    r = requests.get("https://docs.google.com/spreadsheets/d/%s/gviz/tq" % sid, params={"tqx": "out:csv", "gid": gid}, timeout=30)
    r.raise_for_status()
    rows = list(csv.reader(io.StringIO(r.content.decode("utf-8-sig", "ignore"))))
    if not rows: return {}, []
    hdr = [" ".join(h.split()) for h in rows[0]]; d = {}; row_order = []
    for r_ in rows[1:]:
        code = (r_[0] if r_ else "").strip().lower()
        if not code: continue
        # Google Sheet-এর আসল top-to-bottom row order ধরে রাখি।
        row_order.append(code)
        s = d.setdefault(code, {})
        for h, v in zip(hdr[1:], r_[1:]):
            try: n = int(float(v.replace(",", "").strip()))
            except ValueError: continue
            if h and h.lower() not in SKIPH and n > 0: s[h] = s.get(h, 0) + n
    return d, row_order

def match_stock(p, stock):
    """শুধু কোড মিললেই স্টক ধরা হবে (সাধারণ শব্দ দিয়ে আর মেলানো হবে না)।"""
    ks = [p["id"].strip().lower()]
    for part in re.split(r"[-_ /]", p["id"].lower()) + re.findall(r"[a-z0-9]+", p["title"].lower()):
        if len(part) > 2 and re.search(r"\d", part) and re.search(r"[a-z]", part): ks.append(part)
    for k in ks:
        if k in stock: return stock[k]

def excel_cell_text(cell):
    value = cell.value
    if value is None: return ""
    if isinstance(value, (int, float)) and cell.number_format and re.fullmatch(r"0+", cell.number_format):
        return str(int(value)).zfill(len(cell.number_format))
    if hasattr(value, "isoformat"): return value.isoformat()
    return str(value).strip()

def excel_headers(ws):
    return [excel_cell_text(ws.cell(1, col)) for col in range(1, ws.max_column + 1)]

def normalize_stock_size(value):
    """Excel size headings often contain line breaks; use one clean, comparable label."""
    return " ".join(str(value or "").split())

def refresh_stock_from_excel():
    try:
        with EXCEL_LOCK:
            wb = load_workbook(EXCEL_STOCK_FILE, read_only=True, data_only=True)
            bytab, out, tabs = {}, {}, []
            xl_order = {}      # code -> (row নম্বর, tab নম্বর): Excel-এর যত নিচে, ওয়েবে তত উপরে
            tab_no = 0
            for ws in wb.worksheets:
                headers = excel_headers(ws)
                if not headers or not headers[0]: continue
                name = ws.title
                tabs.append(name)
                tab_no += 1
                d = {}
                for rn, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
                    code = excel_cell_text(row[0]).strip().lower() if row else ""
                    if not code: continue
                    _rk = (rn, tab_no)
                    if _rk > xl_order.get(code, (-1, -1)): xl_order[code] = _rk
                    sizes = {}
                    for i, cell in enumerate(row[1:], start=1):
                        h = headers[i] if i < len(headers) else ""
                        if not h or h.lower() in SKIPH: continue
                        try: qty = int(float(cell.value or 0))
                        except (TypeError, ValueError): continue
                        if qty > 0: sizes[h] = sizes.get(h, 0) + qty
                    d[code] = {"sizes": sizes, "tabs": [name]}
                    entry = out.setdefault(code, {"tabs": [], "sizes": {}})
                    if name not in entry["tabs"]: entry["tabs"].append(name)
                    for size, qty in sizes.items(): entry["sizes"][size] = entry["sizes"].get(size, 0) + qty
                bytab[("excel", name)] = d
            wb.close()
        C["bytab"], C["stock"], C["tabs"], C["errors"], C["stock_t"] = bytab, out, tabs, [], time.time()
        apply_stock()
        # Excel-ই একমাত্র সোর্স: Excel-এর সবচেয়ে নিচের code ওয়েবে সবার উপরে, সবচেয়ে উপরের code সবার নিচে
        if xl_order and xl_order != C.get("sheet_order"):
            C["sheet_order"] = xl_order
            C["order_version"] = int(C.get("order_version", 0)) + 1
            _apply_sheet_order()
            _save_catalog_cache(C["products"], xl_order)
        return True
    except Exception as e:
        C["errors"] = ["Excel: " + str(e)]
        print("Excel stock refresh failed:", e)
        return False

_EXCEL_PRODUCT_CACHE={"mtime":None,"products":[]}
def excel_product_catalog_from_stock():
    """Read catalog fields from the Excel stock workbook, when available."""
    if not os.path.isfile(EXCEL_STOCK_FILE): return []
    try:
        mtime=os.path.getmtime(EXCEL_STOCK_FILE)
        if _EXCEL_PRODUCT_CACHE["mtime"]==mtime: return list(_EXCEL_PRODUCT_CACHE["products"])
    except OSError: return []
    output=[]; seen=set()
    aliases={"code":{"code","product code","sku","id","product id","item code","style code"},"title":{"product name","title","name","item name","product title"},"image":{"image","image url","image link","photo","photo url","picture","picture url"},"link":{"link","product link","url","website"},"price":{"price","regular price","sale price","product price","selling price"}}
    try:
        with EXCEL_LOCK:
            wb=load_workbook(EXCEL_STOCK_FILE,read_only=True,data_only=True)
            for ws in wb.worksheets:
                headers=[" ".join(str(x or "").strip().lower().split()) for x in excel_headers(ws)]
                ix={name:next((i for i,h in enumerate(headers) if h in opts),-1) for name,opts in aliases.items()}
                if ix["code"]<0: continue
                for cells in ws.iter_rows(min_row=2,values_only=True):
                    values=["" if v is None else str(v).strip() for v in cells]
                    code=values[ix["code"]] if ix["code"]<len(values) else ""
                    if not code or code.casefold() in seen: continue
                    def get(field): return values[ix[field]] if 0<=ix[field]<len(values) else ""
                    title=get("title") or code
                    raw_price=get("price"); price=num(raw_price)
                    image=get("image"); link=get("link")
                    display=code if re.search(r"[0-9]",code) else ""
                    size_stock={}
                    for i,heading in enumerate(headers[1:],start=1):
                        if i>=len(values) or not heading or any(x in heading for x in ("price","rate","amount","cost","total","name","title","image","photo","link","url")): continue
                        try:
                            qty=int(float(values[i].replace(",","")))
                            if qty>=0: size_stock[headers[i]]=qty
                        except (ValueError,TypeError): pass
                    output.append({"id":code,"display_id":display,"title":title,"store":"Stock Excel","price":price,"old":0,"img":image,"link":link,"desc":"","extra":[],"size_chart_image":"","size_chart_text":"","feed_idx":len(output),"sizes":size_stock,"tabs":[ws.title],"instock":any(v>0 for v in size_stock.values())})
                    seen.add(code.casefold())
            wb.close()
    except Exception as exc: print("Excel catalog fallback failed:",exc)
    _EXCEL_PRODUCT_CACHE.update(mtime=os.path.getmtime(EXCEL_STOCK_FILE),products=list(output))
    return output
def excel_catalog_index():
    index = {}
    for product in view():
        info = {"id": product.get("id", ""), "title": product.get("title", ""), "price": float(product.get("price") or 0),
                "img": product.get("img", ""), "display_id": product.get("display_id", "")}
        for key in (product.get("id", ""), product.get("display_id", ""), display_product_id(product)):
            key = str(key or "").strip().lower()
            if key: index.setdefault(key, info)
    return index

def excel_workbook_data(sheet="__all__", query="", offset=0, limit=100):
    catalog = excel_catalog_index()
    with EXCEL_LOCK:
        wb = load_workbook(EXCEL_STOCK_FILE, read_only=True, data_only=True)
        summaries, all_headers, all_rows = [], [], []
        chosen = None
        for ws in wb.worksheets:
            headers = excel_headers(ws)
            if not headers or not headers[0]: continue
            rows, total, total_value, products = [], 0, 0.0, 0
            for rn, cells in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
                values = [excel_cell_text(c) for c in cells]
                values += [""] * max(0, len(headers)-len(values))
                code = values[0].strip()
                if not code: continue
                products += 1
                row_qty = 0
                for i, h in enumerate(headers[1:], start=1):
                    if not h or h.lower() in SKIPH: continue
                    try:
                        qty = int(float(values[i] or 0))
                        total += qty
                        row_qty += qty
                    except (ValueError, TypeError): pass
                product = catalog.get(code.lower())
                unit_price = float(product.get("price") or 0) if product else 0.0
                line_value = row_qty * unit_price
                total_value += line_value
                if not query or query.lower() in " ".join(values).lower():
                    record = {"row": rn, "values": values, "headers": headers, "sheet": ws.title,
                              "product": product, "total_qty": row_qty, "unit_price": unit_price, "stock_value": line_value}
                    if sheet == "__all__":
                        all_rows.append(record)
                        for header in headers:
                            if header not in all_headers: all_headers.append(header)
                    elif ws.title == sheet: rows.append(record)
            summaries.append({"name": ws.title, "products": products, "total": total, "value": total_value, "headers": headers})
            if ws.title == sheet: chosen = {"name": ws.title, "headers": headers, "rows": rows}
        if sheet == "__all__":
            for record in all_rows:
                indexed = dict(zip(record.pop("headers"), record["values"]))
                record["values"] = [indexed.get(header, "") for header in all_headers]
            chosen = {"name": "__all__", "headers": all_headers, "rows": all_rows}
        wb.close()
    if chosen:
        chosen["total_matches"] = len(chosen["rows"])
        chosen["rows"] = chosen["rows"][offset:offset+limit]
    return {"sheets": summaries, "selected": chosen, "has_file": True, "total_value": sum(x["value"] for x in summaries)}

def snapshot_excel_stock(reason="save", username=""):
    if not os.path.isfile(EXCEL_STOCK_FILE): return ""
    os.makedirs(EXCEL_BACKUP_DIR, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    filename = f"stock-{stamp}-{int(time.time()*1000)%1000:03d}-{reason}.xlsx"
    shutil.copy2(EXCEL_STOCK_FILE, os.path.join(EXCEL_BACKUP_DIR, filename))
    with db() as c:
        c.execute("CREATE TABLE IF NOT EXISTS excel_stock_backups(id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, bytes INTEGER, username TEXT, reason TEXT, created_at REAL)")
        c.execute("INSERT INTO excel_stock_backups(filename,bytes,username,reason,created_at) VALUES(?,?,?,?,?)",
                  (filename, os.path.getsize(os.path.join(EXCEL_BACKUP_DIR, filename)), username, reason, time.time()))
    return filename

def apply_stock():
    has = bool(C["bytab"])
    C["stk"] = {}
    for x in C["products"]:
        m = match_stock(x, C["stock"]) if has else None
        sz = m["sizes"] if m else ({} if has else None)   # শিটে কোড না থাকলে = স্টক নেই
        C["stk"][x["store"] + "|" + x["id"]] = dict(sizes=sz, tabs=m["tabs"] if m else [], instock=bool(sz) if has else None)

def _product_display_id_for_order(p):
    """Sheet-এর code আর catalog-এর UI ID মিলিয়ে newest-first order বের করি।
    SKU দিয়ে order করা হয় না।"""
    pid = str(p.get("id", "") or "").strip().lower()
    title = str(p.get("title", "") or "").strip()
    m = re.match(r"^([A-Za-z0-9]+)(?=\s|[-_/]|$)", title)
    did = m.group(1).strip().lower() if m else ""
    return pid, did

# ---------- নতুন প্রোডাক্ট চেনা (ক্যাটালগে প্রথম কবে দেখা গেছে) ----------
NEW_DAYS = 21          # এত দিন পর্যন্ত Sheet-এ না থাকা নতুন প্রোডাক্ট সবার উপরে থাকবে
_FIRST = {}            # key -> first_seen (0 = আগে থেকেই ছিল)
_FIRST_LOADED = [False]

def _first_seen_load():
    if _FIRST_LOADED[0]: return
    try:
        with db() as c:
            c.execute("CREATE TABLE IF NOT EXISTS product_first_seen(k TEXT PRIMARY KEY, t REAL)")
            for r in c.execute("SELECT k,t FROM product_first_seen"): _FIRST[r["k"]] = float(r["t"] or 0)
    except Exception as e: print("first_seen load failed:", e)
    _FIRST_LOADED[0] = True

def _mark_first_seen(products):
    """নতুন key পেলে এখনকার সময় বসাই। একদম প্রথমবার (টেবিল ফাঁকা) সব প্রোডাক্ট পুরনো (0) ধরা হয়।"""
    _first_seen_load()
    first_run = not _FIRST
    now, new_rows = time.time(), []
    for p in products:
        k = p["store"] + "|" + p["id"]
        if k not in _FIRST:
            _FIRST[k] = 0 if first_run else now
            new_rows.append((k, _FIRST[k]))
    if new_rows:
        try:
            with db() as c: c.executemany("INSERT OR IGNORE INTO product_first_seen(k,t) VALUES(?,?)", new_rows)
        except Exception as e: print("first_seen save failed:", e)
        if not first_run and any(t for _, t in new_rows):
            C["order_version"] = int(C.get("order_version", 0)) + 1

def _sku_num(p):
    m = re.fullmatch(r"(?i)sku[-_ ]?0*(\d+)", str(p.get("id", "") or "").strip())
    return int(m.group(1)) if m else 0

def _max_sku_in_stock(products, order):
    """প্রতিটি স্টোরের জন্য "Excel-এর চেয়েও নতুন" সীমানা বের করি:
    (১) Excel-এ থাকা প্রোডাক্টের সবচেয়ে বড় SKU নম্বর, আর
    (২) ক্যাটালগ ফিডে Excel-এর প্রোডাক্টগুলো কোন দিকে সাজানো (নতুন আগে/পরে) দেখে তার সীমানা।"""
    ms, mt = {}, {}
    for p in products:
        pid, did = _product_display_id_for_order(p)
        r = order.get(pid) if order else None
        if r is None and did and order: r = order.get(did)
        if r is None: continue
        st = p["store"]
        n = _sku_num(p)
        d = ms.setdefault(st, {"sku": 0, "dir": 0, "fr": None})
        if n > d["sku"]: d["sku"] = n
        if p.get("feed_idx") is not None: mt.setdefault(st, []).append(((int(r[0]), int(r[1])), int(p["feed_idx"])))
    for st, lst in mt.items():
        if len(lst) < 8: continue
        lst.sort()
        half = len(lst) // 2
        top = sum(x[1] for x in lst[:half]) / half            # Excel-এর উপরের (পুরনো) অর্ধেক
        bot = sum(x[1] for x in lst[half:]) / (len(lst) - half)  # Excel-এর নিচের (নতুন) অর্ধেক
        idxs = sorted(x[1] for x in lst)
        k = min(4, len(idxs) - 1)
        if bot < top: ms[st]["dir"], ms[st]["fr"] = -1, idxs[k]          # ফিডে নতুন আগে
        elif bot > top: ms[st]["dir"], ms[st]["fr"] = 1, idxs[-1 - k]    # ফিডে নতুন পরে
    return ms

def _order_key(p, order, ms=None):
    """বড় key = উপরে। নতুন (Sheet-এ নেই) > Sheet-এর নিচের row > Sheet-এর উপরের row > বাকি পুরনো।"""
    pid, did = _product_display_id_for_order(p)
    r = order.get(pid) if order else None
    if r is None and did and order: r = order.get(did)
    k = p["store"] + "|" + p["id"]
    t = _FIRST.get(k, 0)
    if r is None:
        n = _sku_num(p)
        if t and time.time() - t < NEW_DAYS * 86400: return (2, t, n)
        d = (ms or {}).get(p["store"])
        if d:
            if n and n > d["sku"]: return (2, 0, n)
            fi = p.get("feed_idx")
            if fi is not None and d["fr"] is not None:
                if d["dir"] < 0 and fi < d["fr"]: return (2, 0, 10**6 - int(fi))
                if d["dir"] > 0 and fi > d["fr"]: return (2, 0, int(fi))
        return (0, -1, -1)
    return (1, int(r[0]), int(r[1]))

def _apply_sheet_order():
    """Google Sheet-এর প্রতিটি tab-এর একদম নিচের/latest row web-এর একদম উপরে আনে।
    Sheet data, stock quantity বা catalog data পরিবর্তন করে না—শুধু display order বদলায়।
    Sheet-এ থাকা product আগে; Sheet-এ না থাকা catalog product পরে থাকবে।"""
    order = C.get("sheet_order") or {}
    if not C.get("products"): return
    _first_seen_load()
    with _CLOCK:
        _ms = _max_sku_in_stock(C["products"], order)
        C["products"] = sorted(C["products"], key=lambda p: _order_key(p, order, _ms), reverse=True)

def refresh_stock():
    if os.path.isfile(EXCEL_STOCK_FILE):
        return refresh_stock_from_excel()
    jobs = []
    for s in sources("sheet"):
        m = re.search(r"/d/([\w-]+)", s["url"])
        if m: jobs += [(m.group(1), n, g, s["name"]) for n, g in discover_tabs(m.group(1))]
    def one(j):
        try: return j, read_tab(j[0], j[2]), None
        except Exception as e: return j, None, str(e)
    with ThreadPoolExecutor(6) as ex: res = list(ex.map(one, jobs))
    seen, errs, row_orders = set(), [], {}
    for j, payload, e in res:
        k = (j[0], j[1]); seen.add(k)
        if payload is not None:
            d, row_order = payload
            C["bytab"][k] = d
            row_orders[k] = row_order
        else:
            errs.append(j[1]); print("tab failed:", j[1], e)
    for k in list(C["bytab"]):
        if k not in seen: del C["bytab"][k]
    out = {}
    # Stock aggregation আগের মতোই থাকবে। শুধু Sheet-এর row position আলাদাভাবে মনে রাখি।
    for k, d in C["bytab"].items():
        sid, n = k
        for code, sz in d.items():
            e = out.setdefault(code, {"tabs": [], "sizes": {}})
            if n not in e["tabs"]: e["tabs"].append(n)
            for h, v in sz.items(): e["sizes"][h] = e["sizes"].get(h, 0) + v

    # প্রতিটি tab-এর একদম নিচের row-কে সবচেয়ে নতুন ধরি।
    # একই ID একাধিকবার থাকলে সেই tab-এর সর্বশেষ row-টাই নেওয়া হবে।
    sheet_order = {}
    for tab_index, (sid, n, gid, sname) in enumerate(jobs):
        rows = row_orders.get((sid, n), [])
        for row_index, code in enumerate(rows, start=1):
            key = str(code).strip().lower()
            rank = (row_index, tab_index)
            if rank > sheet_order.get(key, (-1, -1)):
                sheet_order[key] = rank

    C["stock"], C["tabs"], C["errors"], C["stock_t"] = out, sorted({k[1] for k in seen}), errs, time.time()
    if sheet_order:
        old_order = C.get("sheet_order", {})
        C["sheet_order"] = sheet_order
        _apply_sheet_order()
        if sheet_order != old_order:
            C["order_version"] = int(C.get("order_version", 0)) + 1
        # Order cache-এ রাখা হয় যাতে restart-এর পর live Sheet refresh শেষ হওয়ার আগেই
        # আগের জানা newest-first order দেখা যায়।
        _save_catalog_cache(C["products"], C.get("sheet_order"))
    apply_stock()

# ---------- ইমেজ ফিচার (রঙের গ্রিড + আকৃতির গ্রেডিয়েন্ট + হিউ হিস্টোগ্রাম) ----------
_sigs = {}
def feat(im):
    im = im.convert("RGB"); im.thumbnail((128, 128))
    v = [x / 255 for p in im.resize((6, 6)).getdata() for x in p]          # রঙের গ্রিড
    m = sum(v) / len(v); v = [a - m for a in v]
    px = list(im.convert("L").resize((9, 8)).getdata())                    # আকৃতির গ্রেডিয়েন্ট
    v += [.5 if px[y * 9 + x] > px[y * 9 + x + 1] else -.5 for y in range(8) for x in range(8)]
    gp = list(im.convert("L").resize((32, 32)).getdata()); e = [0] * 64     # এজ ম্যাপ (পকেট, সেলাই, ডিটেইল)
    for y in range(31):
        for x in range(31):
            a = gp[y * 32 + x]; e[(y // 4) * 8 + x // 4] += abs(a - gp[y * 32 + x + 1]) + abs(a - gp[(y + 1) * 32 + x])
    mx, me = max(e) or 1, sum(e) / 64; v += [(a - me) / mx * 1.5 for a in e]
    h = [0.0] * 12
    for r, g, b in im.resize((24, 24)).getdata():
        hh, s, val = rgb_to_hsv(r / 255, g / 255, b / 255)
        if s > .2 and val > .15: h[int(hh * 12) % 12] += s
    t = sum(h) or 1
    v += [2 * x / t - 1 / 6 for x in h]
    n = sum(a * a for a in v) ** .5 or 1
    return [a / n for a in v]

SIGF = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sigs.pkl")
CATF = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalog_cache.pkl")
_CLOCK = threading.Lock()
try:
    _sigs.update({k: v for k, v in pickle.load(open(SIGF, "rb")).items() if isinstance(v, dict)})
except Exception: pass
AI = {"m": None, "err": "", "busy": False, "last_try": 0.0}
_alock = threading.Lock()

def ai_model(force=False):
    """CLIP AI loader — app start-এ background-এ চেষ্টা করে; fail হলে retry করতে পারে."""
    if AI["m"] is not None:
        return AI["m"]
    now = time.time()
    if AI["busy"]:
        return None
    if not force and AI["err"] and now - AI["last_try"] < 30:
        return None
    AI["busy"] = True
    AI["last_try"] = now
    AI["err"] = ""
    try:
        try:
            from fastembed import ImageEmbedding
        except ImportError:
            # Never run a long pip install from the app process. The fast colour/shape
            # image-search index works without CLIP; AI is an optional enhancement.
            AI["err"] = "fastembed is not installed; using standard visual matching"
            print("optional fastembed not installed; standard image search remains enabled")
            return None
        P["stage"] = "AI মডেল প্রস্তুত হচ্ছে…"
        AI["m"] = ImageEmbedding(model_name="Qdrant/clip-ViT-B-32-vision")
        print("AI model ready: Qdrant/clip-ViT-B-32-vision")
    except Exception as e:
        AI["err"] = str(e)[:220] or "fastembed/model লোড হয়নি"
        print("AI model unavailable:", e)
    finally:
        AI["busy"] = False
        P["ai_ready"] = bool(AI["m"])
    return AI["m"]

def _ai_warmup():
    """UI আটকে না দিয়ে startup থেকেই AI package/model প্রস্তুত করতে চেষ্টা করে."""
    while AI["m"] is None:
        ai_model()
        if AI["m"] is not None:
            break
        time.sleep(5)

def emb1(m, im):
    with _alock:
        try: v = next(iter(m.embed([im])))
        except Exception:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as t: im.save(t, "JPEG")
            try: v = next(iter(m.embed([t.name])))
            finally: os.unlink(t.name)
    v = [float(x) for x in v]; n = sum(a * a for a in v) ** .5 or 1
    return [a / n for a in v]

def embed(im):
    """পুরো ছবি + মাঝখানের ক্রপ — জুম/ক্রপ করা ছবিও যেন মিলে যায়।"""
    m = AI["m"]
    if not m: return None, None
    try:
        im = im.convert("RGB"); im.thumbnail((512, 512)); w, h = im.size
        return emb1(m, im), emb1(m, im.crop((int(w * .15), int(h * .15), int(w * .85), int(h * .85))))
    except Exception as e: print("embed failed:", e); return None, None

def sig_url(u):
    if u not in _sigs or not isinstance(_sigs.get(u), dict) or (AI["m"] and _sigs[u] and (_sigs[u].get("e") is None or _sigs[u].get("e2") is None)):
        try:
            r = requests.get(u, timeout=6, headers=UA); r.raise_for_status()
            im = Image.open(io.BytesIO(r.content)); e, e2 = embed(im); _sigs[u] = {"f": feat(im), "e": e, "e2": e2}
        except Exception as ex: print("sig failed:", u, ex); _sigs[u] = None
    return _sigs[u]

def make_q(im):
    e, e2 = embed(im); return {"f": feat(im), "e": e, "e2": e2}

def dot(a, b): return sum(x * y for x, y in zip(a, b))

def score(q, s):
    cf = max(0, dot(q["f"], s["f"]))
    if q["e"]:
        # Candidate images may still be in the fast feature-only index while CLIP
        # embeddings are being built; keep them searchable instead of nearly discarding them.
        if not s.get("e"): return cf
        ff = dot(q["e"], s["e"])
        cross = [dot(x, y) for x, y in ((q["e"], s.get("e2")), (q.get("e2"), s["e"]), (q.get("e2"), s.get("e2"))) if x and y]
        ce = max(ff, .97 * max(cross)) if cross else ff
        return .9 * max(0, (ce - .5) / .5) + .1 * cf
    return cf

def rank(q, exclude=None, n=None):
    d = []
    for x in C["products"]:
        s = _sigs.get(x["img"]); k = x["store"] + "|" + x["id"]
        if s and k != exclude: d.append((score(q, s), k))
    d.sort(reverse=True)
    return [{"key": k, "score": round(min(1, c) * 100, 1)} for c, k in d[:n]]

def _save_catalog_cache(products, sheet_order=None):
    try:
        tmp = CATF + ".tmp"
        with open(tmp, "wb") as f:
            pickle.dump({"t": time.time(), "products": products, "sheet_order": sheet_order or C.get("sheet_order", {}),
                        "order_version": C.get("order_version", 0)}, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, CATF)
    except Exception as e:
        print("catalog cache save failed:", e)

def _load_catalog_cache():
    try:
        with open(CATF, "rb") as f: d = pickle.load(f)
        products = d.get("products") if isinstance(d, dict) else None
        if isinstance(products, list) and products:
            return products, float(d.get("t", 0) or 0), (d.get("sheet_order") or {}), int(d.get("order_version", 0) or 0)
    except Exception as e:
        if os.path.exists(CATF): print("catalog cache load failed:", e)
    return [], 0, {}, 0

def _refresh_catalog_background():
    if _BG["catalog"]: return
    _BG["catalog"] = True
    try:
        srcs = sources("catalog")
        if not srcs:
            C["cat_err"] = []
            P.update(stage="ক্যাটালগ সোর্স নেই", catalog_ready=False, ready=True, cat_done=0, cat_total=0)
            return
        out, errs = [], []
        def one(src):
            try: return fetch_catalog(src), None
            except Exception as e: return [], (src["name"], str(e))
        # প্রতিটি source শেষ হওয়ার সাথে সাথে UI-তে product দেখাই। কোনো এক source ধীর হলে অন্যগুলো আটকে থাকবে না।
        with ThreadPoolExecutor(max_workers=max(1, min(8, len(srcs)))) as ex:
            futures = {ex.submit(one, src): src for src in srcs}
            done_count = 0
            for fut in as_completed(futures):
                items, err = fut.result()
                done_count += 1
                if items:
                    out.extend(items)
                    with _CLOCK:
                        C["products"] = list(out)
                        C["cat_err"] = list(errs)
                    _apply_sheet_order()
                    P.update(catalog_ready=True, ready=True, stage="প্রোডাক্ট লোড হচ্ছে…", cat_done=done_count, cat_total=len(srcs))
                    _save_catalog_cache(C["products"], C.get("sheet_order"))
                if err:
                    errs.append(err[0]); print("catalog failed:", err[0], err[1])
                    with _CLOCK: C["cat_err"] = list(errs)
                else:
                    with _CLOCK: C["cat_err"] = list(errs)
        if out:
            # কোনো সোর্স সাময়ে ফেল করলে তার আগের প্রোডাক্টগুলো যেন হারিয়ে না যায়
            failed = set(errs)
            if failed:
                have = {p["store"] + "|" + p["id"] for p in out}
                out.extend(p for p in list(C.get("products") or []) if p.get("store") in failed and (p["store"] + "|" + p["id"]) not in have)
            _mark_first_seen(out)
            with _CLOCK:
                C["products"] = list(out)
                C["cat_t"] = time.time()
                C["cat_err"] = list(errs)
            _apply_sheet_order()
            _save_catalog_cache(C["products"], C.get("sheet_order"))
            P.update(catalog_ready=True, ready=True, stage="প্রোডাক্ট প্রস্তুত", cat_done=len(srcs), cat_total=len(srcs))
            threading.Thread(target=_refresh_stock_background, daemon=True).start()
        else:
            with _CLOCK: C["cat_err"] = list(errs)
            P.update(catalog_ready=False, ready=True, stage="ক্যাটালগ লোড হয়নি", cat_done=len(srcs), cat_total=len(srcs))
    except Exception as e:
        print("catalog refresh failed:", e)
        P.update(ready=True, stage="ক্যাটালগ লোডে সমস্যা")
    finally:
        _BG["catalog"] = False

def _refresh_stock_background():
    if _BG["stock"]: return
    _BG["stock"] = True
    try:
        P["stage"] = "Syncing stock…"
        refresh_stock()
        P["ready"] = True
        P["catalog_ready"] = bool(C["products"])
        P["stage"] = "প্রোডাক্ট প্রস্তুত"
    except Exception as e:
        print("stock refresh failed:", e)
        P["ready"] = True
    finally:
        _BG["stock"] = False

def _build_ai_background():
    """Index all distinct catalog images off-request; optional CLIP embeddings upgrade the basic visual index."""
    if _BG["ai"]: return
    _BG["ai"] = True
    # Index every distinct catalogue image; the old 120-image cap made products later
    # in the catalogue permanently unsearchable by image.
    urls = list(dict.fromkeys(str(x.get("img") or "").strip() for x in C["products"] if str(x.get("img") or "").strip().startswith(("http://", "https://"))))
    if not urls:
        P.update(ai_ready=bool(AI["m"]), sig_total=0, sig_done=0)
        _BG["ai"] = False
        return
    P.update(sig_total=len(urls), sig_done=0)
    try:
        # Start with cheap colour/shape features; do not wait for package/model installation.
        def warm_feature(u):
            try:
                old = _sigs.get(u)
                if isinstance(old, dict) and old.get("f"):
                    return
                r = requests.get(u, timeout=4, headers=UA); r.raise_for_status()
                im = Image.open(io.BytesIO(r.content)).convert("RGB"); im.thumbnail((512, 512))
                _sigs[u] = {"f": feat(im), "e": None, "e2": None}
            except Exception as e:
                print("lazy image index failed:", str(u)[:100], str(e)[:100])
            finally:
                P["sig_done"] = min(int(P.get("sig_done", 0)) + 1, int(P.get("sig_total", 0)))
        with ThreadPoolExecutor(max_workers=8) as ex:
            list(ex.map(warm_feature, urls))
        try:
            tmp = SIGF + ".tmp"
            with open(tmp, "wb") as fh: pickle.dump(_sigs, fh, protocol=pickle.HIGHEST_PROTOCOL)
            os.replace(tmp, SIGF)
        except Exception as e: print("sigs save failed:", e)
        # Only after the basic index is available do we try the optional model download/load.
        ai_model()
        P["ai_ready"] = bool(AI["m"])
        if AI["m"]:
            def upgrade(u):
                try:
                    old = _sigs.get(u)
                    if isinstance(old, dict) and old.get("f") and old.get("e") is not None and old.get("e2") is not None:
                        return
                    sig_url(u)
                except Exception as e:
                    print("AI signature upgrade failed:", str(u)[:100], str(e)[:100])
            with ThreadPoolExecutor(max_workers=8) as ex:
                list(ex.map(upgrade, urls))
            try:
                tmp = SIGF + ".tmp"
                with open(tmp, "wb") as fh: pickle.dump(_sigs, fh, protocol=pickle.HIGHEST_PROTOCOL)
                os.replace(tmp, SIGF)
            except Exception as e: print("AI signatures save failed:", e)
    except Exception as e:
        print("AI background failed:", e)
    finally:
        P["ai_ready"] = bool(AI["m"])
        P["stage"] = "প্রস্তুত"
        P["ready"] = True
        _BG["ai"] = False

def loader():
    # আগের successful catalog + sheet-order disk-cache থেকে সঙ্গে সঙ্গে দেখাই।
    cached, cache_t, cached_order, cached_order_version = _load_catalog_cache()
    if cached:
        with _CLOCK:
            C["products"] = cached
            C["cat_t"] = cache_t
            C["sheet_order"] = cached_order or {}
            C["order_version"] = cached_order_version
        _mark_first_seen(cached)   # আগের ক্যাশে যা ছিল সেগুলো "পুরনো"; এরপর যা আসবে সেটাই নতুন
        _apply_sheet_order()
        P.update(stage="ক্যাশ থেকে প্রোডাক্ট প্রস্তুত", ready=True, catalog_ready=True,
                 cat_done=1, cat_total=1)
    while True:
        try:
            now = time.time()
            srcs = sources("catalog")
            # cache থাকলেও background-এ fresh catalog আনার চেষ্টা করি; UI অপেক্ষা করবে না।
            if C["force"] or (not C["products"] and srcs) or now - C["cat_t"] > CACHE_TTL:
                C["force"] = False
                P.update(stage="ক্যাটালগ আপডেট হচ্ছে…", cat_total=len(srcs), cat_done=0)
                threading.Thread(target=_refresh_catalog_background, daemon=True).start()
            elif now - C["stock_t"] > STOCK_TTL and C["products"]:
                threading.Thread(target=_refresh_stock_background, daemon=True).start()
        except Exception as e:
            print("loader error:", e); P["ready"] = True
        time.sleep(3)

def pct():
    ct, st = max(P["cat_total"], 1), max(P["sig_total"], 1)
    return 100 if P["ready"] else int(30 * P["cat_done"] / ct + 70 * P["sig_done"] / st)

def prices():
    with db() as c: return {r["k"]: r["price"] for r in c.execute("SELECT * FROM prices")}

def display_product_id(p):
    # UI/copy-র জন্য SKU নয়। Product title-এর শুরুতে থাকা code/ID-টাই ব্যবহার হবে।
    # যেমন: "R208 Grey..." -> R208, "2J266 Whiskered..." -> 2J266
    title = str(p.get("title") or "").strip()
    m = re.match(r"^([A-Za-z0-9]+)(?=\s|[-_/]|$)", title)
    if m and re.search(r"\d", m.group(1)):
        return m.group(1)
    # Title-এ ID না থাকলে SKU-কে কখনোই UI/copy ID হিসেবে দেখাব না।
    return ""

def view():
    pr, out = prices(), []
    # Catalog is the sole source of product cards. Excel supplies stock quantities only;
    # never append Excel-only rows as separate catalog products.
    products=list(C["products"])
    for x in products:
        k = x["store"] + "|" + x["id"]; y = dict(x, key=k, **C["stk"].get(k, {}))
        y["display_id"] = display_product_id(x)
        _g = _GAL.get(k)
        if _g and _g[1]: y["extra"] = list(dict.fromkeys(list(x.get("extra") or []) + _g[1]))
        if k in pr: y.update(price=pr[k], old=x["price"] if x["price"] > pr[k] else 0, edited=True, base=x["price"])
        r = None
        pid, did = _product_display_id_for_order(x)
        if C.get("sheet_order"):
            r = C["sheet_order"].get(pid) or (C["sheet_order"].get(did) if did else None)
        y["sheet_rank"] = list(r) if r is not None else None
        out.append(y)
    # API response-ও newest Sheet row first রাখি, যাতে frontend cache/render
    # বা অন্য কোনো UI layer ভুল করে catalog order-এ ফিরিয়ে না দেয়।
    _o = C.get("sheet_order") or {}
    _ms = _max_sku_in_stock(out, _o)
    out.sort(key=lambda y: _order_key(y, _o, _ms), reverse=True)
    for y in out:
        y["is_new"] = y.get("sheet_rank") is None and _order_key(y, _o, _ms)[0] == 2
    return out

def meta():
    # Product count follows the catalog only; stock workbook is not a product source.
    product_count=len(C["products"])
    return dict(synced=C["stock_t"], tabs=C["tabs"], errors=C["errors"], now=time.time(), pct=pct(), ready=P["ready"], catalog_ready=P["catalog_ready"], stage=P["stage"],
                ai=bool(AI["m"]), ai_ready=P["ai_ready"], ai_busy=AI["busy"], ai_err=AI["err"], products=product_count, has_sheet=bool(sources("sheet")), has_excel_stock=os.path.isfile(EXCEL_STOCK_FILE), cat_err=C.get("cat_err", []), order_version=C.get("order_version", 0))

def auth(f):
    # পুরনো admin-only API-গুলো একইভাবে সুরক্ষিত রাখা হলো।
    return admin_only(f)

# ---------- API ----------
LOGIN_HTML = r"""<!doctype html><html lang="bn"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Alien Tees ERP · লগইন</title><style>*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;background:radial-gradient(ellipse at 15% 15%,#34136d 0,transparent 38%),linear-gradient(135deg,#080910,#171326 58%,#090a11);font:15px system-ui,'Noto Sans Bengali',sans-serif;color:#eef0ff}.login-shell{width:min(940px,100%);min-height:520px;display:grid;grid-template-columns:minmax(340px,.9fr) minmax(0,1.1fr);overflow:hidden;border:1px solid #ffffff20;border-radius:22px;background:#10111beF;box-shadow:0 32px 90px #0009}.login-side{padding:54px 48px;display:flex;flex-direction:column;justify-content:center;background:linear-gradient(155deg,#141622,#0c0d14)}.brand{font-size:12px;font-weight:800;letter-spacing:.18em;color:#c8a45e}.login-side h1{margin:10px 0 4px;font-size:27px;line-height:1.25;color:#fff}.intro{margin:0 0 26px;color:#a4a8bc}.login-side label{display:block;margin:14px 0 7px;color:#d9dbea;font-weight:650}.login-side input{width:100%;height:46px;padding:0 13px;border:1px solid #343747;border-radius:9px;background:#171925;color:#fff;font:inherit;outline:none}.login-side input:focus{border-color:#a548ed;box-shadow:0 0 0 3px #a548ed30}.login-side button{width:100%;height:46px;margin-top:22px;border:0;border-radius:9px;background:linear-gradient(90deg,#7927dc,#ec168f);color:#fff;font:700 15px system-ui,'Noto Sans Bengali',sans-serif;cursor:pointer;box-shadow:0 8px 24px #b024c346}.login-side button:hover{filter:brightness(1.1);transform:translateY(-1px)}.err{min-height:20px;margin-top:10px;color:#ff9b9b;font-size:13px}.help{display:block;margin-top:12px;color:#9da1b5;text-align:center;font-size:12px;line-height:1.6}.welcome{position:relative;isolation:isolate;overflow:hidden;min-height:520px;display:grid;place-items:center;padding:35px;background:radial-gradient(ellipse at 25% 60%,#7b16bb 0,transparent 42%),radial-gradient(ellipse at 83% 45%,#143fbb 0,transparent 44%),linear-gradient(135deg,#171236,#111329 65%,#12101f)}.welcome:before,.welcome:after{content:"";position:absolute;z-index:-1;width:520px;height:220px;left:50%;top:38%;border-radius:50%;filter:blur(18px);transform:translate(-50%,-50%) rotate(-18deg);background:linear-gradient(90deg,transparent,#e329d7 30%,#772cff 52%,#15d8ff 72%,transparent);opacity:.8;animation:wave 8s ease-in-out infinite}.welcome:after{width:460px;height:150px;top:48%;transform:translate(-50%,-50%) rotate(18deg);filter:blur(28px);opacity:.43;animation-delay:-3s}.welcome-copy{max-width:360px;margin:180px 0 0 auto;text-align:right}.welcome-copy b{display:block;font-size:34px;color:#fff;letter-spacing:-.03em}.welcome-copy span{display:block;margin-top:10px;color:#c4c6dc;font-size:13px;line-height:1.7}@keyframes wave{0%,100%{transform:translate(-50%,-50%) rotate(-18deg) scaleX(.9)}50%{transform:translate(-50%,-50%) rotate(-8deg) scaleX(1.12)}}@media(max-width:700px){body{padding:14px}.login-shell{min-height:0;grid-template-columns:1fr}.login-side{padding:34px 26px}.welcome{min-height:190px;order:-1;padding:20px}.welcome-copy{margin:80px 0 0;text-align:center}.welcome-copy b{font-size:25px}.welcome-copy span{font-size:12px}}@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important}}
.courier-head-actions{display:flex;align-items:center;justify-content:flex-end;gap:8px;flex-wrap:wrap}
.courier-head-actions .pill,.courier-head-actions .courier-report-link{white-space:nowrap}
.courier-report-link{text-decoration:none;display:inline-flex;align-items:center;justify-content:center}
</style><style id="login-global-loader-style">#loginGlobalLoader{position:fixed;inset:0;z-index:999999;display:grid;place-items:center;padding:18px;background:rgba(2,17,12,.48);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px)}#loginGlobalLoader[hidden]{display:none!important}.login-loader-card{width:min(340px,calc(100vw - 32px));padding:24px;border-radius:24px;text-align:center;color:#fff;border:1px solid #e2ffed99;background:linear-gradient(145deg,#073123,#0d7552 60%,#082e25);box-shadow:0 25px 80px #0008;animation:loginLoaderPop .28s ease-out}.login-loader-scene{position:relative;width:110px;height:110px;margin:0 auto 12px;display:grid;place-items:center}.login-loader-ring{position:absolute;inset:7px;border:2px solid transparent;border-top-color:#f3ffd7;border-right-color:#6bffc3;border-radius:50%;animation:loginLoaderSpin 1.35s linear infinite}.login-loader-ring.r2{inset:22px;border:1px dashed #ddffed;animation-duration:2.1s;animation-direction:reverse}.login-loader-core{width:38px;height:38px;display:grid;place-items:center;border-radius:13px;color:#075b3a;font-size:24px;background:linear-gradient(145deg,#f4ffe4,#7effbe);box-shadow:0 0 28px #78ffbd99;animation:loginLoaderPulse 1.3s ease-in-out infinite}.login-loader-title{font:800 18px system-ui,'Noto Sans Bengali',sans-serif;margin:4px 0 6px}.login-loader-sub{font:12px/1.65 system-ui,'Noto Sans Bengali',sans-serif;color:#e0ffedcc;margin:0}.login-loader-bar{height:4px;border-radius:99px;overflow:hidden;background:#ffffff30;margin-top:18px}.login-loader-bar i{display:block;width:35%;height:100%;background:linear-gradient(90deg,#d5ffe8,#fff1b7,#e5fff1);animation:loginLoaderRun 1.1s ease-in-out infinite}@keyframes loginLoaderSpin{to{transform:rotate(360deg)}}@keyframes loginLoaderPulse{50%{transform:rotate(8deg) scale(1.08);box-shadow:0 0 40px #78ffbdcc}}@keyframes loginLoaderRun{to{transform:translateX(330%)}}@keyframes loginLoaderPop{from{opacity:0;transform:translateY(8px) scale(.96)}to{opacity:1;transform:none}}@media(prefers-reduced-motion:reduce){.login-loader-card,.login-loader-ring,.login-loader-core,.login-loader-bar i{animation:none!important}}
/* Compact login loader to match the rest of the website. */
#loginGlobalLoader{padding:10px!important;background:rgba(2,17,12,.18)!important;backdrop-filter:blur(1px)!important;-webkit-backdrop-filter:blur(1px)!important}
.login-loader-card{width:auto!important;max-width:calc(100vw - 28px)!important;min-width:0!important;padding:10px 14px!important;border-radius:13px!important;display:grid!important;grid-template-columns:25px minmax(0,1fr)!important;grid-template-rows:auto auto!important;align-items:center!important;column-gap:9px!important;row-gap:2px!important;text-align:left!important;box-shadow:0 7px 24px rgba(0,0,0,.22)!important;animation:loginLoaderPop .18s ease-out!important}
.login-loader-scene{grid-column:1!important;grid-row:1/3!important;width:25px!important;height:25px!important;margin:0!important}
.login-loader-ring{inset:2px!important;border-width:2px!important;animation-duration:.8s!important}
.login-loader-ring.r2,.login-loader-core,.login-loader-sub,.login-loader-bar{display:none!important}
.login-loader-title{grid-column:2!important;grid-row:1/3!important;margin:0!important;font:600 12px/1.35 system-ui,'Noto Sans Bengali',sans-serif!important;letter-spacing:0!important}
@media(prefers-reduced-motion:reduce){.login-loader-card,.login-loader-ring{animation:none!important}}
</style><div id="loginGlobalLoader" role="status" aria-live="polite" aria-busy="true"><div class="login-loader-card"><div class="login-loader-scene"><span class="login-loader-ring"></span><span class="login-loader-ring r2"></span><b class="login-loader-core">✦</b></div><h3 id="loginLoaderTitle" class="login-loader-title">সিস্টেম লোড হচ্ছে</h3><p id="loginLoaderSub" class="login-loader-sub">আপনার অনুরোধ প্রক্রিয়া করা হচ্ছে…</p><div class="login-loader-bar"><i></i></div></div></div><form class="login-shell" id="f"><section class="login-side"><div class="brand">ALIEN TEES · BODYFITTERS</div><h1>ERP Management System</h1><p class="intro">আপনার অ্যাকাউন্টে লগইন করুন</p><label for="u">ইউজারনেম</label><input id="u" name="username" autocomplete="username" required><label for="p">পাসওয়ার্ড</label><input id="p" name="password" type="password" autocomplete="current-password" required><button>লগইন করুন</button><div class="err" id="e" role="status"></div><small class="help">ইউজার পাসওয়ার্ড পেতে IT এডমিন এর সাথে যোগাযোগ করুন</small></section><aside class="welcome" aria-hidden="true"><div class="welcome-copy"><b>স্বাগতম।</b><span>Alien Tees ERP Management System<br>আপনার কাজ, এক জায়গায়।</span></div></aside></form><script>const ll=document.getElementById('loginGlobalLoader');let loginNavigating=false;window.addEventListener('load',()=>{if(!loginNavigating){ll.hidden=true}},{once:true});document.getElementById('f').onsubmit=async e=>{e.preventDefault();const btn=e.target.querySelector('button');const f=new FormData(e.target);ll.hidden=false;document.getElementById('loginLoaderTitle').textContent='লগইন হচ্ছে';document.getElementById('loginLoaderSub').textContent='নিরাপদে আপনার অ্যাকাউন্ট যাচাই করা হচ্ছে…';if(btn)btn.disabled=true;let ok=false;try{const r=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:f.get('username'),password:f.get('password')})});const j=await r.json();if(j.ok){ok=true;loginNavigating=true;document.getElementById('loginLoaderTitle').textContent='লগইন সফল';document.getElementById('loginLoaderSub').textContent='ERP ড্যাশবোর্ড খুলছে…';location.replace('/')}else document.getElementById('e').textContent=j.error||'ইউজারনেম বা পাসওয়ার্ড সঠিক নয়'}catch(err){document.getElementById('e').textContent='সার্ভারের সঙ্গে সংযোগ করা যায়নি। আবার চেষ্টা করুন।'}finally{if(!ok){ll.hidden=true;if(btn)btn.disabled=false}}}</script></html>"""

@app.before_request
def enforce_nine_hour_login():
    if session.get("user"):
        try: started=float(session.get("login_at", 0))
        except (TypeError, ValueError): started=0
        if not started or time.time()-started >= 9*60*60:
            session.clear()
            if request.path.startswith("/api/"): return jsonify(error="session_expired"),401
    if request.path.startswith("/api/") and request.path!="/api/login" and not session.get("user"):
        return jsonify(error="login_required"),401
    return None

@app.get("/")
def home():
    if not session.get("user"):
        return Response(LOGIN_HTML, content_type="text/html; charset=utf-8")
    return Response(HTML.replace("Alien Tees × Bodyfitters", "Alien Tees ERP Management System").replace("ALIEN STOCK WORKSPACE", "ALIEN ERP WORKSPACE").replace("Alien Tees Stock Center", "Alien Tees ERP Management System"), content_type="text/html; charset=utf-8")

@app.get("/api/products")
def products():
    if not session.get("user"): return jsonify(error="login_required"), 401
    if not can_action("view_product"): return jsonify(error="permission_denied"), 403
    items = view()
    if not can_action("view_stock"):
        for p in items:
            p.pop("sizes", None); p.pop("tabs", None); p.pop("instock", None)
    return jsonify(products=items, meta=meta())

@app.get("/api/stock")
def stock_api():
    if not session.get("user"): return jsonify(error="login_required"), 401
    if not can_action("view_stock"):
        r = jsonify(items={}, meta=meta()); r.headers["Cache-Control"] = "no-store"; return r
    r = jsonify(items={k: v for k, v in C["stk"].items()}, meta=meta()); r.headers["Cache-Control"] = "no-store"; return r

def normalize_order_phone(value, cell=None):
    """Keep phone numbers as text and restore common leading zero when Excel stored numeric values."""
    if value is None: return ""
    text = str(value).strip()
    if not text: return ""
    # Preserve formatted Excel values such as 017XXXXXXXX where the cell format contains zeros.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        digits = str(int(value)) if float(value).is_integer() else text
        fmt = str(getattr(cell, "number_format", "") or "") if cell is not None else ""
        zero_pattern = re.search(r"0{8,}", fmt)
        if zero_pattern and len(digits) < len(zero_pattern.group(0)):
            digits = digits.zfill(len(zero_pattern.group(0)))
        # Bangladesh mobile numbers are commonly 11 digits beginning with 01.
        if len(digits) == 10 and digits[0] in "123456789": digits = "0" + digits
        return digits
    return text

@app.post("/api/admin/order-history/upload")
@perm_auth("order_sheet_upload")
def order_history_upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="শুধু .xlsx/.xlsm Excel ফাইল দিন"), 400
    raw = f.read()
    if not raw or len(raw) > 25 * 1024 * 1024:
        return jsonify(error="ফাইল খালি বা 25 MB-এর বেশি"), 400
    import hashlib, uuid
    file_hash = hashlib.sha256(raw).hexdigest()
    with db() as c:
        if c.execute("SELECT 1 FROM excel_stock_imports WHERE sha256=? AND filename LIKE 'ORDER HISTORY:%' LIMIT 1", (file_hash,)).fetchone():
            return jsonify(error="এই একই Excel ফাইল আগে Order History-তে আপলোড হয়েছে; ডুপ্লিকেট রেকর্ড যোগ করা হয়নি"), 409
    try:
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        headers = [re.sub(r"[^a-z0-9]+", " ", excel_cell_text(ws.cell(1, i)).replace("৳", " ").strip().lower()).strip() for i in range(1, ws.max_column + 1)]
        def find_col(*names):
            for n in names:
                for i,h in enumerate(headers):
                    if h == n or n in h: return i
            return -1
        # Header aliases match the uploaded order-report sheet.
        cols = {
          "order_date": find_col("date", "order date"), "source": find_col("source"), "assigned_to": find_col("assigned to", "assigned"),
          "payment_status": find_col("payment status", "payment"), "order_status": find_col("order status", "order status"),
          "delivery_status": find_col("delivery status", "delivery status"), "cons_id": find_col("cons id", "consignment id", "consignment"),
          "address": find_col("address"), "phone": find_col("phone", "mobile", "number"), "customer_name": find_col("customer name", "customer", "name"),
          "product_code": find_col("product code", "code", "sku"), "size": find_col("size"), "product_price": find_col("product price", "price"),
          "delivery_charge": find_col("delivery charge", "delivery"), "advance_paid": find_col("advance paid", "advance"),
          "collectible_amount": find_col("collectable amount", "collectible amount", "collectible price", "collectable price", "collectible amount bdt", "collectable amount bdt", "amount to collect", "cod amount", "cash on delivery", "amount payable", "total due", "collectable", "collectible", "baki", "due amount"), "rating": find_col("rating"), "note": find_col("note", "notes")
        }
        if cols["product_code"] < 0 or cols["customer_name"] < 0 or cols["phone"] < 0:
            wb.close(); return jsonify(error="Excel header-এ Customer Name, Phone এবং Product Code থাকতে হবে"), 400
        batch_id = uuid.uuid4().hex
        product_map = {}
        for prod in C.get("products", []):
            for key in (prod.get("id"), prod.get("code"), prod.get("display_id"), prod.get("sku")):
                if key: product_map[str(key).strip().lower()] = prod
        inserted = 0
        with db() as c:
            for rownum, cells in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                vals = ["" if v is None else (v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v).strip()) for v in cells]
                if not any(vals): continue
                def val(k):
                    ix=cols[k]; return vals[ix] if 0 <= ix < len(vals) else ""
                phone_ix = cols["phone"]
                phone_value = normalize_order_phone(cells[phone_ix], ws.cell(rownum, phone_ix+1)) if 0 <= phone_ix < len(cells) else ""
                code=val("product_code")
                prod=product_map.get(code.lower(), {})
                img=str(prod.get("img") or prod.get("image") or "")
                payload={headers[i] if i < len(headers) and headers[i] else f"column_{i+1}": v for i,v in enumerate(vals)}
                c.execute("""INSERT INTO order_history(batch_id,source_file,row_number,order_date,source,assigned_to,payment_status,order_status,delivery_status,cons_id,address,phone,customer_name,product_code,size,product_price,delivery_charge,advance_paid,collectible_amount,rating,note,image_url,row_json,uploaded_by,uploaded_at,phone_digits)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (batch_id, os.path.basename(f.filename), rownum, val("order_date"), val("source"), val("assigned_to"), val("payment_status"), val("order_status"), val("delivery_status"), val("cons_id"), val("address"), phone_value, val("customer_name"), code, val("size"), val("product_price"), val("delivery_charge"), val("advance_paid"), val("collectible_amount"), val("rating"), val("note"), img, json.dumps(payload, ensure_ascii=False), session.get("user",""), time.time(), re.sub(r"\D", "", str(phone_value or ""))))
                inserted += 1
            c.execute("INSERT INTO excel_stock_imports(filename,sha256,bytes,username,created_at) VALUES(?,?,?,?,?)", ("ORDER HISTORY: "+os.path.basename(f.filename), file_hash, len(raw), session.get("user"), time.time()))
        wb.close()
        log_activity("order_history_import", f.filename, f"{inserted} rows; batch {batch_id}")
        return jsonify(ok=True, inserted=inserted, batch_id=batch_id)
    except Exception as e:
        return jsonify(error="Order Excel পড়া যায়নি: "+str(e)), 400

@app.get("/api/order-lookup")
@login_required
def order_lookup():
    """Quick lookup for order count by consignment ID or customer phone number."""
    if not is_admin_level() and not (has_perm("order_view") or has_perm("pos_orders")):
        return jsonify(error="আপনার Order Search permission নেই"), 403
    kind = (request.args.get("type") or "consignment").strip().lower()
    q = (request.args.get("q") or "").strip()[:120]
    if kind not in ("consignment", "number"):
        return jsonify(error="Search type সঠিক নয়"), 400
    if len(q) < 2:
        return jsonify(ok=True, query=q, kind=kind, total=0, rows=[])
    if kind == "consignment":
        where = "TRIM(COALESCE(cons_id,'')) LIKE ?"
        args = ["%" + q + "%"]
    else:
        digits = re.sub(r"\D", "", q)
        if len(digits) < 3:
            return jsonify(ok=True, query=q, kind=kind, total=0, rows=[])
        # phone values may be stored with +880, spaces, or dashes.
        where = "(phone LIKE ? OR REPLACE(REPLACE(REPLACE(REPLACE(phone,' ',''),'-',''),'+',''),'.','') LIKE ?)"
        args = ["%" + q + "%", "%" + digits + "%"]
    with db() as c:
        total_rows = c.execute("SELECT COUNT(*) n FROM order_history WHERE " + where, args).fetchone()["n"]
        # A customer may have multiple product rows for one consignment; count each consignment only once.
        total_orders = c.execute(
            "SELECT COUNT(DISTINCT CASE WHEN TRIM(COALESCE(cons_id,''))<>'' THEN TRIM(cons_id) ELSE 'row:'||id END) n FROM order_history WHERE " + where,
            args
        ).fetchone()["n"]
        rows = [dict(r) for r in c.execute(
            "SELECT id,order_date,assigned_to,order_status,delivery_status,cons_id,phone,customer_name,product_code,size,product_price,delivery_charge,advance_paid,collectible_amount,row_json FROM order_history WHERE " + where + " ORDER BY id DESC LIMIT 100", args
        ).fetchall()]
    def _amount_number(value):
        try:
            cleaned = re.sub(r"[^0-9.\-]", "", str(value or "").replace(",", ""))
            return float(cleaned) if cleaned and cleaned not in (".", "-", "-.") else 0.0
        except Exception:
            return 0.0
    def _norm_amount_key(value):
        return re.sub(r"[^a-z0-9]+", " ", str(value or "").strip().lower()).strip()
    amount_keys = ("collectible amount", "collectable amount", "collectible amount bdt", "collectable amount bdt",
                   "amount to collect", "cod amount", "cash on delivery", "amount payable", "collectible", "collectable", "baki", "due amount", "total due")
    for row in rows:
        sheet_amount = str(row.get("collectible_amount") or "").strip()
        if not sheet_amount:
            try:
                payload = json.loads(row.get("row_json") or "{}")
            except Exception:
                payload = {}
            if isinstance(payload, dict):
                for alias in amount_keys:
                    for key, value in payload.items():
                        nk = _norm_amount_key(key)
                        if (nk == alias or alias in nk) and str(value or "").strip():
                            sheet_amount = str(value).strip()
                            break
                    if sheet_amount:
                        break
        if sheet_amount:
            row["collectible_amount"] = sheet_amount
            row["collectible_source"] = "sheet"
        else:
            fallback = max(0.0, _amount_number(row.get("product_price")) + _amount_number(row.get("delivery_charge")) - _amount_number(row.get("advance_paid")))
            row["collectible_amount"] = str(round(fallback, 2))
            row["collectible_source"] = "calculated"
        row.pop("row_json", None)
    log_activity("order_lookup", q, f"{kind}: {total_orders} matching orders, {total_rows} rows")
    return jsonify(ok=True, query=q, kind=kind, total=total_orders, total_rows=total_rows, rows=rows)

@app.get("/api/admin/order-history")
@login_required
def order_history_search():
    if not is_admin_level() and not has_perm("order_view"): return jsonify(error="permission_denied"), 403
    q=request.args.get("q","").strip(); start=request.args.get("from","").strip(); end=request.args.get("to","").strip()
    page=max(1, min(100000, int(request.args.get("page", "1") or 1))); limit=100
    where=[]; args=[]
    if q:
        like="%"+q+"%"; digits=re.sub(r"\\D", "", q); where.append("(customer_name LIKE ? OR phone LIKE ? OR product_code LIKE ? OR cons_id LIKE ? OR address LIKE ? OR order_date LIKE ? OR (?<>'' AND phone LIKE ?))"); args += [like]*6+[digits,"%"+digits+"%"]
    if start: where.append("order_date >= ?"); args.append(start)
    if end: where.append("order_date <= ?"); args.append(end)
    clause=(" WHERE "+" AND ".join(where)) if where else ""
    with db() as c:
        total=c.execute("SELECT COUNT(*) n FROM order_history"+clause,args).fetchone()["n"]
        rows=[dict(r) for r in c.execute("SELECT * FROM order_history"+clause+" ORDER BY id DESC LIMIT ? OFFSET ?",args+[limit,(page-1)*limit]).fetchall()]
    return jsonify(ok=True,total=total,page=page,pages=max(1,(total+limit-1)//limit),rows=rows)

@app.get("/api/admin/order-history/summary")
@login_required
def order_history_summary():
    if not is_admin_level() and not has_perm("order_view"): return jsonify(error="permission_denied"), 403
    q=request.args.get("q","").strip(); start=request.args.get("from","").strip(); end=request.args.get("to","").strip()
    where=[]; args=[]
    if q:
        like="%"+q+"%"; digits=re.sub(r"\\D", "", q); where.append("(customer_name LIKE ? OR phone LIKE ? OR product_code LIKE ? OR cons_id LIKE ? OR address LIKE ? OR order_date LIKE ? OR (?<>'' AND phone LIKE ?))"); args += [like]*6+[digits,"%"+digits+"%"]
    if start: where.append("order_date >= ?"); args.append(start)
    if end: where.append("order_date <= ?"); args.append(end)
    clause=(" WHERE "+" AND ".join(where)) if where else ""
    with db() as c:
        records=[dict(r) for r in c.execute("SELECT assigned_to,product_code,collectible_amount,product_price,delivery_charge,advance_paid FROM order_history"+clause,args).fetchall()]
    def amount(v):
        try:
            txt=re.sub(r"[^0-9.\-]", "", str(v or "").replace(",", ""))
            return float(txt) if txt and txt not in (".","-","-.") else 0.0
        except Exception: return 0.0
    assigned={}; codes={}; total=0.0; total_bill=0.0
    for r in records:
        amt=amount(r.get("collectible_amount"));
        if not str(r.get("collectible_amount") or "").strip(): amt=max(0, amount(r.get("product_price"))+amount(r.get("delivery_charge"))-amount(r.get("advance_paid")))
        total += amt
        total_bill += amount(r.get("product_price")) + amount(r.get("delivery_charge"))
        a=str(r.get("assigned_to") or "").strip() or "Unassigned"
        k=str(r.get("product_code") or "").strip() or ""
        assigned.setdefault(a,{"name":a,"orders":0,"collectible":0.0}); assigned[a]["orders"]+=1; assigned[a]["collectible"]+=amt
        if k:
            codes.setdefault(k,{"code":k,"orders":0,"collectible":0.0}); codes[k]["orders"]+=1; codes[k]["collectible"]+=amt
    return jsonify(ok=True,total_orders=len(records),collectible_total=round(total,2),total_bill=round(total_bill,2),assigned=sorted(assigned.values(),key=lambda x:(-x["orders"],x["name"])),codes=sorted(codes.values(),key=lambda x:(-x["orders"],x["code"])))

@app.delete("/api/admin/order-history")
@super_admin_only
def order_history_delete():
    data=request.get_json(silent=True) or {}
    ids=data.get("ids") or []
    delete_all=bool(data.get("all"))
    with db() as c:
        if delete_all:
            count=c.execute("SELECT COUNT(*) FROM order_history").fetchone()[0]
            c.execute("DELETE FROM order_history")
        else:
            clean=[]
            for value in ids:
                try: clean.append(int(value))
                except (TypeError,ValueError): pass
            clean=list(dict.fromkeys(clean))
            if not clean: return jsonify(error="কোনো অর্ডার সিলেক্ট করা হয়নি"),400
            placeholders=",".join("?" for _ in clean)
            count=c.execute(f"SELECT COUNT(*) FROM order_history WHERE id IN ({placeholders})",clean).fetchone()[0]
            c.execute(f"DELETE FROM order_history WHERE id IN ({placeholders})",clean)
    log_activity("order_history_delete", "all" if delete_all else ",".join(map(str,ids)), f"deleted {count} order records")
    return jsonify(ok=True,deleted=count)

@app.post("/api/admin/return-history/upload")
@perm_auth("return_sheet_upload")
def return_history_upload():
    f=request.files.get("file")
    if not f or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="শুধু .xlsx/.xlsm Excel ফাইল দিন"),400
    raw=f.read()
    if not raw or len(raw)>25*1024*1024: return jsonify(error="ফাইল খালি বা 25 MB-এর বেশি"),400
    import hashlib, uuid
    digest=hashlib.sha256(raw).hexdigest()
    with db() as c:
        if c.execute("SELECT 1 FROM excel_stock_imports WHERE sha256=? AND filename LIKE 'RETURN HISTORY:%' LIMIT 1",(digest,)).fetchone(): return jsonify(error="এই Return Excel আগে আপলোড হয়েছে"),409
    try:
        wb=load_workbook(io.BytesIO(raw),read_only=True,data_only=True); ws=wb.active
        headers=[re.sub(r"[^a-z0-9]+"," ",excel_cell_text(ws.cell(1,i)).strip().lower()).strip() for i in range(1,ws.max_column+1)]
        def find_col(*names):
            for n in names:
                nn=re.sub(r"[^a-z0-9]+"," ",n.lower()).strip()
                for i,h in enumerate(headers):
                    if h==nn or nn in h: return i
            return -1
        cols={"assigned_to":find_col("assigned to","assigned"),"cons_id":find_col("cons id","cons. id","consignment id","consignment"),"address":find_col("address"),"phone":find_col("phone","mobile"),"customer_name":find_col("customer name","customer","name"),"product_code":find_col("product code","code","sku"),"size":find_col("size"),"collected_amount":find_col("collected amount","collected amount bdt","collection amount"),"note":find_col("note","notes"),"pathao_bill_status":find_col("pathao bill status","pathao status"),"return_reason":find_col("what s the reason of return","what's the reason of return","reason of return","return reason","reason")}
        if cols["product_code"]<0 or cols["cons_id"]<0: wb.close(); return jsonify(error="Return sheet-এ Cons. ID এবং Product Code কলাম থাকা প্রয়োজন"),400
        batch=uuid.uuid4().hex; count=0
        with db() as c:
            for rownum,cells in enumerate(ws.iter_rows(min_row=2,values_only=True),start=2):
                vals=["" if v is None else (v.strftime("%Y-%m-%d") if hasattr(v,"strftime") else str(v).strip()) for v in cells]
                if not any(vals): continue
                def val(k):
                    i=cols[k]; return vals[i] if 0<=i<len(vals) else ""
                phone_i=cols["phone"]; phone=normalize_order_phone(cells[phone_i]) if 0<=phone_i<len(cells) else ""
                payload={headers[i] if headers[i] else f"column_{i+1}":v for i,v in enumerate(vals)}
                c.execute("INSERT INTO return_history(batch_id,source_file,row_number,assigned_to,cons_id,address,phone,customer_name,product_code,size,collected_amount,note,pathao_bill_status,return_reason,row_json,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(batch,os.path.basename(f.filename),rownum,val("assigned_to"),val("cons_id"),val("address"),phone,val("customer_name"),val("product_code"),val("size"),val("collected_amount"),val("note"),val("pathao_bill_status"),val("return_reason"),json.dumps(payload,ensure_ascii=False),session.get("user",""),time.time()))
                count+=1
            c.execute("INSERT INTO excel_stock_imports(filename,sha256,bytes,username,created_at) VALUES(?,?,?,?,?)",("RETURN HISTORY: "+os.path.basename(f.filename),digest,len(raw),session.get("user"),time.time()))
        wb.close(); log_activity("return_history_import",f.filename,f"{count} rows; batch {batch}")
        return jsonify(ok=True,inserted=count,batch_id=batch)
    except Exception as e: return jsonify(error="Return Excel পড়া যায়নি: "+str(e)),400

@app.post("/api/admin/return-history/manual")
@login_required
def return_history_manual_add():
    if not is_admin_level() and not has_perm("return_sheet_upload"):
        return jsonify(error="permission_denied"),403
    j=request.get_json(silent=True) or {}
    cons=str(j.get("cons_id") or "").strip()[:120]
    code=str(j.get("product_code") or "").strip()[:120]
    if not cons or not code:
        return jsonify(error="Consignment ID এবং Product Code অবশ্যই দিন"),400
    try:
        now=time.time(); batch="manual:"+uuid.uuid4().hex
        with db() as c:
            order=c.execute("SELECT * FROM order_history WHERE TRIM(cons_id)=TRIM(?) ORDER BY id DESC LIMIT 1",(cons,)).fetchone()
            assigned=str(j.get("assigned_to") or (order["assigned_to"] if order else "") or "").strip()
            payload={"Source":"Manual Return Entry","Consignment ID":cons,"Product Code":code,"Return Reason":str(j.get("return_reason") or "").strip()}
            cur=c.execute("INSERT INTO return_history(batch_id,source_file,row_number,assigned_to,cons_id,address,phone,customer_name,product_code,size,collected_amount,note,pathao_bill_status,return_reason,row_json,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
              (batch,"Manual Entry",0,assigned,cons,str(j.get("address") or (order["address"] if order else "") or ""),str(j.get("phone") or (order["phone"] if order else "") or ""),str(j.get("customer_name") or (order["customer_name"] if order else "") or ""),code,str(j.get("size") or (order["size"] if order else "") or ""),str(j.get("collected_amount") or ""),str(j.get("note") or ""),str(j.get("pathao_bill_status") or "Returned"),str(j.get("return_reason") or "Manual return entry"),json.dumps(payload,ensure_ascii=False),session.get("user",""),now))
            rid=cur.lastrowid
        log_activity("return_history_manual_add",cons,json.dumps({"id":rid,"product_code":code},ensure_ascii=False))
        return jsonify(ok=True,id=rid)
    except Exception as e:
        app.logger.exception("Manual return entry failed")
        return jsonify(ok=False,error="Return entry could not be saved: "+str(e)[:300]),500

@app.get("/api/admin/return-history")
@login_required
def return_history_search():
    if not is_admin_level() and not has_perm("return_view"): return jsonify(error="permission_denied"), 403
    q=request.args.get("q","").strip(); where=[]; args=[]
    if q:
        like="%"+q+"%"; where.append("(customer_name LIKE ? OR phone LIKE ? OR product_code LIKE ? OR cons_id LIKE ? OR assigned_to LIKE ? OR return_reason LIKE ? OR pathao_bill_status LIKE ?)"); args += [like]*7
    clause=(" WHERE "+" AND ".join(where)) if where else ""
    with db() as c:
        total=c.execute("SELECT COUNT(*) n FROM return_history"+clause,args).fetchone()["n"]
        rows=[dict(r) for r in c.execute("SELECT * FROM return_history"+clause+" ORDER BY id DESC LIMIT 500",args).fetchall()]
        summary=[dict(r) for r in c.execute("SELECT assigned_to,product_code,pathao_bill_status,return_reason FROM return_history"+clause,args).fetchall()]
        orders=c.execute("SELECT COUNT(*) n FROM order_history").fetchone()["n"]
    assigned={}; reasons={}; statuses={}
    for r in summary:
        a=r.get("assigned_to") or "Unassigned"; assigned[a]=assigned.get(a,0)+1
        k=r.get("return_reason") or "কারণ উল্লেখ নেই"; reasons[k]=reasons.get(k,0)+1
        st=r.get("pathao_bill_status") or "স্ট্যাটাস নেই"; statuses[st]=statuses.get(st,0)+1
    return jsonify(ok=True,total=total,rows=rows,order_total=orders,assigned=sorted([{"name":k,"returns":v} for k,v in assigned.items()],key=lambda x:-x["returns"]),reasons=sorted([{"name":k,"count":v} for k,v in reasons.items()],key=lambda x:-x["count"]),statuses=sorted([{"name":k,"count":v} for k,v in statuses.items()],key=lambda x:-x["count"]))

@app.post("/api/admin/return-history/sync-courier")
@login_required
def return_history_sync_courier():
    """Check recent Order Sheet consignments against configured courier APIs.
    A Returned result is automatically written to Return History by _save_courier_status.
    This is an explicit batch action to avoid unbounded API traffic on every page load.
    """
    if not is_admin_level() and not has_perm("courier_update"):
        return jsonify(error="Courier API sync permission নেই"),403
    j=request.get_json(silent=True) or {}
    wanted=str(j.get("courier") or "all").strip()
    try: settings=json.loads(get_setting("courier_settings") or "{}")
    except Exception: settings={}
    enabled=[]
    for name,cfg in settings.items():
        if wanted not in ("", "all") and name != wanted: continue
        if not cfg.get("enabled") or not cfg.get("base_url") or not cfg.get("api_key"): continue
        auth_type=cfg.get("auth_type","steadfast")
        if auth_type=="steadfast" and not cfg.get("secret_key"): continue
        with db() as c:
            mapping=c.execute("SELECT status_path FROM courier_order_mappings WHERE courier=?",(name,)).fetchone()
        if mapping: enabled.append((name,cfg,mapping["status_path"] or "/status_by_cid/{cons_id}"))
    if not enabled:
        return jsonify(error="Courier Settings-এ API Enable করে API Key/Secret এবং Status Path সেভ করুন"),400
    # Process a bounded batch of most recent consignments; repeated syncs are safe.
    with db() as c:
        orders=[dict(r) for r in c.execute("SELECT * FROM order_history WHERE TRIM(COALESCE(cons_id,''))<>'' ORDER BY id DESC LIMIT 100").fetchall()]
    checked=0; returned=0; errors=[]; results=[]
    for courier,cfg,path in enabled:
        auth_type=cfg.get("auth_type","steadfast")
        headers={"Content-Type":"application/json"}
        if auth_type=="steadfast": headers.update({"Api-Key":cfg.get("api_key",""),"Secret-Key":cfg.get("secret_key","")})
        elif auth_type=="bearer": headers["Authorization"]="Bearer "+cfg.get("api_key","")
        elif auth_type=="api-key": headers["Api-Key"]=cfg.get("api_key","")
        for order in orders:
            cons=str(order.get("cons_id") or "").strip()
            if not cons: continue
            try:
                endpoint=cfg.get("base_url","").rstrip("/")+"/"+str(path).lstrip("/")
                endpoint=endpoint.replace("{cons_id}",requests.utils.quote(cons,safe=""))
                response=requests.get(endpoint,headers=headers,timeout=8)
                response.raise_for_status(); body=response.json()
                if isinstance(body,dict) and isinstance(body.get("data"),dict): body=body["data"]
                if not isinstance(body,dict): continue
                status=str(body.get(cfg.get("status_field","delivery_status")) or body.get("delivery_status") or body.get("status") or body.get("tracking_status") or "").strip()
                if not status: continue
                details=str(body.get("tracking_message") or body.get("message") or "Courier API auto sync")[:1000]
                amount_fields=("collected_amount","received_amount","cod_amount","collectable_amount","collectible_amount","total_amount","amount")
                api_amount=None; api_amount_field=""
                for field in amount_fields:
                    if body.get(field) not in (None,""):
                        try: api_amount=float(re.sub(r"[^0-9.-]","",str(body[field]).replace(",",""))); api_amount_field=field; break
                        except (TypeError,ValueError): pass
                before=str(order.get("delivery_status") or "")
                _save_courier_status(cons,courier,status,str(body.get("rating") or ""),details,"","api",api_amount,api_amount_field)
                checked+=1
                if "return" in status.lower() and "return" not in before.lower(): returned+=1
            except Exception as exc:
                if len(errors)<5: errors.append(f"{courier} / {cons}: {str(exc)[:120]}")
    log_activity("return_history_api_sync",wanted,json.dumps({"checked":checked,"returned":returned,"errors":len(errors)},ensure_ascii=False))
    return jsonify(ok=True,checked=checked,returned=returned,orders_considered=len(orders),couriers=[x[0] for x in enabled],errors=errors)

@app.get("/api/admin/return-history/export.xlsx")
@login_required
def return_history_export():
    if not is_admin_level() and not has_perm("return_view"): return jsonify(error="permission_denied"),403
    with db() as c: rows=[dict(r) for r in c.execute("SELECT * FROM return_history ORDER BY id ASC")]
    wb=Workbook(); ws=wb.active; ws.title="Return History"
    ws.append(["ID","Assigned To","Consignment ID","Address","Phone","Customer","Product Code","Size","Collected Amount","Note","Courier Status","Return Reason","Uploaded By","Uploaded At"])
    for r in rows:
        ws.append([r.get("id"),r.get("assigned_to"),r.get("cons_id"),r.get("address"),r.get("phone"),r.get("customer_name"),r.get("product_code"),r.get("size"),r.get("collected_amount"),r.get("note"),r.get("pathao_bill_status"),r.get("return_reason"),r.get("uploaded_by"),datetime.fromtimestamp(r["uploaded_at"],DHAKA_TZ).strftime("%Y-%m-%d %H:%M:%S") if r.get("uploaded_at") else ""])
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        letter=col[0].column_letter; ws.column_dimensions[letter].width=min(34,max(13,max((len(str(cell.value or "")) for cell in col),default=12)+2))
    output=io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name="Alien-Tees-Return-Report.xlsx")

@app.delete("/api/admin/return-history")
@super_admin_only
def return_history_delete():
    data=request.get_json(silent=True) or {}; ids=data.get("ids") or []; delete_all=bool(data.get("all"))
    with db() as c:
        if delete_all:
            count=c.execute("SELECT COUNT(*) FROM return_history").fetchone()[0]; c.execute("DELETE FROM return_history")
        else:
            clean=[]
            for value in ids:
                try: clean.append(int(value))
                except (TypeError,ValueError): pass
            clean=list(dict.fromkeys(clean))
            if not clean: return jsonify(error="কোনো রিটার্ন সিলেক্ট করা হয়নি"),400
            marks=",".join("?" for _ in clean); count=c.execute(f"SELECT COUNT(*) FROM return_history WHERE id IN ({marks})",clean).fetchone()[0]; c.execute(f"DELETE FROM return_history WHERE id IN ({marks})",clean)
    log_activity("return_history_delete", "all" if delete_all else ",".join(map(str,ids)), f"deleted {count} return records")
    return jsonify(ok=True,deleted=count)

@app.get("/api/admin/order-history/<int:order_id>/comments")
@login_required
def order_comments_get(order_id):
    if not is_admin_level() and not has_perm("order_view"): return jsonify(error="permission_denied"), 403
    with db() as c:
        if not c.execute("SELECT 1 FROM order_history WHERE id=?",(order_id,)).fetchone(): return jsonify(error="অর্ডার পাওয়া যায়নি"),404
        rows=[dict(r) for r in c.execute("SELECT id,message,username,created_at FROM order_history_comments WHERE order_id=? ORDER BY id DESC",(order_id,)).fetchall()]
    return jsonify(ok=True,comments=rows)

@app.post("/api/admin/order-history/<int:order_id>/comments")
@login_required
def order_comments_add(order_id):
    if not is_admin_level() and not has_perm("order_sheet_upload"): return jsonify(error="permission_denied"), 403
    data=request.get_json(silent=True) or {}; msg=str(data.get("message") or "").strip()
    if not msg or len(msg)>2000: return jsonify(error="কমেন্ট ১–২০০০ অক্ষরের হতে হবে"),400
    with db() as c:
        if not c.execute("SELECT 1 FROM order_history WHERE id=?",(order_id,)).fetchone(): return jsonify(error="অর্ডার পাওয়া যায়নি"),404
        c.execute("INSERT INTO order_history_comments(order_id,message,username,created_at) VALUES(?,?,?,?)",(order_id,msg,session.get("user",""),time.time()))
    log_activity("order_comment",str(order_id),msg[:120]); return jsonify(ok=True)

@app.get("/api/admin/order-history/export")
@login_required
def order_history_export():
    if not is_admin_level() and not has_perm("order_view"): return jsonify(error="permission_denied"), 403
    q=request.args.get("q","").strip(); start=request.args.get("from","").strip(); end=request.args.get("to","").strip()
    where=[]; args=[]
    if q:
        like="%"+q+"%"; where.append("(customer_name LIKE ? OR phone LIKE ? OR product_code LIKE ? OR cons_id LIKE ? OR address LIKE ? OR order_date LIKE ?)"); args += [like]*6
    if start: where.append("order_date >= ?"); args.append(start)
    if end: where.append("order_date <= ?"); args.append(end)
    clause=(" WHERE "+" AND ".join(where)) if where else ""
    with db() as c:
        records=[dict(r) for r in c.execute("SELECT * FROM order_history"+clause+" ORDER BY id ASC",args).fetchall()]
    wb=Workbook(); ws=wb.active; ws.title="Final Order Sheet"
    # Reuse the original uploaded sheet's header sequence saved in row_json.
    source_headers=[]
    for rec in records:
        try:
            payload=json.loads(rec.get("row_json") or "{}")
            if payload:
                source_headers=list(payload.keys()); break
        except Exception: pass
    if not source_headers:
        source_headers=["Date","Source","Assigned To","Payment Status","Order Status","Delivery Status","Cons ID","Address","Phone","Customer Name","Product Code","Size","Product Price","Delivery Charge","Advance Paid","Collectable Amount","Rating","Note"]
    ws.append(source_headers)
    map_fields={"date":"order_date","order date":"order_date","source":"source","assigned to":"assigned_to","payment status":"payment_status","order status":"order_status","delivery status":"delivery_status","cons id":"cons_id","consignment id":"cons_id","address":"address","phone":"phone","mobile":"phone","number":"phone","customer name":"customer_name","product code":"product_code","size":"size","product price":"product_price","delivery charge":"delivery_charge","advance paid":"advance_paid","collectable amount":"collectible_amount","collectable amount bdt":"collectible_amount","collectible":"collectible_amount","collectable":"collectible_amount","rating":"rating","note":"note","notes":"note"}
    def order_export_values(rec):
        try: payload=json.loads(rec.get("row_json") or "{}")
        except Exception: payload={}
        out=[]
        for h in source_headers:
            key=str(h).strip().lower(); value=payload.get(h,payload.get(key,None)); field=map_fields.get(key)
            if field in ("delivery_status","rating","cons_id") and rec.get(field) is not None: value=rec.get(field)
            elif value is None and field: value=rec.get(field,"")
            if key in ("phone","mobile","number","phone number","mobile number"):
                value=normalize_order_phone(value if value not in (None,"") else rec.get("phone",""))
            out.append("" if value is None else str(value))
        return out
    all_rows=[]; status_groups={"Delivered":[],"Partial Delivered":[],"Returned":[],"Cancelled":[],"In Way":[],"Pending":[],"Other":[]}; status_counts={}
    def order_status_group(value):
        s=re.sub(r"[_-]+"," ",str(value or "").strip().lower())
        if not s: return "Other"
        if any(x in s for x in ("return","returned","ফেরত")): return "Returned"
        if any(x in s for x in ("cancel","canceled","বাতিল")): return "Cancelled"
        if any(x in s for x in ("in way","in transit","transit","out for delivery","pickup pending","dispatched","shipped","hub")): return "In Way"
        if "partial delivered" in s: return "Partial Delivered"
        if "deliver" in s or "ডেলিভার" in s: return "Delivered"
        if any(x in s for x in ("pending","in review","hold","unknown","অপেক্ষা")): return "Pending"
        return "Other"
    for rec in records:
        values=order_export_values(rec); all_rows.append(values)
        state=str(rec.get("delivery_status") or "No status").strip() or "No status"
        group=order_status_group(state); status_groups[group].append(values); status_counts[state]=status_counts.get(state,0)+1
    for values in all_rows: ws.append(values)
    for group,rows_for_status in status_groups.items():
        state_ws=wb.create_sheet(group)
        state_ws.append(source_headers)
        for values in rows_for_status: state_ws.append(values)
    summary_ws=wb.create_sheet("Status Summary")
    summary_ws.append(["Delivery Status","Orders"])
    for state,count in sorted(status_counts.items(),key=lambda x:(-x[1],x[0].lower())): summary_ws.append([state,count])
    from openpyxl.styles import Font, PatternFill, Alignment
    for sheet in wb.worksheets:
        sheet.freeze_panes="A2"; sheet.auto_filter.ref=sheet.dimensions
        for cell in sheet[1]:
            cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="164A3A"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        sheet.row_dimensions[1].height=30
        if sheet is not summary_ws:
            for col_idx,h in enumerate(source_headers,1):
                if str(h).strip().lower() in ("phone","mobile","number","phone number","mobile number"):
                    for row_idx in range(2,sheet.max_row+1):
                        cell=sheet.cell(row_idx,col_idx); cell.number_format="@"; cell.value=str(cell.value or "")
        for col in sheet.columns:
            letter=col[0].column_letter; maxlen=max((len(str(cell.value or "")) for cell in col[:250]),default=10)
            sheet.column_dimensions[letter].width=min(max(maxlen+2,12),36)
    phone_headers={"phone","mobile","number","phone number","mobile number"}
    for sheet in wb.worksheets:
        if sheet is summary_ws: continue
        for col_idx,h in enumerate(source_headers,1):
            if str(h).strip().lower() in phone_headers:
                for row_idx in range(2,sheet.max_row+1):
                    cell=sheet.cell(row_idx,col_idx); cell.number_format="@"; cell.value=str(cell.value or "")
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    suffix="_filtered" if (q or start or end) else ""
    return send_file(buf,as_attachment=True,download_name=f"order_history{suffix}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post("/api/admin/excel-stock/upload")
@super_admin_only
def excel_stock_upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        return jsonify(error="শুধু .xlsx ফাইল আপলোড করুন"), 400
    raw = f.read()
    if not raw or len(raw) > 25 * 1024 * 1024: return jsonify(error="ফাইল খালি বা 25 MB-এর বেশি"), 400
    tmp = None
    try:
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        valid = any(ws.max_row >= 1 and ws.max_column >= 2 and excel_cell_text(ws.cell(1, 1)) for ws in wb.worksheets)
        wb.close()
        if not valid: return jsonify(error="প্রথম সারিতে কোড ও সাইজের হেডারসহ একটি শিট থাকতে হবে"), 400
        os.makedirs(INSTANCE, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix="stock-", suffix=".xlsx", dir=INSTANCE)
        with os.fdopen(fd, "wb") as out: out.write(raw)
        with EXCEL_LOCK:
            had_existing = os.path.isfile(EXCEL_STOCK_FILE)
            if had_existing: snapshot_excel_stock("before-upload", session.get("user", ""))
            os.replace(tmp, EXCEL_STOCK_FILE)
            if not had_existing: snapshot_excel_stock("initial-upload", session.get("user", ""))
        import hashlib
        with db() as c:
            c.execute("INSERT INTO excel_stock_imports(filename,sha256,bytes,username,created_at) VALUES(?,?,?,?,?)",
                      (os.path.basename(f.filename), hashlib.sha256(raw).hexdigest(), len(raw), session.get("user"), time.time()))
        refresh_stock_from_excel()
        log_activity("excel_stock_upload", f.filename, f"{len(raw)} bytes")
        return jsonify(ok=True, data=excel_workbook_data())
    except Exception as e:
        if tmp and os.path.exists(tmp): os.remove(tmp)
        return jsonify(error="Excel ফাইল পড়া যায়নি: " + str(e)), 400

@app.get("/api/excel-stock/data")
@stock_dashboard_auth
def excel_stock_data():
    if not os.path.isfile(EXCEL_STOCK_FILE): return jsonify(has_file=False, sheets=[], selected=None)
    try:
        data = excel_workbook_data(request.args.get("sheet", "__all__"), request.args.get("q", ""),
                                   max(0, int(request.args.get("offset", 0))), min(250, max(1, int(request.args.get("limit", 100)))))
        with db() as c:
            counts = c.execute("SELECT lower(sheet) sheet_key, lower(product_code) code_key, COUNT(*) n FROM excel_stock_comments GROUP BY lower(sheet), lower(product_code)").fetchall()
        comment_counts = {(r["sheet_key"], r["code_key"]): r["n"] for r in counts}
        for row in (data.get("selected") or {}).get("rows", []):
            row["comment_count"] = comment_counts.get((str(row.get("sheet") or "").lower(), str((row.get("values") or [""])[0]).strip().lower()), 0)
        admin = session.get("role") in ("admin", "super_admin")
        perms = user_perms()
        can_view_qty = admin or perms.get("stock_view") or perms.get("stock_manage")
        can_view_value = admin or perms.get("stock_value_view")
        can_view_history = admin or perms.get("stock_history_view")
        if not can_view_qty:
            headers = (data.get("selected") or {}).get("headers", [])
            for row in (data.get("selected") or {}).get("rows", []):
                for i, header in enumerate(headers[1:], start=1):
                    if header and header.casefold() not in SKIPH and i < len(row["values"]): row["values"][i] = ""
                row["total_qty"] = None
            for summary in data.get("sheets", []): summary["total"] = None
        if not can_view_value:
            data["total_value"] = None
            for summary in data.get("sheets", []): summary["value"] = None
            for row in (data.get("selected") or {}).get("rows", []):
                row["unit_price"] = None; row["stock_value"] = None
                if row.get("product"): row["product"].pop("price", None)
                headers = (data.get("selected") or {}).get("headers", [])
                for i, header in enumerate(headers):
                    label = (header or "").casefold()
                    if any(word in label for word in ("price", "rate", "amount", "cost", "দাম", "মূল্য", "টাকা")) and i < len(row["values"]):
                        row["values"][i] = ""
        today = datetime.now(DHAKA_TZ).date()
        start = datetime.combine(today, datetime_time.min, tzinfo=DHAKA_TZ).timestamp()
        with db() as c:
            edits = c.execute("SELECT COUNT(*) n, COALESCE(SUM(CASE WHEN new_qty>old_qty THEN new_qty-old_qty ELSE 0 END),0) plus_units, COALESCE(SUM(CASE WHEN old_qty>new_qty THEN old_qty-new_qty ELSE 0 END),0) minus_units FROM excel_stock_history WHERE created_at>=?", (start,)).fetchone()
            saves = c.execute("SELECT COUNT(*) FROM excel_stock_batches WHERE created_at>=?", (start,)).fetchone()[0]
        # Dashboard permission grants only aggregate today's cards; row-by-row history stays separately gated.
        can_view_today = admin or perms.get("stock_dashboard") or perms.get("stock_manage")
        data["today"] = ({"saves": saves, "changes": edits["n"], "plus": edits["plus_units"], "minus": edits["minus_units"]}
                          if can_view_history or can_view_today else None)
        return jsonify(data)
    except Exception as e: return jsonify(error=str(e)), 400

@app.post("/api/excel-stock/update")
@app.post("/api/excel-stock/update-batch")
@stock_dashboard_auth
def excel_stock_update():
    j = request.get_json(silent=True) or {}
    changes = j.get("changes") if isinstance(j.get("changes"), list) else [j]
    if not changes or len(changes) > 1000: return jsonify(error="একবারে 1 থেকে 1000টি পরিবর্তন সেভ করা যাবে"), 400
    note = str(j.get("note", ""))[:500]
    tmp = None
    try:
        with EXCEL_LOCK:
            wb = load_workbook(EXCEL_STOCK_FILE)
            edits, changed = [], False
            for change in changes:
                sheet, code, size = str(change.get("sheet", "")), str(change.get("code", "")).strip(), str(change.get("size", "")).strip()
                try: qty = int(change.get("quantity"))
                except (TypeError, ValueError): wb.close(); return jsonify(error="স্টকের পরিমাণ পূর্ণসংখ্যা হতে হবে"), 400
                if qty < 0 or qty > 10000000: wb.close(); return jsonify(error="পরিমাণ 0 থেকে 10,000,000-এর মধ্যে দিন"), 400
                if not sheet or not code or not size or sheet not in wb.sheetnames:
                    wb.close(); return jsonify(error="শিট, কোড বা সাইজ পাওয়া যায়নি"), 404
                ws = wb[sheet]
                headers = excel_headers(ws)
                if size not in headers[1:]: wb.close(); return jsonify(error="সাইজ কলাম পাওয়া যায়নি"), 404
                row = int(change.get("row", 0))
                if row < 2 or row > ws.max_row or excel_cell_text(ws.cell(row, 1)).strip() != code:
                    wb.close(); return jsonify(error="সারি/প্রোডাক্ট কোড মেলেনি; রিফ্রেশ করে আবার চেষ্টা করুন"), 409
                cell = ws.cell(row, headers.index(size) + 1)
                try: old = int(float(cell.value or 0))
                except (TypeError, ValueError): old = 0
                if old == qty: continue
                if session.get("role") not in ("admin", "super_admin"):
                    allowed_stock = user_perms()
                    if not allowed_stock.get("stock_view") and not (allowed_stock.get("stock_set") or allowed_stock.get("stock_manage")):
                        wb.close(); return jsonify(error="পরিমাণ পরিবর্তনের জন্য আগে stock view permission দরকার"), 403
                    if not (allowed_stock.get("stock_set") or allowed_stock.get("stock_manage")) and not (
                        (qty > old and allowed_stock.get("stock_add")) or
                        (qty < old and allowed_stock.get("stock_reduce"))
                    ):
                        wb.close(); return jsonify(error="এই পরিবর্তনের জন্য আপনার স্টক অনুমতি নেই"), 403
                cell.value = qty
                edits.append((sheet, row, code, size, old, qty, session.get("user"), note, time.time()))
                changed = True
            if not changed:
                wb.close(); return jsonify(ok=True, unchanged=True, saved=0)
            fd, tmp = tempfile.mkstemp(prefix="stock-save-", suffix=".xlsx", dir=INSTANCE)
            os.close(fd)
            wb.save(tmp); wb.close()
            snapshot_excel_stock("before-edit", session.get("user", ""))
            os.replace(tmp, EXCEL_STOCK_FILE)
        with db() as c:
            c.executemany("INSERT INTO excel_stock_history(sheet,row_number,product_code,size,old_qty,new_qty,username,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)", edits)
            plus_units = sum(max(0, new_qty-old_qty) for _,_,_,_,old_qty,new_qty,_,_,_ in edits)
            minus_units = sum(max(0, old_qty-new_qty) for _,_,_,_,old_qty,new_qty,_,_,_ in edits)
            c.execute("INSERT INTO excel_stock_batches(username,change_count,plus_units,minus_units,created_at) VALUES(?,?,?,?,?)",
                      (session.get("user"), len(edits), plus_units, minus_units, time.time()))
        refresh_stock_from_excel()
        log_activity("excel_stock_edit", str(len(edits)) + " cells", "ওয়েব থেকে Excel workbook-এ স্টক সেভ")
        return jsonify(ok=True, saved=len(edits))
    except Exception as e:
        if tmp and os.path.exists(tmp): os.remove(tmp)
        return jsonify(error="পরিবর্তন সেভ করা যায়নি: " + str(e)), 400

@app.get("/api/excel-stock/history")
@stock_dashboard_auth
def excel_stock_history():
    if session.get("role") not in ("admin", "super_admin") and not user_perms().get("stock_history_view"):
        return jsonify(error="history_permission_denied"), 403
    with db() as c:
        edits = [dict(r) for r in c.execute("SELECT * FROM excel_stock_history ORDER BY id DESC LIMIT 1000")]
        comments = [dict(r) for r in c.execute("SELECT id,sheet,product_code,size,sold_qty,message,username,created_at FROM excel_stock_comments ORDER BY id DESC LIMIT 1000")]
        imports = [dict(r) for r in c.execute("SELECT * FROM excel_stock_imports ORDER BY id DESC LIMIT 100")]
        backups = [dict(r) for r in c.execute("SELECT * FROM excel_stock_backups ORDER BY id DESC LIMIT 100")] if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='excel_stock_backups'").fetchone() else []
    return jsonify(edits=edits, comments=comments, imports=imports, backups=backups)

@app.route("/api/excel-stock/comments", methods=["GET", "POST"])
def excel_stock_comments():
    if request.method == "GET":
        # Guest/Customer view gets no internal staff sales notes or size targets.
        if not session.get("user"): return jsonify(comments=[], targets=[])
        code = str(request.args.get("code", "")).strip()[:120]
        if not code: return jsonify(comments=[], targets=[])
        targets, aliases = [], {code.casefold()}
        catalog_index = excel_catalog_index()
        catalog_product = catalog_index.get(code.casefold())
        if catalog_product:
            aliases.update(str(catalog_product.get(k) or "").strip().casefold() for k in ("id", "display_id"))
        aliases.discard("")
        if os.path.isfile(EXCEL_STOCK_FILE):
            with EXCEL_LOCK:
                wb = load_workbook(EXCEL_STOCK_FILE, read_only=True, data_only=True)
                for ws in wb.worksheets:
                    headers = excel_headers(ws)
                    if not headers: continue
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        sheet_code = excel_cell_text(row[0]).strip() if row else ""
                        if not sheet_code: continue
                        sheet_info = catalog_index.get(sheet_code.casefold())
                        same_product = sheet_code.casefold() in aliases or (catalog_product and sheet_info and
                                      ((catalog_product.get("id") and sheet_info.get("id") == catalog_product.get("id")) or
                                       (catalog_product.get("display_id") and sheet_info.get("display_id") == catalog_product.get("display_id"))))
                        if same_product:
                            aliases.add(sheet_code.casefold())
                            for col, h in enumerate(headers[1:], start=1):
                                if not h or h.casefold() in SKIPH: continue
                                try: qty = float(row[col] or 0) if col < len(row) else 0
                                except (TypeError, ValueError): qty = 0
                                targets.append({"sheet": ws.title, "size": normalize_stock_size(h), "header": h, "code": sheet_code, "qty": int(qty)})
                wb.close()
        with db() as c:
            placeholders = ",".join("?" for _ in aliases)
            rows = [dict(r) for r in c.execute(
                f"SELECT id,sheet,product_code,size,sold_qty,message,username,created_at FROM excel_stock_comments WHERE product_code COLLATE NOCASE IN ({placeholders}) ORDER BY id DESC LIMIT 200",
                tuple(aliases))] if aliases else []
        can_delete = session.get("role") in ("admin", "super_admin")
        for row in rows: row["can_delete"] = can_delete
        return jsonify(comments=rows, targets=targets)
    if not session.get("user"): return jsonify(error="login_required"), 401
    if session.get("role") not in ("admin", "super_admin") and not user_perms().get("stock_comment"):
        return jsonify(error="permission_denied"), 403
    j = request.get_json(silent=True) or {}
    sheet, code, size = str(j.get("sheet", "")).strip()[:120], str(j.get("product_code", "")).strip()[:120], str(j.get("size", "")).strip()[:80]
    try: sold_qty = int(j.get("sold_qty", 0))
    except (TypeError, ValueError): return jsonify(error="বিক্রির পরিমাণ পূর্ণসংখ্যা দিন"), 400
    message = str(j.get("message", "")).strip()[:500]
    if not sheet or not code or not size or sold_qty < 1 or sold_qty > 10000000:
        return jsonify(error="প্রোডাক্ট, ট্যাব, সাইজ এবং বিক্রির সঠিক পরিমাণ দিন"), 400
    if not os.path.isfile(EXCEL_STOCK_FILE): return jsonify(error="স্টক workbook পাওয়া যায়নি"), 404
    try:
        with EXCEL_LOCK:
            wb = load_workbook(EXCEL_STOCK_FILE, read_only=True, data_only=True)
            valid = False
            if sheet in wb.sheetnames:
                ws = wb[sheet]; headers = excel_headers(ws)
                matched_header = next((h for h in headers[1:] if normalize_stock_size(h).casefold() == normalize_stock_size(size).casefold()), None)
                valid = matched_header is not None and any(
                    excel_cell_text(row[0]).strip().casefold() == code.casefold()
                    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True))
            wb.close()
        fallback_catalog = excel_catalog_index()
        known_product = code.casefold() in fallback_catalog
        if not valid and not (sheet == "স্টক তালিকায় নেই" and known_product):
            return jsonify(error="সাইজটি workbook-এর size heading-এর সঙ্গে মেলেনি। সাইজটি M 38 / L 40 ফরম্যাটে লিখুন।"), 404
        if not valid: size = normalize_stock_size(size)
        now = time.time()
        with db() as c:
            cur = c.execute("INSERT INTO excel_stock_comments(sheet,product_code,size,sold_qty,message,username,created_at) VALUES(?,?,?,?,?,?,?)",
                            (sheet, code, size, sold_qty, message, session.get("user"), now))
            row = dict(c.execute("SELECT id,sheet,product_code,size,sold_qty,message,username,created_at FROM excel_stock_comments WHERE id=?", (cur.lastrowid,)).fetchone())
        log_activity("stock_sold_comment", code, f"{sheet} · {size} · {sold_qty} পিস")
        row["can_delete"] = False
        return jsonify(ok=True, comment=row)
    except Exception as e:
        return jsonify(error="কমেন্ট সেভ হয়নি: " + str(e)), 400

@app.delete("/api/admin/excel-stock/comments/<int:comment_id>")
def excel_stock_comment_delete(comment_id):
    if not session.get("user"): return jsonify(error="login_required"), 401
    if session.get("role") not in ("admin", "super_admin"):
        return jsonify(error="শুধু অ্যাডমিন কমেন্ট মুছতে পারবেন"), 403
    with db() as c: cur = c.execute("DELETE FROM excel_stock_comments WHERE id=?", (comment_id,))
    if not cur.rowcount: return jsonify(error="কমেন্ট পাওয়া যায়নি"), 404
    log_activity("stock_sold_comment_delete", str(comment_id), "অ্যাডমিন বিক্রির কমেন্ট মুছেছেন")
    return jsonify(ok=True)

@app.get("/api/excel-stock/comment-latest")
def excel_stock_comment_latest():
    """লগইন করা সবাই: প্রতিটি প্রোডাক্টের সর্বশেষ বিক্রির কমেন্ট (কার্ডের নিচে দেখানোর জন্য)।"""
    if not session.get("user"): return jsonify(map={})
    out = {}
    with db() as c:
        for r in c.execute("SELECT product_code,size,sold_qty,message,username,created_at,id FROM excel_stock_comments ORDER BY id ASC"):
            k = str(r["product_code"]).strip().casefold()
            o = out.setdefault(k, {"n": 0, "last": None})
            o["n"] += 1
            o["last"] = {"size": r["size"], "sold_qty": r["sold_qty"], "message": r["message"], "username": r["username"], "created_at": r["created_at"]}
    return jsonify(map=out)

@app.get("/api/excel-stock/comment-summary")
@stock_dashboard_auth
def excel_stock_comment_summary():
    """স্টক ম্যানেজমেন্টে দেখানোর জন্য: কয়টি প্রোডাক্টে কমেন্ট হয়েছে + সাম্প্রতিক কমেন্ট।"""
    with db() as c:
        groups = [dict(r) for r in c.execute(
            "SELECT product_code, COUNT(*) AS n, COALESCE(SUM(sold_qty),0) AS sold, MAX(created_at) AS last_at "
            "FROM excel_stock_comments GROUP BY lower(product_code) ORDER BY last_at DESC")]
        recent = [dict(r) for r in c.execute(
            "SELECT id,sheet,product_code,size,sold_qty,message,username,created_at FROM excel_stock_comments ORDER BY id DESC LIMIT 60")]
    return jsonify(products=groups, comments=recent,
                   total_products=len(groups), total_comments=sum(g["n"] for g in groups),
                   total_sold=sum(g["sold"] for g in groups),
                   can_delete=session.get("role") in ("admin", "super_admin"))

@app.delete("/api/admin/excel-stock/history")
@super_admin_only
def excel_stock_history_clear():
    with db() as c:
        c.execute("DELETE FROM excel_stock_history")
        c.execute("DELETE FROM excel_stock_batches")
        c.execute("DELETE FROM excel_stock_imports")
        c.execute("DELETE FROM excel_stock_comments")
        if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='excel_stock_backups'").fetchone():
            c.execute("DELETE FROM excel_stock_backups")
    if os.path.isdir(EXCEL_BACKUP_DIR):
        for filename in os.listdir(EXCEL_BACKUP_DIR):
            if filename.lower().endswith(".xlsx"):
                try: os.remove(os.path.join(EXCEL_BACKUP_DIR, filename))
                except OSError: pass
    log_activity("stock_history_clear", "", "Super Admin stock history ও sales comments clear করেছেন")
    return jsonify(ok=True)

@app.get("/api/admin/excel-stock/download")
@admin_only
def excel_stock_download():
    if not os.path.isfile(EXCEL_STOCK_FILE): return jsonify(error="কোনো workbook আপলোড করা হয়নি"), 404
    return send_file(EXCEL_STOCK_FILE, as_attachment=True, download_name="stock_master.xlsx")

@app.get("/api/admin/excel-stock/backup/<filename>")
@admin_only
def excel_stock_backup_download(filename):
    safe = os.path.basename(filename)
    path = os.path.join(EXCEL_BACKUP_DIR, safe)
    if safe != filename or not os.path.isfile(path): return jsonify(error="ব্যাকআপ পাওয়া যায়নি"), 404
    return send_file(path, as_attachment=True, download_name=safe)

@app.get("/api/admin/excel-stock/latest-backup")
@admin_only
def excel_stock_latest_backup():
    files = [os.path.join(EXCEL_BACKUP_DIR, n) for n in os.listdir(EXCEL_BACKUP_DIR) if n.lower().endswith(".xlsx")] if os.path.isdir(EXCEL_BACKUP_DIR) else []
    if not files: return jsonify(error="এখনো কোনো ব্যাকআপ নেই"), 404
    path = max(files, key=os.path.getmtime)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

@app.get("/api/admin/full-backup")
@admin_only
def full_system_backup():
    """এক ফাইলে অ্যাপের database/config, workbook, আগের stock snapshots, এবং এই app source।"""
    os.makedirs(FULL_BACKUP_DIR, exist_ok=True)
    scope = request.args.get("scope", "all").lower()
    if scope not in {"all", "database", "stock", "history", "users-settings"}:
        return jsonify(error="অজানা ব্যাকআপ ধরন"), 400
    stamp = time.strftime("%Y%m%d-%H%M%S")
    labels = {"all":"পূর্ণ", "database":"ডেটাবেস", "stock":"স্টক", "history":"হিস্ট্রি", "users-settings":"ইউজার-সেটিংস"}
    path = os.path.join(FULL_BACKUP_DIR, f"alien-stock-{scope}-{stamp}.zip")
    db_copy = os.path.join(FULL_BACKUP_DIR, f".data-{stamp}.db")
    try:
        manifest = {"created_at": datetime.now(DHAKA_TZ).isoformat(), "scope": scope}
        if scope in ("all", "database"):
            with sqlite3.connect(DB) as source, sqlite3.connect(db_copy) as target: source.backup(target)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            if scope in ("all", "database"):
                archive.write(db_copy, "data.db")
            if scope in ("all", "stock"):
                if os.path.isfile(EXCEL_STOCK_FILE): archive.write(EXCEL_STOCK_FILE, "instance/stock_master.xlsx")
                if os.path.isdir(EXCEL_BACKUP_DIR):
                    for name in os.listdir(EXCEL_BACKUP_DIR):
                        candidate = os.path.join(EXCEL_BACKUP_DIR, name)
                        if os.path.isfile(candidate) and name.lower().endswith(".xlsx"):
                            archive.write(candidate, "instance/stock_backups/" + name)
            if scope == "all":
                root = os.path.dirname(os.path.abspath(__file__))
                if os.path.isfile(__file__): archive.write(__file__, os.path.basename(__file__))
                for name in ("config.py", "README.md", "requirements.txt", "catalog_cache.pkl", "sigs.pkl"):
                    candidate = os.path.join(root, name)
                    if os.path.isfile(candidate): archive.write(candidate, name)
                for folder in ("media", "static"):
                    base = os.path.join(root, folder)
                    if os.path.isdir(base):
                        for current, _, files in os.walk(base):
                            for filename in files:
                                candidate = os.path.join(current, filename)
                                archive.write(candidate, os.path.relpath(candidate, root).replace(os.sep, "/"))
            if scope in ("history", "users-settings"):
                table_names = ("excel_stock_history", "excel_stock_batches", "excel_stock_comments", "excel_stock_imports", "excel_stock_backups", "activity_logs", "price_history") if scope == "history" else ("users", "settings", "sources", "prices")
                with db() as c:
                    payload = {}
                    existing = {r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                    for table in table_names:
                        if table in existing: payload[table] = [dict(r) for r in c.execute(f"SELECT * FROM {table}")]
                archive.writestr("data-export.json", json.dumps(payload, ensure_ascii=False, indent=2))
            manifest["contents"] = {"all":["database, permissions, settings, histories, Excel, workbook snapshots and source"], "database":["data.db: all app data and configuration"], "stock":["current workbook and Excel snapshots"], "history":["stock, sales-comment, import, price and activity history"], "users-settings":["users, permissions, settings, sources and prices"]}[scope]
            archive.writestr("MANIFEST.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        log_activity("full_system_backup", os.path.basename(path), labels[scope] + " ডেটা ব্যাকআপ তৈরি")
        return send_file(path, as_attachment=True, download_name=os.path.basename(path), mimetype="application/zip")
    except Exception as e:
        if os.path.isfile(path): os.remove(path)
        return jsonify(error="পূর্ণ ব্যাকআপ তৈরি হয়নি: " + str(e)), 500
    finally:
        if os.path.isfile(db_copy): os.remove(db_copy)

@app.post("/api/admin/excel-stock/restore-latest")
@super_admin_only
def excel_stock_restore_latest():
    files = [os.path.join(EXCEL_BACKUP_DIR, n) for n in os.listdir(EXCEL_BACKUP_DIR) if n.lower().endswith(".xlsx")] if os.path.isdir(EXCEL_BACKUP_DIR) else []
    if not files: return jsonify(error="এখনো কোনো ব্যাকআপ নেই"), 404
    return excel_stock_backup_restore(os.path.basename(max(files, key=os.path.getmtime)))

@app.post("/api/admin/excel-stock/restore/<filename>")
@super_admin_only
def excel_stock_backup_restore(filename):
    safe = os.path.basename(filename)
    source = os.path.join(EXCEL_BACKUP_DIR, safe)
    if safe != filename or not os.path.isfile(source): return jsonify(error="ব্যাকআপ পাওয়া যায়নি"), 404
    tmp = None
    try:
        with EXCEL_LOCK:
            snapshot_excel_stock("before-restore", session.get("user", ""))
            fd, tmp = tempfile.mkstemp(prefix="stock-restore-", suffix=".xlsx", dir=INSTANCE)
            os.close(fd); shutil.copy2(source, tmp); os.replace(tmp, EXCEL_STOCK_FILE)
        refresh_stock_from_excel()
        log_activity("excel_stock_restore", safe, "আগের workbook ব্যাকআপ থেকে ফিরিয়ে আনা হয়েছে")
        return jsonify(ok=True)
    except Exception as e:
        if tmp and os.path.exists(tmp): os.remove(tmp)
        return jsonify(error="ব্যাকআপ ফিরিয়ে আনা যায়নি: " + str(e)), 400

@app.get("/api/img")
def img():
    if not can_action("view_product"): return "permission_denied",403
    url = request.args.get("url", ""); host = urlparse(url).hostname or ""
    _gal_load_all()
    if host not in ({urlparse(u).hostname for x in C["products"] for u in [x.get("img","")] + x.get("extra",[]) + ([x.get("size_chart_image")] if x.get("size_chart_image") else []) if u} | {urlparse(u).hostname for v in list(_GAL.values()) for u in v[1]}): return "forbidden", 403
    dl_settings = {}
    if request.args.get("dl"):
        try: dl_settings = json.loads(get_setting("photo_download_settings") or "{}")
        except Exception: dl_settings = {}
        access = dl_settings.get("access", "all")
        if access == "logged" and not current_user(): return "login_required", 403
        if access == "admin" and not is_admin_level(): return "admin_required", 403
        if not can_action("download"): return "permission_denied", 403
    if request.args.get("copy") and not can_action("copy"): return "permission_denied", 403
    r = requests.get(url, timeout=30, headers=UA); h = {"Cache-Control": "public, max-age=86400"}
    if request.args.get("dl"):
        h["Content-Disposition"] = 'attachment; filename="%s.jpg"' % re.sub(r"[^\w.-]", "_", request.args.get("name", "product"))
        log_activity("download", request.args.get("name", "product"), "HD")
        try:
            im = Image.open(io.BytesIO(r.content)).convert("RGB")
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(im)
            code = str(request.args.get("code", "")).strip()[:80] if dl_settings.get("show_code", True) else ""
            price = str(request.args.get("price", "")).strip()[:50] if dl_settings.get("show_price", True) else ""
            label = "   ".join(x for x in (("CODE: " + code) if code else "", ("PRICE: BDT " + price) if price else "") if x)
            if label:
                font_path = next((x for x in ("C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf") if os.path.isfile(x)), "")
                try: font = ImageFont.truetype(font_path, max(16, min(38, im.width // 24))) if font_path else ImageFont.load_default()
                except Exception: font = ImageFont.load_default()
                box = draw.textbbox((0, 0), label, font=font)
                pad = max(8, im.width // 100); tw, th = box[2]-box[0], box[3]-box[1]
                x, y = pad, im.height-th-pad*2
                draw.rounded_rectangle((x-pad//2,y-pad//2,x+tw+pad,y+th+pad),radius=8,fill=(0,0,0))
                draw.text((x,y),label,font=font,fill=(255,255,255))
                out=io.BytesIO(); im.save(out,format="JPEG",quality=95); out.seek(0)
                return Response(out.getvalue(),content_type="image/jpeg",headers=h)
        except Exception as e: print("download label failed:",e)
    elif request.args.get("copy"):
        log_activity("copy", request.args.get("name", "product"), "HD")
    return Response(r.content, content_type=r.headers.get("Content-Type", "image/jpeg"), headers=h)

@app.post("/api/search-image")
def search_image():
    if not can_action("image_search"):
        return jsonify(ok=False, error="permission_denied", message="Image search is not enabled for this account."), 403
    f = request.files.get("file") or request.files.get("image") or request.files.get("photo")
    if not f or not getattr(f, "filename", ""):
        return jsonify(ok=False, matches=[], message="Choose an image file first."), 400
    try:
        raw = f.read()
        if not raw:
            return jsonify(ok=False, matches=[], message="The selected image is empty."), 400
        if len(raw) > 12 * 1024 * 1024:
            return jsonify(ok=False, matches=[], message="Image is too large. Please use an image under 12 MB."), 413
        im = Image.open(io.BytesIO(raw))
        im.verify()
        im = Image.open(io.BytesIO(raw)).convert("RGB")
        im.thumbnail((1400, 1400))
        q = make_q(im)
    except Exception as exc:
        print("image search upload failed:", repr(exc))
        return jsonify(ok=False, matches=[], message="Could not read this image. Use JPG, PNG or WEBP."), 400

    # If the catalog is still starting, trigger a refresh and return a clear retry message.
    if not C.get("products"):
        try: threading.Thread(target=_refresh_catalog_background, daemon=True).start()
        except Exception: pass
        return jsonify(ok=False, matches=[], message="Catalog is still loading. Please retry in a moment."), 503

    image_urls = list(dict.fromkeys(str(p.get("img") or "").strip() for p in C.get("products", [])
                                    if str(p.get("img") or "").strip().startswith(("https://", "http://"))))
    valid_urls = {url for url in image_urls if isinstance(_sigs.get(url), dict) and _sigs[url].get("f")}
    valid_count = len(valid_urls)
    target_count = min(12, len(image_urls))
    # Background indexing normally handles every product. If the cache is cold or only partially
    # built, complete a bounded starter index instead of stopping after just 3 cached products.
    if valid_count < target_count:
        candidates = []
        seen_urls = set()
        needed = target_count - valid_count
        for p in C.get("products", []):
            url = str(p.get("img") or "").strip()
            if url.startswith(("https://", "http://")) and url not in seen_urls and not (isinstance(_sigs.get(url), dict) and _sigs[url].get("f")):
                seen_urls.add(url); candidates.append(p)
        if len(candidates) <= needed:
            products = candidates
        elif needed <= 1:
            products = [candidates[len(candidates)//2]]
        else:
            # Spread the first-use sample across the complete catalogue, not just the first rows.
            indexes = sorted(set(round(i * (len(candidates)-1) / (needed-1)) for i in range(needed)))
            products = [candidates[i] for i in indexes]
        def _warm_one(product):
            url = product.get("img") or ""
            try:
                r = requests.get(url, timeout=3.5, headers=UA)
                r.raise_for_status()
                im2 = Image.open(io.BytesIO(r.content)).convert("RGB")
                im2.thumbnail((512, 512))
                # Fast feature pass: never wait for CLIP model loading in a search request.
                sig = {"f": feat(im2), "e": None, "e2": None}
                with _CLOCK: _sigs[url] = sig
                return True
            except Exception as exc:
                print("lazy image index failed:", url, str(exc)[:120])
                with _CLOCK: _sigs.pop(url, None)
                return False
        if products:
            with ThreadPoolExecutor(max_workers=8) as ex:
                list(ex.map(_warm_one, products))
            try:
                tmp = SIGF + ".tmp"
                with open(tmp, "wb") as fh: pickle.dump(_sigs, fh, protocol=pickle.HIGHEST_PROTOCOL)
                os.replace(tmp, SIGF)
            except Exception as exc: print("lazy image index save failed:", exc)

    matches = rank(q)
    if not matches:
        if not _BG.get("ai"):
            try: threading.Thread(target=_build_ai_background, daemon=True).start()
            except Exception: pass
        return jsonify(ok=False, matches=[], ai=bool(q.get("e")), message="Image index is being prepared in the background. Please try again shortly."), 503
    log_activity("image_search", "[image]", json.dumps({"results": len(matches), "ai": bool(q["e"])}, ensure_ascii=False))
    # Continue indexing beyond the starter batch in the background; this no longer blocks this response.
    if not _BG.get("ai"):
        try: threading.Thread(target=_build_ai_background, daemon=True).start()
        except Exception: pass
    return jsonify(ok=True, matches=matches, ai=bool(q["e"]))

# ---------- প্রোডাক্টের সব ছবি (ওয়েবসাইটের প্রোডাক্ট পেজ থেকে) ----------
_GAL = {}            # key -> (fetched_at, [urls])
_GAL_LOADED = [False]
_GAL_LOCK = threading.Lock()
_IMG_RE = re.compile(r'https?://[^\s"\'<>\\)]+?\.(?:webp|jpe?g|png)(?:\?[^\s"\'<>\\)]*)?', re.I)

def _gal_load_all():
    if _GAL_LOADED[0]: return
    with _GAL_LOCK:
        if _GAL_LOADED[0]: return
        try:
            with db() as c:
                for r in c.execute("SELECT k,v FROM settings WHERE k LIKE 'gallery:%'"):
                    try:
                        d = json.loads(r["v"]); _GAL[r["k"][8:]] = (float(d.get("t", 0)), list(d.get("urls") or []))
                    except Exception: pass
        except Exception: pass
        _GAL_LOADED[0] = True

def _gal_urls_from_page(p, html):
    prim = p.get("img", "") or ""
    m = re.search(r"/images/products/([0-9a-f]+)/", prim)
    folder = m.group(1) if m else ""
    code = (display_product_id(p) or "").lower()
    html = html.replace("\\/", "/").replace("\\u002F", "/").replace("&amp;", "&")
    found = []
    # JSON-LD Product ছবি — সবচেয়ে নির্ভরযোগ্য
    for blk in re.findall(r'<script[^>]+application/ld\+json[^>]*>(.*?)</script>', html, re.S | re.I):
        try: data = json.loads(blk)
        except Exception: continue
        stack = [data]
        while stack:
            o = stack.pop()
            if isinstance(o, list): stack.extend(o)
            elif isinstance(o, dict):
                t = o.get("@type"); t = t if isinstance(t, list) else [t]
                if "Product" in t:
                    im_ = o.get("image"); im_ = im_ if isinstance(im_, list) else [im_]
                    for v in im_:
                        if isinstance(v, dict): v = v.get("url")
                        if isinstance(v, str) and v.startswith("http"): found.append(v)
                stack.extend(v for v in o.values() if isinstance(v, (list, dict)))
    # পেজের ভেতরের CDN ছবি: একই ফোল্ডার বা ফাইলের নামে প্রোডাক্ট কোড থাকলে তবেই (সিমিলার/রিলেটেড বাদ)
    for u in _IMG_RE.findall(html):
        base = u.split("?")[0].rsplit("/", 1)[-1].lower()
        same_folder = bool(folder) and ("/" + folder + "/") in u
        has_code = bool(code) and re.search(r"(^|[_\-])" + re.escape(code) + r"(?![0-9])", base)
        if same_folder or has_code: found.append(u)
    out, seen = [], set()
    for u in found:
        key = u.split("?")[0]
        if key in seen: continue
        seen.add(key); out.append(u)
    prim_key = prim.split("?")[0]
    rest = sorted((u for u in out if u.split("?")[0] != prim_key), key=lambda u: u.split("?")[0].rsplit("/", 1)[-1].split("_", 1)[-1].lower())
    return rest[:30]

def product_gallery(key, force=False):
    _gal_load_all()
    p = next((x for x in C["products"] if x["store"] + "|" + x["id"] == key), None)
    if not p: return None, []
    ttl_ok = 3 * 86400
    hit = _GAL.get(key)
    if hit and not force:
        age = time.time() - hit[0]
        if (hit[1] and age < ttl_ok) or (not hit[1] and age < 3600): return p, hit[1]
    urls = []
    link = p.get("link") or ""
    try:
        if link.startswith(("http://", "https://")):
            r = requests.get(link, timeout=15, headers=UA)
            if r.ok: urls = _gal_urls_from_page(p, r.text)
    except Exception as e:
        print("gallery fetch failed:", key, e)
    if not urls and hit and hit[1]: urls = hit[1]
    _GAL[key] = (time.time(), urls)
    try: set_setting("gallery:" + key, json.dumps({"t": time.time(), "urls": urls}, ensure_ascii=False))
    except Exception: pass
    return p, urls

@app.get("/api/product-gallery")
def product_gallery_api():
    if not can_action("view_product"): return jsonify(error="permission_denied"),403
    key = request.args.get("key", "")[:500]
    p, urls = product_gallery(key, force=bool(request.args.get("refresh")) and session.get("role") in ("admin", "super_admin"))
    if p is None: return jsonify(images=[]), 404
    return jsonify(images=urls)

@app.get("/api/similar")
def similar():
    if not can_action("view_product"): return jsonify(error="permission_denied"),403
    k = request.args.get("key", ""); x = next((p for p in C["products"] if p["store"] + "|" + p["id"] == k), None)
    s = x and _sigs.get(x["img"])
    return jsonify(matches=rank(s, exclude=k, n=8) if s else [])

@app.get("/api/admin/photo-download-settings")
@admin_only
def photo_download_settings_get():
    try: data = json.loads(get_setting("photo_download_settings") or "{}")
    except Exception: data = {}
    return jsonify(show_code=bool(data.get("show_code", True)), show_price=bool(data.get("show_price", True)), access=data.get("access", "all") if data.get("access") in ("all", "logged", "admin") else "all")

@app.post("/api/admin/photo-download-settings")
@admin_only
def photo_download_settings_set():
    j = request.get_json(silent=True) or {}; access = j.get("access", "all")
    if access not in ("all", "logged", "admin"): return jsonify(error="ভুল download access setting"), 400
    data = {"show_code": bool(j.get("show_code", True)), "show_price": bool(j.get("show_price", True)), "access": access}
    set_setting("photo_download_settings", json.dumps(data, ensure_ascii=False))
    log_activity("photo_download_settings", "photo download", json.dumps(data, ensure_ascii=False))
    return jsonify(ok=True)

@app.get("/api/notice")
def notice_get():
    if not session.get("user"): return jsonify(error="login_required"), 401
    try:
        data = json.loads(get_setting("site_notice") or "{}")
        return jsonify(text=str(data.get("text", "")), enabled=bool(data.get("enabled", False)), duration_seconds=max(1, min(300, int(data.get("duration_seconds", 6) or 6))))
    except Exception:
        return jsonify(text="", enabled=False, duration_seconds=6)

@app.post("/api/admin/notice")
@admin_only
def notice_set():
    j = request.get_json(silent=True) or {}
    data = {"text": str(j.get("text", ""))[:2000], "enabled": bool(j.get("enabled", False)), "duration_seconds": max(1, min(300, int(j.get("duration_seconds", 6) or 6)))}
    set_setting("site_notice", json.dumps(data, ensure_ascii=False))
    log_activity("notice_update", "site_notice", "enabled" if data["enabled"] else "disabled")
    return jsonify(ok=True)

@app.get("/api/notifications")
@login_required
def notifications_list():
    with db() as c:
        rows=[dict(r) for r in c.execute("SELECT id,title,message,kind,ref,created_at,read_at FROM notifications ORDER BY id DESC LIMIT 100")]
        unread=c.execute("SELECT COUNT(*) n FROM notifications WHERE read_at IS NULL").fetchone()["n"]
    return jsonify(items=rows,unread=unread)

@app.post("/api/notifications/read")
@login_required
def notifications_read():
    j=request.get_json(silent=True) or {}; nid=j.get("id")
    with db() as c:
        if nid is not None: c.execute("UPDATE notifications SET read_at=? WHERE id=? AND read_at IS NULL",(time.time(),int(nid)))
        else: c.execute("UPDATE notifications SET read_at=? WHERE read_at IS NULL",(time.time(),))
    return jsonify(ok=True)

@app.get("/api/admin/courier-options")
@login_required
def courier_options_get():
    if not is_admin_level() and not (has_perm("courier_view") or has_perm("courier_update")): return jsonify(error="permission_denied"),403
    try: data=json.loads(get_setting("courier_settings") or "{}")
    except Exception: data={}
    names=list(dict.fromkeys(["Steadfast","Pathao","Delivery Express"]+list(data.keys())))
    return jsonify(couriers=names)

@app.get("/api/admin/couriers")
@admin_only
def courier_settings_get():
    try: data=json.loads(get_setting("courier_settings") or "{}")
    except Exception: data={}
    defaults={"Steadfast":"https://portal.packzy.com/api/v1","Pathao":"https://api-hermes.pathao.com","Delivery Express":""}
    couriers={k:{"base_url":v.get("base_url") or defaults.get(k,"") ,**v} for k,v in data.items()}
    for name,base in defaults.items(): couriers.setdefault(name,{"base_url":base,"api_key":"","secret_key":"","enabled":False})
    with db() as c: safe_mappings={r["courier"]:{"status_path":r["status_path"]} for r in c.execute("SELECT * FROM courier_order_mappings")}
    return jsonify(couriers=[{"name":k,"configured":bool(v.get("api_key") and (v.get("secret_key") or v.get("auth_type")=="bearer")),"enabled":bool(v.get("enabled")),"base_url":v.get("base_url",""),"auth_type":v.get("auth_type","steadfast"),"status_field":v.get("status_field","delivery_status"),"fraud_path":v.get("fraud_path",""),"has_api_key":bool(v.get("api_key")),"has_secret_key":bool(v.get("secret_key")),"mapping":safe_mappings.get(k,{})} for k,v in couriers.items()])

@app.post("/api/admin/couriers")
@admin_only
def courier_settings_set():
    j=request.get_json(silent=True) or {}; name=re.sub(r"[^A-Za-z0-9 _-]","",str(j.get("name","")).strip())[:50]
    if not name: return jsonify(error="কুরিয়ারের নাম দিন"),400
    try: data=json.loads(get_setting("courier_settings") or "{}")
    except Exception: data={}
    prior=data.get(name,{}); key=str(j.get("api_key","")).strip() or prior.get("api_key",""); secret=str(j.get("secret_key","")).strip() or prior.get("secret_key",""); base=str(j.get("base_url","")).strip()[:300]
    if base and not base.startswith("https://"): return jsonify(error="API address অবশ্যই HTTPS হতে হবে"),400
    auth_type=str(j.get("auth_type",prior.get("auth_type","steadfast"))).lower()
    if auth_type not in ("steadfast","bearer","api-key","none"): return jsonify(error="ভুল API authentication mode"),400
    data[name]={"api_key":key,"secret_key":secret,"base_url":base or prior.get("base_url",""),"enabled":bool(j.get("enabled",True)),"auth_type":auth_type,"status_field":str(j.get("status_field",prior.get("status_field","delivery_status"))).strip()[:100] or "delivery_status","fraud_path":str(j.get("fraud_path",prior.get("fraud_path",""))).strip()[:300]}
    set_setting("courier_settings",json.dumps(data,ensure_ascii=False)); log_activity("courier_config",name,"credential updated" if j.get("api_key") else "settings updated")
    m=j.get("mapping") or {}
    with db() as c: c.execute("INSERT INTO courier_order_mappings(courier,status_path,updated_at) VALUES(?,?,?) ON CONFLICT(courier) DO UPDATE SET status_path=excluded.status_path,updated_at=excluded.updated_at",(name,str(m.get("status_path", "/status_by_cid/{cons_id}"))[:300],time.time()))
    return jsonify(ok=True)

@app.post("/api/admin/courier-test")
@admin_only
def courier_connection_test():
    j=request.get_json(silent=True) or {}; name=re.sub(r"[^A-Za-z0-9 _-]","",str(j.get("name","Steadfast")).strip())[:50]
    try: settings=json.loads(get_setting("courier_settings") or "{}")
    except Exception: settings={}
    cfg=settings.get(name,{})
    if name!="Steadfast": return jsonify(error="এই quick test শুধু Steadfast-এর জন্য আছে"),400
    if not cfg.get("enabled"): return jsonify(error="আগে Steadfast courier চালু করে Settings Save করুন"),400
    if not cfg.get("api_key") or not cfg.get("secret_key"): return jsonify(error="আগে API Key ও Secret Key দিয়ে Settings Save করুন"),400
    base=str(cfg.get("base_url") or "https://portal.packzy.com/api/v1").rstrip("/")
    if not base.startswith("https://"): return jsonify(error="API Base URL অবশ্যই HTTPS হতে হবে"),400
    try:
        response=requests.get(base+"/get_balance",headers={"Api-Key":cfg["api_key"],"Secret-Key":cfg["secret_key"],"Content-Type":"application/json"},timeout=15)
        if response.status_code in (401,403): return jsonify(error="Steadfast API সংযোগ হয়নি: key/secret ভুল, মেয়াদ শেষ, অথবা API access চালু নেই"),502
        if not response.ok: return jsonify(error=f"Steadfast API HTTP {response.status_code} দিয়েছে"),502
        body=response.json()
        if isinstance(body,dict) and str(body.get("status","200")) not in ("200","success","ok"):
            return jsonify(error="Steadfast API অনুরোধ গ্রহণ করেনি; API credential/access যাচাই করুন"),502
        return jsonify(ok=True,message="Steadfast API connection successful")
    except requests.Timeout: return jsonify(error="Steadfast API সাড়া দিতে দেরি করছে; একটু পরে আবার চেষ্টা করুন"),504
    except requests.RequestException: return jsonify(error="Steadfast API-তে পৌঁছানো যাচ্ছে না; internet ও Base URL যাচাই করুন"),502
    except (ValueError,TypeError): return jsonify(error="Steadfast API থেকে অপ্রত্যাশিত response এসেছে"),502

@app.get("/api/admin/courier-tracking")
@login_required
def courier_tracking_list():
    if not is_admin_level() and not (has_perm("courier_view") or has_perm("courier_update")): return jsonify(error="permission_denied"),403
    q=request.args.get("q","").strip().lower(); courier=request.args.get("courier","Steadfast").strip()[:50] or "Steadfast"
    status_filter=request.args.get("status","").strip().lower(); assigned_filter=request.args.get("assigned_to","").strip().lower()
    start=request.args.get("from","").strip(); end=request.args.get("to","").strip()
    with db() as c:
        orders=[dict(r) for r in c.execute("""SELECT o.id AS order_id,o.cons_id,? AS courier,
            COALESCE(NULLIF(t.status,''),NULLIF(o.delivery_status,''),'Pending') AS status,
            COALESCE(NULLIF(t.rating,''),o.rating,'') AS rating,COALESCE(t.details,'') AS details,
            COALESCE(t.updated_by,o.uploaded_by,'') AS updated_by,
            COALESCE(t.updated_at,o.uploaded_at,0) AS updated_at,o.order_date,o.customer_name,o.phone,o.product_code,o.size,o.assigned_to,t.api_amount,t.api_amount_field,
            o.product_price,o.delivery_charge,o.advance_paid,o.collectible_amount
            FROM order_history o LEFT JOIN courier_tracking t ON t.cons_id=o.cons_id AND t.courier=?
            WHERE TRIM(COALESCE(o.cons_id,''))<>'' AND o.id=(SELECT MAX(o2.id) FROM order_history o2 WHERE TRIM(o2.cons_id)=TRIM(o.cons_id))
            ORDER BY o.id DESC""",(courier,courier)).fetchall()]
        orphan=[dict(r) for r in c.execute("""SELECT t.*,NULL AS order_date,'' AS customer_name,'' AS phone,'' AS product_code,'' AS size,'' AS assigned_to,
            NULL AS product_price,NULL AS delivery_charge,NULL AS advance_paid,NULL AS collectible_amount
            FROM courier_tracking t WHERE t.courier=? AND NOT EXISTS(SELECT 1 FROM order_history o WHERE TRIM(o.cons_id)=TRIM(t.cons_id))
            ORDER BY t.updated_at DESC""",(courier,)).fetchall()]
    rows=orders+orphan
    # Keep a stable, unfiltered snapshot for the summary cards. Search, date,
    # status and assignee filters should only change the table rows; they must
    # never make the whole dashboard look like it has zero orders.
    summary_rows=list(rows)
    if start: rows=[r for r in rows if (r.get("order_date") or datetime.fromtimestamp(float(r.get("updated_at") or 0),DHAKA_TZ).strftime("%Y-%m-%d"))>=start]
    if end: rows=[r for r in rows if (r.get("order_date") or datetime.fromtimestamp(float(r.get("updated_at") or 0),DHAKA_TZ).strftime("%Y-%m-%d"))<=end]
    if assigned_filter: rows=[r for r in rows if assigned_filter in str(r.get("assigned_to") or "").lower()]
    qd=re.sub(r"\D","",q)
    if q: rows=[r for r in rows if q in " ".join(str(r.get(k) or "") for k in ("cons_id","courier","status","customer_name","phone","product_code","order_date","assigned_to","details")).lower() or (len(qd)>=6 and qd in re.sub(r"\D","",str(r.get("phone") or "")))]
    totals={"delivered":0,"partial_delivered":0,"returned":0,"cancelled":0,"pending":0,"inway":0}
    for item in summary_rows:
        s=str(item.get("status") or "").lower().replace("_"," ").replace("-"," ")
        k="returned" if "return" in s else "cancelled" if ("cancel" in s or "canceled" in s) else "inway" if any(x in s for x in ("in way","transit","out for delivery","pickup pending","dispatched","shipped")) else "partial_delivered" if "partial deliver" in s else "delivered" if "deliver" in s else "pending"
        totals[k]+=1
    status_aliases={"partial delivered":"partial_delivered","partial_delivered":"partial_delivered","in way":"inway","inway":"inway"}
    if status_filter:
        normalized=re.sub(r"[_-]+"," ",status_filter).strip()
        wanted=status_aliases.get(normalized,normalized)
        rows=[r for r in rows if status_aliases.get(re.sub(r"[_-]+"," ",str(r.get("status") or "").lower()).strip(),re.sub(r"[_-]+"," ",str(r.get("status") or "").lower()).strip())==wanted]
    def amount(value):
        try:
            raw=re.sub(r"[^0-9.-]","",str(value or "").replace(",",""))
            return float(raw) if raw and raw not in (".","-","-.") else 0.0
        except Exception: return 0.0
    money={"delivered":0.0,"partial_delivered":0.0,"returned":0.0,"cancelled":0.0,"pending":0.0,"inway":0.0}
    for item in summary_rows:
        status=str(item.get("status") or "").lower().replace("_"," ").replace("-"," ")
        key="returned" if "return" in status else "cancelled" if "cancel" in status else "inway" if any(x in status for x in ("transit","in way","out for delivery","pickup pending","dispatched","shipped")) else "partial_delivered" if "partial deliver" in status else "delivered" if "deliver" in status else "pending"
        estimate=amount(item.get("collectible_amount"))
        if not str(item.get("collectible_amount") or "").strip(): estimate=max(0,amount(item.get("product_price"))+amount(item.get("delivery_charge"))-amount(item.get("advance_paid")))
        collectible=amount(item.get("api_amount")) if item.get("api_amount") is not None else estimate
        if key in ("delivered","partial_delivered","returned","cancelled","pending"): money[key]+=collectible
        item["collectible_amount"]=collectible
        item["amount_is_estimate"]=item.get("api_amount") is None
    rows.sort(key=lambda r:float(r.get("updated_at") or 0),reverse=True)
    assigned=sorted({str(r.get("assigned_to") or "").strip() for r in orders if str(r.get("assigned_to") or "").strip()},key=str.lower)
    return jsonify(rows=rows[:2000],counts=totals,totals=totals,money={k:round(v,2) for k,v in money.items()},assigned=assigned,total=len(rows),order_total=len(orders),courier=courier)
@app.get("/api/admin/courier-report.xlsx")
@login_required
def courier_report_excel():
    if not is_admin_level() and not (has_perm("courier_view") or has_perm("courier_update")): return jsonify(error="permission_denied"),403
    start=request.args.get("from","").strip(); end=request.args.get("to","").strip(); courier=request.args.get("courier","Steadfast").strip()[:50] or "Steadfast"
    status_filter=request.args.get("status","").strip().lower(); assigned_filter=request.args.get("assigned_to","").strip().lower()
    order_where=["TRIM(COALESCE(o.cons_id,''))<>''","o.id=(SELECT MAX(o2.id) FROM order_history o2 WHERE TRIM(o2.cons_id)=TRIM(o.cons_id))"]; order_args=[courier,courier]
    if start: order_where.append("COALESCE(NULLIF(o.order_date,''),strftime('%Y-%m-%d',COALESCE(t.updated_at,o.uploaded_at),'unixepoch','+6 hours'))>=?"); order_args.append(start)
    if end: order_where.append("COALESCE(NULLIF(o.order_date,''),strftime('%Y-%m-%d',COALESCE(t.updated_at,o.uploaded_at),'unixepoch','+6 hours'))<=?"); order_args.append(end)
    if assigned_filter: order_where.append("lower(COALESCE(o.assigned_to,'')) LIKE ?"); order_args.append("%"+assigned_filter+"%")
    if status_filter: order_where.append("lower(replace(replace(COALESCE(NULLIF(t.status,''),NULLIF(o.delivery_status,''),'Pending'),'_',' '),'-',' '))=?"); order_args.append(re.sub(r"[_-]+"," ",status_filter).strip())
    event_where=[]; event_args=[]
    if start: event_where.append("strftime('%Y-%m-%d',created_at,'unixepoch','+6 hours')>=?"); event_args.append(start)
    if end: event_where.append("strftime('%Y-%m-%d',created_at,'unixepoch','+6 hours')<=?"); event_args.append(end)
    fraud_where=[]; fraud_args=[]
    if start: fraud_where.append("strftime('%Y-%m-%d',checked_at,'unixepoch','+6 hours')>=?"); fraud_args.append(start)
    if end: fraud_where.append("strftime('%Y-%m-%d',checked_at,'unixepoch','+6 hours')<=?"); fraud_args.append(end)
    event_clause=(" WHERE "+" AND ".join(event_where)) if event_where else ""
    fraud_clause=(" WHERE "+" AND ".join(fraud_where)) if fraud_where else ""
    with db() as c:
        current=[dict(r) for r in c.execute("""SELECT o.cons_id,? AS courier,COALESCE(NULLIF(t.status,''),NULLIF(o.delivery_status,''),'Pending') AS status,COALESCE(NULLIF(t.rating,''),o.rating,'') AS rating,COALESCE(t.details,'') AS details,COALESCE(t.order_id,o.id) AS order_id,o.order_date,o.customer_name,o.phone,o.product_code,o.size,o.assigned_to,COALESCE(t.updated_by,o.uploaded_by,'') AS updated_by,COALESCE(t.updated_at,o.uploaded_at,0) AS updated_at,t.api_amount,t.api_amount_field,o.collectible_amount,o.product_price,o.delivery_charge,o.advance_paid FROM order_history o LEFT JOIN courier_tracking t ON t.cons_id=o.cons_id AND t.courier=? WHERE """+" AND ".join(order_where),order_args).fetchall()]
        events=[dict(r) for r in c.execute("SELECT cons_id,courier,status,details,source,created_at FROM courier_events"+event_clause+" ORDER BY created_at DESC",event_args)]
        fraud_checks=[dict(r) for r in c.execute("SELECT phone,courier,orders,delivered,returned,provider_result,checked_by,checked_at FROM courier_fraud_checks"+fraud_clause+" ORDER BY checked_at DESC",fraud_args)]
        # Include manually tracked consignments that are not matched to an uploaded order.
        orphan_where=["NOT EXISTS(SELECT 1 FROM order_history o WHERE TRIM(o.cons_id)=TRIM(t.cons_id))"]; orphan_args=[courier]
        if status_filter: orphan_where.append("lower(replace(replace(COALESCE(t.status,''),'_',' '),'-',' '))=?"); orphan_args.append(re.sub(r"[_-]+"," ",status_filter).strip())
        if start: orphan_where.append("strftime('%Y-%m-%d',t.updated_at,'unixepoch','+6 hours')>=?"); orphan_args.append(start)
        if end: orphan_where.append("strftime('%Y-%m-%d',t.updated_at,'unixepoch','+6 hours')<=?"); orphan_args.append(end)
        orphans=[dict(r) for r in c.execute("SELECT t.cons_id, t.courier, t.status, t.rating, t.details, t.order_id, '' AS order_date, '' AS customer_name, '' AS phone, '' AS product_code, '' AS size, '' AS assigned_to, t.updated_by, t.updated_at FROM courier_tracking t WHERE t.courier=? AND "+" AND ".join(orphan_where)+" ORDER BY t.updated_at DESC",orphan_args)]
        current.extend(orphans)
    wb=Workbook(); ws=wb.active; ws.title="Courier Status"
    ws.append(["Consignment ID","Courier","Status","API Amount","API Amount Field","Product Price (Order Sheet)","Courier Charge (Order Sheet)","Advance Paid (Order Sheet)","Order Collectible","Rating","Details","Order ID","Courier Booking Date (Order Date)","Assigned To","Customer","Phone","Product Code","Size","Updated By","Updated At"])
    search_query=request.args.get("q", "").strip().lower()
    if search_query:
        current=[r for r in current if search_query in " ".join(str(r.get(k) or "") for k in ("cons_id","courier","status","rating","details","order_id","order_date","customer_name","phone","product_code","size","assigned_to","updated_by")).lower()]
    for r in current:
        estimate=r.get("collectible_amount")
        if not str(estimate or "").strip(): estimate=max(0,float(re.sub(r"[^0-9.-]","",str(r.get("product_price") or "0")) or 0)+float(re.sub(r"[^0-9.-]","",str(r.get("delivery_charge") or "0")) or 0)-float(re.sub(r"[^0-9.-]","",str(r.get("advance_paid") or "0")) or 0))
        ws.append([r.get("cons_id"),r.get("courier"),r.get("status"),r.get("api_amount"),r.get("api_amount_field"),r.get("product_price"),r.get("delivery_charge"),r.get("advance_paid"),estimate,r.get("rating"),r.get("details"),r.get("order_id"),r.get("order_date"),r.get("assigned_to"),r.get("customer_name"),r.get("phone"),r.get("product_code"),r.get("size"),r.get("updated_by"),datetime.fromtimestamp(r["updated_at"],DHAKA_TZ).strftime("%Y-%m-%d %H:%M:%S") if r.get("updated_at") else ""])
    hist=wb.create_sheet("Status History"); hist.append(["Consignment ID","Courier","Status","Details","Source","Updated At"])
    for r in events: hist.append([r["cons_id"],r["courier"],r["status"],r["details"],r["source"],datetime.fromtimestamp(r["created_at"],DHAKA_TZ).strftime("%Y-%m-%d %H:%M:%S") if r.get("created_at") else ""])
    fraud=wb.create_sheet("Fraud Checks"); fraud.append(["Phone","Courier","Orders","Delivered","Returned","Return Rate","Courier API Result","Checked By","Checked At"])
    for r in fraud_checks: fraud.append([r["phone"],r["courier"],r["orders"],r["delivered"],r["returned"],round(r["returned"]/max(1,r["orders"])*100,1),r["provider_result"],r["checked_by"],datetime.fromtimestamp(r["checked_at"],DHAKA_TZ).strftime("%Y-%m-%d %H:%M:%S") if r.get("checked_at") else ""])
    from openpyxl.styles import Font,PatternFill,Alignment
    for sheet in wb.worksheets:
        sheet.freeze_panes="A2"; sheet.auto_filter.ref=sheet.dimensions
        for cell in sheet[1]:
            cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="164A3A"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        sheet.row_dimensions[1].height=30
        for col in sheet.columns:
            letter=col[0].column_letter; sheet.column_dimensions[letter].width=min(38,max(13,max((len(str(cell.value or "")) for cell in col[:250]),default=12)+2))
    output=io.BytesIO(); wb.save(output); output.seek(0); suffix="_filtered" if start or end or status_filter or assigned_filter else ""
    return send_file(output,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name=f"Alien-Tees-Courier-Report{suffix}.xlsx")
@app.get("/api/admin/courier-fraud-check")
@login_required
def courier_fraud_check():
    if not is_admin_level() and not (has_perm("courier_view") or has_perm("courier_update")): return jsonify(error="permission_denied"),403
    phone=normalize_order_phone(request.args.get("phone", "")).strip()
    if not phone: return jsonify(error="ফোন নম্বর দিন"),400
    with db() as c:
        orders=c.execute("SELECT COUNT(*) n FROM order_history WHERE phone LIKE ?",("%"+phone[-10:] ,)).fetchone()["n"]
        delivered=c.execute("SELECT COUNT(*) n FROM order_history WHERE phone LIKE ? AND lower(COALESCE(delivery_status,'')) LIKE '%deliver%'",("%"+phone[-10:],)).fetchone()["n"]
        returned=c.execute("SELECT COUNT(*) n FROM return_history WHERE phone LIKE ?",("%"+phone[-10:],)).fetchone()["n"]
    courier=request.args.get("courier","").strip()[:50]
    try: settings=json.loads(get_setting("courier_settings") or "{}")
    except Exception: settings={}
    if not isinstance(settings,dict): settings={}
    if courier and courier.lower() not in ("all","সব"): targets=[courier]
    else: targets=[k for k,v in settings.items() if isinstance(v,dict) and v.get("enabled") and str(v.get("fraud_path","")).strip()]
    def ask(name):
        cfg=settings.get(name,{}) or {}; path=str(cfg.get("fraud_path","")).strip()
        if not path or not cfg.get("enabled"): return name,None
        try:
            endpoint=(cfg.get("base_url","").rstrip("/")+"/"+path.lstrip("/")).replace("{phone}",requests.utils.quote(phone,safe=""))
            auth_type=cfg.get("auth_type","steadfast"); headers={"Content-Type":"application/json"}
            if auth_type=="steadfast": headers.update({"Api-Key":cfg.get("api_key",""),"Secret-Key":cfg.get("secret_key","")})
            elif auth_type=="bearer": headers["Authorization"]="Bearer "+cfg.get("api_key","")
            elif auth_type=="api-key": headers["Api-Key"]=cfg.get("api_key","")
            response=requests.get(endpoint,headers=headers,timeout=12); response.raise_for_status(); return name,response.json()
        except Exception as exc: return name,{"error":str(exc)[:180]}
    providers=[]
    if targets:
        with ThreadPoolExecutor(max_workers=min(6,len(targets))) as ex: results=list(ex.map(ask,targets))
        providers=[{"courier":n,"result":r} for n,r in results if r is not None]
    checked_at=time.time(); check_id=None
    with db() as c:
        for pv in (providers or [{"courier":"" if courier.lower() in ("","all","সব") else courier,"result":None}]):
            check_id=c.execute("INSERT INTO courier_fraud_checks(phone,courier,orders,delivered,returned,provider_result,checked_by,checked_at) VALUES(?,?,?,?,?,?,?,?)",(phone,pv["courier"],orders,delivered,returned,json.dumps(pv["result"],ensure_ascii=False) if pv["result"] is not None else "",session.get("user",""),checked_at)).lastrowid
    first=providers[0]["result"] if providers else None
    return jsonify(check_id=check_id,phone=phone,orders=orders,delivered=delivered,returned=returned,return_rate=returned/max(1,orders)*100,provider_check=first,providers=providers,checked_couriers=targets)

@app.get("/api/admin/courier-fraud-history")
@login_required
def courier_fraud_history():
    if not is_admin_level() and not (has_perm("courier_view") or has_perm("courier_update")): return jsonify(error="permission_denied"),403
    phone=normalize_order_phone(request.args.get("phone","")).strip()
    try: limit=max(1,min(100,int(request.args.get("limit","20") or 20)))
    except ValueError: limit=20
    with db() as c:
        if phone: rows=c.execute("SELECT * FROM courier_fraud_checks WHERE phone LIKE ? ORDER BY checked_at DESC,id DESC LIMIT ?",("%"+phone[-10:],limit)).fetchall()
        else: rows=c.execute("SELECT * FROM courier_fraud_checks ORDER BY checked_at DESC,id DESC LIMIT ?",(limit,)).fetchall()
    out=[]
    for r in rows:
        d=dict(r)
        try: pr=json.loads(d.get("provider_result") or "null")
        except Exception: pr=None
        if isinstance(pr,dict):
            o=pr.get("data") if isinstance(pr.get("data"),dict) else pr
            brief=pr.get("error") and ("API error: "+str(pr.get("error"))) or ", ".join(f"{k}: {v}" for k,v in list(o.items())[:8] if v is not None and not isinstance(v,(dict,list)))
        else: brief=str(pr) if pr is not None else ""
        d["provider_brief"]=brief[:240]; d.pop("provider_result",None); out.append(d)
    return jsonify(checks=out)

@app.post("/api/admin/courier-tracking")
@login_required
def courier_tracking_update():
    if not is_admin_level() and not has_perm("courier_update"): return jsonify(error="permission_denied"),403
    j=request.get_json(silent=True) or {}; cons=str(j.get("cons_id","")).strip()[:120]; courier=str(j.get("courier","")).strip()[:50]; status=str(j.get("status","Pending")).strip()[:100]
    if not cons or not status: return jsonify(error="Consignment ID এবং status দিন"),400
    result=_save_courier_status(cons,courier,status,str(j.get("rating","")).strip()[:100],str(j.get("details","")).strip()[:1000],normalize_order_phone(j.get("phone","")),"manual")
    return jsonify(ok=True,**result)

def _save_courier_status(cons,courier,status,rating="",details="",phone="",source="api",api_amount=None,api_amount_field=""):
    now=time.time()
    with db() as c:
        prev=c.execute("SELECT status FROM courier_tracking WHERE cons_id=? AND courier=?",(cons,courier)).fetchone()
        order=c.execute("SELECT * FROM order_history WHERE cons_id=? ORDER BY id DESC LIMIT 1",(cons,)).fetchone()
        if not order: order=c.execute("SELECT * FROM order_history WHERE cons_id LIKE ? ORDER BY id DESC LIMIT 1",("%"+cons+"%",)).fetchone()
        if not order and phone: order=c.execute("SELECT * FROM order_history WHERE phone LIKE ? ORDER BY id DESC LIMIT 1",("%"+phone[-10:],)).fetchone()
        c.execute("INSERT INTO courier_tracking(cons_id,courier,status,rating,details,order_id,updated_by,updated_at,api_amount,api_amount_field) VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(cons_id,courier) DO UPDATE SET status=excluded.status,rating=excluded.rating,details=excluded.details,order_id=excluded.order_id,updated_by=excluded.updated_by,updated_at=excluded.updated_at,api_amount=COALESCE(excluded.api_amount,courier_tracking.api_amount),api_amount_field=CASE WHEN excluded.api_amount_field<>'' THEN excluded.api_amount_field ELSE courier_tracking.api_amount_field END",(cons,courier,status,rating,details,order["id"] if order else None,session.get("user",""),now,api_amount,api_amount_field))
        c.execute("INSERT OR IGNORE INTO courier_events(cons_id,courier,status,details,source,created_at) VALUES(?,?,?,?,?,?)",(cons,courier,status,details,source,now))
        if order:
            try: payload=json.loads(order["row_json"] or "{}")
            except Exception: payload={}
            if not isinstance(payload,dict): payload={}
            for key in list(payload):
                normalized=re.sub(r"[^a-z0-9]","",str(key).lower())
                if normalized in ("deliverystatus","status","courierstatus"): payload[key]=status
                elif normalized in ("rating","courierrating") and rating: payload[key]=rating
                elif normalized in ("consid","consignmentid","courierconsignmentid"): payload[key]=cons
            c.execute("UPDATE order_history SET delivery_status=?,rating=CASE WHEN ?<>'' THEN ? ELSE rating END,cons_id=CASE WHEN COALESCE(cons_id,'')='' THEN ? ELSE cons_id END,row_json=? WHERE id=?",(status,rating,rating,cons,json.dumps(payload,ensure_ascii=False),order["id"]))
        if not prev or prev["status"]!=status:
            low=status.lower(); kind="return" if "return" in low else "delivery" if "deliver" in low else "info"; title="রিটার্ন আপডেট" if kind=="return" else "ডেলিভারি সম্পন্ন" if kind=="delivery" else "কুরিয়ার স্ট্যাটাস আপডেট"
            c.execute("INSERT INTO notifications(title,message,kind,ref,created_at) VALUES(?,?,?,?,?)",(title,f"{cons} · {courier or 'Courier'} · {status}",kind,cons,now))
        # Keep Return History recoverable: if a record was manually deleted, the next
        # API sync will recreate it while avoiding duplicates for the same consignment/product.
        if order and "return" in status.lower() and not c.execute("SELECT 1 FROM return_history WHERE TRIM(cons_id)=TRIM(?) AND TRIM(product_code)=TRIM(?)",(cons,order["product_code"] or "")).fetchone():
            c.execute("INSERT INTO return_history(batch_id,source_file,row_number,assigned_to,cons_id,address,phone,customer_name,product_code,size,collected_amount,note,pathao_bill_status,return_reason,row_json,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",("courier:"+cons,"Courier tracking",0,order["assigned_to"],cons,order["address"],order["phone"],order["customer_name"],order["product_code"],order["size"],order["collectible_amount"],details,status,"Courier return status",order["row_json"],session.get("user",""),now))
    log_activity("courier_status",cons,json.dumps({"courier":courier,"status":status,"order_matched":bool(order)},ensure_ascii=False))
    return {"order_matched":bool(order),"report_saved":True}

@app.post("/api/admin/courier-sync")
@login_required
def courier_status_sync():
    if not is_admin_level() and not has_perm("courier_update"): return jsonify(error="permission_denied"),403
    j=request.get_json(silent=True) or {}; courier=str(j.get("courier","")).strip()[:50]; cons=str(j.get("cons_id","")).strip()[:120]
    try: settings=json.loads(get_setting("courier_settings") or "{}")
    except Exception: settings={}
    cfg=settings.get(courier,{})
    with db() as c: mapping=c.execute("SELECT * FROM courier_order_mappings WHERE courier=?",(courier,)).fetchone()
    if not cons or not cfg.get("enabled") or (cfg.get("auth_type","steadfast")!="none" and not cfg.get("api_key")) or (cfg.get("auth_type","steadfast")=="steadfast" and not cfg.get("secret_key")) or not mapping: return jsonify(error="এই courier-এর API চালু করুন এবং প্রয়োজনীয় API key/secret দিন"),400
    path=str(mapping["status_path"] or "/status_by_cid/{cons_id}")
    try:
        endpoint=cfg.get("base_url","").rstrip("/")+"/"+path.lstrip("/")
        endpoint=endpoint.replace("{cons_id}",requests.utils.quote(cons,safe=""))
        auth_type=cfg.get("auth_type","steadfast"); headers={"Content-Type":"application/json"}
        if auth_type=="steadfast": headers.update({"Api-Key":cfg["api_key"],"Secret-Key":cfg["secret_key"]})
        elif auth_type=="bearer": headers["Authorization"]="Bearer "+cfg["api_key"]
        elif auth_type=="api-key": headers["Api-Key"]=cfg["api_key"]
        response=requests.get(endpoint,headers=headers,timeout=15); response.raise_for_status(); body=response.json()
        if isinstance(body,dict) and isinstance(body.get("data"),dict): body=body["data"]
        status=str(body.get(cfg.get("status_field","delivery_status")) or body.get("delivery_status") or body.get("status") or body.get("tracking_status") or "").strip()
        if not status: return jsonify(error="Courier response-এ delivery status পাওয়া যায়নি"),502
        details=str(body.get("tracking_message") or body.get("message") or "Courier API sync")[:1000]
        amount_fields=("collected_amount","received_amount","cod_amount","collectable_amount","collectible_amount","total_amount","amount")
        api_amount=None; api_amount_field=""
        for field in amount_fields:
            if body.get(field) not in (None,""):
                try:
                    api_amount=float(re.sub(r"[^0-9.-]","",str(body[field]).replace(",",""))); api_amount_field=field; break
                except (TypeError,ValueError): pass
        result=_save_courier_status(cons,courier,status,str(body.get("rating") or ""),details,"","api",api_amount,api_amount_field)
        return jsonify(ok=True,status=status,**result)
    except requests.RequestException as exc: return jsonify(error="Courier API request failed: "+str(exc)[:220]),502
    except Exception as exc: return jsonify(error="Courier API response বুঝতে সমস্যা হয়েছে: "+str(exc)[:220]),502

@app.get("/api/theme")
def theme_get():
    try:
        theme=json.loads(get_setting("theme") or "{}")
        theme["language"] = theme.get("language", "en") if theme.get("language") in ("en", "bn") else "en"
        return jsonify(theme)
    except ValueError: return jsonify({})

@app.post("/api/admin/theme")
@auth
def theme_set():
    if not is_admin_level(): return jsonify(error="Admin permission required"), 403
    j = request.get_json(silent=True) or {}
    t = {k: (j[k] if re.fullmatch(r"#[0-9a-fA-F]{6}", str(j.get(k, ""))) else "") for k in ("acc", "bg", "ink", "head", "card")}
    t["font"] = j.get("font") if j.get("font") in ("Hind Siliguri", "Noto Sans Bengali", "Tiro Bangla", "Baloo Da 2") else ""
    t["language"] = j.get("language", "en") if j.get("language") in ("en", "bn") else "en"
    t["title"], t["tag"] = str(j.get("title", ""))[:60], str(j.get("tag", ""))[:80]
    lg = str(j.get("logo", "")); t["logo"] = lg if lg.startswith("data:image/") and len(lg) < 400000 else ""
    set_setting("theme", json.dumps(t)); return jsonify(ok=True)

@app.get("/api/me")
def me():
    u = current_user()
    if u:
        session["role"] = u["role"]  # DB role authoritative
    return jsonify(user=(u["username"] if u else None), role=(u["role"] if u else None), permissions=user_perms(u), guest_permissions=(guest_perms() if not u else {}), expires_at=(float(session.get("login_at",0))+9*60*60 if u else None))

@app.post("/api/login")
def login():
    j = request.get_json(silent=True) or {}
    with db() as c: u = c.execute("SELECT * FROM users WHERE username=?", (j.get("username", ""),)).fetchone()
    if u and check_password_hash(u["pw"], j.get("password", "")):
        session.clear(); session.permanent=True; session["user"], session["role"], session["login_at"] = u["username"], u["role"], time.time(); log_activity("login", "", "success"); return jsonify(ok=True, permissions=user_perms(u), expires_at=session["login_at"]+9*60*60)
    with db() as c: c.execute("INSERT INTO activity_logs(username,action,query,details,created_at) VALUES(?,?,?,?,?)", (j.get("username", "")[:100] or "unknown", "login_failed", "", "failed", time.time()))
    return jsonify(ok=False), 401

@app.get("/api/guest-perms")
def guest_perms_api():
    if not session.get("user"): return jsonify(error="login_required"), 401
    return jsonify(permissions=guest_perms())

@app.get("/api/admin/guest-perms")
@admin_only
def guest_perms_list():
    return jsonify(permissions=guest_perms())

@app.post("/api/admin/guest-perms")
@admin_only
def guest_perms_set():
    j = request.get_json(silent=True) or {}
    perms = {k: bool((j.get("permissions") or {}).get(k, False)) for k in GUEST_PERM_LABELS}
    set_setting("guest_perms", json.dumps(perms, ensure_ascii=False))
    log_activity("guest_perms_update", "", json.dumps(perms, ensure_ascii=False))
    return jsonify(ok=True, permissions=perms)

@app.get("/api/admin/users")
@admin_only
def users_list():
    with db() as c:
        rows = c.execute("""
            SELECT id,username,role,perms FROM users
            ORDER BY CASE WHEN role='super_admin' THEN 0 WHEN role='admin' THEN 1 WHEN role='staff' THEN 2 WHEN role='user' THEN 3 WHEN role='cs' THEN 4 ELSE 5 END, username
        """).fetchall()
    return jsonify(users=[{"id":r["id"],"username":r["username"],"role":r["role"],"permissions":user_perms(r)} for r in rows])

@app.post("/api/admin/users")
@admin_only
def user_add():
    j=request.get_json(silent=True) or {}
    username=str(j.get("username","")).strip()
    password=str(j.get("password",""))
    role=str(j.get("role","staff")).strip().lower()
    if not re.fullmatch(r"[A-Za-z0-9_.-]{2,50}", username): return jsonify(ok=False,error="ইউজারনেম ২-৫০ অক্ষর হতে হবে"),400
    if len(password)<4: return jsonify(ok=False,error="পাসওয়ার্ড কমপক্ষে ৪ অক্ষর দিন"),400
    if role == "user": role = "staff"
    if role not in ("admin","staff","cs"): return jsonify(ok=False,error="Role শুধু Admin, Staff বা CS হবে"),400
    if role == "admin" and session.get("role") != "super_admin":
        return jsonify(ok=False,error="শুধু Super Admin নতুন Admin বানাতে পারবে"),403
    perms={k:bool((j.get("permissions") or {}).get(k,STAFF_DEFAULTS.get(k,False))) for k in PERMS}
    if role=="admin": perms={k:True for k in PERMS}
    try:
        with db() as c:
            c.execute("INSERT INTO users(username,pw,role,perms) VALUES(?,?,?,?)",
                      (username,generate_password_hash(password),role,json.dumps(perms,ensure_ascii=False)))
    except sqlite3.IntegrityError:
        return jsonify(ok=False,error="এই ইউজারনেম আগে থেকেই আছে"),409
    log_activity("user_add", username, json.dumps({"role":role,"permissions":perms},ensure_ascii=False))
    return jsonify(ok=True)

@app.post("/api/admin/users/<int:uid>")
@admin_only
def user_edit(uid):
    j=request.get_json(silent=True) or {}
    new_role=str(j.get("role","staff")).strip().lower()
    if new_role=="user": new_role="staff"
    password=str(j.get("password",""))
    with db() as c:
        u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
        if not u: return jsonify(ok=False,error="ইউজার পাওয়া যায়নি"),404
        if u["role"]=="super_admin":
            return jsonify(ok=False,error="Super Admin account lock করা আছে"),403
        if new_role not in ("admin","staff","cs"): return jsonify(ok=False,error="Role শুধু Admin, Staff বা CS হবে"),400
        if u["role"]=="admin" and session.get("role")!="super_admin":
            return jsonify(ok=False,error="Admin account শুধু Super Admin manage করতে পারবে"),403
        if new_role=="admin" and session.get("role")!="super_admin":
            return jsonify(ok=False,error="শুধু Super Admin Admin role দিতে পারবে"),403
        perms={k:bool((j.get("permissions") or {}).get(k, STAFF_DEFAULTS.get(k,False))) for k in PERMS}
        if new_role=="admin": perms={k:True for k in PERMS}
        sets=["role=?","perms=?"]; vals=[new_role,json.dumps(perms,ensure_ascii=False)]
        if password:
            if len(password)<4: return jsonify(ok=False,error="পাসওয়ার্ড কমপক্ষে ৪ অক্ষর দিন"),400
            sets.append("pw=?"); vals.append(generate_password_hash(password))
        vals.append(uid)
        c.execute("UPDATE users SET "+", ".join(sets)+" WHERE id=?", tuple(vals))
    log_activity("user_edit", u["username"], json.dumps({"role":new_role,"password_changed":bool(password),"permissions":perms},ensure_ascii=False))
    return jsonify(ok=True)

@app.delete("/api/admin/users/<int:uid>")
@super_admin_only
def user_del(uid):
    with db() as c:
        u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
        if not u: return jsonify(ok=False,error="ইউজার পাওয়া যায়নি"),404
        if u["role"]=="super_admin":
            return jsonify(ok=False,error="Super Admin account মুছা যাবে না"),403
        if u["username"] == session.get("user"):
            return jsonify(ok=False,error="নিজের লগইন account মুছবেন না"),400
        if u["role"]=="admin" and session.get("role")!="super_admin":
            return jsonify(ok=False,error="Admin account শুধু Super Admin মুছতে পারবে"),403
        c.execute("DELETE FROM users WHERE id=?",(uid,))
    log_activity("user_delete",u["username"],"")
    return jsonify(ok=True)

@app.get("/api/admin/logs")
@admin_only
def logs_list():
    try: limit=max(1,min(2000,int(request.args.get("limit",1000))))
    except Exception: limit=1000
    username=str(request.args.get("username", "")).strip()[:100]
    action=str(request.args.get("action", "")).strip()[:60]
    sql="SELECT id,username,action,query,details,created_at FROM activity_logs WHERE 1=1"; params=[]
    if username: sql+=" AND username=?"; params.append(username)
    if action: sql+=" AND action=?"; params.append(action)
    sql+=" ORDER BY id DESC LIMIT ?"; params.append(limit)
    with db() as c: rows=c.execute(sql,tuple(params)).fetchall()
    return jsonify(logs=[dict(r) for r in rows])

@app.post("/api/activity/search")
def activity_search():
    j=request.get_json(silent=True) or {}; q=str(j.get("query","" )).strip()[:500]
    if q: log_activity("text_search", q, json.dumps({"results":int(j.get("results",0) or 0)},ensure_ascii=False))
    return jsonify(ok=True)

@app.post("/api/activity/ui")
def activity_ui():
    j=request.get_json(silent=True) or {}
    allowed_actions={"filter_size","filter_stock","sort_price","filter_brand","filter_sheet","product_select","product_view","quick_mode","link_copy","selection_clear","id_copy","name_copy","full_info_copy","share_product","similar_click"}
    action=str(j.get("action",""))[:60]
    if action not in allowed_actions: return jsonify(ok=False,error="invalid_action"),400
    log_activity(action, str(j.get("query",""))[:500], str(j.get("details",""))[:1000])
    return jsonify(ok=True)

@app.get("/api/my/searches")
@login_required
def my_searches():
    with db() as c:
        rows=c.execute("SELECT query,details,created_at FROM activity_logs WHERE username=? AND action='text_search' ORDER BY id DESC LIMIT 20",(session["user"],)).fetchall()
    return jsonify(searches=[dict(r) for r in rows])

@app.post("/api/logout")
def logout(): session.clear(); return jsonify(ok=True)

@app.post("/api/admin/password")
@admin_only
def admin_password_change():
    j = request.get_json(silent=True) or {}
    old = str(j.get("old_password", ""))
    new = str(j.get("new_password", ""))
    if len(new) < 6:
        return jsonify(ok=False, error="নতুন পাসওয়ার্ড কমপক্ষে ৬ অক্ষর দিন"), 400
    u = current_user()
    if not u or not check_password_hash(u["pw"], old):
        return jsonify(ok=False, error="বর্তমান পাসওয়ার্ড সঠিক নয়"), 401
    with db() as c:
        c.execute("UPDATE users SET pw=? WHERE id=?", (generate_password_hash(new), u["id"]))
    log_activity("admin_password_change", u["username"], "password_changed")
    return jsonify(ok=True)

@app.get("/api/admin/sources")
@auth
def src_list(): return jsonify(catalog=sources("catalog"), sheet=sources("sheet"), meta=meta())

@app.post("/api/admin/sources")
@auth
def src_add():
    j = request.get_json(silent=True) or {}; kind, url = j.get("kind"), (j.get("url") or "").strip()
    if kind not in ("catalog", "sheet") or not url.startswith("http"): return jsonify(ok=False, error="সঠিক লিংক দিন"), 400
    if kind == "sheet" and not re.search(r"/d/[\w-]+", url): return jsonify(ok=False, error="এটা গুগল শিটের লিংক না"), 400
    with db() as c: c.execute("INSERT INTO sources(kind,name,url) VALUES(?,?,?)", (kind, (j.get("name") or "").strip() or ("Sheet" if kind == "sheet" else "Store"), url))
    C["force"] = True; return jsonify(ok=True)

@app.delete("/api/admin/sources/<int:i>")
@super_admin_only
def src_del(i):
    with db() as c: c.execute("DELETE FROM sources WHERE id=?", (i,))
    C["force"] = True; return jsonify(ok=True)

@app.post("/api/admin/refresh")
@auth
def refresh(): C["force"] = True; return jsonify(ok=True)

@app.post("/api/admin/price")
@perm_auth("price_edit")
def price_set():
    j = request.get_json(silent=True) or {}; k = str(j.get("key", ""))[:500]
    if not k: return jsonify(ok=False, error="প্রোডাক্ট নির্বাচন করুন"), 400
    p = next((x for x in C.get("products", []) if x.get("store", "") + "|" + x.get("id", "") == k), None)
    if not p: return jsonify(ok=False, error="প্রোডাক্ট পাওয়া যায়নি"), 404
    current = prices().get(k, p.get("price", 0))
    raw = str(j.get("price", "")).strip()
    if raw == "":
        new_effective = num(str(p.get("price", 0))); action = "restore"
    else:
        new_effective = num(raw); action = "set"
    with db() as c:
        if raw == "": c.execute("DELETE FROM prices WHERE k=?", (k,))
        else: c.execute("INSERT OR REPLACE INTO prices(k,price) VALUES(?,?)", (k, new_effective))
        c.execute("INSERT INTO price_history(product_key,display_id,title,store,img,old_price,new_price,username,action,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                  (k, display_product_id(p), p.get("title", ""), p.get("store", ""), p.get("img", ""), current, new_effective, session.get("user") or "guest", action, time.time()))
    log_activity("price_edit", k, json.dumps({"old_price":current,"new_price":new_effective,"action":action,"display_id":display_product_id(p)}, ensure_ascii=False))
    return jsonify(ok=True, old_price=current, new_price=new_effective)

@app.get("/api/admin/price-history")
@admin_only
def price_history():
    try: limit=max(1,min(200,int(request.args.get("limit",100))))
    except Exception: limit=100
    with db() as c: rows=c.execute("SELECT * FROM price_history ORDER BY id DESC LIMIT ?",(limit,)).fetchall()
    return jsonify(history=[dict(r) for r in rows])

# ---------- STAFF / ACTIVITY FEATURES ----------
def low_stock_threshold():
    try:
        return max(1, min(99, int(get_setting("low_stock_threshold") or "2")))
    except Exception:
        return 2

def stock_state(p):
    sizes=p.get("sizes") or {}
    if not sizes:
        return "out", 0
    vals=[int(v or 0) for v in sizes.values() if isinstance(v,(int,float)) and v>0]
    if not vals:
        return "out", 0
    total=sum(vals); threshold=low_stock_threshold()
    return ("low" if min(vals)<=threshold else "in"), total

@app.post("/api/admin/stock-threshold")
@admin_only
def stock_threshold_set():
    j=request.get_json(silent=True) or {}
    try: v=max(1,min(99,int(j.get("threshold",2))))
    except Exception: return jsonify(ok=False,error="১-৯৯ এর মধ্যে threshold দিন"),400
    set_setting("low_stock_threshold",str(v))
    log_activity("stock_threshold",str(v),"")
    return jsonify(ok=True,threshold=v)

@app.get("/api/me/data")
def me_data():
    u = session.get("user")
    if not u: return jsonify(favorites=[], recent=[]), 401
    with db() as c:
        fav=[r["product_key"] for r in c.execute("SELECT product_key FROM favorites WHERE username=? ORDER BY created_at DESC",(u,)).fetchall()]
        recent=[dict(r) for r in c.execute("SELECT product_key, MAX(viewed_at) AS viewed_at FROM recent_views WHERE username=? GROUP BY product_key ORDER BY viewed_at DESC LIMIT 30",(u,)).fetchall()]
    return jsonify(favorites=fav, recent=recent)

@app.post("/api/favorite")
@login_required
def favorite_toggle():
    return jsonify(ok=False,error="favorite_disabled"),404
    if not is_admin_level() and not has_perm("favorite"): return jsonify(ok=False,error="permission_denied"),403
    j=request.get_json(silent=True) or {}; k=str(j.get("key",""))[:500]
    if not k: return jsonify(ok=False),400
    with db() as c:
        exists=c.execute("SELECT 1 FROM favorites WHERE username=? AND product_key=?",(session["user"],k)).fetchone()
        if exists:
            c.execute("DELETE FROM favorites WHERE username=? AND product_key=?",(session["user"],k)); on=False
        else:
            c.execute("INSERT OR IGNORE INTO favorites(username,product_key,created_at) VALUES(?,?,?)",(session["user"],k,time.time())); on=True
    log_activity("favorite_add" if on else "favorite_remove",k,"")
    return jsonify(ok=True,on=on)

@app.post("/api/recent")
@login_required
def recent_add():
    j=request.get_json(silent=True) or {}; k=str(j.get("key",""))[:500]
    if not k: return jsonify(ok=False),400
    p=next((x for x in C.get("products",[]) if x.get("store","")+"|"+x.get("id","")==k),None)
    with db() as c:
        c.execute("DELETE FROM recent_views WHERE username=? AND product_key=?",(session["user"],k))
        c.execute("INSERT INTO recent_views(username,product_key,viewed_at) VALUES(?,?,?)",(session["user"],k,time.time()))
        c.execute("DELETE FROM recent_views WHERE username=? AND id NOT IN (SELECT id FROM recent_views WHERE username=? ORDER BY viewed_at DESC LIMIT 30)",(session["user"],session["user"]))
    log_activity("product_view", display_product_id(p) if p else k, json.dumps({"key":k,"title":(p or {}).get("title","")}, ensure_ascii=False))
    return jsonify(ok=True)

@app.post("/api/report")
@login_required
def report_product():
    if not is_admin_level() and not has_perm("report"): return jsonify(ok=False,error="permission_denied"),403
    j=request.get_json(silent=True) or {}; k=str(j.get("key",""))[:500]; reason=str(j.get("reason","প্রোডাক্ট পাওয়া যাচ্ছে না"))[:500]
    if not k: return jsonify(ok=False),400
    with db() as c: c.execute("INSERT INTO product_reports(username,product_key,reason,created_at,status) VALUES(?,?,?,?,?)",(session["user"],k,reason,time.time(),"open"))
    log_activity("product_report",k,reason); return jsonify(ok=True)

@app.get("/api/admin/dashboard")
@admin_only
def dashboard():
    now=time.time(); day=now-86400
    with db() as c:
        total=c.execute("SELECT COUNT(*) n FROM activity_logs").fetchone()["n"]
        today=c.execute("SELECT COUNT(*) n FROM activity_logs WHERE created_at>=?",(day,)).fetchone()["n"]
        searches=c.execute("SELECT COUNT(*) n FROM activity_logs WHERE created_at>=? AND action IN ('text_search','image_search')",(day,)).fetchone()["n"]
        downloads=c.execute("SELECT COUNT(*) n FROM activity_logs WHERE created_at>=? AND action='download'",(day,)).fetchone()["n"]
        copies=c.execute("SELECT COUNT(*) n FROM activity_logs WHERE created_at>=? AND action='copy'",(day,)).fetchone()["n"]
        reports=c.execute("SELECT COUNT(*) n FROM product_reports WHERE status='open'").fetchone()["n"]
        users=c.execute("SELECT COUNT(*) n FROM users WHERE role IN ('admin','staff','user','cs')").fetchone()["n"]
        byuser=c.execute("SELECT username,COUNT(*) n FROM activity_logs WHERE created_at>=? GROUP BY username ORDER BY n DESC LIMIT 20",(day,)).fetchall()
        ar=c.execute("SELECT action,COUNT(*) n FROM activity_logs WHERE created_at>=? GROUP BY action",(day,)).fetchall()
        action_stats={r["action"]:r["n"] for r in ar}
        topq=c.execute("SELECT query,COUNT(*) n FROM activity_logs WHERE created_at>=? AND action IN ('text_search','image_search') AND TRIM(query)<>'' GROUP BY query ORDER BY n DESC LIMIT 10",(now-7*86400,)).fetchall()
        daily=[]
        for i in range(6,-1,-1):
            ds=now-(i+1)*86400; de=now-i*86400
            n=c.execute("SELECT COUNT(*) n FROM activity_logs WHERE created_at>=? AND created_at<?",(ds,de)).fetchone()["n"]
            label=time.strftime("%d/%m",time.localtime(de))
            daily.append({"label":label,"n":n})
    low=[]; out=[]; threshold=low_stock_threshold()
    for p in view():
        state,stock_total=stock_state(p)
        if state=="low": low.append({"key":p["key"],"id":p.get("display_id") or "", "title":p.get("title",""),"total":stock_total,"sizes":p.get("sizes") or {}})
        elif state=="out": out.append({"key":p["key"],"id":p.get("display_id") or "", "title":p.get("title","")})
    return jsonify(total=total,today=today,searches=searches,downloads=downloads,copies=copies,reports=reports,users=users,byuser=[dict(r) for r in byuser],action_stats=action_stats,top_queries=[dict(r) for r in topq],daily=daily,total_products=len(C.get("products",[])),ai_ready=P.get("ai_ready",False),stage=P.get("stage",""),low_stock_threshold=threshold,low_stock=low[:50],out_of_stock=out[:50])

@app.get("/api/admin/reports")
@admin_only
def reports_list():
    with db() as c: rows=c.execute("SELECT * FROM product_reports ORDER BY id DESC LIMIT 200").fetchall()
    return jsonify(reports=[dict(r) for r in rows])


@app.get("/api/admin/business-dashboard")
@admin_only
def business_dashboard():
    """Financial/operations summary for the current Dhaka day and month.

    Profit is explicitly an estimate and only calculated for POS item units that
    can be matched to recorded purchase cost by product-code/size.
    """
    now_dt = datetime.now(DHAKA_TZ)
    today = now_dt.date().isoformat()
    month_start_date = now_dt.replace(day=1).date().isoformat()
    today_start = now_dt.replace(hour=0, minute=0, second=0, microsecond=0).timestamp()
    tomorrow_start = (now_dt.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)).timestamp()
    month_start_ts = now_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0).timestamp()
    with db() as c:
        sales_today = dict(c.execute("SELECT COUNT(*) n, COALESCE(SUM(total),0) total, COALESCE(SUM(paid),0) paid, COALESCE(SUM(discount),0) discount FROM pos_sales WHERE created_at>=? AND created_at<?", (today_start, tomorrow_start)).fetchone())
        sales_month = dict(c.execute("SELECT COUNT(*) n, COALESCE(SUM(total),0) total, COALESCE(SUM(paid),0) paid, COALESCE(SUM(discount),0) discount, COALESCE(SUM(delivery_charge),0) delivery FROM pos_sales WHERE created_at>=?", (month_start_ts,)).fetchone())
        all_due = c.execute("SELECT COALESCE(SUM(CASE WHEN total>paid THEN total-paid ELSE 0 END),0) due FROM pos_sales").fetchone()["due"]
        by_company = [dict(r) for r in c.execute("SELECT COALESCE(company_key,'alien_tees') company_key, COUNT(*) orders, COALESCE(SUM(total),0) sales, COALESCE(SUM(paid),0) collected FROM pos_sales WHERE created_at>=? GROUP BY COALESCE(company_key,'alien_tees') ORDER BY sales DESC", (month_start_ts,)).fetchall()]
        exp_month = dict(c.execute("SELECT COALESCE(SUM(CASE WHEN status='Approved' THEN amount ELSE 0 END),0) approved, COALESCE(SUM(CASE WHEN status='Pending' THEN amount ELSE 0 END),0) pending, SUM(CASE WHEN status='Pending' THEN 1 ELSE 0 END) pending_count FROM erp_expenses WHERE expense_date>=? AND expense_date<=?", (month_start_date, today)).fetchone())
        purchase_month = dict(c.execute("SELECT COUNT(*) n, COALESCE(SUM(total),0) total FROM erp_purchases WHERE purchase_date>=? AND purchase_date<=?", (month_start_date, today)).fetchone())
        suppliers = c.execute("""SELECT s.id,s.name,COALESCE(s.opening_balance,0)+COALESCE(p.total,0)-COALESCE(pm.total,0) due
            FROM erp_suppliers s
            LEFT JOIN (SELECT supplier_id,SUM(total) total FROM erp_purchases GROUP BY supplier_id) p ON p.supplier_id=s.id
            LEFT JOIN (SELECT supplier_id,SUM(amount) total FROM erp_supplier_payments GROUP BY supplier_id) pm ON pm.supplier_id=s.id
            ORDER BY due DESC""").fetchall()
        supplier_due_rows = [dict(r) for r in suppliers if float(r["due"] or 0) > 0.009]
        supplier_due_count = len(supplier_due_rows)
        supplier_due_total = sum(float(r["due"] or 0) for r in supplier_due_rows)
        supplier_due_rows = supplier_due_rows[:10]
        # Historic average unit cost per product/size. Purchases dated in the future are excluded.
        cost_rows = c.execute("""SELECT lower(trim(pi.item_code)) code, lower(trim(pi.size)) size,
            SUM(pi.quantity) qty, SUM(pi.line_total) amount
            FROM erp_purchase_items pi JOIN erp_purchases p ON p.id=pi.purchase_id
            WHERE pi.quantity>0 AND p.purchase_date<=? AND COALESCE(p.stock_received,0)=1
            GROUP BY lower(trim(pi.item_code)),lower(trim(pi.size))""", (today,)).fetchall()
        avg_cost = {}
        for r in cost_rows:
            qty = float(r["qty"] or 0)
            if qty > 0:
                avg_cost[(_norm_key(r["code"]), _norm_key(r["size"]))] = float(r["amount"] or 0) / qty
        month_sales_rows = c.execute("SELECT items,COALESCE(subtotal,0) subtotal,COALESCE(discount,0) discount FROM pos_sales WHERE created_at>=?", (month_start_ts,)).fetchall()
        due_customers = [dict(r) for r in c.execute("""SELECT customer_name name,phone,COUNT(*) orders,
            SUM(CASE WHEN total>paid THEN total-paid ELSE 0 END) due
            FROM pos_sales WHERE total>paid GROUP BY phone,customer_name HAVING due>0
            ORDER BY due DESC LIMIT 8""").fetchall()]
        due_customer_count = int(c.execute("""SELECT COUNT(*) FROM (
            SELECT phone,customer_name FROM pos_sales WHERE total>paid
            GROUP BY phone,customer_name HAVING SUM(total-paid)>0.009)""").fetchone()[0])
        pending_expenses = [dict(r) for r in c.execute("SELECT expense_no,title,amount,expense_date FROM erp_expenses WHERE status='Pending' ORDER BY expense_date LIMIT 8").fetchall()]
    # Link catalogue keys to real product codes where possible, for cost coverage.
    catalog_by_key = {str(p.get("store", "") + "|" + p.get("id", "")): p for p in C.get("products", [])}
    product_agg = {}
    total_units = matched_units = 0
    matched_revenue = cogs = 0.0
    total_item_gross = matched_item_gross = 0.0
    for sale_row in month_sales_rows:
        try: sale_items = json.loads(sale_row["items"] or "[]")
        except Exception: sale_items = []
        sale_items = [x for x in sale_items if isinstance(x, dict)] if isinstance(sale_items, list) else []
        item_gross_total = sum(max(0, int(_num(x.get("qty")))) * max(0, _num(x.get("price"))) for x in sale_items)
        sale_discount = max(0.0, _num(sale_row["discount"]))
        for item in sale_items:
            qty = max(0, int(_num(item.get("qty"))))
            price = max(0, _num(item.get("price")))
            if qty <= 0: continue
            line_gross = qty * price
            total_item_gross += line_gross
            # Allocate invoice discount proportionally to each item's gross value.
            line_discount = min(line_gross, sale_discount * line_gross / item_gross_total) if item_gross_total > 0 else 0.0
            line_net = max(0.0, line_gross - line_discount)
            total_units += qty
            code_candidates = [str(item.get("id") or ""), str(item.get("code") or "")]
            p = catalog_by_key.get(str(item.get("key") or ""))
            if p:
                code_candidates.extend([str(p.get("id") or ""), str(display_product_id(p) or "")])
                try: code_candidates.append(str(_pos_code_for(p) or ""))
                except Exception: pass
            szkey = _norm_key(item.get("size") or "")
            unit_cost = None
            for code in code_candidates:
                ck = (_norm_key(code), szkey)
                if code and ck in avg_cost:
                    unit_cost = avg_cost[ck]; break
            name = str(item.get("title") or item.get("id") or "Product")
            pk = _norm_key(item.get("id") or name)
            agg = product_agg.setdefault(pk, {"name": name, "code": str(item.get("id") or ""), "units": 0, "sales": 0.0, "_stock_by_size": {}})
            agg["units"] += qty; agg["sales"] += qty*price
            if p:
                stock_key = str(p.get("store") or "") + "|" + str(p.get("id") or "")
                size_stock = (C.get("stk", {}).get(stock_key) or {}).get("sizes") or {}
                wanted_size = normalize_stock_size(item.get("size") or "")
                stock_qty = next((float(_num(v)) for k, v in size_stock.items() if normalize_stock_size(k).casefold() == wanted_size.casefold()), 0.0)
                agg["_stock_by_size"][_norm_key(wanted_size)] = max(0.0, stock_qty)
            if unit_cost is not None:
                matched_units += qty
                matched_item_gross += line_gross
                matched_revenue += line_net
                cogs += qty*unit_cost
    coverage = round(100*matched_units/total_units, 1) if total_units else 0.0
    sales_coverage = round(100*matched_item_gross/total_item_gross, 1) if total_item_gross else 0.0
    profit_estimate = round(matched_revenue-cogs, 2) if matched_units else None
    # Do not present net profit as if complete when part of sold item cost is missing.
    net_estimate = round(profit_estimate-float(exp_month.get("approved") or 0), 2) if matched_units and coverage >= 99.9 else None
    days_elapsed = max(1, int(now_dt.day))
    for product in product_agg.values():
        stock_on_hand = round(sum((product.pop("_stock_by_size", {}) or {}).values()))
        avg_daily = float(product.get("units") or 0) / days_elapsed
        product["stock_on_hand"] = stock_on_hand
        product["avg_daily_units"] = round(avg_daily, 2)
        product["estimated_days_left"] = round(stock_on_hand / avg_daily, 1) if avg_daily > 0 else None
    top_products = sorted(product_agg.values(), key=lambda x: (-x["units"], -x["sales"]))[:8]
    # Build actionable alerts from current operational state, not only KPI counters.
    smart_alerts = []
    try:
        current_products = view()
        low_count = out_count = 0
        for product in current_products:
            try:
                state, _qty = stock_state(product)
                low_count += int(state == "low")
                out_count += int(state == "out")
            except Exception:
                continue
        if out_count:
            smart_alerts.append({"level":"high","title":"Out-of-stock products","detail":f"{out_count} product(s) have no available stock.","panel":"stockrecon"})
        if low_count:
            smart_alerts.append({"level":"medium","title":"Low stock warning","detail":f"{low_count} product(s) are below the low-stock threshold.","panel":"stockrecon"})
    except Exception:
        low_count = out_count = 0
    if int(exp_month.get("pending_count") or 0):
        smart_alerts.append({"level":"medium","title":"Pending expense approvals","detail":f"{int(exp_month.get('pending_count') or 0)} expense(s) need review.","panel":"expenses"})
    if due_customer_count:
        smart_alerts.append({"level":"medium","title":"Customer dues outstanding","detail":f"{due_customer_count} customer account(s) have unpaid POS balances.","panel":"posorders"})
    if supplier_due_count:
        smart_alerts.append({"level":"medium","title":"Supplier balance due","detail":f"{supplier_due_count} supplier(s) have an outstanding balance.","panel":"suppliers"})
    if total_units and coverage < 99.9:
        smart_alerts.append({"level":"medium","title":"Purchase-cost data incomplete","detail":f"Cost is matched for {coverage}% of sold units ({sales_coverage}% of item sales by value); net profit is hidden until coverage is complete.","panel":"purchases"})
    try:
        with db() as c:
            unresolved = int(c.execute("""SELECT COUNT(*) FROM erp_stock_reconciliations r
                WHERE r.action='counted' AND r.variance<>0 AND NOT EXISTS (
                    SELECT 1 FROM erp_stock_reconciliations x WHERE x.sheet=r.sheet AND x.row_number=r.row_number
                    AND x.product_code=r.product_code AND x.size=r.size AND x.id>r.id)""").fetchone()[0])
        if unresolved:
            smart_alerts.append({"level":"high","title":"Unapplied stock-count variances","detail":f"{unresolved} count difference(s) have not been applied to the stock workbook.","panel":"stockrecon"})
    except Exception:
        unresolved = 0
    return jsonify(ok=True, as_of=now_dt.strftime("%Y-%m-%d %H:%M"), month_start=month_start_date,
        today_sales=round(float(sales_today.get("total") or 0),2), today_orders=int(sales_today.get("n") or 0),
        today_collected=round(float(sales_today.get("paid") or 0),2), month_sales=round(float(sales_month.get("total") or 0),2),
        month_orders=int(sales_month.get("n") or 0), month_collected=round(float(sales_month.get("paid") or 0),2),
        month_discount=round(float(sales_month.get("discount") or 0),2), month_delivery=round(float(sales_month.get("delivery") or 0),2),
        customer_due=round(float(all_due or 0),2), month_purchases=round(float(purchase_month.get("total") or 0),2),
        purchase_count=int(purchase_month.get("n") or 0), approved_expenses=round(float(exp_month.get("approved") or 0),2),
        pending_expenses=round(float(exp_month.get("pending") or 0),2), pending_expense_count=int(exp_month.get("pending_count") or 0),
        supplier_due=round(supplier_due_total,2), supplier_due_count=supplier_due_count, suppliers_due=supplier_due_rows,
        top_customer_due=due_customers, by_company=by_company, top_products=top_products, units_sold=total_units,
        cost_coverage=coverage, cost_coverage_sales=sales_coverage, covered_sales=round(matched_revenue,2), estimated_cogs=round(cogs,2) if matched_units else None,
        estimated_gross_profit=profit_estimate, estimated_net_after_approved_expenses=net_estimate,
        profit_note="Gross profit is only for sales items matched to received-purchase cost; sale discount is allocated proportionally. Net estimate is shown only at 99.9%+ unit cost coverage. Returns, other channels, unrecorded costs and tax may not be included.",
        alerts={"pending_expense_count":int(exp_month.get("pending_count") or 0), "due_customer_count":due_customer_count, "due_supplier_count":supplier_due_count, "low_stock_count":int(low_count), "out_of_stock_count":int(out_count), "unresolved_stock_variances":int(unresolved)}, smart_alerts=smart_alerts)


# ---------- Stock Reconciliation / Physical Count ----------
def _stock_recon_allowed(apply_changes=False):
    if is_admin_level():
        return True
    perms = user_perms()
    if apply_changes:
        return bool(perms.get("stock_manage") or perms.get("stock_set"))
    return any(perms.get(k, False) for k in ("stock_dashboard", "stock_manage", "stock_set", "stock_add", "stock_reduce", "stock_history_view"))

@app.get("/api/admin/stock-reconciliation")
@login_required
def stock_reconciliation_list():
    if not _stock_recon_allowed(False):
        return jsonify(ok=False, error="permission_denied"), 403
    if not os.path.isfile(EXCEL_STOCK_FILE):
        return jsonify(ok=False, error="Excel stock workbook is not available. Connect the stock workbook first."), 404
    q = str(request.args.get("q") or "").strip().casefold()[:160]
    sheet_filter = str(request.args.get("sheet") or "").strip()[:120]
    try:
        limit = max(1, min(300, int(request.args.get("limit", 150))))
        offset = max(0, min(1000000, int(request.args.get("offset", 0))))
    except (TypeError, ValueError):
        return jsonify(ok=False, error="Invalid pagination."), 400
    catalog = excel_catalog_index()
    rows = []; sheets = []
    with EXCEL_LOCK:
        wb = load_workbook(EXCEL_STOCK_FILE, read_only=True, data_only=True)
        for ws in wb.worksheets:
            headers = excel_headers(ws)
            if not headers or not headers[0]:
                continue
            sheets.append(ws.title)
            if sheet_filter and ws.title != sheet_filter:
                continue
            # Headers after the code column are counted as size columns unless they look like metadata.
            size_columns = [(i, normalize_stock_size(h)) for i, h in enumerate(headers[1:], start=1)
                            if h and h.strip().casefold() not in SKIPH and not any(k in h.casefold() for k in ("image", "photo", "link", "url", "name", "title", "price", "cost", "amount", "total"))]
            for rn, cells in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
                code = excel_cell_text(cells[0]).strip() if cells else ""
                if not code:
                    continue
                info = catalog.get(code.casefold()) or {}
                title = str(info.get("title") or code)
                for col_idx, size in size_columns:
                    if col_idx >= len(cells):
                        raw_qty = 0
                    else:
                        v = cells[col_idx].value
                        try: raw_qty = int(float(str(v or "0").replace(",", "")))
                        except (TypeError, ValueError): raw_qty = 0
                    probe = f"{ws.title} {code} {title} {size}".casefold()
                    if q and q not in probe:
                        continue
                    rows.append({"sheet":ws.title,"row_number":rn,"code":code,"title":title,"size":size,"system_qty":raw_qty})
        wb.close()
    total = len(rows)
    page = rows[offset:offset+limit]
    with db() as c:
        history = [dict(r) for r in c.execute("SELECT id,sheet,row_number,product_code,size,system_qty,counted_qty,variance,action,note,username,created_at FROM erp_stock_reconciliations ORDER BY id DESC LIMIT 25").fetchall()]
        unresolved = int(c.execute("""SELECT COUNT(*) FROM erp_stock_reconciliations r
            WHERE r.action='counted' AND r.variance<>0 AND NOT EXISTS (
                SELECT 1 FROM erp_stock_reconciliations x WHERE x.sheet=r.sheet AND x.row_number=r.row_number
                AND x.product_code=r.product_code AND x.size=r.size AND x.id>r.id)""").fetchone()[0])
    return jsonify(ok=True, sheets=sheets, rows=page, total=total, offset=offset, limit=limit,
                   has_more=offset+len(page)<total, unresolved_variances=unresolved, history=history)

@app.post("/api/admin/stock-reconciliation")
@login_required
def stock_reconciliation_save():
    data = request.get_json(silent=True) or {}
    action = str(data.get("action") or "counted").strip().lower()
    if action not in ("counted", "apply"):
        return jsonify(ok=False, error="Action must be counted or apply."), 400
    if not _stock_recon_allowed(action == "apply"):
        return jsonify(ok=False, error="You need stock management / direct-set permission to apply a counted quantity."), 403
    sheet = str(data.get("sheet") or "").strip()[:120]
    code = str(data.get("code") or "").strip()[:160]
    size = normalize_stock_size(data.get("size") or "")[:100]
    try:
        row_number = int(data.get("row_number")); counted = int(data.get("counted_qty"))
    except (TypeError, ValueError):
        return jsonify(ok=False, error="Row and counted quantity must be whole numbers."), 400
    if not sheet or not code or not size or row_number < 2 or counted < 0 or counted > 100000000:
        return jsonify(ok=False, error="Sheet, code, size and a non-negative counted quantity are required."), 400
    note = str(data.get("note") or "").strip()[:1000]
    username = str(session.get("user") or "unknown")
    if not os.path.isfile(EXCEL_STOCK_FILE):
        return jsonify(ok=False, error="Excel stock workbook is not available."), 404
    try:
        with EXCEL_LOCK:
            wb = load_workbook(EXCEL_STOCK_FILE)
            if sheet not in wb.sheetnames:
                wb.close(); return jsonify(ok=False, error="Stock sheet not found."), 404
            ws = wb[sheet]
            headers = excel_headers(ws)
            col_idx = next((i+1 for i,h in enumerate(headers) if i > 0 and normalize_stock_size(h).casefold()==size.casefold()), None)
            if not col_idx or row_number > ws.max_row or excel_cell_text(ws.cell(row_number,1)).strip().casefold()!=code.casefold():
                wb.close(); return jsonify(ok=False, error="The stock row changed since it was loaded. Refresh and try again."), 409
            try: system_qty = int(float(ws.cell(row_number,col_idx).value or 0))
            except (TypeError,ValueError): system_qty = 0
            variance = counted - system_qty
            if action == "apply" and variance:
                snapshot_excel_stock("before-reconciliation", username)
                ws.cell(row_number,col_idx).value = counted
                fd, tmp = tempfile.mkstemp(prefix="stock-recon-", suffix=".xlsx", dir=INSTANCE); os.close(fd)
                try:
                    wb.save(tmp); wb.close(); os.replace(tmp, EXCEL_STOCK_FILE); tmp = ""
                finally:
                    if tmp and os.path.exists(tmp): os.remove(tmp)
            else:
                wb.close()
        now = time.time()
        with db() as c:
            c.execute("INSERT INTO erp_stock_reconciliations(sheet,row_number,product_code,size,system_qty,counted_qty,variance,action,note,username,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                (sheet,row_number,code,size,system_qty,counted,variance,action,note,username,now))
            if action == "apply" and variance:
                c.execute("INSERT INTO excel_stock_history(sheet,row_number,product_code,size,old_qty,new_qty,username,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
                    (sheet,row_number,code,size,system_qty,counted,username,"Stock reconciliation applied"+(": "+note if note else ""),now))
                c.execute("INSERT INTO excel_stock_batches(username,change_count,plus_units,minus_units,created_at) VALUES(?,?,?,?,?)",
                    (username,1,max(0,counted-system_qty),max(0,system_qty-counted),now))
        if action == "apply" and variance:
            refresh_stock_from_excel()
        try: log_activity("stock_reconciliation_"+action, code, f"{sheet} / {size}: {system_qty} -> {counted}; variance {variance}; {note}")
        except Exception: pass
        return jsonify(ok=True, action=action, sheet=sheet, code=code, size=size, system_qty=system_qty,
                       counted_qty=counted, variance=variance, stock_updated=bool(action=="apply" and variance))
    except Exception as exc:
        print("stock reconciliation failed:", repr(exc))
        return jsonify(ok=False, error="Stock reconciliation could not be saved: "+str(exc)[:250]), 500


# ---------- Daily Cash Closing ----------
def _cash_close_date(value):
    value = str(value or "").strip()
    if not value:
        return datetime.now(DHAKA_TZ).date().isoformat()
    try:
        return datetime.strptime(value, "%Y-%m-%d").date().isoformat()
    except ValueError:
        return ""

def _cash_close_amounts(closing_date, opening_cash):
    day = datetime.strptime(closing_date, "%Y-%m-%d").replace(tzinfo=DHAKA_TZ)
    start = day.timestamp()
    end = (day + timedelta(days=1)).timestamp()
    with db() as c:
        sales = c.execute("""SELECT COALESCE(SUM(CASE WHEN paid>total THEN total ELSE paid END),0) amount,
            COUNT(*) count FROM pos_sales WHERE created_at>=? AND created_at<? AND lower(trim(COALESCE(method,'')))='cash'""", (start, end)).fetchone()
        expenses = c.execute("""SELECT COALESCE(SUM(amount),0) amount,COUNT(*) count FROM erp_expenses
            WHERE expense_date=? AND lower(trim(COALESCE(payment_method,'')))='cash' AND lower(trim(COALESCE(status,'')))='approved'""", (closing_date,)).fetchone()
        supplier_payments = c.execute("""SELECT COALESCE(SUM(amount),0) amount,COUNT(*) count FROM erp_supplier_payments
            WHERE payment_date=? AND lower(trim(COALESCE(payment_method,'')))='cash'""", (closing_date,)).fetchone()
    cash_sales = round(float(sales["amount"] or 0), 2)
    cash_expenses = round(float(expenses["amount"] or 0), 2)
    cash_supplier_payments = round(float(supplier_payments["amount"] or 0), 2)
    opening = round(float(opening_cash or 0), 2)
    expected = round(opening + cash_sales - cash_expenses - cash_supplier_payments, 2)
    return {"cash_sales":cash_sales,"cash_sale_count":int(sales["count"] or 0),
            "cash_expenses":cash_expenses,"cash_expense_count":int(expenses["count"] or 0),
            "cash_supplier_payments":cash_supplier_payments,"cash_supplier_payment_count":int(supplier_payments["count"] or 0),
            "opening_cash":opening,"expected_cash":expected}

@app.get("/api/admin/cash-closing")
@admin_only
def get_cash_closing():
    closing_date = _cash_close_date(request.args.get("date"))
    if not closing_date:
        return jsonify(ok=False,error="Invalid date. Use YYYY-MM-DD."),400
    with db() as c:
        row = c.execute("SELECT * FROM erp_cash_closings WHERE closing_date=?",(closing_date,)).fetchone()
        prev = c.execute("SELECT actual_cash,closing_date FROM erp_cash_closings WHERE closing_date<? ORDER BY closing_date DESC LIMIT 1",(closing_date,)).fetchone()
    opening = float(row["opening_cash"] if row else (prev["actual_cash"] if prev else 0))
    values = _cash_close_amounts(closing_date, opening)
    saved = dict(row) if row else None
    return jsonify(ok=True,date=closing_date,previous_close_date=(prev["closing_date"] if prev else ""),
        saved=bool(row),actual_cash=round(float(row["actual_cash"] or 0),2) if row else 0,
        difference=round(float(row["actual_cash"] or 0)-values["expected_cash"],2) if row else 0,
        notes=str(row["notes"] or "") if row else "",closed_by=str(row["closed_by"] or "") if row else "",
        **values)

@app.post("/api/admin/cash-closing")
@admin_only
def save_cash_closing():
    data = request.get_json(silent=True) or {}
    closing_date = _cash_close_date(data.get("date"))
    if not closing_date:
        return jsonify(ok=False,error="Invalid date. Use YYYY-MM-DD."),400
    try:
        opening = float(data.get("opening_cash",0) or 0)
        actual = float(data.get("actual_cash",0) or 0)
    except (TypeError,ValueError):
        return jsonify(ok=False,error="Opening and actual cash must be valid numbers."),400
    if not all(map(lambda n: n == n and abs(n) < 100000000000, (opening,actual))) or opening < 0 or actual < 0:
        return jsonify(ok=False,error="Cash amounts must be non-negative valid numbers."),400
    notes = str(data.get("notes") or "").strip()[:2000]
    values = _cash_close_amounts(closing_date, opening)
    difference = round(actual-values["expected_cash"],2)
    now = time.time(); who = str(session.get("user") or "admin")
    with db() as c:
        existing = c.execute("SELECT id FROM erp_cash_closings WHERE closing_date=?",(closing_date,)).fetchone()
        if existing:
            c.execute("""UPDATE erp_cash_closings SET opening_cash=?,expected_cash=?,actual_cash=?,difference=?,notes=?,closed_by=?,updated_at=? WHERE closing_date=?""",
                (round(opening,2),values["expected_cash"],round(actual,2),difference,notes,who,now,closing_date))
        else:
            c.execute("""INSERT INTO erp_cash_closings(closing_date,opening_cash,expected_cash,actual_cash,difference,notes,closed_by,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?)""",(closing_date,round(opening,2),values["expected_cash"],round(actual,2),difference,notes,who,now,now))
    try: log_activity("cash_closing_save",closing_date,json.dumps({"expected":values["expected_cash"],"actual":round(actual,2),"difference":difference},ensure_ascii=False))
    except Exception: pass
    return jsonify(ok=True,date=closing_date,actual_cash=round(actual,2),difference=difference,**values)


LOGO = {"a": base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAG0AAAB4CAYAAAD443x+AAAeg0lEQVR42u19e3QcxZX+V90tWZJly/LbGFsazYxkzCOACY+wIDshgYAxj2DADzbw2yTsZrMQNvzYnN1kOWR3k+wvZEPgQDbZbM4uCzGReYZ3AuuIR8BgY0LAPGwL2xgbZGtGT1ua6arv98dUyaV290iWRrLsqM/xGVmq6a6uW9+te7+69xYwyi+SQn+WknyfpCKZJSkP8l+P/u5d+n4uxq5hE5pLUpA8l7lLcnCX0p87SY63J8ThdjmHQyeFEARwlf6vGuxtAEgAswB8TgvMHRPaMKhGIYQkORnA+frXwYGm/jegW+p/K/RE4JjQCn+5GhEXAKjUSBEhCBqomnN123NJztATQowJrbCX0ohYEYIK839f/yzzoCuoIssBXBSB3LFrCKrR0Z9VJLu1IWGMCV9/3kSyhmQdyesDf4u6zN8b7eeMXYURmqc/v6EHORuwAttJTrHa3x1oZ9p2W9+xf58hWTsmuAIbIfrfugBCsnrQ7zcDTnImyY6AUI1r8E2SmwN/M4L9lj1Bxq4h+mb688SAWrSFd7HleH8p8DfTvlM75T+3nHJboH/UQhdjo1441fj9CNW4yzjIut2aEDSS5IP67+eHOObm59PGGJLCqcViklsCA2yEcZfVPqHXJxUikPP1vcpIbo24148PNxU5GhdgR5v5ZwGo0QyIE+jvLy2VdjmAohCTnwDSmlHZC+D+AKNi7nUpyVIhhD+mJoe+nv0iYh1621KfDsk3QlSf+c6d1n0X5FkfL9CIHFORg1GN+rOC5O4Ii+8Wq/2pgTZB9WiTww7JDRGW6KrDaV0bberR0FafBzA1QFu5+v+/stov158y5L0MOXyO5jAVgPsCLImhtT5PcurhSmuNFtX4aEA1GmS8aLUtJflBnu0ag6IG6zs11r5akF25WqvIMZ9tELTVbJJdEarxry0L84J+aCubOZluPeeZwPfM528PF3bEGYV9uRRAGXJEsNCqzAOwF8BD2hq0SeSo7RVDDk8AsMT6/b0hzyWAs0nWCCHUaBYcB76jMaJIeynCWHjEajuVZGuIERJlHa7p57sGyTeNVp9tVAkrILBj9UCHDfzllnV59QAZfWWRwwnreasi1szXjHM/WgWWrqqrdkaZarzSshKNlecC2A3gKa0WAWBl4PtKt/09gB3Wd42KLAKw1HIr7kXfzVNXtz8RwMlCCI4W85+AIwBy/vzillhiaZaiZzRRVx7JdyOopp8P0AL8NMm7IpzyP1iILiO5I+JZt44WFUm9QZuuqprUUhW/raWm5rjRZuafnYfUXWSpxpsiSOTd2oE+L899TrWee3vgPqbN+yTH2c7+oRFYvQcALfH4nHQs0Ziuii+yBXmohWYoqX+PGMRNJItDWI0gQv5DtykP8d9Mmx9Zzz09j3A/dygZEtbnBJauSp6Yrkl+lI4lLrMFOVpoq3K93RJm0X23H/7QDHS91e7HERNgO8kSIxCSb4YIV5H870MlNCOYj6vi53XFa2VrLPFVAFiHBUWjRTV6ej27NMQaVHowP2G1/2GEMN4jWWStWaflQdHnrcnyrQhV20KyciRVJAHRu4bF4l9ich7bYslbAIALRonAAuvZAxEm+CvWAJfk2RP7Z92mSE8CR+9Kh6HoHuv5tdbvgy7GipGitWiRHOnqxD8zeQxbY4mf2apytDH6YfEdRhg3WNZlWFi4QeNx1ppn1sh/iEBRWge/mn40htBaiuQTI0FrGXS9X1VVkq6O38PkMUxXJx6DRl6YU30o/TSzXlyEXBxikLbqBvCARVut1H9TFrMvALwqhHiTpKOZfPP3BgBZ6znGZ5sEYLGl9n4ZYP4d3XYRybnDSWsR9Z4A5Na5c2dViqLfTCoqXtHuZ9f749xlujMUoykK2lp/nouY6U9abSeTTEWg8bqgX2Xd+3cR5PDTVtsZeZD+9eHy2Yza+7gqeWJHLLmZiTq21yS2tcTjc4Iqc7QJrG4Aa4ogeVVEtNVekkcH1ZilIr8coiJNLGS11T5yTR0GY0QYCzFVk7ygqybZ6tfUsiOWbGueGz951PhieXyzb/dnvel2T+m/BdH4WJhpbq2X0/XWTBiKbrTWy3zW6wmFWtsIiAYtkFQs8dWumlrVEUuqrlhSNsdii0eNL5aHtsrnJ91ttTVh4QyhrZZFWXiWZXp/BIrWR/iJMsxPHKqKtNVdayzx/2S8lq3V8Uw2XsuW6uS1o8oXyzOYZ+RjJKwBjQoL30NyUpT6snzAiyNQpEieZLXPx8gUDYX9N+ru90cfXdoeS/yK8TqmYol9jNdxd3XiXwBg3YJRKrCAarwjYpC22txfnrDw/8rHWlhCH68DfMJU5L9azzkrzyRaOFiGxKi7bXPqjuqMJX/PeB13x+J7Ga9jSyxxtzFKRt1+WchA9suy63Yn5aGtzulvIK0JclfEBGmyeE2P5Dv98JruYATWXJU4qaMm8b4fr+XuWHyfjNeyNZZ4dg3gRfliA9nHGjHfTAvu0wBmo28gqtnTWmW1v9Lyr2C1fx/A8/peKr9m6uOL2UGvCkAMuTADIYTwtW8H655GSEtIThxotBa1hSjQ6O+OJS4sd8TvPIjqtJSZCsct6VJ8K62yX1iYe6+D9sVGWmjsx1H+A4DX9Kwuht64tPpp2t8vhOgB4Fobowfa1vsH+WUA71jCMveyU3mBXHieDHHIpyOXPdpvnrZGjRBo9FOx5NfKhfNrCUzsUio73nGKM+THXVJcGNu2rXV1boPzoHPInZFUjXoQpwI4Tw+IG0DEKr1rLACcrZGQD40DmaFuHhQJABcag0YI8RaAVwLoHnCettllFoBqram9tdJx7uimUhkqWSyEC2Bfh8JFs7a/9z4B9/Lo7NVRgzSjGpcAqLDQZUIKMtDx9oFoqyAaNwB43UyCATxX9YOiKQDOy0NrGeGeQ/IoPfGcMAtRAGrdrFllHTXJ1RWO+EZaSV8BKIJgsXCcdilXzNq6aa2hrwY7kCMpNJM/vTxkUAWA54UQTRqVkwBcmA+NGCBrYLhDIcRGAGsDKDL3XWmpyIeQC9fzsD/OxAcwHsDFYeNmhLDn6OTs2pLx/1sunMtSvu8T8ASpJriu16nkdTO3NT20DguKBBp9jPbLoq3s+I4gbXWNZX4vi/Cteumng2EoLCvyryN8vl46TLeLinB+4QDKTFuIO+PxBV2x5NacsxzPpmIJtlTHM0zUcU8s8YNR7Tz3M2h/FzForXqtM+0fD6GtSPI3g6GUrElzVJ7o5eusSXNFxKTxSR4DAGxocI3AdlfHl+yNJdr31STZHIv76ViCe6rjWcbrmI4lVhnhFsoXGyn1aNaBKwPPlVoFPSmE2KMHbQ6ARVotBft3r157DqrfWkW6QoidAJ5B3xIWZiCXWevpUwD2WIYPrPXwMgDA/PmuQKOfqo7/TbnjPOJDTOgklQfhStKvdF2vTcrnPywt/mKOwmqUhdpmKYjQbgacqFmkB4sATgbwCT0IjjVgwhIGkAsLLw3sr7kAWgE8pu81mEVc9BPzeKrZTBVCtAF41BKWGSsJKb9Cskgcd1ymtSb5b5Wue3s3qTIgPcBRoJzguF4X1Xudmb2XHLdxY0a/aMH2xYYstA9nzSq7BVACYMQekD2Tg46yC+BDAP9rGQLLA9+z0dhiTYKDRruFot0hKHIAXGExN/eGaAUXrvveW83N4zpiiYYK170hLaWvAOEAQgKqRDiuT7WnG2rx0Tt3thirspBqa9BC04EogmVlU1KxxErTOXtrwTANOgLqCxGO8oM6vRYkjwdwioUu014AuGco+1omalgI0Q7g1yEoAoDLSRZp4T4PoEk/O4tclPLPmpYsWV596qeeKve8pS3ZrA8hPAEIBbAoB+dMp1CXTG9q2rRmiKZ9wYWm4S6O3rLlAwqRSsXiz6Zqao4XaPS1QB0AptxDPYAqRORPB2grJ4S2+gDAGj2YhZi190bQWrUAPqUnW0Y75EIL7KYdQtw6649vv1RaXHRmSzbrO0J4egLTA2Sp4zhd5FUzmppeYH29t2iYTPshqcccsuBOadr0BCH+Zxzc1zpqEtcbVgDXXedFOMpGGG8BeFWjrAi5pPcoNO4j6Q1SNfb6hHoSvQBgSwitBQDLe5+RyayGkhkAl+0U4qWp8doNjufFWi2B5TpLv8JxvTYlb5y+dUsDUe+JxlHui5m9oNaaxDdZeww7Y4lnmhPHnKSFUaF3ovNVzBE69DtvOHchgkct9+O7Ee7HRyTLCQi+/vp4ksk9zrirM4l5bK9Jck8sLlOxBM2/PdXxLBN1bInFf3z4+WI6UCVVHb+LyWPYmZy3b0/phOv95uYvRfg8WVObSg/izyO2T94oZFUdy2c7QT/jAEff7+lZqhdCpCZP/R5r5zMVS8g9sYQ6QGC5Dc0HAGANRmZfrGAP4H6/iulY4vHKoqLzOtOtcBtWydKF9Q6lFMJ1bX/nOSFEvR7AiQA2A5gWoI48AP8ghPiuVo1+gQRn3nstgE9affL15wM7hfhiWe0x900CLkxlsz5yhpaw9Kw/yfW8Tsq1aZldWL1tWwYjFPLmFFD67GXEJ1dc2d7R8VZRXRJFp5/mABDCcfrucUm5yhrA87TAgtUMsgBWB9acgpDXet1aFUYOK+AzJWec+dLE7p4L01nfF4AXEJic4HrePuk3dTZtviC2bVt3oX2xEWNEBKBWA87kDRva2nbtuJiXXvKBV1JCSJ8QwgyOB6AFrruaS5eawVuGvvnTRngvCiE2WYGoBSOv9ecDyAXF7ieHpYQDVJYvWnhCuqNDCtfxAl9UJa7r+n62pWP6lAtnkxn9HgqH80XSgRCQ5At91jPflySV/+abb27MbYmAu3fXUUoG4h/N+vdl23gYJhL7yUAfSVLtfXmt3DPraKbjtfvXsFhCtdUk/X2JOn8TcCzJO9mz71+Hq48jJ7CGhly0VWfniZSB+I6ccFTXtX/FjmkzP2hP1F1Gcg7J13PmibINkd5SEsORuWJFa608wFBSijKbYeozn2XqqDlMxWvZEkuodCyRZd187sztv/1FzlqSfzjsyw9aJvX3+liDOYEw+9HH/HjesaqzuoZ+ch7bj5pz98dLVyZIfpNSSmsAVxfKzM83uVJkhezp6RsWnkMbO269lS2VU5lKzmOqOp7hvGP5EZxrrWQQs810+JYfDJT929zHdM/m1E7H3ffIlsqpbEnUyZbquM/kPHZVx5tb4VzRkyvK+az+3kXDlWrUm6nyxS+WpMoqGjIPPuSTlMxmbY3AnnfeYcvcGFti8QzrjmVz6YTv51wFmdUTK6P7etthqyKtQNQDHWWZm8RdS6/g7umzVDrn33BPddzvqEmSdfPZOXXGg7uO+cSxJJeTnNywP0ShYNeanNGBnYnEtLbkvOc5YzZTK6/K6QFp+fVaM7RcstTn7LlsPbqqQUc777Hezaj+D0iWDpcqH1br0bpW9LHSlCIcAeX7m7pfeXXjxPIJQiqpAMARws2QbOnpkWWTJ19S3t31StotmrRaiLbLAQkhnEI5rES9twjwd8RitRMUnitX/LNUyThfvfyK8LdvBxwHUNoIlBIA1ISLFrttXV2vbdixbTmAo5EzoJRFZCv9+0WD2es7pEKzoq0mYn9Zo5yOdxwJAI7n/YLbms71Pbd9XM7WV9pVEI4QbiqTkVKIsspk7Z3n1yRfaq+pPVMAUm/7uEMVmECj/1FV/IxK4T1XLJx5rdL3RXGxJ/a0IPPkU8YJy316HgEIZ/EFPW0Lz7l8EYQPIV5GLswvSGoH40wOG6S5/TjKEun0E1OBHa1FzuWu48DLCaP3RR0hXEky1dPjF0F8UoDPt9fE/21LTU2FACQB9+ZB9Jn1OnC0Kn7xBNd5BsCMDiWlEMKDUhAl45B5/HGQBNze2wulFMWUKcVzH7qvllRC9C1dGFZ+cMpIlB8spNAYUWjMCO8VMXnyG/zKV4pmv/XW0+27dq2c6HmuIGXAQRdCCK+DSmVITBDeDTPorkvXJC8RgLwFUGsGmA6U2yJa6orGRr+lJvGX413nIR8o20cqB8LVqhsoK4Pa8AdkNm4EhOhVkY7jKCfncF9joWg19meY2lWBJiEXQznsBw85BVKNjp5hRwE4B+Ghb7l9s5/+1Pc7WpfN6Gzr3L1l8w2VJSUeyWxIxxwBiJT0fSGQGC/Eg+2xxKqd1dVVi/ru2eXjQoXAapmOxW+ZJJyf9JAqm/Oq+37PdSG6upB55Ne5r7IX/K7++7kkp+U4ZLFZb+2EhaSvKOCe37AjzdznEgyg7J9bXrEcwMPTZPaRlo92/WxySUmRIkPJYCGE102qdqVUueNcWSG89a2x+NfMnl0Ys25F+rItlvjZJMf7xzalpAkLOJDUUhClpcg8/TRVJiPgumaqGRRNRC4O01xR5QfPIhkb7vKDhbpx7wZiYPfAxHc8K4T4UKNyJoCzkFMjvxuXbvl2e2vbs1OKiz0ZUJWWynQcwElJKX0hpkx0vDs6axKNrbHkqYvQ6NuGigl72DVjxviO6uQjEx33yynl+wyw9EGhsbRUjXt3s8iuW587cEjJKIsYyAX9tIeoyHEw0VrDaEU6BVKNiuR8AKehb3xHWLTVxQAqkEugmFsOPNK55PwrO/fufXeC57kKjFQtjhCuTzIlfVksnLM98MWOWOIHvYZK/f5qAeVlE54td8WFLdL3BYSXzzJQpD/Z85yejvZfjfvECV8FIOC6MoCiPyMZ1+hvBvA0IkLx9LvK0Yw0c48rEF3278mQaCtPq9HTj/rpT2/rOGrGpUqp1mLh9LoCEagTjhBuh1IyA3jljnvjTLrrOmKJpaKx0d8zNzF/mjvuuVLHOa1Fyj5hAeEqgnKi63rtfvaPe/d2XIPx41cBaAtBUbFBkRbKPYguP3jScJYfLITQpKZvgvEdRniP6ggokKwDcAb2xz4awa2YtXbtxN3bP1gyznWFG3AFIoSX4w61oVLiOA2p6vhjwuUzxY5ItEop+xcYVJlw3B6ldnZSLp5DdgshOgE8mQdF5rCHZwHs1MIKJoksC3xv9AjNikE8HcA8hEdb2arxCi0oGTBUPsKGDVuq/O7nW5X62kTHdQcaeiaE8PZpQ2WC615QJJxZbUopR4j+8shUcW7XsqtbcsnspqbtuP32Yr3vFxXQegJyQbcQQnQBeDiwppt3vkynIA+LzzZUpIlQ2mq/8LYAeMFSFVdEoPFxcfLJu3nzzSUztrx35x4pf1jpul6URRlmqAjAaVdKhpr0BwqMLsAiCKedXDZt+5b1RL2H667LaHP/WeQqtEaiKCKg1UR3VQM4W09oZ9QIzQpEDUsBMi+6WgiR0S94GoD5UWgEACxc6BP13vStm29M+/LBKa43YMEZlSkGoJIEICc6jtsp5V/Nen/zoyb9SE8uTwixz7goISj6AskSLZC1AN5GRIbpaDREDG31WQAzsT+0GtYMvc/4ZpYBEkTjZgAvkBRi0SIfaJQEnE6/e2U75asVruspsHCWGJmtdFyvRarvTdvW9O+sr/dOwfpsXyBaZMCBKKoCUG8lNf4q8F6G1lpMctJw0FpDEVowfzpIW70G4A2NyvHa8Y5CY2/hMRMcM3fHjn2ZrHvxPsUdZcJxVQFYBkX6la5XlJL+PdO2bv57ot5DY6MMrJFmkF9FLpg2DEU2OdyA6AzT84eD1nKGoBolyRkAPof+86c/A+CofGi0qR8TuTztg3d3dkt1EYG9xbkGHJrAXK9NyTWTq7Zckyt3FJl+5GoU3ReBogtMGSghxNvIJeLny9MeFQkYdtm/CRh42b8gGtcLId6wDvDps+4Q9d707Vte61BqebFwHC/H9B+04BQocym06u02mfkCGiGX5o9RtMsP+iEoqtSCEwFaK8j8LyI5p9C01mBvFEVbGfXxOyHENo3K6QDOzUsiR6gPgUZ/HRYUzdy65ZF2qW6ocFxPgP5BdlSVCsfNKjbvk2Jx1fbt6f5KQVh52u8hV+s/LE/bLmXxMIAuHJinXYpcvl1BaS1nEKrR0FZJAJ8K0FYIoa2WaMI1DI0PBlVj8DoF67NEvTd92+bbUkreUel6RWG7AhECy5XjAbo7FS+auX1T00GUgghm9QSLeC4kWaWFvAvAbyMc8qjjwkYUab25XOh7rKMRXhrA49YsXBGBxjVCiO0DC0RtlGtQ7015f/N1Kd9/bLLrFfXnCpj0o3HCcTqlXDFz2+aXD7IUhLRQ1BmCot6cu34yTE8hebxe351DJbT+8qcfE0Kk9cskAZyJvim7QTT22wcBcKF2BfyyccvalXx9Yj+ugABkLv1IfX3GtqYHD7YUhJWE+DGA3+Qjh/UEfRpAM8IzTK8spIo82CoBhrb6JIDjQhzloGpcaqHRzp9OAXjiYPKnjdEwY+PGzk6hLuqm+qg0whUgma10Xa9Fyh9O37r5x+uwoCjgiw1Ydv2QwydragtCiA4AjwRQasZmqc6/k4cCaUE9HXSUtwFotFRBFBofF0KkDzZ/2rgCs5uatncpcTGA7qKAK6BIf7LrFaWlbJi6dfONBLwFWD/YbBuTp/1bAB+F0FpOgNaKcsiTAM4sFPN/MAVUDG0VZhH1JjQIIbr1z58EcHwEGgedP21cgVlbN63tIv+81HEcV8/gXl+M8oXK0uKrdDjCoEtBWLRWp4WiMHK4WAv3RQCbEJ5hOmy0Vl7VmOdYRxOzf4rV/rbhPIDHZFy2VNfcxHgdW6rj3d01SbbHEu99WFs7VRsjhdjkNQG4+Yp4ftqqZfkvERmmH5OcMNT3Hmznfxlx1NXrViZKqT7XZViPuuqtuh2L/4TJeeyIJZubY7FaLTC3gO8u+ini+Z9W2+OjMkxJXjZih8Vas2hKnqMZ/856wXxoXGBPgqEJbX9EVkcs8cDH1fFzgf2h3wV8f5NU8p0IFDXrIF3Tfm1ECd+HC/XuA+q0FsbVIfXxlc4eiVvto9C4odDHN4ZFYg3D+x/s8ZfXRwi3k+Qs+57DKTTT6d9GnCjxrNX2UByUKkwxmhEYg5ciTuz4tdV2jq5sFzYGXx2mMQjtbCzPsY7/x2qfD401IzLLhknb6M+/iUBRF8nZVvuoSnovDPsYWJ39vxGdbTPRt7pd1OHfzxyuAgtM3nyHp3/NWteX5yk/OG9Yx8JKFHwtYnG9z2qb75DVa4ZdLYycBf1oBIp+b7WdZOWyBYV787CNhdXJk/Mc/H3hAA5Z7S3CeTjnJlsGWb4insda7f8rwiDbaO41nKrxBxGO8g6SZQbq1iGrQTSuGjFTd2Rcnwp92m8Yir5jaajP5nHIzyz4mFgdHKdZjDCn8g6r/YI8aFw8Yk7lyKnIX0RM5HesyZ5v7O4suIq0aKvP5ZktZ1jC/eEA0CiOIKF9Js+4nGWNS5SW+lAHPBVuXKzO/XeEXn7TapPvkNXbD3cDJMIwKya5JeKdf2K1z2cPFK6CgzVLJuUp+/dt6yXOzTPrTj8S1rOItf77ESjaRbLcGp8oy7uhYGNjWUnL85T9q7OE+z8RnX/TqFkcQZelYU7Mg6JLLN8un49bmKpE1sOeiPBHnrMEVprnkNVvHUmqMUJNhp3vRpL3W7ZBTB8EEea/fmnIY2QJbA7Jff14/g7JH0WcP91bhPNwZUEGqCK/EXFYbKe1NAjNFIUBYM2Qx8jqzNfzHPFRptXj8yECM51pPFIFFpjcYWeWms99JK/X7VZELDUZkokhjZXVmbURHOI9ehuiPaQTtpCvPVJVY8hYPR1AEQPr3D0k51qnMwY1198Peqz6qe9rft4WgiqGCHjGkeKbDUArXWVVpWPIMkGSb+l6zLZmMp+DLz+Yp5I2QzqjIhCWtQ7bcXAEX5YxUmn5bNmQsfEZfakhVUTXHSjSxwQH16ow2Js29ow5/U9BYCE+7UySDXkEJUPGzp7sPzpoFWn5HgvzCIx5ZtDtFl3l4U/ostUayWut9T47gDE047xdH+8y8CXFEtp/DOCBtq7eQfKi4H3+1C6tpcwYzte+rBkrOUDBfd6+z0AhPkHH5zECykF03W8FqnhHstExCLvAJXlLiKXICBWpSN474Ilv0VZL+1k4zYM7SP7lnzq68lnhFhDqrXhJPwIM5ndpkpMHpCItWD8cUH1hi+gLZnf2SOQVhwl1FdaxLFGgML/7836Z/8B5mZ0hqtFmRP7JErA3JpaB2wr65yt1YGuYa2CE9nS/1nc/J9OaG71LcpG14Dpj4hi0kVKlw+uCRoga8AnDFtJetCRuw/c/SVaMGRuFU5f657+1OMts4PPGfrUZyXlaUNL6YjPJZWPGxrAYKQYoC6ytHXvs1/cbPk/yHwNfesKC6JixMbxGyjgrvoaa9ZckT4oEi1Z5JvStm+Q3wuA8dg27kXJ+IM4m+uAhkmdbkDwlCOGxa0SMFIO6aSRXWcmXbqiaJHkfyZ+QuZIbY+gaFai7WrtfF4eqSJJnjRkbo9I1OFXnuR2g9f4/OVSMzV+96TQAAAAASUVORK5CYII="), "b": base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAN4AAABaCAYAAADXRuh5AAAkFUlEQVR42u19eZicVZX371R1Zw9JCAkBAmEXAsi+LwMoYAYlg34qOIAPzjA6jmwygMgOggRGVlEHUFmGYWT7EAURh8UH8rGJIJAggYQlZAUMkJClu6t+3x/3nPTpm/d9662tU+n0fZ56qrr6rbue7Z5VSHaheY0ACgAWA5goInNJFkSkvMqDpIgISa4DYHsAZf1toYYxrUn0PwGwDMAiAAtFZHE0hyKAsogQTW4kCwBge0FyOIC9AOwHYDcA4wGsC2CozrsE4GMA7wOYCeApAE+IyAv1zt/t/UYANgPQ1YC9t/0WAJ0APtS5/y0+f5JtOu9ylfNdF8C2Ds7qbSUARQBvici7Nk7CuKN1XHs+794IgCUAXm7TCUsTEU9ybkpBF7IngIfd9x3uAPM0icYvRfMoAVgO4H2SbwJ4EcCjCsSfGACLSKmJSLeyf5I7AzgOwCQAW1dY52hFjN0BHAVgGckXANwN4A4RmV/j/IuKbP8K4Gz3/Yoq4CPpmbK+Sop8SwC8S/J1AE8DmArgLyLSVSXhsHOcBOC/9Hw76oBl6jzbtY9rAZzsxon36fMAbq5ijwwGBwGYD2BbIVnuBcRbDGCbnBxvHIDJALYAcCiAHascr8shejFjTnGbCeB2AD8RkQVp86wT4US5HEluC+AcAF8B0BY9+hGAJwA8C2CuAtVQAFsC2FeJU7yGBQCuB3CViCypBvlsrSR3BHCAUvNJADatYnmGYAWHiIUcv/kzgDsB/LeIzMlDONx8bQ8PArBBA47oLQC/BXCziDyfwfEmAPiWwum2OXEAAP4E4BcAbgTJMpvXrO+PSW7oRaycANFG8uskP9S+Sq5f/+rS97tJbktyJ5I7kNyb5JEkv0/yQZJL3O+79NXp+iXJ2SSPtY02ZGmUaKmfTyX5kRuzQ987SV5JcvMs5CW5P8nfu990ur5eInmg4yC1zncdkhe6fS+nvGzvzyW5Dcmdde93IvkZkv9M8gaSr0Xr7YhgZSHJS0mOsf3Ku/ck1yN5pjvfNDhJepX0dQHJEVXu0SA9yxXRHnkYI8k/kTyiB+y3GuIpYLWRLNrGk5ykB5U2V1vgDRX63prkNdHB0B2WB+ArqwWASkhHchjJO6J5lxzgHeJ/o3tQdPtRjPo9L+qn0yHjyYZ8VQBwwcZy350dAXPa3h9Xoe8hJD9P8qGE9Xe5794i+bUkgpUCK36u+5F8L+Fs05qt6fyI2OeFUYPP0xP2yNZ0HcnB0T0crc7xhGS7fr4rWlDS4f9Sgac9AtxixHEmkVyUckAegKc0gHPY4Ywg+ZhDDE9xF5Pcu4qDLzpkPifagy63pgtrnb/bvwGOW5Uy9v4E/c0Afe+x/1HfR5N8N/p9OSJ8N5AcVAWxNjg5gOQyx4EqwebfSI6x+da4R6NILnB71OX3P/EMWhnxPMXWg62EeL/IAjRDSod8nSkH5IHg8DqAV3TMgST/EImVPYBWn2+vsu82/fzrBCC2z2fWMf+ivt/gCEba3v+z01CmcidHMCaQfDLhTD3gPuZFzxzztbP9TgasxNxuFskhnkjWQFTbVcT35/uzTIljTUA8fZ+cg+r+Ig+QuQO6MQcyv6BUXGo4GJv7f2Yg3WPVioQRIRGSm+udMe2OcUQtyGfiFMkf1ot4CeYDkBzupID4DDrd/WjdKpDP+v5dBeSzffqE5JY1wqa4O+YC1/czSmxTrymNsH2sia2sG/Jjp4Zmgoq9DGAnAPurdqsaMbkoIiWSXwXwL6ptbY/U7wTwg5VfVGl/U61rQURmAbhF+yxHNjQCuJHkeF13ocoxGm7TFJEu3Z/FAL4MYIbbb2ttaoLYFcCvlFjmIX7Wx0mqTZeEs4WzjQ4BcHCGSSTTrKHz2R3AWB17KYATRGRFWGry/q2ViCciJd2QVwC8mnE4Zf3+kGoORoG7rGLSlQkGXnMOeB7A43p4tZouqL+/MYGIFLTfsQAu7w3HgGrOQJHvfbVjrkgwwrcrwfqszr9UCWbVzFAUkdcB/NDtQVb7pnJKVil1UPf0n9x+XyEiL1UyiaytHM9sQSUA0xIOHBHX2DqipjngSgjg3wFs6BAN0Vh3GjDVihTO1vgKgGcirueNzUeRPMAAs4WQr01EngFwTQqSmNH6FJKTDGFzSDQF7XN6AjeNpZpdABxrEkQVEk2Z5B5qzyOA1wBcYYQ3k1Vi7W229jkZiGdtRI5neiC0inb/ksDtqAfeiW4PnXo5kSHugynEw97PatB4jWwlBdTL9CxioPWG+KtJDq3EmXQvRESWAjg9x5oJ4HySI6vgetbfD1QsFgDnqPeTVCKkazPiWVuW45lq/Fnt0I4HMFKBSBIObKaKuY1EhCeQ7Ldo4uchJHdVSt0SZ++QZBGA/0wR+41rbw3gm3k4kxNlH0RwqSuip/sXInF8AoDT8vTtuN0R7hryCIB7nCSFfsTLbkNzPPNenv1Sl6KS2p++hmT3NAOqF0SkQw+qXsQzDjEdwRcwBl4TP4sAjqlRkdBkyZ8C4DYEf84ikh2uCeAkksN1nyVnv99DcDBPu8sb8p2sGs5UwqT9keRAABeh20/0rGrOcW1GPAPW8TkA8cWcwGr7uQeAbSrs8YuNQgD1HzSu8UbKfdTGOZzkkJyA21tcr6xc7y0Aj6XM37j2BABHujtapX4LIjJTRdm0u5ch5HAAFxkXzhDryyrR7KjP3SIizxkn7Ee8bM5UJjkA3U7YSZypoGLm73MqV6yPSXZ/yXhmRpPurK+niK8GuFsimEha7fxNNf/bCvcqAji6CmWXca+rcypavqpKqFKCx41of+siRHGUEUKdLjRO2M/x8h3yHnpvSLoXlRRJHhCRl3NGKxgX2bcCMpcAzG6SouP1DE5q9839W1DctHCgJxFMC8UU4iEA9iW5cZ67qrtDLnOKlizELgC4TJEuVrTYteA0lZQKCGaOOY4T9iNeCrcrOK3TGSnih92HliKEneTlokSwmU2ssL+LEcJ4mtHm5nhmjyo4Rq8djb7PAvB2ClGyu+pwAPvkJR4JipZCijRiCpi9ARzjFS3ONrsZgH/Tub0K4Po85oO1EvGcz2SbiJTVc+JMAF9wSOY5nVHWb4nIKzm5ne3l5ggBq8hQrCxGiLlrBsdbmAGQ9t1WJNuVY7TKPY+6z8tVJEzbGzuHvWpU4JyFyh4tBGBhQvY7I6znIZiXBMD31WQh1SrI+hriWZhI0XnIi4jQIdxIkv+hl+2y22iLEi7qwRwnIrdVc2HWtlkObvIR8pkxamkfR3e6JMQbD2C9VrwCVLin+jVsXw3XdoqWNyooWuz7TdFtXmhTrrkHgH/U5x4Wkfvymg/6OuItV3ewDkW0suVxIbk7yfMQ3LROc/K8eacU9f0hBN/M26pMoWAAsUkGQBggLbF0B00Q1ZYgGOez2ggA41rwnmdtdg6uvSnJQabRzXuPdIqWVzMULYZ8p5DcAt1G/osQ3Ng60O2MUFNr6yMIZwTkQJLX6+YMAjAYwBil8Js6kbLkLu+i96I7AdxqyYNIDhCRjhoAf/0cAL00uhc2si1TxGvPEKMKCImUWg3xbC/ez4F466pIPydDbEwTZ5eSPB1Bg5rGVe0ueYmIHEXyHwAcpv//uYj8uZ7cPH0F8ewwtkG3/SzvIRvibgzgMDV+P2PGbXQ7wuZtI3M8U2riXnRU6N+IzfAWPs+PcxCFIQCG1XCXNEXLAyTvAfAlJGcLM1H9KyRvBPBd/fs9ABdXaz7oq4hn2qenELJODYjubyMQkiftCGA7xw1sw8fpAXxJv59K8koRudc0WlXc8/IEsy6POFAziFAj5rm62ic55z+odn3bSo+WQxG8l2Kjuf98jyK6AJgiIvPqzUTXVxDPgPcvIvKTjN0u6KX8aAAnqKhScjI9dU/2RbAV/QrASSKyMAfyVYNAXS20Z63Y8gB0sVbi4UKH3iB5GYBLkJ0j05zkpwH4aS3mg76uXBmsUdMD9d2/iqpseUlEzkKwZT3gLthFR4hK+voqgEdIbprDWFvNXWlgCwC/tPA55kGoLnTH8NUkJeX0aPHw8H01xNd9N+9riFdWbWGXiMSvUmTPm4Vgx7sJq3quF9EdurM9gN9o9uA8GrTlq1nMywsQy1v4HPPc3TrqWUMUOnRWBaQrAnheRO43R/h6F7hWea5E9jwTK76J4JibRPHaHfJdV8F51trfciDAAHf4jW6DKlwh7MwXt/BRjXB396x74Md1Sg5GSF9CegoQa0saucC1NjrBRX6XAXxbDzJJ2dGuVO9okgdnRHAbQs7LId4NcYlyGi3yDTHEzuCGnegOdWqlu57txfoZc7Pv3gPwQSOIsRKrSudQ7Ee8BiKf3v3+CuB/0J0AJ62dmgMg3s6xt8OaIG6K69vbKNO48vwWRDybyyY5npmhkksj4xl7TRnVHwjbzXVuz6Bs3kA/IUXR4p18y8h22RqhnKkZbWQGoBiAvYNQManl7uj6vm2GAsjW9fKaDMP9iNcdjvKiiolJGb/su2Ho9uxPc4B+E5XzuKyD5nmOjMtBoV91KvWW4HiuIMgIdDtBSIbIN7UFOXY/4lUj4ydEb2dxi+2SgMK5I30E4C8p/dgdchi6q9s0GvE2zvHMU00auxGwOBEhM1tSjKQ5RcxFyKiWV0zsR7xWxT99X5jj2Q0zkNP28/Ec4tSEJt2Pts74v5lInmhBoLUzOBjp8XI234dE5MNW4tj9iFffoS/J8czQHFzxIQXwpChqOsreSDHN7p1bpXAz88x5AcBf7TctdAaWamFyBje2e/OtazrA9SNeT2QYkuOZZWmA4QJLX1VxjkhPOrRrA7mO9bkRQk6VNMAVAHflycjcy/c7u7ftqftCrKrksjlPBfBEowzZ/YjXGog3Lsezcyvcj8w2eBOSS0jbnu9IcmSV8WSVEG9nVdwk5fIsIATg3tGKSgkVGU9C5ZTrF1eT8bkf8VoV43oWlP9Uxr7Yd69UAFzjevciRFLHWlL7ewMAe+qzjTqHz6RwUUPEm0VkTg1R9U3ldmpP3R3AF5GeeKoI4G4RebjWqO9+xGuxPVDg3xvdFV/SsoMtQbc2jRmUu6CpvC9GsjeMAf0/1KsccCnnhgA4POFczZC+CMCP6o0jazTR0/cCgB8hOBXE+28O7PMRIsKlTwBdP96tRJZv5VCaPCYib1cKEVIKXkAwyj+CZCdsADhSOW09SYfMc+MwhJjDWAyzu9GFIjIbVaaha3JrU851DkLKwTg0h+6e/E+1pNHrR7zG3cMaSXHbFUkmAfh7rJpxLG5XV7jfraJsAfCvym383cVc09ZHqKXGWs4i8vdMcmfrQnCY/j2AH6sSo7y6996q2YpIJ8ljAVzoCIQndqZkOUlEHqw3+DTHHbkf8RIAbGiDD75dD35zhGIZaa1LD/92EXk07/0iqtP2DXe4Zbf3BHCaJmct1VA+ywDxWOUYnnCUFOlmAThen6s2jYXt/eAG7ntRo0S6SH4dwC/RM/GU59IFhExf1zcL6RxstVdAyiFOHyBrA+KZKLVLiuIg74FbDe42PfhOktshlLbaOOVSbxzjNYSCFlVtuHPCvk9F2YLjfHbXWg+hYmuhGuRTwtFF8lMIxS/93ciIxVwAkzVVQS0imiHqztVyBrfntu+GcCWSgzTy+2a357YfNvfFCDXrrmwipzMD/KfQbbSXFMTbCsBYS4jcCIxvuRroLiem1SofR3JOVOM7qQ73L+2gK/Q/hOSJJD9IqatedjW4Z5HcJu/c0zR3+n48yRWuxrevU36LW29RXwVX59w+tzkRcyuSr7o1+P5eI7m9Hz8nshR0bJvLQW6uzKqBbiJkRv8DSX5F68rTnWcpqlX+/0juUs3cqyQIBVcrXUhOTYGDeI3n62/a3PlIn0G8aH4bkfxjzo25wd99POBrgfh9SZ5Pcrr7bcnNtSsCgMdIbloP0iUg30EkZ7gxOh2SP0Zy15zAcxTJua4PP+/7SI6rF3BJ7kny7Zx7f3y8TyQHkNxA13wxyZfc7zoSEPpdkqd6AtTkK8wQkj+vsD464rDC1ln3pZJkuYmXS1NlLwawjYjMrSTyKCWyXJiHISQl2gTBBatQYZx5CF4j7dH/hiF4dYxzfXSgu9pobEubp+LbNSqSNkTUcTar0QDORagYa/cny4W5HMD9AH6NEPrygYpAA3UNeyNkQ9vH3YcMQN8EcKmI3GRIUI14qakNx6h29IsIpaiG5dz71xCiMgY4zeooBHvlmEhpEvf1CkJtvFtEZEEtc6/iPjda57S/wtZOFdYXr7OoZ3MrQuT6QhH5eI3leI4jHMzV0xaTfILkySTX92Jvgw+/6D5vR/J6x7nitkLF4fdILsqY+19Inqblo1ber6qYk4ld5/bSXi8i+RTJK5QbDoikE2nwnlvhkY1JftTgtTzr+pc1juPZ91qR89sIGaS8divveOUM7VTswtWFkELgDQDTtDCiR5ByM7zfzVvFuCjJMQheJ59VJdImCK5fSZq2ZQgR5K8DeFpNBVNFpNNz1WoBU/f+EAQj/CdOlZ87PTqSw6DMdPKRKnveBjBLROYmEKSm7bdqI0cilOqyNOzVwFYMY0Tw7X0bwLXVzrvlRM0WMF00DQBSKHEPZ18df30Vi0fr4ZohebESivki8sHqmncTuH95TQ3vqbW1teBhNNJ3MRfxMcDWDGS95gNoBMivWcefiwp17qJ9KjfoDtoYVXn2PsPtd2k1Insjz7G0xiOeUr412gG2njVbLbbIhhRLEWzGPvUFV6xmIMlawfHW9qYIxf6d6Nut30m6v/W3fo7XK3cN2H2ujnHo7mfFXtiSsvMRLPTieM1eW7P2ka5/NgAevb2X7mXXAYn/V2nclkS83rprNGKcXlbG9Or9t5fX1pSxHHEt1+AgXnTEoVQByZN+i7RxW5Xj7Yzg0dGBbFtSFiAKVrXd+Rrn0xFsZm3oTkwkFaioefwLgD/pb3ZGcqCrDwEqZChKvP2rkDDfdgBviMh8kkMB7JDEoRo83msIdrc9dG8qhUv58eMzQMo5tAGYJyKzSA5DqF3YieR8K3nHsv6XI9RVmCciKyJkqIiAZvJydtYNdX7bA9gcwdxjdtaPEcK+3gUwQ+HqVRFZnDluC3mumONvG8mnmzQPa39VH9DpNfa3kORY9WVs5r7Z+3m6Nwc1eTzzV/wuyfFN8PKI13W3EVqSSxo4zjKSH+o536HO6aNjWKvAqUByMsl79LzjtjxlziV1rL+N5JEkB0fct7U4nlEDDXX5PIDNAHwZwClKWRhRTEEwJt+GnmE2g5QabaDUaQv0jIUrIBimAeAAhHCPf0TwlokpNSJZ/mZ9zdRilT7HiafUMxBSRLwBYDaAYxDyRdr49v4OQnqI9RCyg+2mlFUcZS+iO5enubKV0NO/NG28uNb7HB1vpBtvp2g8QcjNOUcp/JYIPqVHYdV6DLaOVxBqDfpzGoyQqn689rGhG8dCsCy51CsI6Q73AHAGgN2xqk+n/f0nAP+L4LtqnG+w7s3W2s8gfY1ACPk5CsBsktcCuNKywcWcz/nS7oSQiuLgCEzvRagO+wJCnfaSwto2CJVlv4TgdbSZvo4BMI3kjxHqpneuHHcNiE44Wj3gS64/84h/scJvh5Hcm+SNjppbCMoO0bMnRuE1sff9eQn9n+aoH9X7/kjNf+Kfu85FEaTOX7n935H8vfPVJMnf6P9Pjb6fRvL/ZIzXEY03PWG8/Un+Lur3twlrvTrqy6/nhgrnMJrk4SQfivqY4TmCPjvErb8rimagxvGljdNOckeda0dK5MYD6jrWg/M5f8sjHKe3376nzKASrI4heQHJpdGcqb6p21tYUqvG460MWtW/43ipkgO8wfasi2MrJoQGHa7ih81pko4z0J53omc8zjsKEBYPZ2ErV7m1PkJylKeerjLtjREgWb+vaFBoWxzH5gC9TPJZ/e4/opCldaPxBkTjdUbjzSC5Tsp4V7rxnvMxkbo3Y1z8YkwAb3PjF6N4wvgcTnd9LCS5gZ+/ft5eiZmPv7S1XJc2VjTOoYpAPubPEOF/9WxWxjrqb/ZUMdXGMyL9OUeoiu53Pn7Ri6j7kXzT9WME7Wpba0va8SxSGd2FA2elaJAIoKTPlkTEv+gQuF1EHgBwnBNPxjktoal/Z8bqaH2fLSJLnXbLFDrj9JnXAHxZRBZZoKrOqUsr1DJDYWPPlNx8CyJyioo2AmCMAoeN94aO9zdDDLcHtY73XQB363ijAAyz9ereLEJ37T8mmB+SzqAcnUNRRK4A8AMdZ4SK2dZnp573bAALkJKhLW0shwjtIvIwgBNdHwV0Fxr9DIBTdX1FhbOBAH6qIqoX5R8RkYeUUPVYlyt0WnIVh9tE5EkVU19317kyXAqNljeg66EX60Fgla3bReR+hESzQL7ktdZiDmrAMF4P6HRFgjYt+8w6CY6F9ZysGtixeocZo+P9u4i8bwmDGjjeKYpgG6K7KmuS1rDWcayo50UI1ZkGuHurROctNY5TBmAVf/8LoYCMTzRlCZ9O1MpERlg+pxrqONPZVEuJWGmfXS6ZNhF5E8CRCJEkRZ3DyDUG8bJsJVW2km7gZQjhQJtU07ffdKXiAwFsCuDPAH6rXKOrgTatgoi8q4RisCqLNkYIvrxfOWCpwePNQUj8NMhxImnUGTgFWieASzIIIBs0ThmhjgXQM8kUlLjs48710JRxSzWMb8g3DSHXjim47CrCtcllzBLNvoVQzWffGg/ZANGi5H9Sa3q+ylddCoCf6Zi7KqL/zGqxNziUxsaz1PMTEzS89Q4gjrv+Rm1guzV6nKi/1zM01du477ZOmcOna9lnh3x3Keddma1NRMprDeKZ87Ee+gUIpZfr8V7pAnA5gP9bK2WsMF+7R8wA8D1V71+q6mygwSW23HgzAXxfDcKNkjZ6iGJKNFaoKP1cM9bj2oqM/w1xBGFQ9D9Lv3gYyc31Dlet+c0SFZ8NYCmAUdbHWhOdoGrr5YqAU9FdUbRqANX3+QDOTBJFm0A0pujHB5s9nipqfhjZzxrRJ0luAWAX5QIQkZv9vjYh5YMguwLUUnd1+CCBY5b1rnsTycNFZJlTsjAPrKhC6R2StyAkNx4MYHGhjyObqYm3RkiTMMGprov1Hmoz8oOkjFV0auumjmfO0Q0ex/o7BsBP4nU1a9sUOSYmcG5b2wz33YsJfZhS5iAAD5DcwpRnVczdpKzLVWIpr0nKlVqb2ZF2B/Bpryio1ynXaUybHjvnVNi9OV4jxzE/xb0QfED99+UmECpRgB+EkJrfw7p55yxUYmztHr0+xEhqSqyDADxD8kyS60UmjFQkdCL8WyJythaz6fOIZ1TvQGQnQupvzZU6RL1F9qlw52oU0rW7uhVboaf7mdnobhGRD5wd82UANyJoIGPttBWdGY2gFX+B5BSSOzoCXM6SSuJEy30B8cRry7wRVS+36wD4gm5mf+Bv8ySLNgBt5hVjXh7GQRH8PdcBsLxOMdYAuBh5rrS5NPEdJI9wop2NZ7lLpwO41EqcKXcs6J39SX0mdkQwk0AJQZt9BoDnSD5K8jsqhpad80YPLuhsmegrypXOhHQJK7kbySkIRtrFWA1VYdaStky1lV1JHI/kMQDOQ3J9imrbCvNcSeF4oxFsZ+crfHunbAt5miwiH/qMdyQhIosVYW8GcIR2abUczIGg6OCrXUXQgwAsIflHBO+fB0VkodczxCJ1X0C8USRNfLA4rxEAtkOo0jO5QQfe3xK4j75/muS/obuwJAAMV4K3J0LUQfybWsfahORe6Kn+H4TgELGb3uk2Qs/Mz3a3ux3AGUlFXJRLFURkEYDJJE9BMOP4iBCTEmMEBELG7cP1tYDkvQgRCc+bIqkvcDxDom0BvBVxuvboUGIXoP7W+HPYU19ZratOeLMz/Iq+Mq96ihxLFT4eAfDfIvK0caEkpY43aYjI1STvQggX+7oiM6J7okdsj4Tr6/3yBJJ3ArhERKYr96OIcE1FPJ8o9/FIjh8EYF0EL4+x/UjX1GZKiwcQ/C8HOOAbosC6D4CvKUeopxkBvR/ALQhZt4/X8/ZR6TMAXAvgeYWPN0Wkw4l9rFDNl45DzQFwNsmrEPwuj0KoudCegISFBCRs07V/geT5InLVyvT6rRyP5+op3J4SVvNixm/HaR2Epa42wuZJc3D2vt9E49j70zVq2Gz+N6TM/2VfLqoBGr14vDgs6DWL3cs7XpQZYFrUn63npgp97KRVh8ok/xxVsk0aa5SGYvmxbC2Xu+cP1DoMJRf+Uyb5llaaXVl9thZ7YVLJN5KfJnm21kwoR5HnSeXefCzgT+2c1vR7T7vGsxV8PTmlXPNF5BoAJ/Uzpqa3ASkxckWSA0TkRdUCNkK5NUj7HSwijyNEVBQiqWcCgFtJXmCKt0r2wiRC5ApprnSWEJGXROQSFa33AjAFIUzLFx31WQvsLtgJ4FskzxWR0pqOeES3Ebas6tyy06YVAfwcwDQVffrteE06h5QYuRJCjF0BwH0A5gMYXKdxnq7fNhG5BcAf0G1rMxNBF4DzSR6Xp9Ju1py8s4TBlX73rIh8DyGc6GgEN8SCm4NXDJmG9RySu/VZTZ8ioHnw/9rdSfpb756DRVKsAPAEGlRPHT1zW56GkFnMB70a97lai4uWsxJtaaazXHDluKAh4RIR+R+9/30NIZC3kIB80HvwKX3ec0Xfn9KNGNCPCqsH//T9SUS2vnr8Ql2N+ZcBXBMBu0UXjAJwnRGAFPFyFII72EF5Snk7LuiR0JwF7gCwH0LURRLyEcABfR3xbNHPAbgeIf7LI2R/6x2uZ8h2K7oN0yuRp07R0zjZDxFShMTR5iUAnyf5jQyRcwWCHfCsWmpXRKJou4i8o1rQeZHYaSaIsYU+fuAWibxARL5j3gRrWy22FjqPD0Vkuj8fkl8muW3EGWsRZT9CMHjHeVoM8C8nuUmKyLlCFSSHkNzHwnlqXKOlGZmDEIkhCVectrXCmyOv+NDfeuUcCi5L13gAd6I7O3ahRmA3kfMuBJuicTovco5Gd+VWcYgvypE/1OcvraW0cgIXFhWtjfP6tnStQLzYQbW/rdZz8Eqvr+u/pjXgCmC/PR2hlLTnfIaIkxO0nIZchnh/B+BkhZdiPWt191mfUJkA3l5TEG91OzdLbwS8tvheNmT95j2iOUFPBLDEAX3NY7lo71cRskAnKTbKAH5EchNFPl9j4n1971Kut5/lTamVuyNkkwZ6OnQLgHtbHvF0AV0VqFyjxknrr+aUfRUQVhpNVCqsI63gR16OUq5nXsplRLneZQg+jfPQHRxr0dqljLGYsad2f7sCwXUsScu5HkLcne1/IUI8KwPwK5ITDfmq8PQpJHB0Q74Cgqnh+kKrIlu0gI1TKOI4AMP12arV0hEwEJoaAquWQh4XeWVIlfPfqhnzTxivoONtlrKOsQCG5h1P/29xZRZtkMSRLB6vmOC5YoGmdIlfLwJwggL5ByKyRMdq0/mPRcjixoQ17GD3tDhTtVO0LFGRMyZCJnIeSnJKlOtzsUO8EkL6vz+QPDBK9xBnkvbZpNvUxNBF8gyExLnl6J73bRFZuCbUTpikqbfTahpc2CBkP9al+o598EjypBr7rTT/ixpMtA7VlOFp411aY78XRPvh+7whx++HkfwsyQejugJ3xRyD5N0JtRPKzh/yGxXGMp/Ve1P6sb8vt9oNWjOBCbUalpO8mOR6OfdpPZcOv8ONtdz5jxZbtXbCQJKba1GQj6JiIuUoHz5J/lQdcYdXwzW07sJWJM/VnPnljHE6SV5GciLJoRVKPQ0iuVkV8/9xLfOP9mtTkqeo03Cl8X5WaTyl5ENJ7kDympQ+7e9ntaDKmVob4XT9fCHJa0neT/KNCGE79LdX6Xij1On5oYz5lx3iTCG5reZVQQLyCskt1Tk+6VxXFo4h+W2SW5D8HMmHE5ydqU7elyvx2EhraQxQgrKxErwp6qAdt6kk9/VEoVXr4z3qvNGrafO1qGWecdZVT/muGsaZqW5IiFXPehiP1Dj/eZXmn7COgSQfjyrTVDPejtE67H0iydlsfjtJa/G9V+Pv/+iKyEgC1zszZz8rrIIUyV2UGD9KckECIq7QvZup7ytSaijeR/KLbk+LXlvXheY185tbDGCiiMxNC0KMAGoSQvzWUuTLo2+VWgcCeFjrGGRmWtbD+nsEN7JlVYwzQJ/7nYh8Eo+j8z8MIb9INfNv19cf8sw/uswfpnewuvfLvY/QfrvQXTG3HsVLXL+dupdPI9Rr+IKeQzmnPa8Lwe/zEz0Lpii1Big8pSmdyk6h8gSABdF5jkfIZrANQrbpCXr/GwZgqPZrmtm5AP6KkNb/GRF52xMCb9L6/zh25SbQaQcbAAAAAElFTkSuQmCC")}

@app.get("/logo/<k>.png")
def logo(k): return Response(LOGO.get(k, b""), content_type="image/png", headers={"Cache-Control": "public, max-age=86400"})

HTML = r"""<!doctype html><html lang="bn"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Alien Tees × Bodyfitters</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600&family=Hind+Siliguri:wght@400;500;600&family=Noto+Sans+Bengali:wght@400;600&family=Tiro+Bangla&family=Baloo+Da+2:wght@500&family=Anek+Bangla:wght@500;700&display=swap" rel="stylesheet">
<style>
:root{--acc:#c9a45c;--bg:#f4f2ee;--ink:#1a1512;--head:#1b1512;--card:#fff;--pk:#ff4fa3;--f:'Hind Siliguri','Noto Sans Bengali',sans-serif}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--f)}
header{background:linear-gradient(#0f0b09,var(--head));text-align:center;padding:22px 12px 14px;border-bottom:2px solid var(--acc);position:relative}
header h1{margin:0;font:600 34px 'Playfair Display',serif;color:var(--acc)}header img{max-height:48px}
#tag{color:#bdb3a6;letter-spacing:.3em;font-size:12px;margin-top:4px}#clk{color:var(--acc);margin-top:8px;font-size:15px}
#menu{position:absolute;left:14px;top:14px;background:none;border:1px solid var(--acc);color:var(--acc);border-radius:10px;font-size:20px;padding:4px 12px;cursor:pointer}
#prog{max-width:760px;margin:14px auto 0;padding:0 14px}#prog .b{height:10px;background:#ddd;border-radius:9px;overflow:hidden}#prog i{display:block;height:100%;width:0;background:var(--acc);transition:width .4s}#prog p{margin:6px 0 0;font-size:13px;text-align:center;color:#777}
main{max-width:1180px;margin:0 auto;padding:18px 14px 80px}
#dz{border:2px dashed var(--acc);border-radius:16px;padding:18px;text-align:center;background:var(--card);cursor:pointer;max-width:760px;margin:0 auto 12px}#dz.on{background:#fff6e0;border-style:solid}#dz small{display:block;color:#888}
#q{display:block;width:100%;max-width:760px;margin:0 auto 12px;padding:13px 16px;border-radius:14px;border:1px solid #d8d2c8;font:inherit;background:var(--card);color:var(--ink)}
.tabs{display:flex;gap:8px;justify-content:center;flex-wrap:wrap;margin-bottom:12px}.tabs button,.pill{border:1px solid #d8d2c8;background:var(--card);color:var(--ink);border-radius:30px;padding:9px 20px;font:inherit;cursor:pointer}.tabs .on{background:var(--ink);color:var(--acc);border-color:var(--ink)}
#info{text-align:center;color:#777;margin-bottom:12px;font-size:14px}
#grid,.sg{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px}
.c{background:var(--card);border:2px solid transparent;border-radius:16px;overflow:hidden;cursor:pointer;position:relative;box-shadow:0 2px 10px #0001;transition:box-shadow .2s}
.c img{width:100%;aspect-ratio:3/4;object-fit:cover;display:block;background:#eee}.c .m{padding:10px 12px}.c b{display:block;font-weight:500;font-size:14px;line-height:1.3;height:2.6em;overflow:hidden}
.c span{font-size:12px;color:#888}.pr{font-weight:600;margin-top:4px}.pr s{color:#aaa;font-weight:400;font-size:12px;margin-left:6px}.pr em{font-size:11px;color:var(--pk);font-style:normal;margin-left:6px}.wm{display:none}.stockstate{display:inline-block;margin-top:6px;font-size:11px;font-weight:700;border-radius:8px;padding:3px 7px;background:#eef6ea;color:#2f6b25}.stockstate.low{background:#fff1dc;color:#8a4b00}.stockstate.out{background:#fdecec;color:#a11}.quick .c .m b{font-size:15px}.quick .c{border-radius:12px}.quick .c img{aspect-ratio:1/1}.quick .st{display:none}
.c.pk,.c.sel{border-color:var(--pk)}.c.pk{animation:bn .45s}@keyframes bn{30%{transform:translateY(-7px) scale(1.02)}60%{transform:translateY(2px)}}
.ck{position:absolute;top:8px;left:8px;background:#fffd;border-radius:8px;padding:4px 6px}.sc{position:absolute;top:8px;right:8px;background:var(--ink);color:var(--acc);border-radius:20px;font-size:12px;padding:3px 10px}
.st{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}.st i{font-style:normal;font-size:11px;background:#eef6ea;color:#2f6b25;border-radius:6px;padding:1px 6px}.st .no{background:#fdecec;color:#b33}
footer{text-align:right;color:#777;font-size:13px;padding:30px 18px 90px}
#dr{position:fixed;top:0;left:0;bottom:0;width:310px;max-width:88vw;background:var(--card);box-shadow:4px 0 30px #0005;transform:translateX(-105%);transition:transform .3s;z-index:20;padding:18px;overflow:auto}#dr.on{transform:none}
#dr h3{margin:0 0 12px}#dr label{display:block;margin:12px 0 4px;font-size:13px;color:#777}#dr select,#dr button{width:100%;padding:10px;border-radius:10px;border:1px solid #d8d2c8;font:inherit;background:var(--card);color:var(--ink);margin-bottom:6px;cursor:pointer}
#bd{position:fixed;inset:0;background:#0007;display:none;z-index:15}#bd.on{display:block}
#bar{position:fixed;left:50%;bottom:16px;transform:translateX(-50%);background:var(--ink);color:var(--acc);border-radius:40px;padding:10px 16px;display:none;gap:8px;align-items:center;z-index:10;flex-wrap:wrap;justify-content:center;max-width:96vw}#bar.on{display:flex}
#bar button{background:var(--acc);color:#1a1512;border:0;border-radius:20px;padding:7px 14px;font:inherit;cursor:pointer}
.fav{position:absolute;right:8px;bottom:8px;border:0;border-radius:50%;width:34px;height:34px;background:#fff;color:#a8843d;box-shadow:0 2px 10px #0003;cursor:pointer;font-size:18px;z-index:5}.fav.on{background:#c9a45c;color:#1a1512}.copycode{border:1px solid #d8d2c8;background:var(--card);border-radius:8px;padding:4px 8px;cursor:pointer;font:inherit}.dashgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:10px 0}.stat{border:1px solid #e4dfd7;border-radius:12px;padding:12px;background:color-mix(in srgb,var(--card) 94%,var(--acc))}.stat b{font-size:24px}.report{background:#fff1f1!important;color:#a11!important;border-color:#e3b4b4!important}
.ov{position:fixed;inset:0;background:#000a;display:none;align-items:center;justify-content:center;z-index:30;padding:12px}.ov.on{display:flex}
.box{background:var(--card);color:var(--ink);border-radius:20px;max-width:900px;width:100%;max-height:94vh;overflow:auto;position:relative;padding:22px;border:1px solid var(--acc);box-shadow:0 0 60px #c9a45c55;animation:pop .25s}@keyframes pop{from{transform:scale(.94);opacity:0}}.admin-full{padding:0!important;align-items:stretch!important;background:#0b0807ee!important}.admin-full .box.ad{position:fixed!important;inset:0!important;width:100vw!important;max-width:none!important;height:100dvh!important;max-height:none!important;min-height:100vh!important;border:0!important;border-radius:0!important;padding:0!important;box-shadow:none!important;animation:none!important;overflow:hidden!important}.admin-head{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:18px 24px;background:linear-gradient(135deg,#0f0b09,#1b1512);color:#fff;border-bottom:1px solid var(--acc);flex:0 0 auto}.admin-body.admin-body-no-tabs{display:block!important;min-height:0;overflow:hidden}.admin-body-no-tabs .admin-content{height:100%;max-height:100%;box-sizing:border-box;overflow:auto}.admin-head-actions{display:flex;align-items:center;gap:8px}.admin-menu-return{border-color:#ffffff35!important;background:#ffffff10!important;color:#fff!important}.admin-menu-return:hover{background:#ffffff20!important;border-color:var(--acc)!important;color:var(--acc)!important}.admin-head .x{flex:0 0 auto}@media(max-width:800px){.admin-body-no-tabs .admin-content{padding:15px}.admin-head-actions{gap:5px}.admin-menu-return{padding:7px 10px!important;font-size:12px!important}}
.admin-head h2{margin:0;color:var(--acc);font-size:24px}.admin-head p{margin:3px 0 0;color:#c8beb2;font-size:12px}.box.ad{display:flex;flex-direction:column}.admin-body{display:grid;grid-template-columns:230px minmax(0,1fr);min-height:0;flex:1;overflow:hidden}.admin-nav{padding:14px;background:#f7f4ef;border-right:1px solid #ddd6cb;overflow:auto;min-height:0}.admin-nav button{display:block;width:100%;text-align:left;border:1px solid transparent;background:transparent;color:var(--ink);border-radius:12px;padding:11px 12px;margin-bottom:6px;font:inherit;cursor:pointer}.admin-nav button:hover{background:#eee7dc}.admin-nav button.on{background:var(--ink);color:var(--acc);border-color:var(--ink)}.admin-content{min-width:0;min-height:0;overflow:auto;padding:24px;background:var(--card);-webkit-overflow-scrolling:touch}.admin-content h3:first-child{margin-top:0}.admin-section{max-width:1350px;margin:0 auto}.dash-hero{display:flex;justify-content:space-between;align-items:flex-end;gap:18px;flex-wrap:wrap;margin-bottom:16px}.dash-hero h3{margin:0;font-size:28px}.dash-hero small{color:#777}.dash-actions{display:flex;gap:8px;flex-wrap:wrap}.dash-actions button{border:1px solid #ddd2c3;background:#fffaf2;border-radius:10px;padding:8px 12px;font:inherit;cursor:pointer}.dash-actions button:hover{background:#f1e5d4}.dashgrid{grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}.stat{min-height:90px;display:flex;flex-direction:column;justify-content:center}.stat small{color:#777}.stat b{font-size:28px;line-height:1.1}.chart-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px;margin:14px 0}.chart-card{border:1px solid #e4dfd7;border-radius:16px;padding:16px;background:#fff;box-shadow:0 3px 15px #00000008}.chart-card h4{margin:0 0 12px}.barrow{display:grid;grid-template-columns:minmax(90px,150px) 1fr 44px;align-items:center;gap:8px;margin:9px 0;font-size:13px}.bartrack{height:12px;background:#eee7dd;border-radius:10px;overflow:hidden}.barfill{height:100%;background:linear-gradient(90deg,var(--acc),#e8c77f);border-radius:10px;min-width:2px}.barrow b{text-align:right}.mini-line{display:flex;align-items:flex-end;gap:10px;height:150px;padding-top:6px}.daycol{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;gap:5px;height:100%}.daybar{width:100%;max-width:42px;background:linear-gradient(#d8bd7b,var(--acc));border-radius:8px 8px 2px 2px;min-height:3px}.daycol small{color:#777;font-size:11px}.stock-columns{display:grid;grid-template-columns:1fr 1fr;gap:14px}.price-history{border:1px solid #eee5da;border-radius:12px;overflow:auto;max-height:430px;background:#fff}.ph-row{display:flex;gap:12px;padding:10px;border-bottom:1px solid #eee5da;align-items:center}.ph-row:last-child{border-bottom:0}.ph-row img{width:68px;height:82px;object-fit:cover;border-radius:8px;background:#eee}.ph-info{min-width:0;flex:1}.ph-info>b{display:block}.ph-info small{display:block;color:#777;margin:2px 0 5px}.ph-info strong{color:#1f5a14}.ph-info em{font-size:11px;color:#8a4b00;margin-left:6px}.toplist{max-height:330px;overflow:auto}.toplist .item{display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid #eee}.toplist .item:last-child{border-bottom:0}.admin-table-wrap{overflow:auto;border:1px solid #eee5da;border-radius:12px}.admin-table-wrap table{min-width:700px}.admin-note{background:#fbf7ef;border:1px dashed #d8c19b;border-radius:12px;padding:10px 12px;color:#6d6257;font-size:13px;margin:10px 0}@media(max-width:800px){.admin-body{grid-template-columns:1fr}.admin-nav{display:flex;gap:6px;overflow:auto;border-right:0;border-bottom:1px solid #ddd6cb;padding:8px}.admin-nav button{width:auto;min-width:max-content;margin:0}.stock-columns{grid-template-columns:1fr}.admin-content{padding:16px}.admin-head{padding:14px 16px}.dash-hero h3{font-size:22px}}
.x{position:absolute;top:10px;right:12px;width:36px;height:36px;border-radius:50%;border:0;background:var(--ink);color:var(--acc);font-size:18px;cursor:pointer;z-index:2}
.mg{display:grid;grid-template-columns:1fr 1fr;gap:20px}@media(max-width:700px){.mg{grid-template-columns:1fr}#brand-watermark{right:8px;font-size:10px;padding:6px 8px;max-width:42vw;overflow:hidden;text-overflow:ellipsis}}
#mi{width:100%;max-height:70vh;object-fit:contain;border-radius:14px;background:#eee}.th{display:flex;gap:6px;margin-top:8px;overflow:auto;padding-bottom:4px}.th img{width:62px;height:78px;object-fit:cover;border-radius:8px;cursor:pointer;border:2px solid transparent}.thumb-wrap{position:relative;flex:0 0 auto}.thumb-copy{position:absolute!important;right:3px;bottom:3px;width:auto!important;margin:0!important;padding:2px 5px!important;border:0!important;border-radius:6px!important;background:#111c!important;color:#fff!important;font-size:11px!important;cursor:pointer}.gallery-label{font-size:12px;color:#777;margin-top:7px}.sizechart{margin-top:14px;border:1px solid #e4dfd7;border-radius:12px;padding:10px;background:#fffaf4}.sizechart h4{margin:0 0 8px}.sizechart img{width:100%;max-height:420px;object-fit:contain;border-radius:8px;background:#eee}.sizechart p{white-space:pre-line;margin:0;color:#555;font-size:13px}
.acts{display:flex;flex-wrap:wrap;gap:8px;margin:14px 0}.acts button,.acts a{background:var(--ink);color:var(--acc);border:0;border-radius:20px;padding:8px 14px;font:inherit;cursor:pointer;text-decoration:none;font-size:14px}
.ad table{width:100%;border-collapse:collapse;font-size:14px}.ad td{padding:6px;border-bottom:1px solid #eee}.ad input,.ad select{padding:8px;border-radius:8px;border:1px solid #ccc;font:inherit;background:var(--card);color:var(--ink)}.ad .row{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0}
.userbox{border:1px solid #e4dfd7;border-radius:12px;padding:12px;margin:10px 0;background:color-mix(in srgb,var(--card) 92%,var(--acc))}.perms{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:7px;margin:8px 0}.perms label{margin:0!important;color:var(--ink)!important}.ad h4{margin:18px 0 6px}.ad .tb{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}.ad .tb button{border:1px solid #ccc;background:var(--card);color:var(--ink);border-radius:20px;padding:6px 14px;font:inherit;cursor:pointer}.ad .tb .on{background:var(--ink);color:var(--acc)}
#toast{position:fixed;top:92px;right:18px;left:auto;transform:translateY(-16px);background:var(--ink);color:var(--acc);padding:10px 18px;border-radius:30px;z-index:60;display:none;box-shadow:0 5px 20px #0003;font-weight:600;max-width:min(520px,calc(100vw - 36px));text-align:center;opacity:0;transition:transform .22s ease,opacity .22s ease}#toast.show{display:block;transform:translateY(0);opacity:1}#toast.danger{background:#c62828;color:#fff}
#brand-watermark{position:fixed;right:14px;bottom:16px;top:auto;transform:none;background:#fff;color:#555;border:1px solid #d8d2c8;border-radius:8px;padding:7px 10px;font-size:11px;font-weight:700;line-height:1.25;letter-spacing:.15px;box-shadow:0 3px 12px #0002;z-index:12;white-space:nowrap;pointer-events:none}
.n{font-family:'Anek Bangla','Noto Sans Bengali',sans-serif;font-weight:600}
.hl{display:flex;align-items:center;justify-content:center;gap:18px}.hg{height:50px;width:auto}.hg.b{height:38px}
#sync{font-size:13px;margin:0 0 10px;color:#2f6b25;line-height:1.5;background:#eef6ea;border-radius:10px;padding:8px 10px}
.tabs.sz button{padding:6px 15px}.tabs .t2{border-style:dashed}.tabs .t2.on{background:var(--acc);color:#1a1512;border-color:var(--acc)}
.iw{position:relative}.c .iw .lg{position:absolute;left:8px;bottom:8px;width:auto;height:26px;aspect-ratio:auto;background:#000b;padding:4px 7px;border-radius:8px;object-fit:contain}
.tg{position:absolute;right:8px;bottom:8px;background:#145da0!important;color:#fff!important;border:1px solid #fff!important;font-size:12px;font-weight:700;border-radius:8px;padding:4px 10px;max-width:60%;box-shadow:0 2px 8px #0005}
.st{gap:6px!important;margin-top:8px!important}.st i{display:inline-flex!important;align-items:center;gap:7px;background:#e9f6e4!important;border:1px solid #bfe3b3;color:#1f5a14!important;border-radius:10px!important;padding:3px 4px 3px 10px!important;font-size:15px!important;font-weight:700}
.st i .q{text-decoration:none;background:#2f8a1e;color:#fff;border-radius:7px;padding:1px 8px;font-size:14px}.st i.lo{background:#fff1dc!important;border-color:#f3c98a;color:#8a4b00!important}.st i.lo .q{background:#e08a00}
.st i.no{background:#fdecec!important;border-color:#f3b9b9;color:#b33!important;padding:4px 12px!important}
.st i{font-size:18px!important;padding:4px 5px 4px 12px!important;gap:9px!important;-webkit-text-stroke:.4px currentColor;text-shadow:0 0 1px #fff}
.st i .q{font-size:18px!important;min-width:34px;text-align:center;padding:2px 9px!important;font-family:system-ui,Arial,sans-serif;font-weight:800;letter-spacing:.5px;-webkit-text-stroke:0;text-shadow:0 1px 2px #000c}
#dsz{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}#dr #dsz button{width:auto;margin:0;padding:7px 12px;font-weight:600}#dsz small{margin-left:6px;color:#999;font-weight:400}#dr #dsz .on{background:var(--ink);color:var(--acc)}#dsz .on small{color:var(--acc)}
.ck{z-index:4;padding:6px 8px!important;border-radius:10px!important;box-shadow:0 2px 8px #0003}.ck input{width:20px;height:20px;accent-color:var(--pk);cursor:pointer;display:block;margin:0}
.c .iw .lg{left:auto!important;right:8px;top:8px;bottom:auto!important;height:24px;z-index:3}.c .iw .tg{z-index:3}
.c .iw .sc{top:auto;right:auto;left:8px;bottom:8px;z-index:3}
.personal-box{margin-top:10px;border:1px solid #e4dfd7;border-radius:14px;background:color-mix(in srgb,var(--card) 96%,var(--acc));overflow:hidden}.personal-toggle{width:100%!important;border:0!important;border-radius:0!important;background:transparent!important;color:var(--ink)!important;text-align:left;padding:11px 12px!important;margin:0!important;font-weight:700;display:flex;align-items:center;justify-content:space-between}.personal-toggle:hover{background:#f4eee5!important}.personal-toggle span{font-size:18px;color:#9b7840}.personal-box #personal{padding:0 10px 10px}.personal-box #personal label{margin:10px 0 5px!important;font-weight:700;color:#665b50}.personal-box #personal button{font-size:12px;padding:8px!important;margin-bottom:5px!important;text-align:left}.photo-count{position:absolute;left:8px;top:8px;z-index:3;background:#fffdf8dd;color:#6f5528;border:1px solid #dec99c;border-radius:20px;padding:3px 8px;font-size:11px;font-weight:700;backdrop-filter:blur(2px)}
.c{transition:transform .28s cubic-bezier(.2,.9,.3,1.3),box-shadow .28s,border-color .2s}.c:hover{transform:translateY(-4px);box-shadow:0 14px 30px #0002;border-color:var(--acc)}
/* Product cards: clean image frame and readable match badge */
#grid .c{border:1px solid #e6e9e7!important;border-radius:22px!important;box-shadow:0 8px 22px #193c2b12!important}
#grid .c:hover{box-shadow:0 15px 32px #193c2b20!important;border-color:#b8d8c4!important}
#grid .c .m{padding:12px 15px 15px!important}
#grid .c .m>b{height:auto!important;min-height:1.3em;max-height:2.6em;font-weight:700!important;color:#24332b!important}
#grid .c .m>span{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-top:7px}
#grid .c .copycode{background:#f4f8f5!important;border-color:#dce9df!important;border-radius:999px!important;padding:5px 10px!important;font-size:12px!important;font-weight:700!important}
#grid .c .sc{top:auto!important;left:9px!important;right:auto!important;bottom:9px!important;max-width:calc(100% - 18px);padding:6px 10px!important;border:1px solid #ffffffb8!important;border-radius:999px!important;background:#103e30eF!important;color:#fff!important;font-size:12px!important;font-weight:750!important;box-shadow:0 4px 12px #0003!important;white-space:nowrap}
.c.hit,.c.sel{border-color:var(--pk)!important;box-shadow:0 0 0 3px #ff4fa338,0 16px 34px #ff4fa347}
.c.pk{animation:bn2 .55s cubic-bezier(.3,1.6,.5,1)}@keyframes bn2{0%{transform:none}35%{transform:translateY(-11px) scale(1.04)}68%{transform:translateY(1px) scale(.995)}100%{transform:none}}
.excel-hero{max-width:1500px;margin:0 auto;color:#17251f}.stepper{display:flex;align-items:center;gap:10px;color:#9aa49d;font:600 11px system-ui;letter-spacing:.08em;flex-wrap:wrap}.stepper span{padding:8px 12px;border:1px solid #e1e7e2;border-radius:30px;background:#fff}.stepper span.active{background:#163c31;color:#f5d68d;border-color:#163c31}.stepper span.done{color:#49715d;background:#edf5ef}.stepper i{height:1px;width:26px;background:#dfe5df}.welcome{padding:30px 4px 20px}.welcome .eyebrow{font:700 10px system-ui;letter-spacing:.18em;color:#9e8043}.welcome h2{font:600 32px system-ui;margin:8px 0;color:#18382e}.welcome p{color:#758179;max-width:720px;line-height:1.7;margin:0}.upload-card{position:relative;display:flex;min-height:190px;align-items:center;justify-content:center;flex-direction:column;gap:9px;border:1px dashed #b7c8bc;background:linear-gradient(145deg,#f6faf6,#fff);border-radius:20px;padding:28px;text-align:center;cursor:pointer}.upload-card:hover{border-color:#708c70;box-shadow:0 12px 30px #193b2c12}.upload-card span{font-size:12px;color:#7d8980}.upload-card input{position:absolute;inset:0;width:100%;height:100%;opacity:0;cursor:pointer}.upload-icon{width:44px;height:44px;border-radius:14px;background:#183e32;color:#f3d38b;display:grid;place-items:center;font:30px system-ui}.excel-tip{display:grid;gap:5px;margin-top:15px;padding:16px 18px;border:1px solid #e6e9e2;border-radius:14px;background:#fff;color:#718077;font-size:13px}.excel-tip b{color:#254536}.excel-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:12px;margin:4px 0 22px}.excel-stats>div{padding:16px 18px;border:1px solid #e4e9e4;border-radius:16px;background:linear-gradient(145deg,#fff,#f7faf6);box-shadow:0 5px 18px #183b2a08}.excel-stats small,.excel-stats em{display:block;color:#7d8980;font-style:normal;font-size:11px}.excel-stats b{display:inline-block;margin:7px 7px 2px 0;font:650 27px system-ui;color:#173c30}.excel-tools{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin:16px 0 10px}.sheet-tabs,.excel-actions{display:flex;align-items:center;gap:7px;flex-wrap:wrap}.sheet-tabs button{border:1px solid #e2e7e1;background:#fff;border-radius:10px;padding:9px 13px;cursor:pointer;color:#46564a;font:inherit}.sheet-tabs button.on{background:#193e32;color:#f4d58e;border-color:#193e32}.sheet-tabs small{opacity:.7;margin-left:5px}.excel-actions input{border:1px solid #dce3dc;border-radius:10px;padding:9px 12px;font:inherit;width:180px}.excel-actions .pill,.excel-hero>.pill{display:inline-block;text-decoration:none;border:1px solid #dce3dc;background:#fff;color:#31473a;border-radius:10px;padding:8px 12px;cursor:pointer}.upload-replace{position:relative;overflow:hidden}.upload-replace input{position:absolute;inset:0;opacity:0;cursor:pointer}.excel-table-wrap{width:100%;overflow:auto;border:1px solid #dfe6df;border-radius:14px;background:white;max-height:60vh}.excel-grid{border-collapse:separate;border-spacing:0;width:max-content;min-width:100%;font:13px system-ui,'Hind Siliguri',sans-serif}.excel-grid th,.excel-grid td{border-right:1px solid #edf0ec;border-bottom:1px solid #edf0ec;padding:8px 10px;white-space:nowrap;text-align:left}.excel-grid thead th{position:sticky;top:0;background:#edf3ed;color:#51645a;font-size:10px;letter-spacing:.06em;z-index:2}.excel-grid .rowhead{position:sticky;left:0;background:#f6f8f5;color:#89958c;z-index:1}.excel-grid thead .rowhead{z-index:3}.excel-grid tbody tr:hover td{background:#f8fbf7}.code-cell{color:#173d31}.excel-cell{width:78px;border:1px solid transparent;background:transparent;border-radius:6px;padding:5px 6px;text-align:right;font:inherit;color:#223d30}.excel-cell:hover,.excel-cell:focus{border-color:#c6d6c8;background:#f5faf5;outline:none}.excel-cell:disabled{opacity:.55}.empty-sheet{text-align:center!important;color:#88938c;padding:30px!important}.table-foot{padding:8px 4px;color:#859088;font-size:11px}@media(max-width:760px){.excel-stats{grid-template-columns:repeat(2,minmax(0,1fr))}.welcome h2{font-size:25px}.excel-tools{align-items:stretch}.excel-actions{width:100%}.excel-actions input{flex:1}.stepper{gap:5px}.stepper i{width:10px}}
.stock-visuals{display:grid;grid-template-columns:minmax(0,1.5fr) minmax(260px,1fr);gap:14px;margin:0 0 20px}.visual-card{padding:18px;border:1px solid #e4e9e4;border-radius:16px;background:linear-gradient(145deg,#fff,#f8faf7);box-shadow:0 7px 22px #183b2a08}.visual-title{font:700 13px system-ui;color:#31483a;margin-bottom:16px}.visual-row{display:grid;grid-template-columns:minmax(70px,130px) minmax(40px,1fr) 58px;align-items:center;gap:10px;margin:12px 0;font:12px system-ui;color:#66766a}.visual-row span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.visual-row b{text-align:right;color:#274c38}.visual-track{height:10px;border-radius:10px;background:#eaf0ea;overflow:hidden}.visual-track i{display:block;height:100%;border-radius:inherit;background:linear-gradient(90deg,#225541,#9fc18c);min-width:4px}.donut-card{display:grid;grid-template-columns:150px 1fr;align-items:center;column-gap:15px}.donut-card .visual-title{grid-column:1/-1;margin-bottom:8px}.donut{width:142px;height:142px;border-radius:50%;display:grid;place-items:center;background:conic-gradient(var(--donut,#dce8d9 0 100%));position:relative}.donut:before{content:"";position:absolute;inset:13px;border-radius:50%;background:#fff}.donut>div{position:relative;display:grid;text-align:center}.donut b{font:700 21px system-ui;color:#193d30}.donut small{font-size:10px;color:#859188}.donut-caption{display:grid;gap:8px}.donut-caption span{display:flex;align-items:center;gap:7px;font-size:11px;color:#68776d}.donut-caption i{width:8px;height:8px;border-radius:50%;background:#4f805f}.donut-caption b{margin-left:auto;color:#334f3c}.backup-note{margin-top:12px;color:#7a897e;background:#f4f8f2;border:1px solid #e3ebe0;border-radius:10px;padding:10px 12px;font-size:12px}.save-excel{background:#193f32!important;color:#f2d893!important;border-color:#193f32!important;font-weight:700!important}.save-excel:disabled{opacity:.5;cursor:not-allowed}.excel-cell.pending{background:#fff5d9!important;border-color:#d7a943!important;color:#714d0c!important}.table-foot button:disabled{opacity:.45;cursor:not-allowed}@media(max-width:780px){.stock-visuals{grid-template-columns:1fr}.donut-card{grid-template-columns:130px 1fr}.donut{width:120px;height:120px}}.stock-thumb{width:38px;height:38px;object-fit:cover;border-radius:8px;border:1px solid #e2e8e1;background:#f1f4f0;display:block}.stock-thumb-empty{display:inline-grid;place-items:center;width:38px;height:38px;border-radius:8px;background:#f2f4f1;color:#aab3ab;font:12px system-ui}.excel-grid td:has(.stock-thumb){padding:4px 8px}.admin-full .admin-head{display:flex!important;position:sticky;top:0;z-index:30;flex-shrink:0}.admin-nav-top{display:none;align-items:center;justify-content:space-between;padding:7px 8px 14px;border-bottom:1px solid #ded8ce;margin-bottom:12px}.admin-full .admin-nav-top{display:flex}.admin-nav-top b{font:700 12px system-ui;letter-spacing:.1em;color:#8d7446}.admin-nav-top small{display:block;margin-top:5px;color:#68746b;font:11px system-ui}.admin-nav-top button{width:auto!important;margin:0!important;padding:5px 10px!important;border:1px solid #ddd6cb!important;background:#fff!important;border-radius:8px!important;color:#433!important;cursor:pointer}.excel-cell::-webkit-outer-spin-button,.excel-cell::-webkit-inner-spin-button{-webkit-appearance:none;margin:0}.excel-cell[type=number]{appearance:textfield;-moz-appearance:textfield}.stock-product-name{max-width:220px;white-space:normal!important;color:#405849;font-weight:600}.admin-full .admin-body{flex:1;min-height:0;height:calc(100dvh - 78px);overflow:hidden}.admin-full .admin-content{min-width:0;overflow:auto}.admin-full .admin-nav{overflow-y:auto;overscroll-behavior:contain}#brand-watermark{position:fixed;right:14px;bottom:12px;top:auto;transform:none;background:rgba(15,20,18,.3);color:#fff;border:1px solid rgba(255,255,255,.82);border-radius:8px;padding:7px 10px;font-size:11px;font-weight:700;letter-spacing:.12px;text-shadow:0 1px 3px #1118;box-shadow:0 3px 12px #0003;z-index:80;opacity:.9;white-space:nowrap;pointer-events:none}.stock-price,.stock-value{color:#080808!important;font-weight:850!important;font-size:14px!important}.stock-price{background:#fff8df!important;border-radius:6px}.today-title{font:700 14px system-ui;color:#294b39;margin:0 0 9px}.today-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:0 0 18px}.today-grid>div{padding:13px 15px;border:1px solid #e4e9e4;border-radius:14px;background:#fff;box-shadow:0 4px 14px #183b2a08}.today-grid small{display:block;color:#77847b;font:11px system-ui}.today-grid b{display:block;margin-top:5px;color:#173c30;font:750 23px system-ui}.today-grid .plus b{color:#187345}.today-grid .minus b{color:#b53a34}@media(max-width:760px){.today-grid{grid-template-columns:repeat(2,minmax(0,1fr))}} .stock-price,.stock-value{color:#080808!important;font-weight:800!important;font-size:15px!important;font-family:"Hind Siliguri","Noto Sans Bengali",sans-serif!important;font-variant-numeric:tabular-nums!important;letter-spacing:.1px}.excel-cell{color:#080808!important;font-weight:750!important;font-family:"Hind Siliguri","Noto Sans Bengali",sans-serif!important;font-size:15px!important;font-variant-numeric:tabular-nums!important}#grid .pr{color:#080808!important;font-family:"Hind Siliguri","Noto Sans Bengali",sans-serif!important;font-weight:850!important;font-size:20px!important;font-variant-numeric:tabular-nums!important}.stock-comments{margin:16px 0;padding:16px;border:1px solid #e1e8e1;border-radius:14px;background:#f8faf7}.stock-comments h4{margin:0 0 5px;color:#173c30}.stock-comment-help{margin:0 0 12px;color:#758179;font-size:12px}.stock-comment-form{display:flex;align-items:flex-end;gap:9px;flex-wrap:wrap;padding:12px;background:#fff;border:1px solid #e5e9e4;border-radius:12px}.stock-comment-form label{display:grid;gap:5px;color:#637168;font-size:11px}.stock-comment-form input,.stock-comment-form select{min-width:140px;border:1px solid #d8e0d8;border-radius:8px;padding:9px;font:14px Arial,"Noto Sans Bengali",sans-serif;color:#111;background:white}.stock-comment-form select{appearance:auto;cursor:pointer;min-height:38px}.eye-mask{display:inline-flex;align-items:center;justify-content:center;min-width:42px;color:#76837a;font:18px system-ui;border:1px solid #e2e8e2;border-radius:8px;background:#f6f8f6;padding:3px 7px}.stock-comment-list{display:grid;gap:8px;margin-top:10px}.stock-comment-row{padding:11px 12px;background:#fff;border:1px solid #e6ebe5;border-radius:10px}.stock-comment-row b{display:block;color:#173c30}.stock-comment-row small{display:block;margin-top:3px;color:#78857c;font-size:11px}.stock-comment-row p{margin:6px 0 0;color:#222}#dz{display:flex!important;align-items:center!important;justify-content:center!important;width:min(820px,100%)!important;max-width:100%!important;min-height:82px!important;height:82px!important;padding:8px!important;box-sizing:border-box!important;overflow:hidden!important;font-size:0!important;line-height:0!important;background:var(--card)!important;border:1.5px dashed #5fb58b!important;border-radius:18px!important;margin:4px auto 12px!important;cursor:pointer!important}#dz>*{font-size:0!important;line-height:0!important}#dz small,#dz b{display:none!important}#dz img{display:block!important;max-height:64px!important;max-width:calc(100% - 8px)!important;width:auto!important;height:auto!important;object-fit:contain!important;border-radius:8px!important;margin:0 auto!important}#dz:focus-visible{outline:2px solid #24835b!important;outline-offset:3px!important}#dz.on{background:#eaf8ef!important;border-style:solid!important;box-shadow:0 0 0 3px #d5f2df!important}.search-wrap{position:relative;width:min(680px,94%);margin:14px auto 20px}.search-wrap #q{width:100%;box-sizing:border-box;padding-right:54px}.search-wrap #imageSearchButton{position:absolute;right:6px;top:50%;transform:translateY(-50%);width:40px;height:40px;border:0;border-radius:12px;background:#183f32;color:#f3d894;font-size:19px;cursor:pointer}.backup-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin:18px 0}.backup-card{display:grid;gap:8px;min-height:150px;padding:18px;border:1px solid #e1e7df;border-radius:16px;background:linear-gradient(145deg,#fff,#f7faf6);color:#193d30;text-decoration:none;box-shadow:0 5px 18px #173c300a}.backup-card:hover{transform:translateY(-2px);box-shadow:0 12px 24px #173c3015}.backup-card.primary{background:linear-gradient(140deg,#173d30,#255a42);color:#fff;border-color:#173d30}.backup-card span{font-size:22px}.backup-card b{font-size:16px}.backup-card small{color:#718078}.backup-card.primary small{color:#d6e3d9}.backup-card em{align-self:end;font-style:normal;font-weight:700;color:#a27d37}.backup-card.primary em{color:#f2d893}.mg{align-items:start}.mg>.st{align-self:start;min-height:0;height:auto;max-height:180px;overflow:auto;align-content:start}.mg>.st i{align-self:start;min-height:0;height:auto}.product-detail-grid{display:grid;grid-template-columns:minmax(280px,1fr) minmax(300px,1fr);gap:22px;align-items:start}.product-media #mi{width:100%;height:min(62vh,620px);object-fit:cover;border-radius:14px;background:#eee}.product-summary h2{font-size:clamp(22px,2.3vw,30px);line-height:1.35;margin:0 34px 14px 0;color:#101311}.product-meta{display:flex;align-items:center;gap:8px;color:#7a817d;flex-wrap:wrap}.product-modal-price{font-size:28px;font-weight:850;color:#111;margin:18px 0 10px}.product-modal-price s{font-size:15px;color:#929892;font-weight:500;margin-left:6px}.product-modal-sizes{margin:14px 0}.product-modal-sizes .st{align-content:start;align-items:start}.product-modal-tabs{font-size:12px;color:#788078;margin:10px 0}.product-summary .acts{margin:18px 0 12px}.product-description{font-size:14px;line-height:1.65;color:#68716b;white-space:pre-line;max-height:150px;overflow:auto}.similar-section{margin-top:26px}.similar-section h3{margin:0 0 12px;font-size:19px}@media(max-width:720px){.product-detail-grid{grid-template-columns:1fr;gap:14px}.product-media #mi{height:min(52vh,480px)}.product-summary h2{font-size:22px}.product-modal-price{font-size:24px}}.search-primary{grid-column:1/-1;display:grid!important;grid-template-columns:minmax(180px,28%) minmax(0,1fr);align-items:stretch;border:2px solid #d2ad5f!important;box-shadow:0 12px 32px #92713a20!important}.search-primary .iw>img:first-child{height:100%;min-height:220px;aspect-ratio:1/1;object-fit:cover}.search-primary .m{align-self:center;padding:18px}.search-primary .m b{font-size:20px;height:auto;max-height:3em}.search-primary .sc{font-weight:800}.c .iw .sc .match-percent{display:inline!important;visibility:visible!important;opacity:1!important;color:inherit!important;font:inherit!important;font-weight:900!important;letter-spacing:.02em!important}.c .iw .sc{white-space:nowrap!important;max-width:calc(100% - 16px)!important}.search-primary .stockstate.out+.st{display:none!important}@media(max-width:620px){.search-primary{grid-template-columns:1fr}.search-primary .iw>img:first-child{height:auto;min-height:0;aspect-ratio:4/3}.search-primary .m{padding:12px}}#siteNoticeShade{position:fixed;inset:0;background:#0009;z-index:99999;display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box}#siteNoticeCard{position:relative;max-width:520px;width:100%;max-height:85dvh;overflow:auto;box-sizing:border-box;background:var(--card,#fff);color:var(--ink,#222);border:2px solid var(--acc,#c9a45c);border-radius:18px;padding:24px;box-shadow:0 18px 70px #0005;text-align:center;white-space:pre-wrap;overflow-wrap:anywhere}#siteNoticeCard h3{margin:0 36px 12px;color:var(--acc,#c9a45c)}#siteNoticeClose{position:absolute;right:10px;top:10px;width:34px;height:34px;border:0;border-radius:50%;background:#f0f0f0;color:#222;font-size:22px;line-height:1;display:grid;place-items:center;cursor:pointer;touch-action:manipulation}#siteNoticeClose:focus-visible{outline:3px solid var(--acc,#c9a45c);outline-offset:2px}
/* Sticky premium header */
header{position:sticky!important;top:0;z-index:9000;isolation:isolate;overflow:hidden;backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);box-shadow:0 12px 32px #0004;animation:headerReveal .8s cubic-bezier(.2,.8,.2,1) both}
header::before{z-index:-1;animation:headerGlow 7s ease-in-out infinite alternate}header::after{height:2px;box-shadow:0 0 14px color-mix(in srgb,var(--acc) 75%,transparent)}
header .hl{position:relative;z-index:1;animation:brandFloat 5s ease-in-out infinite}header h1{background:linear-gradient(100deg,var(--acc),#fff2c6,var(--acc));background-size:220% auto;-webkit-background-clip:text;background-clip:text;color:transparent!important;animation:premiumShine 6s linear infinite;text-shadow:0 0 26px color-mix(in srgb,var(--acc) 15%,transparent)}
header .hg{transition:transform .5s cubic-bezier(.2,.8,.2,1),filter .5s;animation:logoBreathe 4s ease-in-out infinite alternate}header .hg:hover{transform:translateY(-3px) rotate(-3deg) scale(1.06);filter:drop-shadow(0 0 12px var(--acc))}
@keyframes premiumShine{to{background-position:220% center}}@keyframes headerGlow{from{opacity:.55;transform:translateX(-3%)}to{opacity:1;transform:translateX(3%)}}@keyframes brandFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-2px)}}@keyframes logoBreathe{from{filter:drop-shadow(0 2px 6px #0005)}to{filter:drop-shadow(0 4px 12px color-mix(in srgb,var(--acc) 35%,transparent))}}@keyframes headerReveal{from{opacity:.65;transform:translateY(-8px)}to{opacity:1;transform:translateY(0)}}
@media(prefers-reduced-motion:reduce){header,header::before,header .hl,header h1,header .hg{animation:none!important;scroll-behavior:auto!important}}
@media(max-width:600px){header{padding:14px 10px 12px!important}header h1{font-size:clamp(18px,5vw,25px)!important;line-height:1.2}header .hg{max-height:34px}#tag{font-size:9px;letter-spacing:.22em}#clk{margin-top:6px;font-size:11px;padding:4px 10px}#menu{left:8px;top:8px;width:38px;height:38px}}
@media(max-width:600px){#siteNoticeShade{padding:12px}#siteNoticeCard{width:100%;max-width:440px;padding:20px 16px 18px;border-radius:16px;font-size:15px;line-height:1.65}#siteNoticeCard h3{font-size:19px;text-align:left;margin-right:36px}#siteNoticeClose{right:8px;top:8px;width:38px;height:38px}}
/* Final clean white sidebar */
.order-history .visual-card,.order-history .excel-stats>div{background:#fff!important;border:1px solid #e8edf3!important;box-shadow:0 4px 16px #17203308!important;color:#202938!important}
.order-history .excel-grid th{background:#f4f7fb!important;color:#354258!important}
.order-history .excel-grid td{background:#fff!important;color:#344054!important}
.order-history .pill{cursor:pointer}

#dr{background:#fff!important;color:#202938!important;backdrop-filter:none!important;box-shadow:8px 0 32px #17203318!important;border-right:1px solid #e7eaf0!important}
.dr-head{background:#fff!important;color:#202938!important;border-bottom:1px solid #edf0f4!important;border-radius:0!important}
.dr-brand b{color:#202938!important}.dr-brand small{color:#7b8494!important}
#dr .dr-x{background:#f4f6f9!important;color:#263244!important;border-color:#e2e6ed!important}
#dr label{color:#667085!important}#dr select{background:#fff!important;color:#202938!important;border-color:#dfe4ec!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){color:#293548!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover{background:#f3f6fb!important;border-color:#dfe6f1!important}
#dr button::before{background:#f4f6fa!important;border-color:#e2e7ef!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle)::after{background:#52637a!important;filter:none!important}
#dr #dsz .on{background:#e9eef7!important;color:#1d2b42!important}



header .hl{perspective:700px}header h1{display:inline-block;background:linear-gradient(105deg,#b47b17 5%,#ffe8a3 28%,#d49b2d 48%,#fff0b5 68%,#a86c0b 92%);background-size:220% auto;background-clip:text;-webkit-background-clip:text;color:transparent!important;-webkit-text-fill-color:transparent;animation:headerGoldSweep 7s ease-in-out infinite;text-shadow:0 2px 14px rgba(239,190,88,.2)}header .hg{animation:headerLogoFloat 4s ease-in-out infinite}header .hg.b{animation-delay:.8s}@keyframes headerGoldSweep{0%,100%{background-position:0% 50%;filter:brightness(.96)}50%{background-position:100% 50%;filter:brightness(1.18)}}@keyframes headerLogoFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-3px)}}.tg{background:#145da0!important;color:#fff!important;border-color:#fff!important}.similar-section .sg .c .sc{display:block!important}@media(prefers-reduced-motion:reduce){header h1,header .hg{animation:none!important}}
header{background:#0f7655!important;background-image:linear-gradient(135deg,#0f7655,#145a42)!important}
/* ERP polish: remove unused quick-mode and personal-list controls */
#quickbtn,.personal-box{display:none!important}
/* Match admin and main header brand */
.admin-full .admin-head{background:linear-gradient(135deg,#0f7655,#145a42)!important;border-bottom-color:#0b6247!important}
.admin-full .admin-head h2{color:#fff!important}
/* Consistent graphical status cards across dashboards */
.courier-stats article,.excel-stats>div,.order-summary-grid .visual-card,.return-summary-grid .visual-card{border-top:4px solid var(--dash-accent,#17865b)!important;box-shadow:0 5px 18px #123b2b12!important}
.courier-stats article:nth-child(2),.excel-stats>div:nth-child(2){--dash-accent:#4d83cf}
.courier-stats article:nth-child(3),.excel-stats>div:nth-child(3){--dash-accent:#d55454}
.courier-stats article:nth-child(4),.excel-stats>div:nth-child(4){--dash-accent:#d89a2b}
.courier-stats article:nth-child(5),.excel-stats>div:nth-child(5){--dash-accent:#8b6fc0}
.similarity-badge,.sc,.sim-score{white-space:nowrap!important;overflow:visible!important;min-width:max-content!important;font-size:12px!important;font-weight:800!important;color:#fff!important;background:#145da0!important;padding:5px 9px!important;border-radius:999px!important;display:inline-flex!important;align-items:center!important;gap:3px!important}

/* ===== Header: Notifications + POS বাটন (সমান আকার) ===== */
header .hdr-actions{position:absolute;right:16px;top:12px;display:flex;gap:10px;z-index:3}
header .hdr-actions #notifyBtn,header .hdr-actions #posBtn{position:static!important;width:132px;height:44px;display:inline-flex;align-items:center;justify-content:center;gap:8px;border-radius:14px;padding:0 10px;font-family:inherit;font-size:14px;font-weight:800;cursor:pointer;transition:transform .15s,box-shadow .15s,filter .15s;box-shadow:0 5px 16px #0000003a,inset 0 1px 0 #ffffff33}
header .hdr-actions #notifyBtn{border:1px solid #ffffff4d;background:linear-gradient(145deg,#ffffff26,#ffffff0d);color:#fff}
header .hdr-actions #posBtn{border:1px solid #f4d58e;background:linear-gradient(135deg,#ffe8a3,#d49b2d);color:#3d2a00}
header .hdr-actions #notifyBtn:hover,header .hdr-actions #posBtn:hover{transform:translateY(-2px);filter:brightness(1.08)}
header .hdr-actions #posBtn[hidden]{display:none!important}
header .hdr-actions .hb-ic{font-size:17px;line-height:1}
@media(max-width:760px){header .hdr-actions{top:8px;right:8px;gap:6px}header .hdr-actions #notifyBtn,header .hdr-actions #posBtn{width:44px;padding:0}header .hdr-actions .hb-tx{display:none}}

/* ===== POS ===== */
#pos{position:fixed;inset:0;z-index:100000;background:#eef3f0;display:none;flex-direction:column;color:#172a22;font-family:inherit}
#pos.on{display:flex}
#pos *{box-sizing:border-box}
#pos button{font-family:inherit;cursor:pointer}
.pos-top{display:flex;align-items:center;gap:12px;padding:10px 18px;background:linear-gradient(135deg,#0f7655,#145a42);color:#fff;flex-wrap:wrap}
.pos-top b{font-size:18px}.pos-top .sp{flex:1}
.pos-chip{background:#ffffff22;border:1px solid #ffffff3d;border-radius:999px;padding:6px 13px;font-size:13px;font-weight:700}
#pos .pos-back{border:1px solid #f4d58e;background:linear-gradient(135deg,#ffe8a3,#d49b2d);color:#3d2a00;border-radius:12px;padding:8px 16px;font-weight:800}
.pos-main{flex:1;display:grid;grid-template-columns:minmax(0,1fr) 390px;min-height:0}
.pos-left{display:flex;flex-direction:column;min-height:0;padding:12px;gap:8px}
.pos-search{position:relative}
.pos-search>span{position:absolute;left:13px;top:50%;transform:translateY(-50%);opacity:.6}
#pos #posQ{width:100%;padding:13px 14px 13px 40px;border:2px solid #17865b;border-radius:12px;font-size:16px;background:#fff;color:#172a22;outline:none}
#pos #posQ:focus{box-shadow:0 0 0 4px #17865b26}
.pos-hints{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:#60756a}
.pos-hints kbd{background:#dfece4;border-radius:5px;padding:1px 6px;font-size:11px;margin-right:4px;font-family:inherit;font-weight:700}
.pos-results{background:#fff;border:1px solid #d4e5db;border-radius:14px;max-height:40vh;overflow:auto;box-shadow:0 6px 18px #123b2b12}
.pos-row{display:flex;gap:12px;align-items:center;padding:9px 12px;border-bottom:1px solid #edf3ef;cursor:pointer}
.pos-row:last-child{border-bottom:0}
.pos-row.hl{background:#e6f6ed;box-shadow:inset 4px 0 0 #17865b}
.pos-row img,.pos-item img{width:48px;height:48px;object-fit:cover;border-radius:9px;background:#eef3f0;flex:0 0 auto}
.pos-row .pi,.pos-item .pt{flex:1;min-width:0}
.pos-row .pi b,.pos-item .pt b{display:block;font-size:14px}
.pos-row .pi small,.pos-item small{display:block;color:#60756a;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pos-store{font-size:10px;font-weight:700;background:#e3efe8;color:#12613f;border-radius:6px;padding:1px 6px;margin-left:4px}
.pos-row .pp{font-weight:800;color:#12613f;white-space:nowrap}
.pos-sz{display:flex;flex-wrap:wrap;gap:6px;margin-top:5px}
#pos .pos-sz button{border:1px solid #b9dfc9;background:#fff;color:#172a22;border-radius:9px;padding:4px 10px;font-size:12px;font-weight:700}
#pos .pos-sz button:hover{background:#17865b;color:#fff;border-color:#17865b}
.pos-sz button i{font-style:normal;opacity:.65;margin-left:5px;font-weight:600}
.pos-sz button em{font-style:normal;background:#d89a2b;color:#fff;border-radius:5px;padding:0 5px;margin-right:5px}
.pos-out{font-size:12px;font-weight:700;color:#b42318;background:#fdecea;border-radius:8px;padding:3px 9px}
.pos-empty{padding:22px;text-align:center;color:#7a8c82;font-size:14px}
.pos-cart{flex:1;min-height:120px;overflow:auto;background:#fff;border:1px solid #d4e5db;border-radius:14px}
.pos-cart-h{position:sticky;top:0;background:#f6faf8;padding:10px 14px;font-weight:800;border-bottom:1px solid #e1ece5;z-index:1}
.pos-badge{display:inline-grid;place-items:center;min-width:22px;height:22px;border-radius:22px;background:#17865b;color:#fff;font-size:12px;padding:0 6px;margin-left:6px}
.pos-item{display:flex;gap:12px;align-items:flex-start;margin:10px;padding:10px;background:#f8fbf9;border:1px solid #e1ece5;border-radius:12px}
.pos-qty{display:inline-flex;align-items:center;gap:5px;margin-top:6px}
#pos .pos-qty button{width:30px;height:30px;border:1px solid #cfe0d6;background:#fff;border-radius:8px;font-size:16px;font-weight:800;color:#172a22}
#pos .pos-qty input{width:52px;height:30px;text-align:center;border:1px solid #cfe0d6;border-radius:8px;font:inherit;font-weight:800}
.pos-price-edit{display:flex!important;align-items:center;gap:7px;margin:5px 0;color:#24533f!important;font-size:11px!important}.pos-price-edit input{width:108px;max-width:100%;height:30px;padding:4px 7px;border:1px solid #b7d9c4;border-radius:7px;background:#fff;font:600 13px system-ui;color:#183d2b}.pos-item .pl{text-align:right;display:flex;flex-direction:column;align-items:flex-end;gap:3px}
.pos-item .pl b{font-size:15px}
#pos .pos-del{border:0;background:transparent;color:#d92d20;font-size:17px;padding:0}
.pos-right{background:#fff;border-left:1px solid #d4e5db;padding:14px 16px;overflow:auto;display:flex;flex-direction:column;gap:9px}
.bizdash-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:12px 0 16px}.bizdash-tile{min-width:0;padding:15px;border:1px solid #dfe9e1;border-radius:14px;background:linear-gradient(145deg,#fff,#f6faf7);display:flex;flex-direction:column;gap:6px}.bizdash-tile.primary{background:linear-gradient(135deg,#174b36,#24684b);border-color:#174b36;color:#fff}.bizdash-tile small{font-size:12px;color:#687d70}.bizdash-tile.primary small,.bizdash-tile.primary em{color:#d9ebdf}.bizdash-tile b{font-size:clamp(17px,2vw,24px);font-weight:850;color:inherit;overflow-wrap:anywhere}.bizdash-tile em{font-style:normal;font-size:11px;color:#79877f}.bizdash-columns{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.bizdash-columns .chart-card{min-width:0}.bizdash-note{display:block;margin-top:10px;font-size:11px;line-height:1.5;color:#7c8981;overflow-wrap:anywhere}@media(max-width:900px){.bizdash-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.bizdash-columns{grid-template-columns:1fr}}@media(max-width:480px){.bizdash-grid{gap:7px}.bizdash-tile{padding:11px;border-radius:11px}.bizdash-tile small{font-size:11px}.bizdash-tile b{font-size:17px}.bizdash-tile em{font-size:10px}.pos-customer-history{max-height:36vh;overflow:auto;-webkit-overflow-scrolling:touch}.pos-right input,.pos-right select{min-height:42px}.pos-customer-suggestions button{min-height:42px;touch-action:manipulation}}
.pos-customer-suggestions{display:flex;flex-direction:column;gap:4px;padding:6px 0 9px;border-bottom:1px solid #e0eae3;margin-bottom:8px}.pos-customer-suggestions>b{font-size:11px;color:#587565;padding:0 2px}.pos-customer-suggestions button{display:flex;flex-direction:column;align-items:flex-start;gap:2px;width:100%;padding:7px 9px;text-align:left;border:1px solid #dce9df;border-radius:8px;background:#fff;color:#173c2c;cursor:pointer}.pos-customer-suggestions button:hover{background:#edf7f0}.pos-customer-suggestions small{font-size:11px;color:#60766a}.pos-customer-history{margin:-2px 0 5px;border:1px solid #d6e6dc;border-radius:11px;background:#f7fbf8;overflow:hidden;color:#203d2e}.pos-customer-history[hidden]{display:none!important}.pos-history-title{padding:8px 10px;background:#e3f2e8;color:#20533a;font-size:11px;font-weight:800}.pos-history-muted{padding:9px 10px;color:#677b6d;font-size:11px}.pos-history-row{padding:9px 10px;border-top:1px solid #e3ebe5}.pos-history-order,.pos-history-money{display:flex;align-items:center;justify-content:space-between;gap:6px}.pos-history-order b{font-size:11px;color:#174d36}.pos-history-order small,.pos-history-products{font-size:10px;color:#68786e}.pos-history-products{margin:4px 0;overflow-wrap:anywhere}.pos-history-money{font-size:10px;color:#52695a}.pos-history-money b{color:#183d2d}.pos-history-money .is-due{color:#b42318}.pos-history-row button{margin-top:6px!important;padding:5px 8px!important;border:1px solid #b9d8c4!important;border-radius:7px!important;background:#fff!important;color:#1c6944!important;font-size:10px!important;cursor:pointer}
.pos-right h4{margin:8px 0 0;font-size:12px;letter-spacing:.08em;color:#12613f;text-transform:uppercase}
.pos-right label{display:block;font-size:12px;color:#60756a;font-weight:700}
#pos .pos-right input,#pos .pos-right select,#pos .pos-right textarea{display:block;width:100%;margin-top:4px;padding:9px 10px;border:1px solid #cfe0d6;border-radius:10px;font:inherit;font-size:14px;color:#172a22;background:#fff}
#pos .pos-right input:focus,#pos .pos-right select:focus,#pos .pos-right textarea:focus{outline:none;border-color:#17865b;box-shadow:0 0 0 3px #17865b22}
.pos-two{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.pos-sum{background:#f3f8f5;border:1px solid #e1ece5;border-radius:12px;padding:10px 13px}
.pos-sum div{display:flex;justify-content:space-between;padding:3px 0;font-size:14px}
.pos-sum .tot{border-top:1px solid #d4e5db;margin-top:5px;padding-top:8px;font-size:19px;font-weight:800;color:#0f7655}
#pos #posCh.due{color:#b42318;font-weight:800}
#pos .pos-pay{margin-top:4px;background:linear-gradient(135deg,#0f7655,#22a06b);color:#fff;border:0;border-radius:13px;padding:14px;font-size:16px;font-weight:800;box-shadow:0 6px 16px #0f765540}
#pos .pos-pay:disabled{opacity:.5;cursor:not-allowed}
#pos .pos-clear{background:#f1f4f2;border:1px solid #dfe6e2;border-radius:12px;padding:10px;font-weight:700;color:#344}
.pos-msg{background:#fdecea;color:#b42318;border:1px solid #f4b8b1;border-radius:10px;padding:9px 11px;font-size:13px;font-weight:600}
#posRcpt{position:absolute;inset:0;background:#0a1a13b8;display:flex;align-items:center;justify-content:center;z-index:5}
#posRcpt[hidden]{display:none}
.pos-rc-box{background:#fff;border-radius:18px;width:min(420px,94vw);max-height:92vh;overflow:auto;padding:18px;box-shadow:0 20px 60px #0007}
.pos-rc-ok{background:#e6f6ed;color:#0f7655;border-radius:10px;padding:9px 12px;font-weight:800;margin-bottom:12px;text-align:center}
#posRcBody table{width:100%;border-collapse:collapse;font-size:13px}
#posRcBody td{padding:5px 0;border-bottom:1px dashed #cdd8d2;vertical-align:top}
#posRcBody .r-h{text-align:center;margin-bottom:8px}#posRcBody .r-h b{display:block;font-size:16px}
#posRcBody .r-s{display:flex;justify-content:space-between;margin:5px 0;font-size:14px}
#posRcBody .r-s.tot{font-size:17px;font-weight:800}
#posRcBody .r-f{text-align:center;color:#60756a;margin:10px 0 0}
.pos-rc-act{display:flex;gap:10px;margin-top:14px}
#pos .pos-rc-act button{flex:1;padding:12px;border-radius:12px;border:1px solid #cfe0d6;background:#f6faf8;font-weight:800;font-size:14px}
#pos .pos-rc-act .pri{background:linear-gradient(135deg,#0f7655,#22a06b);color:#fff;border:0}
#toast{z-index:9500!important}
@media(max-width:900px){.pos-main{grid-template-columns:1fr;overflow:auto}.pos-left{min-height:auto}.pos-right{border-left:0;border-top:1px solid #d4e5db}.pos-results{max-height:34vh}}


/* Size filter is collapsed by default; expand only when the user asks. */
#dr .size-filter-head{display:flex;align-items:center;justify-content:space-between;gap:10px;cursor:pointer;user-select:none;margin:15px 4px 6px;padding:10px 12px;border:1px solid rgba(255,255,255,.24);border-radius:12px;background:rgba(255,255,255,.08);color:inherit;font:700 11px system-ui,'Hind Siliguri',sans-serif;letter-spacing:.1em;text-transform:uppercase}
#dr .size-filter-head:focus-visible{outline:2px solid #c6f0d7;outline-offset:2px}
#dr .size-filter-head small{font-size:10px;letter-spacing:0;text-transform:none;font-weight:500;opacity:.8}
#dr .size-filter-chevron{font-size:17px;line-height:1;opacity:.9}
#dr #dsz[hidden]{display:none!important}
</style></head><body>
<style id="global-mobile-friendly-fixes">
*{min-width:0}html{scrollbar-gutter:stable;-webkit-text-size-adjust:100%}body{overflow-x:hidden}main{width:100%;max-width:100%;box-sizing:border-box}
button,input,select,textarea,a{touch-action:manipulation}img,video,canvas{max-width:100%}
.admin-content,.admin-section,.erp-module,.hr-shell,.excel-hero,.visual-card{min-width:0;max-width:100%;box-sizing:border-box}
.excel-table-wrap,.admin-table-wrap,.hr-table-wrap,.hr-table-wrap,.table-wrap{max-width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}
.stock-entry-grid{grid-template-columns:repeat(auto-fit,minmax(min(100%,180px),1fr))!important}
.erp-pi-stock-hint{display:block;color:#6b7b71;font-size:11px;line-height:1.5;margin-top:5px;overflow-wrap:anywhere}
.erp-pi-name{width:100%;min-width:160px;box-sizing:border-box}.erp-pi-size{min-width:110px;max-width:100%;box-sizing:border-box}
@media(max-width:760px){
 header{box-sizing:border-box;padding:12px 12px 62px!important;min-height:112px!important;gap:8px!important;overflow:visible!important}
 header .hl{width:100%;box-sizing:border-box;padding:36px 8px 0 42px!important;min-width:0!important}
 header .hl h1,#ttl{font-size:clamp(15px,4vw,20px)!important;line-height:1.35!important;overflow-wrap:anywhere}
 header #tag,header #clk{display:none!important}
 header .hdr-actions{top:8px!important;right:8px!important;display:flex!important;gap:5px!important}
 header .hdr-actions button{padding:7px 9px!important;min-height:36px!important;font-size:12px!important}
 header .hdr-actions .hb-tx{display:none!important}
 header #quickOrderLookup{position:absolute!important;left:12px!important;right:12px!important;top:auto!important;bottom:8px!important;width:auto!important;height:40px!important;margin:0!important}
 header #quickOrderLookup .quick-order-control{width:100%!important;height:40px!important;display:flex!important}
 header #quickOrderLookup input{min-width:0!important;width:100%!important;font-size:13px!important}
 main{padding:10px!important}
 #dz{width:100%!important;max-width:100%!important;min-height:62px!important;margin:4px auto 10px!important;box-sizing:border-box}
 .search-wrap{width:100%!important;max-width:100%!important;margin:8px auto 14px!important}
 #grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:9px!important}
 #grid>.c{min-width:0!important}
 .admin-full{padding:0!important;align-items:stretch!important}
 .admin-full>.box.ad{width:100vw!important;max-width:100vw!important;height:100dvh!important;max-height:100dvh!important;margin:0!important;border-radius:0!important}
 .admin-full .admin-head{padding:10px 12px!important;gap:8px!important}
 .admin-full .admin-head h2{font-size:17px!important}
 .admin-full .admin-body{display:flex!important;flex-direction:column!important;height:auto!important;min-height:0!important;max-height:calc(100dvh - 62px)!important;overflow:hidden!important}
 .admin-full .admin-nav{display:flex!important;flex:0 0 auto!important;flex-wrap:nowrap!important;gap:5px!important;max-height:74px!important;overflow:auto!important;padding:7px!important;border-right:0!important;border-bottom:1px solid #ded8ce!important}
 .admin-full .admin-nav button{width:auto!important;min-width:max-content!important;padding:7px 9px!important;margin:0!important;font-size:12px!important}
 .admin-full .admin-content{flex:1 1 auto!important;width:100%!important;max-width:100%!important;min-height:0!important;padding:10px!important;overflow:auto!important}
 .admin-full .admin-nav-top{padding:5px!important;margin:0!important}
 .admin-content .dash-hero,.admin-content .row,.admin-content .excel-actions,.admin-content .dash-actions{gap:8px!important;flex-wrap:wrap!important;align-items:stretch!important}
 .admin-content .excel-tools{max-width:100%!important}
 .admin-content input:not([type=checkbox]):not([type=radio]):not([type=file]),.admin-content select,.admin-content textarea{max-width:100%!important;box-sizing:border-box!important}
 .admin-content .hr-form-grid{grid-template-columns:repeat(1,minmax(0,1fr))!important}
 .admin-content .hr-form-grid .wide{grid-column:auto!important}
 .admin-content .hr-stats{grid-template-columns:repeat(2,minmax(0,1fr))!important}
 .admin-content .co-grid,.admin-content .order-summary-grid{grid-template-columns:minmax(0,1fr)!important}
 .admin-content .excel-stats{grid-template-columns:repeat(2,minmax(0,1fr))!important}
 .admin-content .chart-grid,.admin-content .stock-visuals{grid-template-columns:minmax(0,1fr)!important}
 .admin-content .stock-entry-grid{grid-template-columns:minmax(0,1fr)!important}
 .admin-content .erp-pi-name{min-width:170px!important}
 .admin-content .excel-grid{min-width:720px}
 .admin-content .order-grid{min-width:1050px}
 .admin-content .pill,.admin-content button{min-height:38px}
 #brand-watermark{right:7px!important;bottom:7px!important;max-width:72vw!important;white-space:normal!important;font-size:9px!important;padding:5px 7px!important}
}
@media(max-width:420px){#grid{gap:7px!important}.c .m{padding:8px!important}.c .m>b{font-size:12px!important}.admin-content .excel-stats{grid-template-columns:1fr 1fr!important}.admin-content .dash-hero h3{font-size:18px!important}}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{scroll-behavior:auto!important;animation-duration:.01ms!important;transition-duration:.01ms!important}}
#dz{display:none!important;width:0!important;height:0!important;min-height:0!important;margin:0!important;padding:0!important;border:0!important;visibility:hidden!important}
</style>
<div id="brand-watermark">XEROX-S1AM · © 2026 All Rights Reserved</div>
<header><button id="menu" onclick="dr(1)">☰</button><div class="hdr-actions"><button id="notifyBtn" onclick="openNotifications()" title="Notifications"><span class="hb-ic">🔔</span><span class="hb-tx">Notifications</span><span id="notifyCount">0</span></button><button id="posBtn" type="button" onclick="event.preventDefault();window.openPOS ? window.openPOS() : alert('POS is still loading. Please try again.')" title="POS Sales (F9)" hidden><span class="hb-ic">🛒</span><span class="hb-tx">POS</span></button></div><div class="hl"><h1 id="ttl">Alien Tees ERP Management System</h1></div><div id="tag">ERP MANAGEMENT SYSTEM</div><div id="clk"></div>
<section id="quickOrderLookup" aria-label="Order search"><div class="quick-order-control"><input id="orderQuickSearch" placeholder="Consignment ID / Customer Number" autocomplete="off" inputmode="search" onkeydown="if(event.key==='Enter'){event.preventDefault();runOrderLookup()}"><button type="button" onclick="runOrderLookup()" aria-label="Search orders" title="Search orders">⌕</button></div><div id="quickOrderResults" hidden></div></section>
</header>
<div id="prog"><div class="b"><i id="pb"></i></div><p id="pt"></p></div>
<main>
<input type="file" id="fi" accept="image/*" hidden>
 <div class="search-wrap"><div id="imageSearchPreview" hidden aria-label="Selected image preview" title="Selected image"><img id="imageSearchPreviewImg" alt="Selected image"></div><input id="q" placeholder="Search by code or name (e.g. 1R01 or Polo)"><button id="imageSearchButton" type="button" title="Search by image" aria-label="Search by image" onclick="document.getElementById('fi').click()">▧</button></div>
<style>
/* Selected image preview is shown inside the search field; no empty drop-zone is required. */
#imageSearchPreview[hidden]{display:none!important}
#imageSearchPreview:not([hidden]){position:absolute!important;left:7px!important;top:50%!important;transform:translateY(-50%)!important;width:40px!important;height:40px!important;display:grid!important;place-items:center!important;z-index:6!important;overflow:hidden!important;border:1px solid #b8d8c4!important;border-radius:10px!important;background:#f3faf5!important;box-shadow:0 2px 7px #173b2a14!important;pointer-events:none!important}
#imageSearchPreview img{display:block!important;width:100%!important;height:100%!important;max-height:none!important;max-width:none!important;object-fit:contain!important;border-radius:8px!important}
.search-wrap.has-image #q{padding-left:55px!important}
/* Compact order search stays inside the existing header, directly beneath Notifications/POS. */
/* Header quick search / POS click reliability fix; preserve existing layout and artwork. */
header{position:relative!important;top:auto!important;overflow:visible!important;isolation:isolate;z-index:2147482000!important}
/* Three matching header controls: Consignment Search, Notifications, and POS. */
header #quickOrderLookup{position:absolute!important;right:300px!important;top:12px!important;width:132px!important;height:44px!important;margin:0!important;z-index:2147482001!important;box-sizing:border-box;pointer-events:auto!important;isolation:isolate}
header #quickOrderLookup .quick-order-control,header #quickOrderLookup #quickOrderResults{pointer-events:auto!important}
header .hdr-actions{position:absolute!important;right:16px!important;top:12px!important;z-index:2147482002!important;pointer-events:auto!important}
header .hdr-actions button,header #posBtn,header #notifyBtn{position:relative!important;pointer-events:auto!important;z-index:2147482003!important;touch-action:manipulation;cursor:pointer}
header .quick-order-control{pointer-events:auto!important;z-index:511!important}
header .quick-order-control input,header .quick-order-control button{pointer-events:auto!important;touch-action:manipulation}
header .quick-order-control{display:flex;gap:0;min-width:0;width:132px;height:44px;border:0!important;border-radius:14px;overflow:hidden;background:#fff;box-shadow:0 5px 16px #0000003a,inset 0 1px 0 #ffffff33;position:relative;z-index:2147482003;pointer-events:auto!important}
header .quick-order-control input{display:block;flex:1 1 auto;width:1px;min-width:0;height:44px;padding:0 7px;border:0!important;border-radius:0;background:#fff!important;color:#111!important;-webkit-text-fill-color:#111!important;font:11px system-ui,sans-serif;outline:none;box-shadow:none!important;pointer-events:auto!important;position:relative;z-index:2147482004!important;user-select:text!important;cursor:text!important}
header .quick-order-control input::placeholder{color:#111!important;-webkit-text-fill-color:#111!important;opacity:1!important}
header .quick-order-control button{width:34px;height:44px;flex:0 0 34px;border:0;border-radius:0;background:linear-gradient(135deg,#ffe8a3,#d49b2d);color:#263449;font-size:20px;line-height:1;cursor:pointer;position:relative;z-index:2147482004!important;pointer-events:auto!important}
#quickOrderResults{position:absolute;right:-284px;top:48px;width:min(1200px,calc(100vw - 32px));max-height:76vh;overflow:auto;padding:18px;background:#fff;color:#17212d;border:1px solid #dce4ed;border-radius:10px;box-shadow:0 10px 26px #0002;z-index:2147482005;pointer-events:auto}
.quick-order-summary{font:700 13px system-ui,sans-serif;color:#173d31;margin-bottom:10px}.quick-order-table{width:100%;min-width:1100px;border-collapse:collapse;font:13px system-ui,sans-serif}.quick-order-table th,.quick-order-table td{text-align:left;padding:10px 9px;border-bottom:1px solid #edf0ed;white-space:nowrap}.quick-order-table th{background:#f3f7f4;color:#52665a}.quick-order-empty{color:#7c867f;padding:10px 0}.quick-order-summary{display:flex;align-items:center;justify-content:space-between;gap:12px}.quick-order-close{flex:0 0 auto;border:0;background:#e5e7eb;color:#111827;border-radius:7px;width:28px;height:28px;font-size:20px;line-height:1;cursor:pointer}.quick-order-close:hover{background:#d1d5db}
@media(max-width:980px) and (min-width:761px){header #quickOrderLookup{right:16px!important;top:62px!important;width:264px!important;height:36px!important}header .quick-order-control{width:264px;height:36px;border-radius:10px}header .quick-order-control input{height:36px;font-size:12px;padding:0 9px}header .quick-order-control button{height:36px}#quickOrderResults{right:0;top:40px;width:min(1200px,calc(100vw - 32px))}}
@media(max-width:760px){header #quickOrderLookup{right:8px!important;top:58px!important;width:220px!important;height:36px!important}header .quick-order-control{width:220px;height:36px;border-radius:10px}header .quick-order-control input{height:36px;font-size:11px;padding:0 7px}header .quick-order-control button{height:36px}#quickOrderResults{right:-4px;top:40px;width:calc(100vw - 16px);max-height:76vh;padding:10px}.quick-order-table{min-width:1050px}}
</style>
<style>
#dz{display:flex!important;min-height:78px!important;height:auto!important;align-items:center!important;justify-content:center!important;padding:12px!important;font-size:0!important;line-height:0!important;color:transparent!important;cursor:pointer!important;touch-action:manipulation}
#dz small{display:none!important}
#dz:focus-visible{outline:3px solid var(--acc,#c8a45f);outline-offset:3px}
#dz.on{border-style:solid!important;box-shadow:0 0 0 3px color-mix(in srgb,var(--acc,#c8a45f) 18%,transparent)!important}
.has-stock-comments{background:#fff9e7!important}.commented-code{background:#ffefb8;color:#111!important;padding:3px 6px;border-radius:6px}.stock-comment-jump{margin:3px;padding:4px 7px;border:1px solid #d5bd78;border-radius:7px;background:#fff8df;color:#604a16;font:inherit;font-size:12px;cursor:pointer}
</style>

<style id="cool-search-loader-style">
#searchLoader{position:fixed;inset:0;z-index:99990;display:flex;align-items:center;justify-content:center;padding:18px;background:rgba(3,18,13,.30);backdrop-filter:blur(5px);-webkit-backdrop-filter:blur(5px);pointer-events:none;opacity:1}
#searchLoader[hidden]{display:none!important}
.search-loader-card{position:relative;isolation:isolate;overflow:hidden;width:min(360px,calc(100vw - 34px));padding:24px 26px 22px;border-radius:26px;border:1px solid rgba(220,255,238,.55);background:linear-gradient(145deg,rgba(7,49,35,.98),rgba(13,117,82,.97) 58%,rgba(8,50,39,.98));box-shadow:0 25px 85px rgba(0,0,0,.45),inset 0 1px 0 rgba(255,255,255,.26);color:#fff;text-align:center;transform:translateY(8px) scale(.96);opacity:0;transition:opacity .24s ease,transform .28s cubic-bezier(.2,.8,.2,1)}
#searchLoader.is-active .search-loader-card{transform:translateY(0) scale(1);opacity:1}
.search-loader-card:before{content:"";position:absolute;inset:-60% -40%;z-index:-1;background:linear-gradient(108deg,transparent 38%,rgba(255,255,255,.095) 48%,rgba(208,255,226,.15) 51%,transparent 61%);transform:translateX(-65%);animation:searchGlassSweep 2.8s ease-in-out infinite}
.search-loader-scene{position:relative;width:148px;height:128px;margin:0 auto 10px;display:grid;place-items:center;filter:drop-shadow(0 0 18px rgba(111,255,185,.16))}
.search-loader-ring{position:absolute;left:50%;top:50%;border-radius:50%;transform:translate(-50%,-50%);box-sizing:border-box}
.search-loader-ring.r1{width:104px;height:104px;border:2px solid transparent;border-top-color:#efffd0;border-right-color:#66ffc0;box-shadow:0 0 13px #5dffb855;animation:searchRingSpin 1.5s linear infinite}
.search-loader-ring.r2{width:78px;height:78px;border:1px dashed rgba(222,255,238,.75);animation:searchRingSpin 2.2s linear infinite reverse}
.search-loader-ring.r3{width:130px;height:52px;border:1px solid rgba(255,226,147,.9);transform:translate(-50%,-50%) rotate(-32deg);animation:searchOrbit 3s ease-in-out infinite}
.search-loader-core{position:relative;width:43px;height:43px;display:grid;place-items:center;border-radius:15px;background:linear-gradient(145deg,#f4ffe4,#93ffd0 48%,#20bd7c);border:1px solid #f4ffef;box-shadow:0 0 13px rgba(122,255,190,.75),0 0 36px rgba(49,255,160,.35);animation:searchCorePulse 1.4s ease-in-out infinite}
.search-loader-core:before{content:"✦";color:#075b3a;font-size:25px;line-height:1;text-shadow:0 1px 0 #fff8}
.search-loader-spark{position:absolute;width:7px;height:7px;border-radius:50%;background:#fff4c0;box-shadow:0 0 12px #ffe9a0;animation:searchSpark 1.8s ease-in-out infinite}
.search-loader-spark.s1{top:18px;left:30px}.search-loader-spark.s2{right:20px;top:50px;width:5px;height:5px;animation-delay:-.6s}.search-loader-spark.s3{bottom:14px;left:48px;width:4px;height:4px;animation-delay:-1.1s}
.search-loader-title{margin:4px 0 5px;font:800 19px/1.35 system-ui,'Noto Sans Bengali',sans-serif;letter-spacing:.01em;color:#fff}
.search-loader-sub{margin:0;color:rgba(238,255,246,.82);font:12px/1.65 system-ui,'Noto Sans Bengali',sans-serif}
.search-loader-progress{height:4px;margin:18px 4px 0;border-radius:99px;overflow:hidden;background:rgba(255,255,255,.18);box-shadow:inset 0 1px 2px rgba(0,0,0,.15)}
.search-loader-progress i{display:block;width:34%;height:100%;border-radius:inherit;background:linear-gradient(90deg,#d5ffe8,#fff1b7,#e5fff1);box-shadow:0 0 12px #d8ffe8;animation:searchProgressRun 1.15s cubic-bezier(.65,0,.35,1) infinite}
.search-loader-foot{margin-top:12px;display:flex;align-items:center;justify-content:center;gap:7px;color:rgba(239,255,245,.72);font:700 9px/1.3 system-ui,sans-serif;letter-spacing:.2em;text-transform:uppercase}
.search-loader-foot i{width:5px;height:5px;border-radius:50%;background:#c9ffe0;box-shadow:0 0 9px #a5ffcb;animation:searchDotPulse 1s ease-in-out infinite}
@keyframes searchRingSpin{to{transform:translate(-50%,-50%) rotate(360deg)}}
@keyframes searchOrbit{0%,100%{transform:translate(-50%,-50%) rotate(-32deg) scale(.96)}50%{transform:translate(-50%,-50%) rotate(32deg) scale(1.05)}}
@keyframes searchCorePulse{0%,100%{transform:rotate(-8deg) scale(.94);box-shadow:0 0 13px rgba(122,255,190,.65),0 0 25px rgba(49,255,160,.25)}50%{transform:rotate(8deg) scale(1.07);box-shadow:0 0 20px rgba(188,255,208,.95),0 0 42px rgba(49,255,160,.48)}}
@keyframes searchSpark{0%,100%{opacity:.28;transform:scale(.65)}50%{opacity:1;transform:scale(1.45)}}
@keyframes searchProgressRun{0%{transform:translateX(-120%)}100%{transform:translateX(330%)}}
@keyframes searchGlassSweep{0%,15%{transform:translateX(-65%)}75%,100%{transform:translateX(65%)}}
@keyframes searchDotPulse{0%,100%{opacity:.3;transform:scale(.75)}50%{opacity:1;transform:scale(1.25)}}
@media(prefers-reduced-motion:reduce){.search-loader-card,.search-loader-card:before,.search-loader-ring,.search-loader-core,.search-loader-spark,.search-loader-progress i,.search-loader-foot i{animation:none!important;transition:none!important}}
</style>
<div id="searchLoader" hidden role="status" aria-live="polite" aria-busy="false">
 <div class="search-loader-card"><div class="search-loader-scene" aria-hidden="true"><span class="search-loader-ring r1"></span><span class="search-loader-ring r2"></span><span class="search-loader-ring r3"></span><span class="search-loader-core"></span><i class="search-loader-spark s1"></i><i class="search-loader-spark s2"></i><i class="search-loader-spark s3"></i></div><h3 id="searchLoadingTitle" class="search-loader-title">প্রোডাক্ট সার্চ চলছে</h3><p id="searchLoadingSub" class="search-loader-sub">প্রোডাক্ট কোড, নাম ও ক্যাটালগ স্ক্যান করা হচ্ছে…</p><div class="search-loader-progress"><i></i></div><div class="search-loader-foot"><i></i> ALIEN TEES · SMART SEARCH <i></i></div></div>
</div>
<div class="tabs" id="tabs"></div><div id="info"></div><div id="grid"></div><div id="sent"></div>
</main>
<footer><b>XEROX-S1AM</b> · © 2026 All Rights Reserved</footer>

<div id="bd" onclick="dr(0);"></div>
<aside id="dr"><div class="dr-head"><div class="dr-brand"><img src="/logo/a.png" alt=""><div><b>Smart ERP Management Panel</b><small>Alien Tees</small></div></div><button class="dr-x" type="button" onclick="dr(0)" aria-label="Close sidebar">✕</button></div><div id="sync"></div>
 <label>Category</label><select id="ft" onchange="tabf=this.value;drawTabs();drawSizes();render(1)"><option value="">All categories</option></select>
<div class="size-filter-head" data-perm="size-filter-label" role="button" tabindex="0" aria-controls="dsz" aria-expanded="false" onclick="toggleSizeFilters()" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleSizeFilters()}"><span>SIZE <small id="szh"></small></span><span class="size-filter-chevron" id="sizeFilterChevron" aria-hidden="true">＋</span></div><div id="dsz" hidden></div>
 <label>Stock</label><select id="fk" onchange="render(1)"><option value="">All products</option><option value="1">In stock only</option><option value="0">Out of stock</option></select>
 <label>Sort</label><select id="so" onchange="render(1)"><option value="">Default</option><option value="a">Price: low to high</option><option value="d">Price: high to low</option></select>
<label>Selection</label><button onclick="selAll();dr(0)">☑ Select All / Clear</button><button data-perm="download" onclick="dr(0);dlSel()">⬇ Download Selected Images</button><button data-perm="copy" onclick="copySel()">Copy Images</button>
<div id="quickbtn" hidden></div><div id="personal" hidden></div>





<div class="dr-menu-group" data-menu-group="operations" hidden>
<label class="dr-sec">Sales & Operations</label>
<button class="direct-admin-link" data-direct="posorders" onclick="adminTab('posorders')">🧾 POS Orders & Memos</button>
<button class="direct-admin-link" data-direct="couriers" onclick="adminTab('couriers')">🚚 Courier Tracking</button>
<button class="direct-admin-link" data-direct="purchases" onclick="adminTab('purchases')">🛒 Purchase Management</button>
<button class="direct-admin-link" data-direct="suppliers" onclick="adminTab('suppliers')">🏭 Suppliers</button>
<button class="direct-admin-link" data-direct="expenses" onclick="adminTab('expenses')">💸 Expense Management</button>
<button class="direct-admin-link" data-direct="erpReports" onclick="adminTab('erpReports')">📊 Purchase & Expense Reports</button>
<button class="direct-admin-link" data-direct="price" onclick="adminTab('price')">💰 Price Editor</button>

<button class="direct-admin-link" data-direct="orders" onclick="adminTab('orders')">🧾 Order History</button>
<button class="direct-admin-link" data-direct="returns" onclick="adminTab('returns')">↩ Return History</button>
</div>
<div class="dr-menu-group" data-menu-group="hr" hidden>
<label class="dr-sec">People & HR</label>
<button class="direct-admin-link" data-direct="hr" onclick="adminTab('hr')">👥 HR Management</button>
</div>
<div class="dr-menu-group" data-menu-group="inventory" hidden>
<label class="dr-sec">Stock & Catalog</label>
<button class="direct-admin-link" data-direct="excelstock" onclick="adminTab('excelstock')">📦 Stock Management</button>
<button class="direct-admin-link" data-direct="stockentry" onclick="adminTab('stockentry')">➕ Stock Entry</button>
<button class="direct-admin-link" data-direct="dashboard" onclick="adminTab('dashboard')">📊 Dashboard</button>
<button class="direct-admin-link" data-direct="cat" onclick="adminTab('cat')">🛍 Catalog</button>

</div>
<div class="dr-menu-group" data-menu-group="settings" hidden>
<label class="dr-sec">Administration & Settings</label>
<button class="direct-admin-link" data-direct="company" onclick="adminTab('company')">🏢 My Company</button>
<button class="direct-admin-link" data-direct="notice" onclick="adminTab('notice')">📢 Notices</button>
<button class="direct-admin-link" data-direct="photosettings" onclick="adminTab('photosettings')">📸 Image Download Settings</button>
<button class="direct-admin-link" data-direct="design" onclick="adminTab('design')">🎨 Design Settings</button>
<button class="direct-admin-link" data-direct="users" onclick="adminTab('users')">👥 User Management</button>
<button class="direct-admin-link" data-direct="guest" onclick="adminTab('guest')">👤 Guest / Customer Access</button>
<button class="direct-admin-link" data-direct="logs" onclick="adminTab('logs')">📋 Activity Log</button>
<button class="direct-admin-link" data-direct="reports" onclick="adminTab('reports')">📈 Reports</button>
<button class="direct-admin-link" data-direct="backup" onclick="adminTab('backup')">🗄 Backup & Restore</button>
<button class="direct-admin-link" data-direct="password" onclick="adminTab('password')">🔑 Change Password</button>
</div>
<label class="dr-sec">Account</label><button id="adbtn" onclick="adminOpen()"> Sign In</button><button id="staffprice" onclick="staffPriceOpen()" hidden> Price Editor</button><button id="stockdash" onclick="adminTab('excelstock')" hidden> Excel Stock Dashboard</button><button id="useradmin" onclick="adminTab('users')" hidden>👥 User Management</button><button id="dashadmin" onclick="adminTab('dashboard')" hidden>📊 Dashboard</button><button id="logadmin" onclick="adminTab('logs')" hidden> Activity Log</button><button id="reportadmin" onclick="adminTab('reports')" hidden> Reports</button><button id="guestadmin" onclick="adminTab('guest')" hidden title="Guest / Customer Access Control" style="display:none"> Guest / Customer Access</button><button id="passadmin" onclick="adminTab('password')" hidden>🔑 Password</button><button id="lo" onclick="logout()" hidden>Sign Out</button></aside>

<div class="ov" id="mo" onclick="if(event.target==this)closeM()"><div class="box" id="mb"></div></div>
<div id="pos"></div>
<div class="ov" id="ao"><div class="box ad" id="ab"></div></div>
<div id="bar"><span id="bc"></span><button data-perm="download" onclick="dlSel()">HD ডাউনলোড</button><button data-perm="copy" onclick="copySel()"> আলাদা আলাদা ছবি কপি</button><button data-perm="link_copy" onclick="copyLinks()">লিংক কপি</button><button onclick="sel.clear();logUiAction('selection_clear','','all');render(1)">✕</button></div>
<div id="toast"></div>
<style>
/* ===== PREMIUM LAYER (XEROX-S1AM) ===== */
:root{--r:20px;--line:color-mix(in srgb,var(--ink) 10%,transparent);--soft:color-mix(in srgb,var(--ink) 5%,var(--card));--sh1:0 1px 2px rgba(24,18,10,.05),0 6px 18px rgba(24,18,10,.06);--sh2:0 2px 6px rgba(24,18,10,.06),0 22px 48px rgba(24,18,10,.16);--ok:#16a34a;--warn:#d97706;--bad:#dc2626}
html{scroll-behavior:smooth}
body{background:radial-gradient(1100px 420px at 50% -140px,color-mix(in srgb,var(--acc) 18%,transparent),transparent 72%),var(--bg);-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
::selection{background:color-mix(in srgb,var(--acc) 40%,transparent)}
/* header */
header{background:linear-gradient(180deg,color-mix(in srgb,var(--head) 100%,#000) 0%,var(--head) 100%);border-bottom:0;padding:26px 14px 18px;box-shadow:0 18px 44px -14px #062617a8,0 6px 18px #0005;overflow:hidden}
header::before{content:"";position:absolute;inset:0;background:radial-gradient(600px 160px at 50% 0,color-mix(in srgb,var(--acc) 22%,transparent),transparent 70%);pointer-events:none}
header::after{content:"";position:absolute;left:0;right:0;bottom:0;height:1px;background:linear-gradient(90deg,transparent,var(--acc),transparent)}
header h1{font-size:clamp(22px,4.4vw,36px);letter-spacing:.01em;position:relative}
#tag{letter-spacing:.42em;font-size:11px;opacity:.85}
#notifyBtn{position:absolute;right:16px;top:14px;border:1px solid #ffffff33;background:#ffffff12;color:#f4d58e;border-radius:12px;padding:8px 11px;font:inherit;cursor:pointer;z-index:3}#notifyCount{display:inline-grid;place-items:center;min-width:20px;height:20px;border-radius:20px;background:#b42318;color:#fff;font-size:11px}
.notification-panel{position:relative;background:#fff!important;color:#202938;padding:24px;border:1px solid #e5e7eb;border-radius:14px;box-shadow:0 8px 28px #17203314;backdrop-filter:none!important;-webkit-backdrop-filter:none!important}.notification-panel h3{margin:0 42px 18px 0}.notification-list{display:grid;gap:10px;max-height:68vh;overflow:auto}.notification-item{display:flex;align-items:center;justify-content:space-between;gap:16px;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;box-shadow:none!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important}.notification-item.unread{border-left:4px solid #b88738;background:#fffcf5}.notification-item.read{color:#697586}.notification-item:focus-visible{outline:3px solid #17865b;outline-offset:2px}.notification-item small{display:block;color:#7b8494;font-size:11px;margin-top:3px}.notification-item p{margin:6px 0 0;white-space:pre-wrap}.read-label{white-space:nowrap;color:#52705d;font-size:13px}.notification-panel .pill{flex:0 0 auto}
#clk{display:inline-block;margin-top:10px;padding:5px 14px;border-radius:999px;background:#ffffff10;border:1px solid #ffffff1c;backdrop-filter:blur(8px);font-size:13px}
#menu{width:44px;height:44px;border-radius:14px;background:#ffffff10;border:1px solid #ffffff22;backdrop-filter:blur(8px);display:grid;place-items:center;padding:0;transition:.2s;z-index:2}
#menu:hover{background:var(--acc);color:#1a1512;transform:scale(1.06)}
.hg{filter:drop-shadow(0 4px 10px #0006)}
/* search */
#q{border-radius:18px;border:1px solid var(--line);padding:15px 20px;font-size:16px;box-shadow:var(--sh1);transition:.2s}
#q:focus{outline:0;border-color:var(--acc);box-shadow:0 0 0 4px color-mix(in srgb,var(--acc) 22%,transparent),var(--sh1)}
#dz{border:1.5px dashed color-mix(in srgb,var(--acc) 70%,transparent);border-radius:18px;padding:14px 18px;background:linear-gradient(180deg,var(--card),var(--soft));transition:.2s;font-size:14px}
#dz:hover{transform:translateY(-2px);box-shadow:var(--sh1);border-style:solid}
.tabs button,.pill{border:1px solid var(--line);border-radius:999px;padding:8px 18px;font-weight:500;transition:.2s;box-shadow:0 1px 0 #0000000a}
.tabs button:hover,.pill:hover{transform:translateY(-1px);border-color:var(--acc);box-shadow:var(--sh1)}
.tabs .on{background:var(--ink);color:var(--acc);border-color:var(--ink);box-shadow:0 6px 16px -6px #0006}
#info{font-size:13px;letter-spacing:.01em}
#prog .b{height:6px;background:var(--line)}#prog i{background:linear-gradient(90deg,var(--acc),color-mix(in srgb,var(--acc) 55%,#fff))}
/* product cards */
#grid,.sg{gap:18px}
.c{border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--sh1);background:var(--card)}
.c:hover{transform:translateY(-5px);box-shadow:var(--sh2);border-color:color-mix(in srgb,var(--acc) 60%,transparent)}
.c .iw{overflow:hidden}.c .iw>img:first-child{transition:transform .5s cubic-bezier(.2,.8,.2,1)}.c:hover .iw>img:first-child{transform:scale(1.045)}
.c .m{padding:12px 14px 14px}
.c b{font-weight:600;font-size:14px;height:2.6em;letter-spacing:.005em}
.c span{color:color-mix(in srgb,var(--ink) 55%,transparent)}
.pr{font-size:18px;font-weight:700;margin-top:6px}.pr s{font-size:12px;font-weight:400}
.copycode{border-radius:10px;border:1px solid var(--line);background:var(--soft);padding:3px 9px;font-size:12px;font-weight:600;transition:.15s}.copycode:hover{background:var(--acc);color:#1a1512;border-color:var(--acc)}
.stockstate{border-radius:999px;padding:4px 10px;font-size:11px;font-weight:700;background:color-mix(in srgb,var(--ok) 12%,var(--card));color:#166534}
.stockstate.low{background:color-mix(in srgb,var(--warn) 14%,var(--card));color:#92400e}.stockstate.out{background:color-mix(in srgb,var(--bad) 12%,var(--card));color:#991b1b}
.photo-count,.c .iw .tg,.c .iw .sc{backdrop-filter:blur(6px);border-radius:999px}
.ck{border-radius:12px!important;background:#fffffff0;backdrop-filter:blur(6px)}
.fav{display:none!important}
/* SIZE CHIPS — পরিষ্কার সাইজ + পরিমাণ */
.st{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
.st .zc{display:inline-flex;align-items:center;gap:6px;padding:3px 4px 3px 10px;border-radius:999px;font:600 12px/1 system-ui,'Hind Siliguri',sans-serif;background:color-mix(in srgb,var(--ok) 9%,var(--card));border:1px solid color-mix(in srgb,var(--ok) 30%,transparent);color:#14532d;white-space:nowrap}
.st .zc .sl{font-weight:800;letter-spacing:.02em;font-size:13px}
.st .zc .sc2{font-style:normal;font-weight:500;font-size:11px;opacity:.6;margin-left:-2px}
.st .zc .q{display:inline-block;min-width:26px;text-align:center;padding:4px 8px;border-radius:999px;background:var(--ok);color:#fff;font:800 12px/1 system-ui,sans-serif;text-decoration:none;font-variant-numeric:tabular-nums}
.st .zc.lo{background:color-mix(in srgb,var(--warn) 11%,var(--card));border-color:color-mix(in srgb,var(--warn) 38%,transparent);color:#7c2d12}.st .zc.lo .q{background:var(--warn)}
.st .zc.no{background:color-mix(in srgb,var(--bad) 9%,var(--card));border-color:color-mix(in srgb,var(--bad) 32%,transparent);color:#991b1b;padding:5px 12px}.st .zc.no .q{background:var(--bad)}
.product-modal-sizes{margin:16px 0;padding:14px;border:1px solid var(--line);border-radius:18px;background:var(--soft)}
.product-modal-sizes::before{content:"সাইজ ও স্টক";display:block;font:700 11px system-ui,'Hind Siliguri';letter-spacing:.12em;text-transform:uppercase;color:color-mix(in srgb,var(--ink) 50%,transparent);margin-bottom:10px}
.product-modal-sizes .st{margin-top:0;gap:10px}
.product-modal-sizes .zc{font-size:15px;padding:6px 6px 6px 14px;gap:9px}
.product-modal-sizes .zc .sl{font-size:16px}.product-modal-sizes .zc .sc2{font-size:13px}.product-modal-sizes .zc .q{font-size:15px;min-width:34px;padding:5px 10px}
.c .st .zc{font-size:11px;padding:2px 3px 2px 9px;gap:5px}.c .st .zc .sl{font-size:12px}.c .st .zc .sc2{font-size:10px}.c .st .zc .q{min-width:22px;padding:3px 6px;font-size:11px}.c .st{gap:5px;margin-top:8px}
.st .zc .sl{display:inline!important;height:auto!important;overflow:visible!important;line-height:1!important}
/* modal */
.ov{background:#0a0705b8;backdrop-filter:blur(8px)}
.box{border-radius:26px;border:1px solid var(--line);box-shadow:var(--sh2);padding:26px}
.product-summary h2{padding-right:48px;margin:0 0 8px;font-size:22px;line-height:1.3}
.product-modal-price{font-size:28px;font-weight:800;margin:6px 0}.product-modal-price s{font-size:15px;font-weight:400;opacity:.5}
.acts{gap:8px}
.acts button,.acts a{background:var(--soft);color:var(--ink);border:1px solid var(--line);border-radius:12px;padding:9px 14px;font-size:13.5px;font-weight:500;transition:.18s}
.acts button:hover,.acts a:hover{background:var(--ink);color:var(--acc);border-color:var(--ink);transform:translateY(-1px)}
.acts .report{background:#fff1f1!important}
#mi{border-radius:20px;box-shadow:var(--sh1)}
.th img{border-radius:12px}.th img:hover{border-color:var(--acc)}
/* সব ✕ বাটন সবসময় দেখা ও কাজ করবে */
.x{display:grid!important;place-items:center;width:40px;height:40px;font-size:16px;line-height:1;background:var(--ink);color:var(--acc);box-shadow:0 4px 14px #0004;transition:.2s;z-index:50;top:12px;right:12px}
.x:hover{transform:rotate(90deg) scale(1.08);background:var(--bad);color:#fff}
.admin-head .x{position:static!important}
.direct-admin-link{display:none!important}.direct-admin-link.direct-visible{display:flex!important;align-items:center;gap:8px;width:100%;text-align:left;margin:3px 0;padding:10px 11px;border:1px solid #d9e9df;border-radius:11px;background:linear-gradient(120deg,#fff,#f1f8f3);color:#214c39;font:600 12px system-ui,'Noto Sans Bengali',sans-serif;cursor:pointer}.direct-admin-link.direct-visible:hover{background:#e6f4eb;border-color:#9ac9ad;transform:translateX(2px)}
.dr-menu-group{margin:9px 0 12px;padding:0 0 8px;border-bottom:1px solid #e8edf0}
.dr-menu-group[hidden]{display:none!important}
.dr-menu-group>.dr-sec{display:block;margin:12px 4px 7px;color:#667085;font:800 10px/1.4 system-ui,sans-serif;letter-spacing:.13em;text-transform:uppercase}
.dr-menu-group .direct-admin-link.direct-visible{font:650 12.5px/1.35 system-ui,sans-serif;min-height:42px}
.direct-admin-link.direct-visible:active{transform:scale(.985)}
@media(max-width:600px){.dr-menu-group .direct-admin-link.direct-visible{padding:10px;font-size:12px}}

/* sidebar */
#dr{width:340px;background:color-mix(in srgb,var(--card) 94%,transparent);backdrop-filter:blur(18px);border-radius:0 28px 28px 0;box-shadow:12px 0 60px -10px #0007;padding:0 0 24px;z-index:9000}
#bd{z-index:8900;background:#0a07058c;backdrop-filter:blur(4px)}
.dr-head{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:18px 18px 14px;position:sticky;top:0;z-index:3;background:linear-gradient(180deg,var(--head),color-mix(in srgb,var(--head) 85%,#000));color:var(--acc);border-radius:0 28px 0 0}
.dr-brand{display:flex;align-items:center;gap:10px}.dr-brand img{height:34px;width:auto}.dr-brand b{display:block;font:600 17px 'Playfair Display',serif;color:var(--acc);line-height:1.1}.dr-brand small{font-size:10px;letter-spacing:.2em;color:#fff9}
#dr>select,#dr>button{width:calc(100% - 32px)!important}#dr .dr-x{width:38px!important;height:38px;border-radius:12px!important;margin:0!important;padding:0!important;display:grid;place-items:center;background:#ffffff14!important;border:1px solid #ffffff2a!important;color:#fff!important;font-size:16px;transition:.2s}
#dr .dr-x:hover{background:var(--bad)!important;transform:rotate(90deg)}
#dr>*:not(.dr-head){margin-left:16px;margin-right:16px}
#dr>#sync{margin-top:14px;border-radius:14px}
#dr label{font:700 11px system-ui,'Hind Siliguri';letter-spacing:.1em;text-transform:uppercase;color:color-mix(in srgb,var(--ink) 48%,transparent);margin:18px 4px 6px}
#dr select{border-radius:12px;background:var(--soft);border:1px solid var(--line);padding:11px 12px}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){position:relative;display:flex;align-items:center;gap:12px;text-align:left;padding:11px 14px 11px 12px;border-radius:14px;border:1px solid transparent;background:transparent;font-weight:500;font-size:14.5px;transition:.18s}
#dr button[hidden]{display:none!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover{background:color-mix(in srgb,var(--acc) 16%,var(--card));border-color:color-mix(in srgb,var(--acc) 40%,transparent);transform:translateX(3px)}
#dr button::before{content:"";flex:0 0 auto;width:34px;height:34px;border-radius:11px;background:var(--soft);border:1px solid var(--line);
  -webkit-mask:none;mask:none}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle)::after{content:"";position:absolute;left:12px;width:34px;height:34px;background:var(--acc);filter:brightness(.78);-webkit-mask-repeat:no-repeat;mask-repeat:no-repeat;-webkit-mask-position:center;mask-position:center;-webkit-mask-size:19px;mask-size:19px;-webkit-mask-image:var(--ic);mask-image:var(--ic)}
#dr #dsz button::before,#dr .dr-x::before,#dr .personal-toggle::before{display:none}
#dr #dsz button{display:inline-block;border-radius:999px;border:1px solid var(--line);background:var(--soft);font-size:13px}
#dr #dsz .on{background:var(--ink);color:var(--acc)}
#dr #lo{margin-top:8px;color:var(--bad);--ic:var(--i-out)}#dr #lo:hover{background:color-mix(in srgb,var(--bad) 12%,var(--card))!important;border-color:color-mix(in srgb,var(--bad) 35%,transparent)!important}#dr #lo::after{background:var(--bad)!important;filter:none!important}
.personal-box{border-radius:16px;border:1px solid var(--line)}
:root{
--i-dot:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round'><circle cx='12' cy='12' r='4'/></svg>");
--i-user:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='8' r='4'/><path d='M4 21a8 8 0 0 1 16 0'/></svg>");
--i-gear:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='3'/><path d='M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z'/></svg>");
--i-box:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M21 8 12 3 3 8v8l9 5 9-5z'/><path d='m3 8 9 5 9-5M12 13v8'/></svg>");
--i-users:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='9' cy='8' r='3.5'/><path d='M2 21a7 7 0 0 1 14 0M17 4.5a3.5 3.5 0 0 1 0 7M22 21a7 7 0 0 0-4-6.3'/></svg>");
--i-chart:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 3v18h18'/><rect x='7' y='12' width='3' height='6' rx='1'/><rect x='12' y='7' width='3' height='11' rx='1'/><rect x='17' y='10' width='3' height='8' rx='1'/></svg>");
--i-list:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='5' y='3' width='14' height='18' rx='2'/><path d='M9 8h6M9 12h6M9 16h4'/></svg>");
--i-flag:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M5 21V4M5 4h11l-2 4 2 4H5'/></svg>");
--i-key:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='8' cy='15' r='4'/><path d='m11 12 9-9M16 7l3 3M14 9l2 2'/></svg>");
--i-out:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9'/></svg>");
--i-tag:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M20.6 13.4 13.4 20.6a2 2 0 0 1-2.8 0L3 13V3h10l7.6 7.6a2 2 0 0 1 0 2.8z'/><circle cx='7.5' cy='7.5' r='1.3'/></svg>");
--i-check:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='3' width='18' height='18' rx='4'/><path d='m8 12 3 3 5-6'/></svg>");
--i-dl:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3v12M7 10l5 5 5-5M4 21h16'/></svg>");
--i-copy:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='9' y='9' width='12' height='12' rx='2'/><path d='M5 15V5a2 2 0 0 1 2-2h8'/></svg>");
--i-zap:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M13 2 4 14h7l-1 8 9-12h-7z'/></svg>");
--i-shield:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3 4 6v6c0 5 3.4 8.3 8 9 4.6-.7 8-4 8-9V6z'/><path d='m9 12 2 2 4-4'/></svg>")}
#dr button{--ic:var(--i-dot)}#dr #adbtn{--ic:var(--i-user)}#dr #staffprice{--ic:var(--i-tag)}#dr #stockdash{--ic:var(--i-box)}#dr #useradmin{--ic:var(--i-users)}#dr #dashadmin{--ic:var(--i-chart)}#dr #logadmin{--ic:var(--i-list)}#dr #reportadmin{--ic:var(--i-flag)}#dr #guestadmin{--ic:var(--i-shield)}#dr #passadmin{--ic:var(--i-key)}
#dr button[onclick*="selAll"]{--ic:var(--i-check)}#dr button[onclick*="dlSel"]{--ic:var(--i-dl)}#dr button[onclick*="copySel"]{--ic:var(--i-copy)}#dr #quickbtn{--ic:var(--i-zap)}
/* ওয়াটারমার্ক — সব পেজে, সব সময়, ডান-নিচের কোণে */
#brand-watermark{display:block!important;visibility:visible!important;opacity:1!important;position:fixed!important;right:12px!important;bottom:12px!important;top:auto!important;left:auto!important;z-index:2147483000!important;pointer-events:none;background:color-mix(in srgb,var(--card) 82%,transparent)!important;color:color-mix(in srgb,var(--ink) 78%,transparent)!important;border:1px solid var(--line)!important;backdrop-filter:blur(10px);border-radius:999px!important;padding:6px 13px!important;font:700 11px/1.2 'Hind Siliguri',system-ui,sans-serif!important;letter-spacing:.06em;box-shadow:0 6px 20px -6px #0005!important;white-space:nowrap}
#brand-watermark::before{content:"";display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--acc);margin-right:7px;vertical-align:middle;box-shadow:0 0 0 3px color-mix(in srgb,var(--acc) 25%,transparent)}
@media(max-width:700px){#brand-watermark{right:8px!important;bottom:8px!important;font-size:9.5px!important;padding:5px 10px!important;max-width:none!important;overflow:visible!important}}
main{padding-bottom:110px}footer{padding-bottom:110px}#bar{bottom:54px;z-index:9500;border-radius:22px;box-shadow:var(--sh2)}#toast{z-index:2147482000}
.admin-content{padding-bottom:70px}
/* admin */
.admin-head{padding:16px 22px;background:linear-gradient(135deg,var(--head),color-mix(in srgb,var(--head) 80%,#000))}
.admin-nav{background:var(--soft);border-right:1px solid var(--line)}
.admin-nav button{border-radius:14px;padding:11px 14px;font-weight:500;transition:.18s}.admin-nav button:hover{background:color-mix(in srgb,var(--acc) 18%,var(--card));transform:translateX(3px)}
.ad input,.ad select{border-radius:12px;border:1px solid var(--line);padding:10px 12px}.ad input:focus,.ad select:focus{outline:0;border-color:var(--acc);box-shadow:0 0 0 3px color-mix(in srgb,var(--acc) 22%,transparent)}
.ad table{border-radius:14px;overflow:hidden}.ad td{border-bottom:1px solid var(--line)}
.stat{border-radius:16px;box-shadow:var(--sh1)}
/* কমেন্ট */
.stock-comments{border-radius:20px!important;border:1px solid var(--line)!important;background:var(--soft)!important;padding:18px!important}
.stock-comments h4{font-size:16px;color:var(--ink)!important;display:flex;align-items:center;gap:8px}
.stock-comment-form{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;align-items:end;margin:12px 0}
.stock-comment-form label{display:flex;flex-direction:column;gap:5px;font-size:12px;font-weight:600;color:color-mix(in srgb,var(--ink) 60%,transparent)}
.stock-comment-form input{padding:10px 12px;border-radius:12px;border:1px solid var(--line);font:inherit;background:var(--card);color:var(--ink)}
.stock-comment-form .pill{background:var(--ink);color:var(--acc);border:0;padding:11px 16px}
.stock-comment-list{display:grid;gap:10px}
.stock-comment-row{position:relative;padding:12px 14px;border-radius:14px;background:var(--card);border:1px solid var(--line);border-left:4px solid var(--acc)}
.stock-comment-row b{display:block;font-size:14px}.stock-comment-row small{color:color-mix(in srgb,var(--ink) 50%,transparent)}.stock-comment-row p{margin:6px 0 0;font-size:14px}
.stock-comment-row .pill{position:absolute;right:10px;top:10px;padding:5px 12px;font-size:12px;color:var(--bad);border-color:color-mix(in srgb,var(--bad) 35%,transparent)}
.stock-comment-row .pill:hover{background:var(--bad);color:#fff}
#cmtSummary{margin:0 0 22px;padding:18px;border-radius:22px;border:1px solid var(--line);background:linear-gradient(145deg,var(--card),var(--soft));box-shadow:var(--sh1)}
#cmtSummary h3{margin:0 0 4px}.cs-sub{margin:0 0 14px;font-size:13px;color:color-mix(in srgb,var(--ink) 55%,transparent)}
.cs-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:14px}
.cs-stats div{padding:12px 14px;border-radius:16px;background:var(--card);border:1px solid var(--line)}.cs-stats b{display:block;font-size:26px;line-height:1.1}.cs-stats>div>span{font-size:12px;color:color-mix(in srgb,var(--ink) 55%,transparent)}
.cs-chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
.cs-chip{display:inline-flex;align-items:center;gap:8px;padding:6px 8px 6px 12px;border-radius:999px;background:color-mix(in srgb,var(--acc) 16%,var(--card));border:1px solid color-mix(in srgb,var(--acc) 45%,transparent);cursor:pointer;font:600 13px system-ui;color:var(--ink);transition:.15s}
.cs-chip:hover{transform:translateY(-2px);box-shadow:var(--sh1)}.cs-chip i{font-style:normal;background:var(--ink);color:var(--acc);border-radius:999px;padding:2px 8px;font-size:12px}
.cs-list{display:grid;gap:8px;max-height:340px;overflow:auto;padding-right:4px}
#dr:not(.on){box-shadow:none!important}#dr label:has(+ [style*='display: none']){display:none}
@media(max-width:700px){#dr{width:88vw}.box{padding:18px;border-radius:22px}.c:hover{transform:none}}

/* ===== মোবাইল/ইউজার-ফ্রেন্ডলি ফিক্স ===== */
#grid{grid-template-columns:repeat(auto-fill,minmax(205px,1fr))}
.acts{display:grid;grid-template-columns:repeat(auto-fit,minmax(128px,1fr));gap:8px}
.acts button,.acts a{display:inline-flex;align-items:center;justify-content:center;gap:6px;text-align:center;min-height:40px;line-height:1.2;white-space:normal;box-sizing:border-box}
.acts .report{color:#b91c1c!important;border-color:#f3b9b9!important;background:#fff1f1!important}
.acts .report:hover{background:#dc2626!important;color:#fff!important}
/* সিমিলার প্রোডাক্ট: ছোট কার্ডের পাশে-স্ক্রল স্ট্রিপ */
.similar-section{margin-top:22px;padding-top:16px;border-top:1px solid var(--line)}
.similar-section h3{margin:0 0 12px;font-size:17px}
.similar-section .sg{display:flex!important;gap:12px;overflow-x:auto;scroll-snap-type:x mandatory;padding:4px 2px 14px;-webkit-overflow-scrolling:touch}
.similar-section .sg .c{flex:0 0 150px;width:150px;scroll-snap-align:start;border-radius:16px}
.similar-section .sg .c img{aspect-ratio:3/4}
.similar-section .sg .c .m{padding:8px 10px 10px}
.similar-section .sg .c b{font-size:12.5px;height:2.6em}
.similar-section .sg .c .pr{font-size:14px;margin-top:3px}
.similar-section .sg .c .st,.similar-section .sg .c .stockstate,.similar-section .sg .c .m>span,.similar-section .sg .c .ck,.similar-section .sg .c .tg,.similar-section .sg .c .photo-count{display:none!important}
.similar-section .sg .c:hover{transform:none}
@media(max-width:700px){
 #grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
 main{padding:14px 10px 110px}
 .c{border-radius:16px}.c .m{padding:9px 10px 11px}.c b{font-size:12.5px}.pr{font-size:15px}
 .c .copycode{font-size:11px;padding:2px 7px}
 .c .st .zc{font-size:10px}.c .st .zc .q{min-width:19px;padding:2px 5px;font-size:10px}
 .c .stockstate{font-size:10px;padding:3px 8px}
 .ov{padding:0;align-items:flex-end}
 .ov>.box:not(.ad){max-height:94vh;border-radius:24px 24px 0 0;width:100%;padding:16px 14px 20px}
 .product-detail-grid{gap:12px}
 #mi{max-height:42vh;width:100%;object-fit:contain}
 .th img{width:52px;height:66px}
 .product-summary h2{font-size:18px}.product-modal-price{font-size:24px}
 .acts{grid-template-columns:1fr 1fr}
 .acts button,.acts a{font-size:12.5px;padding:8px 8px}
 .stock-comments{padding:12px!important}
 header h1{font-size:20px}.hg{height:34px}.hg.b{height:26px}
}

.cmt-note{display:flex;gap:8px;align-items:flex-start;margin-top:10px;padding:8px 10px;border-radius:12px;background:color-mix(in srgb,var(--warn) 12%,var(--card));border:1px solid color-mix(in srgb,var(--warn) 35%,transparent);border-left:4px solid var(--warn);font-size:12px;line-height:1.35}
.cmt-note .cmt-ic{font-size:14px;line-height:1.2}.cmt-note b{display:block;font-weight:600;color:#7c2d12;word-break:break-word}.cmt-note small{display:block;margin-top:2px;color:color-mix(in srgb,var(--ink) 50%,transparent);font-size:11px}
.similar-section .sg .c{flex:0 0 118px;width:118px;border-radius:14px}
.similar-section .sg .c b{font-size:11.5px}.similar-section .sg .c .pr{font-size:12.5px}.similar-section .sg .c .cmt-note,.similar-section .sg .c .copycode,.similar-section .sg .c .lg{display:none!important}.similar-section .sg .c .sc{display:block!important;top:8px!important;left:8px!important;right:auto!important;bottom:auto!important;z-index:6;background:#17324d!important;color:#fff!important;border:1px solid #fff!important;font-size:11px!important;font-weight:800!important;box-shadow:0 2px 8px #0007!important}
.similar-section h3{font-size:15px}

.th{display:flex;gap:8px;overflow-x:auto;padding:6px 2px 10px;scroll-snap-type:x proximity;-webkit-overflow-scrolling:touch}
.th .thumb-wrap{flex:0 0 auto;scroll-snap-align:start}
.th .thumb-wrap img{width:68px;height:86px;object-fit:cover;border-radius:12px;border:2px solid transparent;cursor:pointer;transition:.15s}
.th .thumb-wrap img.cur,.th .thumb-wrap img:hover{border-color:var(--acc);transform:translateY(-2px)}
#modalCmtTop{display:grid;gap:8px;margin:10px 0}
.newtag{position:absolute;left:10px;bottom:10px;z-index:3;padding:4px 11px;border-radius:999px;background:linear-gradient(135deg,#f59e0b,#ef4444);color:#fff;font:800 11px/1 'Hind Siliguri',system-ui;letter-spacing:.04em;box-shadow:0 4px 12px #ef444466}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}

/* ===== FULL PREMIUM GLASSMORPHISM THEME ===== */
:root{--glass-bg:rgba(255,255,255,.58);--glass-bg-strong:rgba(255,255,255,.74);--glass-border:rgba(255,255,255,.72);--glass-shadow:0 12px 36px rgba(15,23,42,.10),inset 0 1px 0 rgba(255,255,255,.8)}
body{background:radial-gradient(ellipse at 12% 0%,rgba(34,197,94,.13),transparent 34%),radial-gradient(ellipse at 90% 16%,rgba(56,189,248,.12),transparent 32%),linear-gradient(135deg,#eef7f2 0%,#f4f7fb 48%,#edf4f8 100%)!important;background-attachment:fixed!important}
/* Keep the original site header styling/controls intact for reliable navigation. */
header{background:var(--head)!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important;box-shadow:0 4px 14px rgba(0,0,0,.12)!important;border-radius:0!important;border:0!important}
main,footer{position:relative}
.box,.c,.personal-box,.stock-comments,#cmtSummary,.cs-stats>div,.stock-comment-row,.similar-section,.admin-head,.admin-nav,.stat,.admin-content,.ad,.ov>.box,#dr,#bar,#siteNoticeCard,.modal,.panel,.card,.section-card,.product-summary,.product-detail-grid,.stock-comment-form,.th .thumb-wrap,.search-box,.filter-box,.toolbar,.pagination,.empty-state{background:var(--glass-bg)!important;border:1px solid var(--glass-border)!important;box-shadow:var(--glass-shadow)!important;backdrop-filter:blur(18px) saturate(145%)!important;-webkit-backdrop-filter:blur(18px) saturate(145%)!important}
.c,.box,.personal-box,.stock-comments,#cmtSummary,.stat,.ad,.section-card,.filter-box,.toolbar{border-radius:22px!important;overflow:hidden}
.c{background:rgba(255,255,255,.58)!important;transition:transform .28s ease,box-shadow .28s ease,border-color .28s ease!important}
.c:hover{transform:translateY(-5px)!important;box-shadow:0 20px 42px rgba(15,23,42,.16),inset 0 1px 0 rgba(255,255,255,.9)!important;border-color:rgba(255,255,255,.95)!important}
.c .m,.product-summary,.product-detail-grid{background:rgba(255,255,255,.34)!important;backdrop-filter:blur(12px)!important}
button,.pill,.acts a,.acts button,#menu,.admin-nav button,.personal-toggle{border-radius:14px!important;transition:transform .2s ease,box-shadow .2s ease,background .2s ease!important}
button:hover,.pill:hover,.acts a:hover,.acts button:hover{box-shadow:0 8px 22px rgba(16,185,129,.18)!important}
input,select,textarea,.stock-comment-form input,.ad input,.ad select{background:rgba(255,255,255,.64)!important;border:1px solid rgba(255,255,255,.9)!important;border-radius:14px!important;box-shadow:inset 0 1px 3px rgba(15,23,42,.04)!important;backdrop-filter:blur(10px)!important}
input:focus,select:focus,textarea:focus,.ad input:focus,.ad select:focus{border-color:rgba(16,185,129,.7)!important;box-shadow:0 0 0 4px rgba(16,185,129,.13)!important;outline:none!important}
#dr{background:rgba(14,30,27,.78)!important;border:1px solid rgba(255,255,255,.18)!important;backdrop-filter:blur(28px) saturate(160%)!important;-webkit-backdrop-filter:blur(28px) saturate(160%)!important;box-shadow:18px 0 60px rgba(0,0,0,.24),inset 0 1px 0 rgba(255,255,255,.14)!important}
.admin-head{background:linear-gradient(135deg,rgba(8,50,36,.86),rgba(8,110,72,.68))!important;color:#fff!important;border-radius:0 0 22px 22px!important}
.admin-nav{background:rgba(255,255,255,.42)!important}
.admin-nav button:hover,.personal-toggle:hover{background:rgba(16,185,129,.14)!important}
#bar{background:rgba(255,255,255,.72)!important;backdrop-filter:blur(24px) saturate(150%)!important;-webkit-backdrop-filter:blur(24px) saturate(150%)!important;border:1px solid rgba(255,255,255,.86)!important}
#siteNoticeShade,.ov{backdrop-filter:blur(12px) saturate(120%)!important;-webkit-backdrop-filter:blur(12px) saturate(120%)!important}
#siteNoticeCard{background:rgba(255,255,255,.82)!important;border:1px solid rgba(255,255,255,.95)!important;border-radius:26px!important;box-shadow:0 28px 80px rgba(0,0,0,.24),inset 0 1px 0 #fff!important;backdrop-filter:blur(28px) saturate(160%)!important;color:var(--ink)!important}
.photo-count,.c .iw .tg,.c .iw .sc,.ck,#clk,#brand-watermark,.newtag,.cmt-note,.cs-chip{backdrop-filter:blur(14px) saturate(150%)!important;-webkit-backdrop-filter:blur(14px) saturate(150%)!important;box-shadow:0 5px 18px rgba(15,23,42,.10)!important}
.stock-comment-row{background:rgba(255,255,255,.56)!important}
.cs-stats>div{background:rgba(255,255,255,.54)!important}
hr,.similar-section{border-color:rgba(255,255,255,.7)!important}
/* Soft animated glass sheen; respects reduced-motion preference above */
header:after{pointer-events:none}
@keyframes glassSheen{0%,65%,100%{background-position:140% 0}30%{background-position:-40% 0}}
@media(max-width:700px){body{background-attachment:scroll!important}.box,.c,.personal-box,.stock-comments,#cmtSummary,.stat,.ad,.section-card{border-radius:18px!important}.c:hover{transform:none!important}header{backdrop-filter:blur(22px) saturate(150%)!important;-webkit-backdrop-filter:blur(22px) saturate(150%)!important}}

.order-summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px;margin:16px 0}.order-summary-grid .visual-card{min-width:0;background:linear-gradient(145deg,#ffffff,#eefaf4);border:1px solid #b9dfc9;border-radius:14px;padding:14px;box-shadow:0 8px 22px rgba(18,75,49,.08)}.order-summary-grid h3{margin:0 0 10px;color:#124b35}.order-summary-grid .excel-table-wrap{max-height:360px;overflow:auto}.order-grid{min-width:1550px!important}.order-bulk-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:12px 0}.order-bulk-actions input{width:auto!important;min-width:18px}.order-bulk-actions small{color:#9a3412;font-weight:600}.order-grid td{color:#172a22!important;background:rgba(255,255,255,.9)!important}.order-grid thead th{position:sticky;top:0;background:#163f32!important;color:#fff!important;z-index:4}.order-grid tbody tr:nth-child(even) td{background:#eef8f1!important}.order-thumb{width:54px;height:54px;object-fit:contain;background:#fff;border:1px solid #b9d8c5;border-radius:8px}.order-address{white-space:normal!important;min-width:220px;max-width:320px}.order-uploader{display:block;color:#60756a;font-size:10px;margin-top:4px}.order-history .admin-note{background:#eaf7ff;color:#123e59;border-color:#9bd8f3;font-weight:600}

/* ===== SCROLL + PERFORMANCE FIX ===== */
html{scroll-behavior:auto!important;overflow-y:auto!important;overscroll-behavior-y:auto!important}
body{overflow-y:auto!important;overflow-x:hidden!important;min-height:100%;height:auto!important;background-attachment:scroll!important;scroll-behavior:auto!important}
/* Restore original header behavior: it scrolls away instead of staying pinned. */
header{position:relative!important;top:auto!important;z-index:10!important;animation:none!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important;isolation:auto!important}
header::before,header::after,header .hl,header h1,header .hg{animation:none!important}
/* Heavy blur filters on every card can make scrolling sluggish on phones. Keep the glass look, reduce GPU work. */
.box,.c,.personal-box,.stock-comments,#cmtSummary,.cs-stats>div,.stock-comment-row,.similar-section,.admin-head,.admin-nav,.stat,.admin-content,.ad,.ov>.box,#dr,#bar,#siteNoticeCard,.modal,.panel,.card,.section-card,.product-summary,.product-detail-grid,.stock-comment-form,.th .thumb-wrap,.search-box,.filter-box,.toolbar,.pagination,.empty-state{backdrop-filter:none!important;-webkit-backdrop-filter:none!important}
.c .m,.product-summary,.product-detail-grid,input,select,textarea,.stock-comment-form input,.ad input,.ad select,.photo-count,.c .iw .tg,.c .iw .sc,.ck,#clk,#brand-watermark,.newtag,.cmt-note,.cs-chip,#bar,#siteNoticeShade,.ov,#siteNoticeCard,#dr{backdrop-filter:none!important;-webkit-backdrop-filter:none!important}
body,main,#app,#root{touch-action:pan-y;overscroll-behavior-y:auto}
.admin-full .admin-body{min-height:0;overflow:hidden!important}
.admin-full .admin-content,.admin-full .admin-nav{overscroll-behavior:contain;-webkit-overflow-scrolling:touch;touch-action:pan-y}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important;scroll-behavior:auto!important}}

/* Readable controls and subtle premium live header animation */
input[type="range"]{appearance:none;-webkit-appearance:none;width:100%;height:12px;border-radius:999px;background:linear-gradient(90deg,#16a36a,#56c9e5 55%,#dce8e2);border:1px solid #b5d7c6;box-shadow:inset 0 1px 3px #153d2b22;cursor:pointer;outline:none}
input[type="range"]::-webkit-slider-thumb{-webkit-appearance:none;appearance:none;width:22px;height:22px;border-radius:50%;background:#fff;border:4px solid #11845b;box-shadow:0 2px 9px #0b513a55}
input[type="range"]::-moz-range-thumb{width:15px;height:15px;border-radius:50%;background:#fff;border:4px solid #11845b;box-shadow:0 2px 9px #0b513a55}
input[type="range"]:focus-visible{outline:3px solid #58c99b;outline-offset:4px}
/* keep header in its normal document flow; animate only a restrained light sweep */
header,.site-header,#mainHeader,.main-header,.topbar{position:relative;isolation:isolate}
header::after,.site-header::after,#mainHeader::after,.main-header::after,.topbar::after{content:"";position:absolute;inset:0;pointer-events:none;z-index:0;background:linear-gradient(110deg,transparent 0%,rgba(255,255,255,0) 35%,rgba(255,255,255,.15) 48%,rgba(255,255,255,0) 60%,transparent 100%);background-size:220% 100%;animation:headerLightSweep 8s ease-in-out infinite}
header>* ,.site-header>* ,#mainHeader>* ,.main-header>* ,.topbar>*{position:relative;z-index:1}
@keyframes headerLightSweep{0%,35%{background-position:120% 0}70%,100%{background-position:-120% 0}}
@media(prefers-reduced-motion:reduce){header::after,.site-header::after,#mainHeader::after,.main-header::after,.topbar::after{animation:none}}

/* FINAL UI RESTORE: clean white theme, clear slide controls, product selection and scrollable comments */
:root{color-scheme:light!important;--bg:#f6f8fb!important;--card:#fff!important;--ink:#202938!important;--muted:#667085!important;--line:#e5eaf1!important}
html,body,#app,main,.admin-full,.admin-body,.admin-content{color:#202938!important}
body{background:#f6f8fb!important}
#dr{background:#fff!important;color:#202938!important;border-right:1px solid #e5eaf1!important;box-shadow:8px 0 32px rgba(15,23,42,.12)!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important;border-radius:0 18px 18px 0!important}
#dr .dr-head,#dr .dr-head *{background-color:#fff!important;color:#202938!important}
#dr a,#dr button,#dr .nav-item,#dr label{color:#344054!important}
#bd{background:rgba(15,23,42,.32)!important}
.admin-section,.admin-card,.visual-card,.excel-stats>div,.order-summary-grid>div,.comment-modal>div{background:#fff!important;color:#202938!important;border-color:#e5eaf1!important}
.admin-section h3,.visual-card h3,.excel-grid th{color:#263448!important}
.admin-full .admin-content{background:#f6f8fb!important;overflow-y:auto!important;overscroll-behavior-y:contain!important;-webkit-overflow-scrolling:touch!important}
.admin-full .admin-body{overflow:visible!important;min-height:0!important}
.admin-full .admin-content,.admin-full .admin-nav{touch-action:pan-y!important}
input,select,textarea{color:#202938}
input[type="range"]{-webkit-appearance:none;appearance:none;height:8px;border-radius:999px;background:linear-gradient(90deg,#16a36a,#8bd5b4)!important;border:1px solid #c9e7d8!important;cursor:pointer}
input[type="range"]::-webkit-slider-thumb{-webkit-appearance:none;appearance:none;width:20px;height:20px;border-radius:50%;background:#fff!important;border:3px solid #16865c!important;box-shadow:0 2px 7px rgba(16,24,40,.2)!important}
input[type="range"]::-moz-range-thumb{width:16px;height:16px;border-radius:50%;background:#fff!important;border:3px solid #16865c!important;box-shadow:0 2px 7px rgba(16,24,40,.2)!important}
input[type="checkbox"],input[type="radio"]{accent-color:#16865c}
.code-cell{display:inline-block;border-radius:7px!important;padding:5px 8px!important;transition:background .2s,border-color .2s,box-shadow .2s,transform .2s!important}
.code-cell:hover{background:#eff6ff!important;box-shadow:0 0 0 2px #bfdbfe!important;transform:translateY(-1px)}
.code-cell.is-selected{background:#dbeafe!important;color:#1d4ed8!important;box-shadow:0 0 0 3px #60a5fa,0 0 18px #60a5fa55!important;animation:orderCodePulse .75s ease-out}
.order-grid tr.order-row-selected td{background:#f0f7ff!important;border-top:1px solid #93c5fd!important;border-bottom:1px solid #93c5fd!important}
.order-grid tr.order-row-selected td:first-child{box-shadow:inset 4px 0 #3b82f6}
@keyframes orderCodePulse{0%{transform:scale(1);box-shadow:0 0 0 0 #60a5fa88}45%{transform:scale(1.08);box-shadow:0 0 0 5px #60a5fa55}100%{transform:scale(1);box-shadow:0 0 0 3px #60a5fa}}
.comment-modal{align-items:flex-start!important;overflow-y:auto!important;overscroll-behavior:contain!important;touch-action:pan-y!important}
.comment-modal>div{margin:auto 0!important;max-height:calc(100dvh - 40px)!important;overflow-y:auto!important;touch-action:pan-y!important}
@media(max-width:700px){#dr{width:88vw!important;border-radius:0 14px 14px 0!important}.comment-modal{padding:10px!important}.comment-modal>div{max-height:calc(100dvh - 20px)!important}}
@media(prefers-reduced-motion:reduce){.code-cell,.order-grid tr.order-row-selected td{animation:none!important;transition:none!important}}

/* LABABEL-inspired refresh: white + rich green, with warm gold and subtle mint accents */
:root{color-scheme:light!important;--acc:#16845b!important;--bg:#f2f8f4!important;--ink:#173b2c!important;--head:#123b2b!important;--card:#ffffff!important;--line:#d9e9df!important;--muted:#60786a!important}
body{background:radial-gradient(ellipse at 12% 0%,#e2f5e9 0,transparent 38%),linear-gradient(145deg,#f6faf7,#edf6f0)!important;color:#183b2c!important}
#dr{background:linear-gradient(180deg,#ffffff 0%,#f1faf4 100%)!important;color:#183b2c!important;border-right:1px solid #cfe5d7!important;box-shadow:8px 0 34px rgba(18,82,51,.14)!important;backdrop-filter:blur(14px)!important;-webkit-backdrop-filter:blur(14px)!important}
#dr .dr-head{background:linear-gradient(135deg,#e5f7ec,#fff)!important;border-bottom:1px solid #d5eadc!important}
#dr .dr-brand b{color:#12613f!important}#dr .dr-brand small{color:#688675!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){border-radius:12px!important;transition:background .18s,border-color .18s,transform .18s!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover{background:linear-gradient(100deg,#e3f7eb,#f4fbf6)!important;border-color:#b8e0c7!important;color:#125c3d!important;transform:translateX(2px)}
#dr button.on,#dr .on{background:linear-gradient(110deg,#13784f,#239866)!important;color:#fff!important;border-color:#13784f!important;box-shadow:0 5px 14px #16845b30!important}
#dr button::before{background:#e9f7ee!important;border-color:#cbe9d6!important}#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle)::after{background:#16845b!important;filter:none!important}
#dr #dsz .on{background:linear-gradient(110deg,#13784f,#239866)!important;color:#fff!important}
/* Liquid-inspired pill navigation, scoped to the slide-out sidebar */
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){isolation:isolate;overflow:hidden;border:1px solid #d7e9df!important;border-radius:17px!important;background:linear-gradient(145deg,#fff 5%,#f3fbf6 48%,#e8f7ee 100%)!important;box-shadow:inset 0 1px 0 #fff,0 5px 13px #174a3010!important;transition:transform .22s ease,box-shadow .22s ease,border-color .22s ease,background .22s ease!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover{transform:translateX(3px) translateY(-1px)!important;border-color:#9bd8b4!important;background:linear-gradient(145deg,#fff 0%,#e8f8ef 55%,#d4f1e1 100%)!important;box-shadow:inset 0 1px 0 #fff,0 8px 19px #13784f20!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle).on,#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):is(.active){color:#fff!important;border-color:#14794f!important;background:linear-gradient(120deg,#116743,#16885a 58%,#28a36d)!important;box-shadow:inset 0 1px 0 #ffffff35,0 8px 22px #13784f38!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle).on::before,#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):is(.active)::before{background:#ffffff20!important;border-color:#ffffff45!important}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle).on::after,#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):is(.active)::after{background:#fff!important}
@media(prefers-reduced-motion:reduce){#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){transition:none!important}}
.admin-full .admin-head{background:linear-gradient(115deg,#103d2b,#16784f 72%,#249767)!important;border-bottom:2px solid #e8c56e!important}
.admin-full .admin-head h2{color:#f4d889!important}.admin-full .admin-head p{color:#e0f4e8!important}
.admin-full .admin-nav{background:linear-gradient(180deg,#f4fbf6,#eaf6ef)!important;border-right:1px solid #d1e7d9!important}
.admin-full .admin-nav{padding:14px 12px!important}
.admin-full .admin-nav button{display:flex!important;align-items:center!important;gap:11px!important;min-height:50px!important;border:1px solid #d7e9df!important;border-radius:16px!important;margin:0 0 7px!important;padding:7px 11px!important;color:#315846!important;background:linear-gradient(145deg,#fff,#f0f9f4)!important;box-shadow:0 4px 12px #174a3010!important;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease!important}
.admin-full .admin-nav button::before{display:none!important}.admin-full .admin-nav button{gap:0!important;padding-left:14px!important}
.admin-full .admin-nav button::before{content:'◇';display:grid;place-items:center;flex:0 0 34px;width:34px;height:34px;border:1px solid #cbe9d6;border-radius:11px;background:#e9f7ee;color:#16845b;font-size:17px;font-weight:700}
.admin-full .admin-nav button[onclick*="dashboard"]::before{content:'▥'}.admin-full .admin-nav button[onclick*="orders"]::before{content:'▤'}.admin-full .admin-nav button[onclick*="returns"]::before{content:'↩'}.admin-full .admin-nav button[onclick*="couriers"]::before{content:'↗'}.admin-full .admin-nav button[onclick*="users"]::before{content:'♙'}.admin-full .admin-nav button[onclick*="backup"]::before{content:'⬡'}.admin-full .admin-nav button[onclick*="notice"]::before{content:'⚑'}.admin-full .admin-nav button[onclick*="design"]::before{content:'✦'}.admin-full .admin-nav button[onclick*="excelstock"]::before{content:'⬢'}.admin-full .admin-nav button[onclick*="reports"]::before{content:'▧'}.admin-full .admin-nav button[onclick*="password"]::before{content:'⌑'}
.admin-full .admin-nav button:hover{background:linear-gradient(145deg,#fff,#e5f6eb)!important;border-color:#a9d8ba!important;transform:translateY(-1px)!important;box-shadow:0 7px 16px #16845b1b!important}
.admin-full .admin-nav button.on{background:linear-gradient(110deg,#14784e,#239766)!important;color:#fff!important;border-color:#13784f!important;box-shadow:0 5px 14px #16845b25!important}
.admin-full .admin-nav button.on::before{color:#fff;background:#ffffff20;border-color:#ffffff40}
.acts button,.acts a{box-shadow:0 4px 12px #174a3010!important;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease!important}.acts button:hover,.acts a:hover{transform:translateY(-2px)!important;box-shadow:0 9px 20px #174a3022!important}
.admin-full .admin-content{background:linear-gradient(135deg,#f6faf7,#edf6f0)!important;overflow-y:auto!important;overscroll-behavior-y:contain!important;-webkit-overflow-scrolling:touch!important}
.admin-section,.admin-card,.visual-card,.chart-card,.userbox,.excel-tip,.upload-card,.order-summary-grid .visual-card{border-color:#d4e9dc!important;box-shadow:0 7px 22px rgba(22,92,57,.07)!important}
.excel-stats>div{background:linear-gradient(145deg,#fff,#eef9f2)!important;border-color:#cfe7d8!important;box-shadow:0 6px 18px #155e3b0c!important}
.excel-stats b,.welcome h2,.order-summary-grid h3{color:#12613f!important}
.excel-grid thead th{background:linear-gradient(180deg,#e4f5e9,#d8efdf)!important;color:#205c40!important;border-bottom-color:#c1e0cc!important}
.excel-grid tbody tr:nth-child(even) td{background:#f8fcf9!important}.excel-grid tbody tr:hover td{background:#eaf8ef!important}
.excel-grid td{color:#294b39!important;border-color:#e2eee6!important}
.pill,.dash-actions button{border-color:#bcdcc8!important;background:linear-gradient(145deg,#fff,#f1faf4)!important;color:#176443!important;transition:transform .15s,box-shadow .15s!important}
.pill:hover,.dash-actions button:hover{box-shadow:0 4px 13px #16845b24!important;transform:translateY(-1px)}
.upload-replace,.order-upload-pill,.return-upload-pill{background:linear-gradient(110deg,#14784e,#239766)!important;color:#fff!important;border:1px solid #14784e!important;box-shadow:0 5px 14px #16845b24!important}
.upload-replace:hover{background:linear-gradient(110deg,#0f6842,#1d865a)!important;color:#fff!important}
.upload-locked{background:#fff7e5!important;color:#8a6218!important;border-color:#f0d99c!important}
.order-history .excel-stats>div:nth-child(4n+1){border-top:3px solid #16845b!important}.order-history .excel-stats>div:nth-child(4n+2){border-top:3px solid #4b8bd8!important}.order-history .excel-stats>div:nth-child(4n+3){border-top:3px solid #d6a640!important}.order-history .excel-stats>div:nth-child(4n){border-top:3px solid #d9788c!important}
.order-history .visual-card{background:linear-gradient(145deg,#fff,#f0faf4)!important}
input[type=checkbox],input[type=radio]{accent-color:#16845b!important}
input[type=range]{background:linear-gradient(90deg,#16845b,#a7dfba)!important;border-color:#b9dfc8!important}
input[type=range]::-webkit-slider-thumb{border-color:#16845b!important}
#toast{background:linear-gradient(110deg,#123f2d,#16845b)!important;color:#fff!important;border:1px solid #b8e4c8!important}#toast.danger{background:#b83232!important;color:#fff!important}
@media(max-width:800px){.admin-full .admin-nav{background:#eef8f2!important}}

/* FINAL FIX: restore scrollable product detail and compact, image-led product tiles */
.ov:not(.admin-full){overflow-y:auto!important;overflow-x:hidden!important;align-items:flex-start!important;justify-content:center!important;-webkit-overflow-scrolling:touch!important;overscroll-behavior-y:contain!important;padding:12px!important}
.ov:not(.admin-full)>.box:not(.ad){flex:0 0 auto!important;margin:0 auto!important;max-height:calc(100dvh - 24px)!important;overflow-x:hidden!important;overflow-y:auto!important;overscroll-behavior-y:contain!important;-webkit-overflow-scrolling:touch!important;touch-action:pan-y!important}
#mb{min-height:0!important;overflow:visible!important}
.similar-section{display:block!important;clear:both!important;position:relative!important;overflow:visible!important;margin-top:26px!important;padding-top:16px!important}
.similar-section .sg{display:flex!important;flex-flow:row nowrap!important;align-items:stretch!important;width:100%!important;min-width:0!important;max-width:100%!important;overflow-x:auto!important;overflow-y:hidden!important;gap:12px!important;padding:5px 3px 16px!important;scroll-snap-type:x proximity!important;scrollbar-width:thin!important;touch-action:pan-x pan-y!important;-webkit-overflow-scrolling:touch!important}
.similar-section .sg .c{display:block!important;position:relative!important;flex:0 0 148px!important;width:148px!important;min-width:148px!important;max-width:148px!important;height:auto!important;min-height:0!important;overflow:hidden!important;align-self:stretch!important;scroll-snap-align:start!important}
.similar-section .sg .c .iw{height:auto!important;min-height:0!important;overflow:hidden!important}
.similar-section .sg .c .iw>img:first-child{display:block!important;width:100%!important;height:auto!important;aspect-ratio:3/4!important;object-fit:cover!important}
.similar-section .sg .c .m{display:block!important;min-height:0!important;height:auto!important;padding:8px 9px 10px!important;overflow:hidden!important}
.similar-section .sg .c .m>b{display:-webkit-box!important;height:auto!important;max-height:2.8em!important;line-height:1.35!important;-webkit-line-clamp:2!important;-webkit-box-orient:vertical!important;overflow:hidden!important;overflow-wrap:anywhere!important}
.similar-section .sg .c .pr{display:block!important;margin-top:5px!important;font-size:13px!important;white-space:nowrap!important}
/* The main product list remains a responsive card grid, not uniform square tiles. */
#grid{display:grid!important;grid-template-columns:repeat(auto-fill,minmax(205px,1fr))!important;align-items:start!important;gap:16px!important}
#grid>.c{height:auto!important;min-height:0!important;align-self:start!important}
#grid>.c .iw{height:auto!important;min-height:0!important;aspect-ratio:auto!important;overflow:hidden!important}
#grid>.c .iw>img:first-child{display:block!important;width:100%!important;height:auto!important;aspect-ratio:3/4!important;object-fit:cover!important}
#grid>.c .m{height:auto!important;min-height:0!important;overflow:visible!important}
@media(max-width:700px){
 .ov:not(.admin-full){padding:8px!important;align-items:flex-start!important}
 .ov:not(.admin-full)>.box:not(.ad){width:100%!important;max-height:calc(100dvh - 16px)!important;margin:0 auto!important;border-radius:20px!important;padding:16px 12px 20px!important}
 #grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:10px!important}
 .similar-section .sg .c{flex-basis:124px!important;width:124px!important;min-width:124px!important;max-width:124px!important}
}

/* Solid, readable product detail modal; keep the rest of the ERP theme intact. */
.ov:not(.admin-full){background:rgba(15,23,42,.72)!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important}
.ov:not(.admin-full)>.box:not(.ad),#mb{background:#fff!important;color:#1f2937!important;border:1px solid #d8dee8!important;box-shadow:0 18px 55px rgba(0,0,0,.25)!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important;opacity:1!important}
#mb .product-summary,#mb .product-detail-grid,#mb .stock-comments,#mb .similar-section,#mb .th .thumb-wrap{background:#fff!important;color:#1f2937!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important;opacity:1!important}
#mb input,#mb select,#mb textarea{background:#fff!important;border:1px solid #d1d5db!important;backdrop-filter:none!important}

.courier-dashboard{max-width:1450px;margin:0 auto;color:#202a36}.courier-heading{display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:16px}.courier-heading h2{margin:4px 0;font-size:27px}.courier-heading p,.courier-card p{margin:4px 0;color:#667085}.courier-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:16px}.courier-stats article,.courier-card{background:#fff!important;border:1px solid #e1e6ed!important;border-radius:14px!important;box-shadow:0 5px 18px #1720330c!important;backdrop-filter:none!important;-webkit-backdrop-filter:none!important;padding:16px}.courier-stats small{display:block;color:#667085}.courier-stats b{font-size:28px;color:#172b3b}.courier-meter{height:7px;background:#edf1f5;border-radius:9px;margin-top:9px;overflow:hidden}.courier-meter i{display:block;height:100%;border-radius:9px}.courier-card{margin:14px 0}.courier-card h3{margin:0 0 10px}.courier-card h4{margin:16px 0 8px}.courier-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}.courier-tabs button{border:1px solid #d9e0e8;background:#fff;color:#344054;border-radius:9px;padding:8px 12px;cursor:pointer}.courier-tabs button.on{background:#176b4c;color:#fff;border-color:#176b4c}.courier-tabs i{font-size:10px;font-style:normal;opacity:.75}.courier-fields{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;align-items:end}.courier-fields label{display:grid;gap:5px;font-size:12px;font-weight:650;color:#526071;min-width:0}.courier-fields label.wide{grid-column:span 2}.courier-fields input,.courier-fields select{width:100%;min-width:0;padding:10px 11px;background:#fff!important;border:1px solid #d5dce5!important;border-radius:9px!important;box-shadow:none!important;color:#202a36}.courier-fields .courier-toggle{display:flex;align-items:center;grid-template-columns:20px 1fr;padding:10px}.courier-fields .courier-toggle input{width:18px}.courier-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}.courier-actions #fraudResult{flex-basis:100%;color:#344054;font-weight:600}.courier-dashboard .excel-table-wrap{background:#fff!important;backdrop-filter:none!important}.courier-dashboard .excel-grid{min-width:820px}@media(max-width:900px){.courier-fields{grid-template-columns:repeat(2,minmax(0,1fr))}.courier-stats{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:560px){.courier-fields{grid-template-columns:1fr}.courier-fields label.wide{grid-column:auto}.courier-stats{gap:8px}.courier-stats article{padding:12px}.courier-heading h2{font-size:22px}}

/* Compact Excel cells and prevent content from escaping panels on narrow screens. */
.excel-table-wrap{max-width:100%;min-width:0;overflow:auto;overscroll-behavior:contain}.excel-grid .excel-cell{display:block;width:76px!important;min-width:62px!important;max-width:80px!important;height:36px!important;margin:0 auto;padding:5px 8px!important;border:1px solid #d5e0d8!important;border-radius:7px!important;background:#fff!important;box-shadow:none!important;text-align:center;font-size:14px!important}.excel-grid td{vertical-align:middle}.admin-content,.admin-body,.admin-content .admin-section,.admin-content .excel-hero,.admin-content .order-history,.admin-content .courier-dashboard{min-width:0;max-width:100%;box-sizing:border-box}.admin-note{max-width:100%;box-sizing:border-box;overflow-wrap:anywhere;word-break:normal;border-style:solid!important;border-color:#e6dfd1!important;line-height:1.55}.ov:not(.admin-full)>.box:not(.ad){box-sizing:border-box;max-width:min(980px,calc(100vw - 28px))!important;border-radius:22px!important;overflow-x:hidden!important}.ov:not(.admin-full)>.box:not(.ad) :is(#mb,.product-detail-grid,.product-summary,.stock-comments,.similar-section){min-width:0;max-width:100%;box-sizing:border-box}.ov:not(.admin-full)>.box:not(.ad) .similar-section{overflow:hidden!important;border-radius:14px}.ov:not(.admin-full)>.box:not(.ad) .similar-section .sg{min-width:0;max-width:100%;overflow-x:auto!important;overflow-y:hidden!important}.courier-dashboard .courier-heading,.courier-dashboard .courier-card{min-width:0;max-width:100%;box-sizing:border-box}

/* Consistent opaque surfaces: remove the old glass overlay from cards and controls. */
*,*::before,*::after{-webkit-backdrop-filter:none!important;backdrop-filter:none!important}.c,.personal-box,.stock-comments,#cmtSummary,.cs-stats>div,.stock-comment-row,.similar-section,.stat,.section-card,.visual-card,.chart-card,.search-box,.filter-box,.toolbar,.pagination,.empty-state,.th .thumb-wrap{background:#fff!important;border:1px solid #e1e6eb!important;box-shadow:0 3px 14px rgba(24,39,54,.07)!important}.c .m,.product-summary,.product-detail-grid,.stock-comment-form{background:#fff!important}.admin-content,.admin-nav{background:#fff!important}input,select,textarea{background:#fff!important;border:1px solid #d3dbe4!important;box-shadow:none!important}


/* Product card polish and high-contrast similarity text. */
#grid>.c,.similar-section .sg>.c{border:1px solid #dce6df!important;border-radius:22px!important;box-shadow:0 8px 22px #193c2b12!important;overflow:hidden!important}
#grid>.c .m,.similar-section .sg>.c .m{padding:12px 14px 15px!important}
#grid>.c .m>b,.similar-section .sg>.c .m>b{font-weight:700!important;color:#24332b!important}
#grid>.c .copycode,.similar-section .sg>.c .copycode{background:#f4f8f5!important;border-color:#dce9df!important;border-radius:999px!important;padding:5px 10px!important;font-size:12px!important;font-weight:700!important}
#grid>.c .sc,.similar-section .sg>.c .sc{top:auto!important;left:8px!important;right:auto!important;bottom:8px!important;z-index:6!important;max-width:calc(100% - 16px);padding:6px 10px!important;border:1px solid #ffffffd9!important;border-radius:999px!important;background:#103e30ef!important;color:#fff!important;font-size:12px!important;font-weight:750!important;line-height:1.25!important;box-shadow:0 4px 12px #0004!important;white-space:nowrap}
@media(max-width:700px){#grid>.c .sc,.similar-section .sg>.c .sc{font-size:11px!important;padding:5px 8px!important}}
/* Premium catalog filters, echoing the sidebar's rounded icon cards */
.tabs{gap:9px!important}
.tabs button{display:inline-flex!important;align-items:center;gap:10px;min-height:52px;padding:8px 17px 8px 10px!important;border:1px solid #d7e9df!important;border-style:solid!important;border-radius:20px!important;background:linear-gradient(145deg,#fff,#f1faf5)!important;color:#183b2c!important;box-shadow:0 4px 12px #174a3010!important;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease,background .18s ease!important}
.tabs button::before{content:'◇';display:inline-grid;place-items:center;flex:0 0 34px;width:34px;height:34px;border:1px solid #cbe9d6;border-radius:11px;background:#e9f7ee;color:#16845b;font-size:18px;font-weight:700}
/* Keep category filters text-only; their labels are sufficient. */
.tabs button::before{content:none!important;display:none!important}.tabs button{padding-left:18px!important}
.tabs button[data-t="all"]::before{content:'✦'}
.tabs button[data-t="new"]::before{content:'★'}
.tabs button[data-t="s"]::before{content:'♧'}
.tabs button[data-t="t"]::before{content:'◇'}
.tabs button:hover{transform:translateY(-2px)!important;border-color:#9bc9aa!important;background:linear-gradient(145deg,#fff,#eaf7ef)!important;box-shadow:0 9px 20px #173d2b24!important}
.tabs button.on{border:1px solid #9bd8b4!important;border-style:solid!important;background:linear-gradient(145deg,#fff,#e8f7ee)!important;color:#125c3d!important;box-shadow:0 7px 17px #16845b24!important}
.tabs button.on::before{background:linear-gradient(140deg,#13784f,#239866)!important;color:#fff!important;border-color:#13784f!important}
.tabs button:active{transform:translateY(0) scale(.98)!important;box-shadow:0 3px 8px #173d2b24!important}
button:not(:disabled),a.pill{transition:transform .18s ease,box-shadow .18s ease!important}
button:not(:disabled):hover,a.pill:hover{box-shadow:0 7px 17px #173d2b20!important}
button:not(:disabled):active,a.pill:active{transform:translateY(1px)!important;box-shadow:0 3px 8px #173d2b20!important}
#grid>.c,.similar-section .sg>.c{transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease!important}
#grid>.c:hover,.similar-section .sg>.c:hover{transform:translateY(-4px)!important;border-color:#a9cfb6!important;box-shadow:0 16px 30px #173d2b28!important}
#grid>.c:active,.similar-section .sg>.c:active{transform:translateY(-1px) scale(.995)!important;box-shadow:0 8px 18px #173d2b24!important}
button:focus-visible,a:focus-visible,.c:focus-visible{outline:3px solid #77bd91!important;outline-offset:3px!important}
/* Premium midnight-blue header with warm gold brand contrast */
:root{--head:#142235!important}
header{background:linear-gradient(120deg,#0d1726 0%,#172b43 52%,#101b2d 100%)!important;background-image:linear-gradient(120deg,#0d1726 0%,#172b43 52%,#101b2d 100%)!important;border-bottom:2px solid #c9a45c!important;box-shadow:0 10px 30px #0b14254d!important}
.excel-product-placeholder{display:grid;place-content:center;gap:8px;aspect-ratio:3/4;background:radial-gradient(circle at 25% 20%,#e9f6ee,#dce9e1 65%,#cbd9d0);color:#1a5b40;text-align:center;padding:16px;box-sizing:border-box}.excel-product-placeholder b{font-size:clamp(22px,4vw,38px);line-height:1.15;overflow-wrap:anywhere}.excel-product-placeholder small{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#61776b}.courier-config{padding:0!important;overflow:hidden}.courier-config-toggle{display:flex;align-items:center;justify-content:space-between;width:100%;min-height:66px;padding:14px 18px;border:0;background:linear-gradient(120deg,#fff,#f3faf6);color:#173d2d;text-align:left;cursor:pointer}.courier-config-toggle b,.courier-config-toggle small{display:block}.courier-config-toggle b{font-size:16px}.courier-config-toggle small{margin-top:4px;color:#748477;font-weight:400}.courier-chevron{display:grid;place-items:center;width:34px;height:34px;border:1px solid #cde5d5;border-radius:11px;color:#16845b;font-size:20px}.courier-config-body{padding:0 16px 16px}.courier-config-body[hidden]{display:none!important}.courier-money{display:block;margin-top:3px;color:#19845b;font-size:12px;font-weight:600}.courier-date{display:grid;gap:4px;color:#68766f;font-size:11px;font-weight:600}.courier-date input{min-width:145px}
.courier-provider-tabs{display:flex;gap:9px;flex-wrap:wrap;margin:0 0 14px}.courier-provider-tabs button{border:1px solid #d8e0ea;border-radius:999px;background:#fff;padding:9px 16px;color:#243447;font-weight:700;cursor:pointer;box-shadow:0 4px 12px #17203310;transition:transform .16s,box-shadow .16s}.courier-provider-tabs button:hover{transform:translateY(-2px);box-shadow:0 7px 18px #17203320}.courier-provider-tabs button.on{background:#164a3a;color:#fff;border-color:#164a3a}.courier-stats{grid-template-columns:repeat(6,minmax(0,1fr))!important}@media(max-width:1050px){.courier-stats{grid-template-columns:repeat(3,minmax(0,1fr))!important}}@media(max-width:650px){.courier-stats{grid-template-columns:repeat(2,minmax(0,1fr))!important}}
/* Dashboard polish: consistent cards, compact comments, aligned courier actions. */
.courier-head-actions{display:flex;align-items:center;justify-content:flex-end;gap:9px;flex-wrap:wrap;margin-left:auto}
.courier-head-actions #courierSyncProgress{flex-basis:100%;text-align:right;font-size:12px;color:#667085}
.courier-head-actions .courier-report-link{white-space:nowrap}
.comments-overview{margin:12px 0 18px;border:1px solid #d7e8df;border-radius:18px;background:linear-gradient(135deg,#f8fcf9,#fff);box-shadow:0 8px 24px #1233220b;overflow:hidden}
.comments-overview>summary{display:flex;align-items:center;justify-content:space-between;gap:12px;cursor:pointer;list-style:none;padding:17px 22px;color:#145d40}
.comments-overview>summary::-webkit-details-marker{display:none}
.comments-overview>summary span{display:flex;flex-direction:column;gap:4px}
.comments-overview>summary b{font-size:17px}
.comments-overview>summary small{font-size:12px;color:#7b8e83}
.comments-overview>summary i{font-style:normal;border:1px solid #cde4d7;border-radius:10px;padding:4px 10px;font-size:18px}
.comments-overview[open]>summary i{transform:rotate(45deg)}
.comments-overview-body{padding:0 20px 20px}
/* Keep Stock Excel available in admin, but don't show its source as a customer-facing category chip. */
/* Requested sidebar color palette: deep Alien Tees green #0F3325 */
#dr{background:#0F3325!important;background-image:none!important;color:#fff!important;border-right:1px solid #ffffff26!important;}
#dr .dr-head{background:#0F3325!important;background-image:none!important;color:#fff!important;border-bottom:1px solid #ffffff26!important;border-radius:0!important;box-shadow:none!important;}
#dr .dr-brand b,#dr .dr-brand small,#dr label{color:#fff!important;}
#dr .dr-x{background:#ffffff12!important;color:#fff!important;border-color:#ffffff35!important;}
/* Prevent the stock status, category selector and size controls from rendering as white cards. */
#dr #sync{background:#174535!important;background-image:none!important;color:#e7f5ed!important;border:1px solid #ffffff20!important;box-shadow:none!important;}
#dr #sync .n{color:#e7f5ed!important;}
#dr select{background:#174535!important;background-image:none!important;color:#fff!important;border-color:#ffffff35!important;}
#dr select option{background:#0F3325!important;color:#fff!important;}
#dr #ft,#dr #fk,#dr #so{background-color:#174535!important;color:#fff!important;}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){color:#fff!important;background:transparent!important;border-color:transparent!important;}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover{background:#ffffff12!important;border-color:#ffffff26!important;}
#dr button::before{background:#ffffff10!important;border-color:#ffffff26!important;}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle)::after{background:#fff!important;filter:none!important;}
#dr #dsz{background:transparent!important;background-image:none!important;}
#dr #dsz button{background:#174535!important;background-image:none!important;color:#fff!important;border-color:#ffffff2b!important;box-shadow:none!important;}
#dr #dsz button:hover{background:#205b43!important;border-color:#ffffff55!important;}
#dr #dsz .on{background:#2b7956!important;color:#fff!important;border-color:#84c6a3!important;}
#dr #dsz small{color:#d4e8dc!important;}
#dr .personal-box{background:#174535!important;background-image:none!important;border-color:#ffffff26!important;color:#fff!important;}
#dr .personal-toggle{color:#fff!important;background:transparent!important;}
#dr #lo{color:#ffe3e3!important;}
/* Searchable Bangladesh district/thana autocomplete. */
#posGeoSelects .pos-geo-field{position:relative;width:100%;min-width:0;}
#posGeoSelects .pos-geo-field>input{width:100%;box-sizing:border-box;min-width:0;background:#fff;color:#17251f;border:1px solid #d5e3d9;border-radius:10px;padding:10px 12px;font:inherit;outline:none;}
#posGeoSelects .pos-geo-field>input:focus{border-color:#0D7552;box-shadow:0 0 0 3px #0D755224;}
#posGeoSelects .pos-geo-field>input:disabled{background:#f1f4f2;color:#89948c;cursor:not-allowed;}
#posGeoSelects .pos-geo-suggestions{position:absolute;left:0;right:0;top:calc(100% + 4px);z-index:2147482000;max-height:220px;overflow-y:auto;border:1px solid #cbded2;border-radius:10px;background:#fff;box-shadow:0 12px 28px #10291b30;padding:4px;}
#posGeoSelects .pos-geo-suggestions[hidden]{display:none!important;}
#posGeoSelects .pos-geo-option{display:block!important;width:100%;padding:9px 10px!important;border:0!important;border-radius:7px!important;background:#fff!important;color:#17251f!important;text-align:left!important;box-shadow:none!important;font:inherit!important;cursor:pointer;}
#posGeoSelects .pos-geo-option:hover,#posGeoSelects .pos-geo-option:focus{background:#eaf6ef!important;color:#0D7552!important;outline:none;}
#posGeoSelects .pos-geo-option small{display:block;color:#718177;font-weight:400;font-size:11px;margin-top:1px;}
#posGeoSelects select[hidden]{display:none!important;}

/* FINAL REQUESTED VISUAL UPDATE: solid Alien Tees green sidebar, transparent brand header, subtle glass-card borders. */
#dr{background:#0D7552!important;background-color:#0D7552!important;background-image:none!important;color:#fff!important;border-right:1px solid rgba(255,255,255,.28)!important;}
/* Keep the Smart ERP Management Panel / Alien Tees brand area clear; no white header block. */
#dr .dr-head{background:transparent!important;background-color:transparent!important;background-image:none!important;color:#fff!important;border-bottom:1px solid rgba(255,255,255,.24)!important;box-shadow:none!important;}
#dr .dr-head .dr-brand,#dr .dr-head .dr-brand>div,#dr .dr-head .dr-brand b,#dr .dr-head .dr-brand small,#dr .dr-head .dr-brand img{background:transparent!important;background-color:transparent!important;background-image:none!important;box-shadow:none!important;}
#dr .dr-head .dr-brand b,#dr .dr-head .dr-brand small{color:#fff!important;}
#dr .dr-head .dr-x{background:rgba(255,255,255,.10)!important;background-color:rgba(255,255,255,.10)!important;color:#fff!important;border:1px solid rgba(255,255,255,.35)!important;}
/* Keep all sidebar controls in the green palette rather than allowing inherited white backgrounds. */
#dr #sync,#dr select,#dr #ft,#dr #fk,#dr #so,#dr #dsz button,#dr .personal-box{background:#075D42!important;background-color:#075D42!important;background-image:none!important;color:#fff!important;border-color:rgba(255,255,255,.24)!important;}
#dr select option{background:#0D7552!important;color:#fff!important;}
#dr #dsz{background:transparent!important;background-image:none!important;}
#dr #dsz .on{background:#16845F!important;background-color:#16845F!important;color:#fff!important;border-color:#b8ead0!important;}
#dr .personal-toggle{background:transparent!important;background-color:transparent!important;color:#fff!important;}
#dr .dr-x{color:#fff!important;}
/* A subtle 1px white outline around the existing glass-style cards; preserve card layout and behavior. */
.box,.c,.personal-box,.stock-comments,#cmtSummary,.cs-stats>div,.stock-comment-row,.similar-section,.admin-head,.admin-nav,.stat,.admin-content,.ad,.ov>.box,#bar,#siteNoticeCard,.modal,.panel,.card,.section-card,.product-summary,.product-detail-grid,.stock-comment-form,.th .thumb-wrap,.search-box,.filter-box,.toolbar,.pagination,.empty-state{border:1px solid rgba(255,255,255,.62)!important;border-style:solid!important;}

/* SIDEBAR GLASS OPTION HIGHLIGHT: brighter, clearer controls without changing layout or behavior. */
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle){
  color:#fff!important;
  background:linear-gradient(135deg,rgba(255,255,255,.22),rgba(255,255,255,.09) 55%,rgba(0,40,25,.10))!important;
  background-color:rgba(255,255,255,.12)!important;
  border:1px solid rgba(255,255,255,.46)!important;
  border-radius:14px!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.24),0 5px 14px rgba(0,25,16,.18)!important;
  text-shadow:0 1px 2px rgba(0,35,20,.35)!important;
  transition:background .18s ease,border-color .18s ease,box-shadow .18s ease,transform .18s ease!important;
}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover,
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):focus-visible{
  color:#fff!important;
  background:linear-gradient(135deg,rgba(255,255,255,.34),rgba(255,255,255,.19) 55%,rgba(0,50,29,.10))!important;
  border-color:rgba(255,255,255,.82)!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.38),0 0 0 2px rgba(224,255,238,.12),0 8px 22px rgba(0,25,16,.24)!important;
  outline:none!important;
}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle).on,
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle).active,
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle)[aria-current="page"]{
  color:#fff!important;
  background:linear-gradient(120deg,rgba(211,255,229,.32),rgba(255,255,255,.18))!important;
  border-color:#e4fff0!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.42),0 0 0 1px rgba(230,255,240,.18),0 7px 20px rgba(0,28,16,.26)!important;
}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle)::before{
  background:rgba(255,255,255,.13)!important;
  border-color:rgba(255,255,255,.32)!important;
}
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):hover::before,
#dr button:not(.dr-x):not(#dsz button):not(.personal-toggle):focus-visible::before{
  background:rgba(255,255,255,.23)!important;
  border-color:rgba(255,255,255,.60)!important;
}
/* Make the Category, Stock, and Sort controls read as glass chips against green. */
#dr #sync,#dr select,#dr #ft,#dr #fk,#dr #so{
  color:#fff!important;
  background:linear-gradient(135deg,rgba(255,255,255,.20),rgba(255,255,255,.08))!important;
  border:1px solid rgba(255,255,255,.52)!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.22),0 5px 14px rgba(0,25,16,.16)!important;
}
#dr #sync{border-radius:14px!important;}
#dr select:focus-visible{outline:2px solid rgba(255,255,255,.85)!important;outline-offset:2px!important;}
#dr #dsz button{
  color:#fff!important;
  background:linear-gradient(135deg,rgba(255,255,255,.20),rgba(255,255,255,.08))!important;
  border:1px solid rgba(255,255,255,.48)!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.24),0 4px 12px rgba(0,25,16,.16)!important;
}
#dr #dsz button:hover,#dr #dsz button:focus-visible{
  background:linear-gradient(135deg,rgba(255,255,255,.34),rgba(255,255,255,.17))!important;
  border-color:rgba(255,255,255,.85)!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.4),0 0 0 2px rgba(255,255,255,.10),0 7px 17px rgba(0,25,16,.22)!important;
}
#dr #dsz button.on,#dr #dsz .on{
  color:#fff!important;
  background:linear-gradient(120deg,#1c986b,#37b985)!important;
  border-color:#e7fff1!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.42),0 0 0 1px rgba(255,255,255,.20),0 7px 18px rgba(0,25,16,.24)!important;
}
#dr .personal-box{background:linear-gradient(135deg,rgba(255,255,255,.17),rgba(255,255,255,.07))!important;border:1px solid rgba(255,255,255,.44)!important;box-shadow:inset 0 1px 0 rgba(255,255,255,.20),0 6px 16px rgba(0,25,16,.15)!important;}
#dr .personal-toggle:hover,#dr .personal-toggle:focus-visible{background:rgba(255,255,255,.15)!important;outline:1px solid rgba(255,255,255,.35)!important;outline-offset:-2px!important;}
@media(prefers-reduced-motion:reduce){#dr button{transition:none!important}}


/* FIX: keep the sidebar brand header fully opaque while the menu scrolls beneath it. */
#dr{overflow-x:hidden!important;overflow-y:auto!important;overscroll-behavior:contain!important;}
#dr .dr-head{
  position:sticky!important;
  top:0!important;
  z-index:2147483000!important;
  background:#0D7552!important;
  background-color:#0D7552!important;
  background-image:none!important;
  opacity:1!important;
  backdrop-filter:none!important;
  -webkit-backdrop-filter:none!important;
  isolation:isolate!important;
  box-shadow:0 8px 0 #0D7552!important;
}
#dr .dr-head::before,#dr .dr-head::after{background:none!important;pointer-events:none!important;}
#dr .dr-head .dr-brand,#dr .dr-head .dr-brand>div,#dr .dr-head .dr-brand b,#dr .dr-head .dr-brand small,#dr .dr-head .dr-brand img{background:transparent!important;background-color:transparent!important;background-image:none!important;}

/* Compact standard loader — keep all global loading triggers and logic unchanged. */
#searchLoader{padding:10px!important;background:rgba(3,18,13,.12)!important;backdrop-filter:blur(1px)!important;-webkit-backdrop-filter:blur(1px)!important;}
.search-loader-card{width:auto!important;max-width:calc(100vw - 28px)!important;min-width:0!important;padding:10px 14px!important;border-radius:13px!important;display:grid!important;grid-template-columns:25px minmax(0,1fr)!important;grid-template-rows:auto auto!important;align-items:center!important;column-gap:9px!important;row-gap:1px!important;text-align:left!important;border:1px solid rgba(230,255,240,.45)!important;background:rgba(7,49,35,.96)!important;box-shadow:0 7px 24px rgba(0,0,0,.22),inset 0 1px 0 rgba(255,255,255,.1)!important;transform:translateY(3px)!important;}
#searchLoader.is-active .search-loader-card{transform:translateY(0)!important;}
.search-loader-card:before{display:none!important;animation:none!important;}
.search-loader-scene{grid-column:1!important;grid-row:1/3!important;width:25px!important;height:25px!important;margin:0!important;filter:none!important;}
.search-loader-ring.r1{width:22px!important;height:22px!important;border-width:2px!important;box-shadow:none!important;animation:searchRingSpin .8s linear infinite!important;}
.search-loader-ring.r2,.search-loader-ring.r3,.search-loader-core,.search-loader-spark{display:none!important;animation:none!important;}
.search-loader-title{grid-column:2!important;grid-row:1!important;margin:0!important;font:600 12px/1.35 system-ui,'Noto Sans Bengali',sans-serif!important;letter-spacing:0!important;white-space:normal!important;color:#fff!important;}
.search-loader-sub{display:none!important;}
.search-loader-progress{grid-column:1/3!important;grid-row:3!important;height:2px!important;margin:7px 0 0!important;background:rgba(255,255,255,.15)!important;box-shadow:none!important;}
.search-loader-progress i{width:30%!important;box-shadow:none!important;animation:searchProgressRun .95s ease-in-out infinite!important;}
.search-loader-foot{display:none!important;}
@media(prefers-reduced-motion:reduce){.search-loader-card,.search-loader-ring,.search-loader-progress i{animation:none!important;transition:none!important}}

</style>
<script>
/* Global overlay disabled: do not cover/block the whole app for every API request.
   The search pane itself continues to show its own inline feedback. */
const __siteLoaderState={pending:0,manual:0,initial:false,navigating:false,kind:'global',showTimer:null,hideTimer:null};
function __siteLoaderBusy(){return false}
function __siteLoaderMessage(){const t=document.getElementById('searchLoadingTitle'),d=document.getElementById('searchLoadingSub');if(!t||!d)return;if(__siteLoaderState.manual>0){const im=__siteLoaderState.kind==='image';t.textContent=im?'ছবির মিল খোঁজা হচ্ছে':'প্রোডাক্ট সার্চ চলছে';d.textContent=im?'রঙ, ডিজাইন ও ছবির মিল বিশ্লেষণ করা হচ্ছে…':'প্রোডাক্ট কোড, নাম ও ক্যাটালগ স্ক্যান করা হচ্ছে…'}else if(__siteLoaderState.initial){t.textContent='Alien Tees ERP প্রস্তুত হচ্ছে';d.textContent='ড্যাশবোর্ড, ক্যাটালগ ও সেটিংস লোড করা হচ্ছে…'}else if(__siteLoaderState.navigating){t.textContent='পৃষ্ঠা পরিবর্তন হচ্ছে';d.textContent='আপনার নির্বাচিত পৃষ্ঠা প্রস্তুত করা হচ্ছে…'}else{t.textContent='সিস্টেম লোড হচ্ছে';d.textContent='তথ্য আনা, সংরক্ষণ, Sync ও রিপোর্ট প্রস্তুত করা হচ্ছে…'}}
function __siteLoaderShow(){const el=document.getElementById('searchLoader');if(!el||!__siteLoaderBusy())return;clearTimeout(__siteLoaderState.hideTimer);el.hidden=false;__siteLoaderMessage();if(!el.classList.contains('is-active')){el.classList.remove('is-active');void el.offsetWidth;requestAnimationFrame(()=>{if(__siteLoaderBusy()){el.hidden=false;el.classList.add('is-active')}})}else el.classList.add('is-active')}
function __siteLoaderRefresh(immediate=false){const el=document.getElementById('searchLoader');if(!el)return;if(__siteLoaderBusy()){clearTimeout(__siteLoaderState.hideTimer);__siteLoaderMessage();if(immediate||__siteLoaderState.initial||__siteLoaderState.manual>0){clearTimeout(__siteLoaderState.showTimer);__siteLoaderShow()}else if(el.hidden){clearTimeout(__siteLoaderState.showTimer);__siteLoaderState.showTimer=setTimeout(__siteLoaderShow,150)}else el.classList.add('is-active')}else{clearTimeout(__siteLoaderState.showTimer);if(el.hidden)return;el.classList.remove('is-active');clearTimeout(__siteLoaderState.hideTimer);__siteLoaderState.hideTimer=setTimeout(()=>{if(!__siteLoaderBusy())el.hidden=true},270)}}
function __siteLoaderBegin(){__siteLoaderState.pending++;__siteLoaderRefresh(false);let closed=false;return()=>{if(closed)return;closed=true;__siteLoaderState.pending=Math.max(0,__siteLoaderState.pending-1);__siteLoaderRefresh(false)}}
const __siteLoaderBootDone=()=>{__siteLoaderState.initial=false;__siteLoaderRefresh(false)};if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',__siteLoaderBootDone,{once:true});else setTimeout(__siteLoaderBootDone,0);
window.addEventListener('beforeunload',()=>{__siteLoaderState.navigating=true;__siteLoaderRefresh(true)});
/* Track all Fetch calls until their response body has also been consumed, not merely until headers arrive. */
if(typeof window.fetch==='function'&&!window.fetch.__siteLoaderWrapped){const __nativeFetch=window.fetch.bind(window);const __wrappedFetch=function(...args){const release=__siteLoaderBegin();let fp;try{fp=__nativeFetch(...args)}catch(e){release();throw e}return Promise.resolve(fp).then(response=>{let bodyUsed=false,done=false,timer=null;const finish=()=>{if(done)return;done=true;clearTimeout(timer);release()};['json','text','blob','arrayBuffer','formData'].forEach(method=>{if(typeof response[method]!=='function')return;const nativeMethod=response[method].bind(response);try{Object.defineProperty(response,method,{configurable:true,value:(...bodyArgs)=>{bodyUsed=true;clearTimeout(timer);let result;try{result=nativeMethod(...bodyArgs)}catch(e){finish();throw e}return Promise.resolve(result).finally(finish)}})}catch(_){}});timer=setTimeout(()=>{if(!bodyUsed)finish()},90);return response},error=>{release();throw error})};__wrappedFetch.__siteLoaderWrapped=true;window.fetch=__wrappedFetch}
/* Track XHR too, so older widgets and uploads get the same animation. */
if(window.XMLHttpRequest&&!XMLHttpRequest.prototype.__siteLoaderWrapped){const __xhrOpen=XMLHttpRequest.prototype.open,__xhrSend=XMLHttpRequest.prototype.send;XMLHttpRequest.prototype.open=function(...a){this.__siteLoaderURL=a[1];return __xhrOpen.apply(this,a)};XMLHttpRequest.prototype.send=function(...a){const release=__siteLoaderBegin();let done=false;const finish=()=>{if(done)return;done=true;release()};this.addEventListener('loadend',finish,{once:true});try{return __xhrSend.apply(this,a)}catch(e){finish();throw e}};XMLHttpRequest.prototype.__siteLoaderWrapped=true}

const $=id=>document.getElementById(id),bp=n=>Number(n).toLocaleString(uiLanguage==='en'?'en-US':'en-US'),bt=s=>'<span class="n">'+s+'</span>',bn=n=>bt(bp(n)),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const im=u=>'/api/img?url='+encodeURIComponent(u),DZ0='',bd=s=>String(s).replace(/\d/g,c=>'০১২৩৪৫৬৭৮৯'[c]);
function bnTime(d){const h=d.getHours(),p=h<4?'রাত':h<6?'ভোর':h<12?'সকাল':h<15?'দুপুর':h<17?'বিকাল':h<19?'সন্ধ্যা':'রাত',z=n=>String(n).padStart(2,'0');return p+' '+bd((h%12||12)+':'+z(d.getMinutes())+':'+z(d.getSeconds()))+' মিনিট'}
let aiOn=false;const dzTxt=()=>{};
let __imagePreviewUrl='';
function dzReset(){
 const box=$('imageSearchPreview'),img=$('imageSearchPreviewImg'),wrap=$('q')?.closest('.search-wrap');
 if(__imagePreviewUrl&&__imagePreviewUrl.startsWith('blob:')){try{URL.revokeObjectURL(__imagePreviewUrl)}catch(_){}}
 __imagePreviewUrl='';if(img)img.removeAttribute('src');if(box)box.hidden=true;if(wrap)wrap.classList.remove('has-image');
}
function showPrev(f){
 if(!f||!String(f.type||'').startsWith('image/'))return;
 const box=$('imageSearchPreview'),img=$('imageSearchPreviewImg'),wrap=$('q')?.closest('.search-wrap');if(!box||!img||!wrap)return;
 if(__imagePreviewUrl&&__imagePreviewUrl.startsWith('blob:')){try{URL.revokeObjectURL(__imagePreviewUrl)}catch(_){}}
 try{__imagePreviewUrl=URL.createObjectURL(f);img.src=__imagePreviewUrl;box.title=f.name||'Selected image';box.hidden=false;wrap.classList.add('has-image')}catch(_){box.hidden=true;wrap.classList.remove('has-image')}
}
const GUEST_PERM_LABELS={search:' টেক্সট সার্চ',image_search:' ইমেজ সার্চ',view_product:' প্রোডাক্ট দেখা',view_stock:' স্টক দেখা',size_filter:' সাইজ ফিল্টার',download:' HD ডাউনলোড',copy:' HD ছবি কপি',link_copy:' লিংক কপি',id_copy:' ID কপি',name_copy:' নাম কপি',full_info_copy:' Full Info',share_product:' Share',website_view:' ওয়েবসাইটে দেখুন'};
const CS_DEFAULTS_JS={search:true,image_search:true,view_product:true,website_view:true,price_edit:false,view_stock:false,size_filter:false,download:false,copy:false,link_copy:false,id_copy:false,name_copy:false,full_info_copy:false,share_product:false,favorite:false,report:false};
let ALL=[],MAP={},store='',tabf='',szf='',sizeFiltersExpanded=false,newOnly=false,lastSig='',lastOrderVersion=-1,hitK='',sel=new Map(),res=null,shown=0,list=[],me=null,meRole='',metaD={},perms={},guestPerms={},favorites=new Set(),recentKeys=[];
const jf=async(u,o)=>{try{const r=await fetch(u,o);let j={};try{j=await r.json()}catch(e){j={error:(r.status>=500?'Server error ('+r.status+'). Check the app console/log for details.':'Server returned an unreadable response ('+r.status+').')}};return Object.assign(j,{_ok:r.ok,_s:r.status})}catch(e){return {_ok:false,_s:0,error:'Network error: '+String(e&&e.message||e)}}};
const post=(u,b,m)=>jf(u,{method:m||'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b||{})});
async function loadPersonal(){
 if(!me){favorites=new Set();recentKeys=[];$('personal').hidden=true;return}
 const [d,sj]=await Promise.all([jf('/api/me/data'),jf('/api/my/searches')]);
 favorites=new Set(d.favorites||[]);recentKeys=(d.recent||[]).map(x=>x.product_key);
 const fav=recentKeys.filter(k=>favorites.has(k));
 const recent=recentKeys;
  const btn=k=>MAP[k]?`<button onclick="openM('${esc(k)}')">🕘 ${esc(productId(MAP[k])||MAP[k].title)} — ${esc(MAP[k].title).slice(0,28)}</button>`:'';
 const searchBtn=x=>`<button type="button" class="replay-search" data-query="${esc(x.query||'')}">🔎 ${esc(x.query||'')}</button>`;
 $('personal').hidden=true;
 $('personal').innerHTML=`<label>Recently Viewed</label>${recent.slice(0,10).map(btn).join('')||'<small style="color:#999">None</small>'}<label>Recent Searches</label>${(sj.searches||[]).slice(0,10).map(searchBtn).join('')||'<small style="color:#999">None</small>'}`
 $('personal').onclick=e=>{const b=e.target.closest('.replay-search');if(b)replaySearch(b.dataset.query||'')};
}
function togglePersonal(){const el=$('personal'),ar=$('personalArrow');if(!el)return;el.hidden=!el.hidden;if(ar)ar.textContent=el.hidden?'＋':'−'}

async function toggleFav(k){return;if(!me)return needLogin();if(!(allowed('favorite')||meRole==='admin'))return toast('আপনার ফেভারিট permission নেই',true);const j=await post('/api/favorite',{key:k});if(j._ok){j.on?favorites.add(k):favorites.delete(k);render(1);loadPersonal();toast(j.on?'ফেভারিটে রাখা হয়েছে':'ফেভারিট থেকে সরানো হয়েছে')}}
function productId(p){
 // ID একমাত্র backend-এর display_id থেকে আসবে; p.id/SKU কখনো ব্যবহার হবে না।
 return String(p?.display_id||'').trim();
}
function findStockCatalogProduct(code){const target=String(code||'').trim().toLowerCase();if(!target)return null;return ALL.find(p=>[p.id,p.display_id,productId(p)].some(v=>String(v||'').trim().toLowerCase()===target)||String(p.title||'').trim().toLowerCase().startsWith(target+' '))||null}
function stockThumb(code){const p=findStockCatalogProduct(code);return p?.img?`<img class="stock-thumb" loading="lazy" src="${im(p.img)}" alt="${esc(code)}" title="${esc(p.title||code)}" onerror="this.style.display='none'">`:`<span class="stock-thumb-empty" title="ক্যাটালগে ছবি মেলেনি">—</span>`}
function openStockProductFromRow(code){const p=findStockCatalogProduct(code);if(p?.key)return openM(p.key);toast('এই কোডের প্রোডাক্ট ক্যাটালগে পাওয়া যায়নি',true)}
function stockState(p){
 const sz=p?.sizes||{}, vals=Object.values(sz).map(Number).filter(v=>v>0), total=vals.reduce((a,b)=>a+b,0);
 if(!vals.length)return {state:'out',total:0};
 const threshold=Math.max(1,Number(localStorage.getItem('stockThreshold')||2));
 return {state:Math.min(...vals)<=threshold?'low':'in',total};
}
async function copyText(t){if(!t)return false;try{if(navigator.clipboard&&window.isSecureContext){await navigator.clipboard.writeText(t);return true}}catch(e){}try{const ta=document.createElement('textarea');ta.value=t;ta.setAttribute('readonly','');ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();ta.setSelectionRange(0,ta.value.length);const ok=document.execCommand('copy');ta.remove();return ok}catch(e){return false}}
async function copyCode(k){if(!canDo('id_copy'))return toast('ID কপি permission বন্ধ আছে',true);const p=MAP[k];if(!p)return;const id=productId(p);if(await copyText(id)){logUiAction('id_copy',id,'ID');toast(id+' কপি হয়েছে')}else toast('ID কপি হয়নি — ব্রাউজারে clipboard permission দিন',true)}
async function copyProductName(k){if(!canDo('name_copy'))return toast('নাম কপি permission বন্ধ আছে',true);const p=MAP[k];if(!p)return;if(await copyText(p.title||'')){logUiAction('name_copy',productId(p)||k,'Product Name');toast('Product Name কপি হয়েছে')}else toast('কপি হয়নি',true)}
async function copyFullInfo(k){if(!canDo('full_info_copy'))return toast('Full Info permission বন্ধ আছে',true);const p=MAP[k];if(!p)return;const id=productId(p),st=p.sizes||{};const stock=canDo('view_stock')?(Object.keys(st).length?Object.entries(st).map(([a,b])=>a+': '+b).join(' | '):'স্টক নেই'):'';const txt=[id?'ID: '+id:'',p.title||'',p.store?'Brand: '+p.store:'','Price: ৳'+Number(p.price||0).toLocaleString('en-US'),stock?'Stock: '+stock:''].filter(Boolean).join('\n');if(await copyText(txt)){logUiAction('full_info_copy',id||k,'Full Info');toast('Full Info কপি হয়েছে')}else toast('কপি হয়নি',true)}
async function shareProduct(k){if(!canDo('share_product'))return toast('Share permission বন্ধ আছে',true);const p=MAP[k];if(!p)return;const id=productId(p),txt=[id||'',p.title||'',p.price?'৳'+Number(p.price||0).toLocaleString('en-US'):'',p.link||''].filter(Boolean).join('\n');try{if(navigator.share){await navigator.share({title:p.title||id||'Product',text:txt,url:p.link||location.href});logUiAction('share_product',id||k,'native share');toast('Share করা হয়েছে');return}}catch(e){}const wa='https://wa.me/?text='+encodeURIComponent(txt);if(p.link){logUiAction('share_product',id||k,'whatsapp');window.open(wa,'_blank','noopener');return}if(await copyText(txt)){logUiAction('share_product',id||k,'copy share text');toast('Share text কপি হয়েছে')}}
 function toggleQuickMode(){const on=!document.body.classList.contains('quick');document.body.classList.toggle('quick',on);localStorage.setItem('quickMode',on?'1':'0');logUiAction('quick_mode',on?'on':'off',on?'on':'off');$('quickbtn').textContent=on?'⚡ Quick mode: On':'⚡ Quick mode: Off';toast(on?'Quick mode on':'Quick mode off')}
function reportProduct(k){if(!canDo('report'))return toast('আপনার রিপোর্ট permission নেই',true);const reason=prompt('কী সমস্যা?','প্রোডাক্ট পাওয়া যাচ্ছে না / তথ্য ভুল');if(reason===null)return;post('/api/report',{key:k,reason}).then(j=>j._ok?toast('রিপোর্ট Admin-এর কাছে গেছে'):toast('রিপোর্ট করা যায়নি',true))}
 function toast(t,danger=false){const e=$('toast');e.textContent=UI_EN[t]||t;translateUi(e);e.classList.toggle('danger',!!danger);e.style.display='block';requestAnimationFrame(()=>e.classList.add('show'));clearTimeout(e._t);e._t=setTimeout(()=>{e.classList.remove('show');setTimeout(()=>{if(!e.classList.contains('show'))e.style.display='none'},220)},2600)}
function needLogin(){toast('এই ফিচার ব্যবহার করতে ইউজার দিয়ে লগইন করুন',true)}
function dr(o){$('dr').classList.toggle('on',!!o);$('bd').classList.toggle('on',!!o)}
setInterval(()=>{const d=new Date(),loc=uiLanguage==='en'?'en-US':'en-US';$('clk').innerHTML=bt(d.toLocaleDateString(loc,{day:'numeric',month:'long',year:'numeric'})+' · '+d.toLocaleTimeString(loc,{hour:'numeric',minute:'2-digit',second:'2-digit'}))},1000);

// ---- থিম ----
const UI_EN={
 'কুরিয়ার ট্র্যাকিং':'Courier Tracking','কুরিয়ার ট্র্যাকিং':'Courier Tracking','স্টক ম্যানেজমেন্ট':'Stock Management','ড্যাশবোর্ড':'Dashboard','ক্যাটালগ':'Catalog','দাম এডিট':'Price Editor','নোটিশ':'Notices','Notifications':'Notifications','Photo Download':'ছবি ডাউনলোড','Order History':'অর্ডার হিস্ট্রি','Return History':'রিটার্ন হিস্ট্রি','ডিজাইন':'Design','ব্যাকআপ':'Backup','ইউজার ম্যানেজমেন্ট':'User Management','সব Activity':'All Activity','পাসওয়ার্ড':'Password','Excel স্টক ড্যাশবোর্ড':'Excel Stock Dashboard','প্রডাক্ট স্টক':'Product Stock','অ্যাকাউন্ট ও ম্যানেজমেন্ট':'Account & Management','প্রডাক্ট':'Product','স্টক':'Stock','ফিল্টার':'Filter','ফটো ডাউনলোড':'Photo Download','সাম্প্রতিক':'Recent','আমার তালিকা':'My list','দ্রুত মোড':'Quick mode','বন্ধ':'Off','চালু':'On','সব প্রোডাক্ট':'All products','নতুন আগমন':'New arrivals','শিট':'Sheet','সেভ পরিবর্তন':'Save changes','মোট অর্ডার':'Total orders','মোট রিটার্ন':'Total returns','অ্যাসাইনি':'Assignee','স্টক সেন্টার':'Stock Center','অ্যাডমিন ম্যানেজমেন্ট':'Admin Management','অনুমতি':'Permissions','প্রোডাক্ট রিপোর্ট':'Product reports','এক্সেল রিপোর্ট':'Excel report',
 'Drop an image here, click to browse, or paste an image':'Drop, click, or paste an image here','কোড বা নাম লিখে খুঁজুন (যেমন 1R01 বা Polo)':'Search by code or name (e.g. 1R01 or Polo)','সিমিলার প্রোডাক্ট':'Similar Products','মিলের হার':'Similarity','মিল':'Match','নতুন':'New','স্টকে':'In stock','কম স্টক':'Low stock','স্টক নেই':'Out of stock','ID নেই':'No ID','খুঁজুন':'Search','রিসেট':'Reset','নতুন ফাইল':'New file','সব সিলেক্ট':'Select all','সিলেক্ট করা ডিলিট':'Delete selected','সেভ করুন':'Save','সেভ করুন':'Save','সেটিংস সেভ করুন':'Save settings','লগইন':'Log in','লগআউট':'Log out','বন্ধ করুন':'Close','দেখেছি':'Mark as read','সবগুলো দেখেছি':'Mark all as read','কোনো Notifications নেই':'No notifications','কোনো প্রোডাক্ট নেই':'No products','কোনো activity নেই':'No activity','কোনো রিপোর্ট নেই':'No reports','কোনো সারি নেই':'No rows','ডাটা নেই':'No data','ছবি':'Images','সাইজ':'Size','দাম':'Price','প্রোডাক্ট':'Product','প্রোডাক্ট কোড':'Product Code','কাস্টমারের নাম':'Customer Name','ফোন':'Phone','ঠিকানা':'Address','সময়':'Time','কারণ':'Reason','সংখ্যা':'Count','অর্ডার':'Orders','রিটার্ন':'Returns','ডেলিভারি':'Delivery','ডেলিভারড':'Delivered','অপেক্ষমাণ':'Pending','ইন ওয়ে':'In Way','সব':'All','সব ট্যাব':'All sheets','মোট স্টক':'Total stock','মোট প্রোডাক্ট':'Total products','স্টক ট্যাব':'Stock sheets','মোট স্টক মূল্য':'Total stock value','অপেক্ষমাণ পরিবর্তন':'Pending changes','আজকের স্টক কার্যক্রম':'Today’s stock activity','ট্যাব অনুযায়ী স্টক':'Stock by sheet','স্টকের বণ্টন':'Stock distribution','স্টক পরিবর্তনগুলো Save Changes দিলে Excel-এ সেভ হবে।':'Stock changes are saved to Excel when you select Save Changes.','ইমেজ সার্চ':'Image Search','নাম কপি':'Copy name','ID কপি':'Copy ID','Full Info':'Full Info','Share':'Share','রিপোর্ট':'Report','ওয়েবসাইটে দেখুন':'View website','HD ডাউনলোড':'Download HD','এই ছবি কপি':'Copy this image','সব ছবি এক এক করে কপি':'Copy images one by one','ওয়েবসাইটের ভাষা':'Website language','বাংলা':'Bengali','Professional English':'Professional English','রেডিমেড টেমপ্লেট':'Ready-made themes','সাইটের নাম':'Website name','ট্যাগলাইন':'Tagline','লোগো':'Logo','সেভ':'Save','রিসেট প্রিভিউ':'Reset preview','কুরিয়ার':'Courier','কুরিয়ার':'Courier','Status dashboard':'স্ট্যাটাস ড্যাশবোর্ড','Courier API configuration':'কুরিয়ার API কনফিগারেশন','Courier name':'কুরিয়ারের নাম','API Base URL':'API বেস URL','API Key':'API কী','Secret Key':'সিক্রেট কী','Fraud Check':'ফ্রড চেক','Excel রিপোর্ট':'Excel report','Excel ডাউনলোড':'Download Excel','স্ট্যাটাস সেভ হয়নি':'Status could not be saved','নতুন ইউজারনেম':'New username','পাসওয়ার্ড':'Password','পাসওয়ার্ড আবার লিখুন':'Confirm password','ইউজার পাসওয়ার্ড পেতে IT এডমিন এর সাথে যোগাযোগ করুন':'Contact the IT Admin to get your user password','প্রোডাক্ট স্টক':'Product stock','ছবি মিলিয়ে দেখছি…':'Matching image…','লোড হচ্ছে…':'Loading…','আগের':'Previous','পরের':'Next','অর্ডার দেখুন':'View order','রিটার্নের কারণ':'Return reasons','কার কত রিটার্ন':'Returns by assignee','সম্পূর্ণ ব্যাকআপ':'Full backup','অ্যাপ ডেটাবেস':'App database','স্টক ফাইল':'Stock file','হিস্ট্রি ও কমেন্ট':'History and comments','ইউজার ও সেটিংস':'Users and settings','ডাউনলোড করুন':'Download','প্রথম Excel রিপোর্ট আপলোড করুন।':'Upload the first Excel report.'
};
Object.assign(UI_EN,{
 'স্টক ম্যানেজমেন্ট':'Stock Management','ড্যাশবোর্ডে ফিরুন':'Back to Dashboard','সার্চ':'Search','নাম / ফোন / কোড / কনসাইনমেন্ট':'Name / Phone / Code / Consignment','নাম / ফোন / Product Code / Cons. ID / কারণ':'Name / Phone / Product Code / Cons. ID / Reason','শুধু স্টকে আছে':'In stock only','স্টক নেই':'Out of stock','সাইজ বাছাই করুন':'Choose a size','ক্যাটাগরি':'Category','সব ক্যাটাগরি':'All categories','সর্বশেষ':'Latest','দাম: কম থেকে বেশি':'Price: low to high','দাম: বেশি থেকে কম':'Price: high to low','নতুন প্রোডাক্ট':'New products','প্রোডাক্ট লোড হচ্ছে…':'Loading products…','প্রোডাক্ট প্রস্তুত':'Products ready','ক্যাশ থেকে প্রোডাক্ট প্রস্তুত':'Products loaded from cache','প্রোডাক্টের ছবি এখনো লোড হচ্ছে, একটু পরে চেষ্টা করুন':'Product images are still loading. Please try again shortly.','এখনো পাওয়া যায়নি':'No matches found yet','কমেন্ট':'Comment','খুলুন':'Open','মুছুন':'Delete','এডিট':'Edit','যোগ করুন':'Add','ইউজার যোগ করুন':'Add user','ইউজার আপডেট করুন':'Update user','ইউজার মুছে ফেলুন':'Delete user','ইউজারনেম':'Username','নাম':'Name','ভূমিকা':'Role','অনুমতির তালিকা':'Permission list','অনুমতি সেভ করুন':'Save permissions','অনুমতি আপডেট হয়েছে':'Permissions updated','আপনি':'You are','হিসেবে লগইন আছেন':'signed in as','শুধু অ্যাডমিন মুছতে পারবেন':'Only an administrator can delete this','সব Activity':'All Activity','আজ':'Today','তারিখ':'Date','ইউজার':'User','অ্যাকশন':'Action','বিস্তারিত':'Details','ফাইল':'File','শিট':'Sheet','ট্যাব':'Tab','কোড':'Code','পরিবর্তন':'Change','পুরোনো':'Previous','নতুন':'New','নোট':'Note','আপলোড':'Upload','আপলোড করেছেন':'Uploaded by','আকার':'Size','এখনো কোনো পরিবর্তন নেই':'No changes yet','কোনো আপলোড নেই':'No uploads','কোনো বিক্রির কমেন্ট নেই':'No sales comments','বিক্রির কমেন্ট':'Sales comments','বিক্রির কমেন্ট — এক নজরে':'Sales comments at a glance','সবাই কমেন্ট দেখতে পারবে; মুছতে পারবেন শুধু অ্যাডমিন।':'Everyone can view comments; only admins can delete them.','প্রোডাক্টে কমেন্ট':'Products with comments','মোট কমেন্ট':'Total comments','মোট বিক্রি (পিস)':'Total sold (pieces)','সাম্প্রতিক কমেন্ট দেখুন':'View recent comments','ফাইল আপলোডের ইতিহাস':'File upload history','স্টক পরিবর্তনের ইতিহাস':'Stock change history','কে, কখন, কোন শিটের কোন পণ্যের পরিমাণ বদলেছেন—সব এখানে রেকর্ড থাকে।':'This records who changed a product quantity, when, and on which sheet.','ড্যাশবোর্ডে ফিরুন':'Back to dashboard','Excel রিপোর্ট ডাউনলোড':'Download Excel report','Excel রিপোর্ট':'Excel report','Return Excel আপলোড':'Upload Return Excel','রিটার্ন পাওয়া অ্যাসাইনি':'Assignees with returns','আলাদা কারণ':'Distinct reasons','Return record নেই। Return Excel আপলোড করুন।':'No return records. Upload a Return Excel file.','মোট রিটার্ন রেকর্ড':'Total return records','অর্ডার রেকর্ড নেই। প্রথম Excel রিপোর্ট আপলোড করুন।':'No order records. Upload the first Excel report.','অর্ডার দেখুন':'View order','সব অর্ডার ডিলিট':'Delete all orders','সিলেক্ট':'Select','ক্রমিক':'No.','তারিখ':'Date','কাস্টমারের নাম':'Customer name','ঠিকানা':'Address','পরিশোধ':'Paid','বাকি/Collectible':'Due / Collectible','ডেলিভারি':'Delivery','নোট':'Note','সেভ হয়নি':'Could not save','ডিজাইন ও ভাষা':'Design and language','অ্যাকসেন্ট':'Accent','ব্যাকগ্রাউন্ড':'Background','হেডার':'Header','কার্ড':'Card','ওয়েবসাইটের ভাষা':'Website language','Professional English':'Professional English','বাংলা':'Bengali','সেভ হয়েছে':'Saved','রিসেট প্রিভিউ':'Reset preview','পণ্য':'Product','ছবি':'Image','মোট':'Total','মূল্য':'Value','রিটার্নের কারণ':'Return reason','কার কত রিটার্ন':'Returns by assignee','অপেক্ষমাণ':'Pending','ফেরত':'Returned','ডেলিভারড':'Delivered','চলমান':'In progress','বাতিল':'Cancelled','যাচাই করুন':'Check','সেভ পরিবর্তন':'Save changes','পরিবর্তন সেভ করুন':'Save changes','স্টক সেভ হয়েছে':'Stock saved','স্টক হিস্ট্রি দেখার অনুমতি নেই':'You do not have permission to view stock history','এই কোডের প্রোডাক্ট ক্যাটালগে পাওয়া যায়নি':'This product code was not found in the catalog','এই ফিচার ব্যবহার করতে ইউজার দিয়ে লগইন করুন':'Log in with your user account to use this feature','লগইন করুন':'Log in','ইউজারনেম ও পাসওয়ার্ড পেতে IT এডমিনের সাথে যোগাযোগ করুন।':'Contact the IT Admin for a username and password.','অনুমতি নেই':'Permission denied','হ্যাঁ':'Yes','না':'No','বন্ধ':'Close','পিছনে':'Back','পরবর্তী':'Next','আগের':'Previous','ধন্যবাদ':'Thank you'
});
Object.assign(UI_EN,{
 'অ্যাকাউন্ট':'Account','অ্যাডমিন প্যানেল':'Admin Panel','আমার অ্যাকাউন্ট':'My Account','লগইন':'Sign In','লগআউট':'Sign Out','দাম এডিট':'Price Editor','Excel স্টক ড্যাশবোর্ড':'Excel Stock Dashboard','ইউজার ম্যানেজমেন্ট':'User Management','ড্যাশবোর্ড':'Dashboard','সব Activity':'All Activity','রিপোর্ট':'Reports','পাসওয়ার্ড':'Password','কুরিয়ার ট্র্যাকিং':'Courier Tracking','কুরিয়ার ট্র্যাকিং':'Courier Tracking','স্টক ম্যানেজমেন্ট':'Stock Management','স্টক এন্ট্রি':'Stock Entry','অর্ডার হিস্ট্রি':'Order History','রিটার্ন হিস্ট্রি':'Return History','নোটিশ':'Notices','ছবি ডাউনলোড':'Photo Download','ডিজাইন':'Design','ব্যাকআপ':'Backup','গেস্ট/কাস্টমার অ্যাক্সেস':'Guest / Customer Access','POS অর্ডার ও মেমো':'POS Orders & Memos','স্টক':'Stock','প্রডাক্ট স্টক':'Product Stock','গুগল শিট':'Google Sheets','ক্যাটালগ':'Catalog','শিট':'Sheet','পাসওয়ার্ড পরিবর্তন':'Change Password','কোনো সাইজ নেই (শিট যুক্ত আছে কিনা দেখুন)':'No sizes available (check whether a sheet is connected)','বন্ধ করুন':'Close','সব সিলেক্ট / বাতিল':'Select All / Clear','HD ডাউনলোড (সিলেক্ট করা)':'HD Download (Selected)','ছবি কপি করুন':'Copy Images','সিলেকশন':'Selection','স্বাগতম':'Welcome','অ্যাডমিন প্যানেল':'Admin Panel','আমার অ্যাকাউন্ট':'My Account'
});
Object.assign(UI_EN,{
 'সব':'All','সব ধরন':'All types','ডিফল্ট':'Default','সিলেক্ট':'Selected','সিলেক্ট করা':'Selected','অ্যাসাইনি':'Assignee','নাম / ফোন / কোড / কনসাইনমেন্ট':'Name / Phone / Code / Consignment','কোড / সাইজ খুঁজুন':'Search by code or size','কোড বা নাম খুঁজুন':'Search by code or name','প্রোডাক্ট লোড হচ্ছে…':'Loading products…','প্রোডাক্ট প্রস্তুত':'Products ready','প্রোডাক্ট নেই':'No products','এই ট্যাবে কোনো সারি নেই':'No rows in this sheet','অর্ডার দেখুন':'View order','অর্ডার আপলোড':'Upload orders','Return Excel আপলোড করুন; মূল Order History-এর সাথে Cons. ID / Product Code দিয়ে তুলনা করুন।':'Upload a Return Excel file and compare it with Order History by Cons. ID or Product Code.','আজকের Excel আপলোড':'Upload today’s Excel file','Excel রিপোর্ট ডাউনলোড':'Download Excel report','ডেলিভারি স্ট্যাটাস':'Delivery status','কুরিয়ার চার্জ':'Courier charge','পণ্যের দাম':'Product price','পরিশোধ':'Advance paid','বাকি/Collectible':'Balance / Collectible','ঠিকানা':'Address','নোট':'Note','ক্রমিক':'No.','তারিখ':'Date','ছবি':'Image','ফোন':'Phone','সাইজ':'Size','Source':'Source','Assigned To':'Assigned To','Payment Status':'Payment Status','Order Status':'Order Status','Delivery Status':'Delivery Status','Cons ID':'Cons. ID','না':'No','দেখুন':'View','পরিবর্তন':'Changes','সেভ না হওয়া':'Unsaved','সেভ না হওয়া পরিবর্তন':'Unsaved changes','নতুন সংখ্যা':'New quantity','বর্তমান স্টক গোপন':'Current stock is hidden','স্টক দেখার অনুমতি নেই':'You do not have permission to view stock','স্টক মূল্য দেখার অনুমতি নেই':'You do not have permission to view stock value','দাম দেখার অনুমতি নেই':'You do not have permission to view prices','শুধু Admin courier API settings বদলাতে পারবে':'Only an admin can change courier API settings','Courier API settings সেভ হয়েছে':'Courier API settings saved','ডাটা নেই':'No data','কোনো রেকর্ড নেই':'No records','কোনো courier tracking record নেই':'No courier tracking records','সার্চ করতে':'Search','সেভ হয়েছে':'Saved','সেভ হয়নি':'Could not save','সেটিংস সেভ':'Save settings','ডাউনলোড':'Download','আপলোড':'Upload','মুছুন':'Delete','বন্ধ':'Off','চালু':'On','হ্যাঁ':'Yes'
});
Object.assign(UI_EN,{
 'ডিজাইন ও ভাষা':'Design & Language','রেডিমেড টেমপ্লেট':'Ready-made themes','অ্যাকসেন্ট':'Accent','ব্যাকগ্রাউন্ড':'Background','লেখা':'Text','হেডার':'Header','কার্ড':'Card','সাইটের নাম':'Website name','ট্যাগলাইন':'Tagline','নোটিশ লিখুন...':'Write a notice…','নোটিশ চালু থাকবে':'Enable notice','নোটিশ কত সেকেন্ড থাকবে':'Notice duration in seconds','নোটিশ সেভ করুন':'Save notice','ছবিতে Product Code/Price দেখানো এবং কারা ছবি ডাউনলোড করতে পারবে—এখান থেকে নিয়ন্ত্রণ করুন। সার্ভারেও অনুমতি যাচাই হবে।':'Choose whether downloads show the Product Code or price and who can download images. Permissions are also checked on the server.','ডাউনলোড ছবিতে CODE দেখাবে':'Show CODE on downloaded images','ডাউনলোড ছবিতে PRICE দেখাবে':'Show PRICE on downloaded images','কারা ছবি ডাউনলোড করতে পারবে?':'Who can download images?','সব ইউজার (বর্তমান permission সাপেক্ষে)':'All users (subject to current permissions)','শুধু লগইন করা ইউজার':'Logged-in users only','শুধু Admin / Super Admin':'Admins and Super Admins only','নতুন পাসওয়ার্ড কমপক্ষে ৬ অক্ষর দিন':'The new password must be at least 6 characters','নতুন পাসওয়ার্ড দুইবার একই নয়':'The new passwords do not match','Admin পাসওয়ার্ড পরিবর্তন হয়েছে':'Admin password changed','Admin পাসওয়ার্ড পরিবর্তন হয়নি':'Admin password was not changed','দাম সেভ হয়নি':'Price could not be saved','দাম সেভ হয়েছে':'Price saved','আসল দাম ফিরেছে':'Original price restored','সাইট নোটিশ':'Site notice','নোটিশ ও প্রদর্শনের সময় ঠিক করুন।':'Set the notice and how long it appears.','ইউজার পাসওয়ার্ড পেতে IT এডমিন এর সাথে যোগাযোগ করুন':'Contact your IT Admin to get your login credentials','আপনার অ্যাকাউন্টে লগইন করুন':'Sign in to your account','ইউজারনেম বা পাসওয়ার্ড সঠিক নয়':'Incorrect username or password','শুধু অ্যাডমিন courier API settings বদলাতে পারবে':'Only admins can change courier API settings','স্টক সিঙ্ক হচ্ছে…':'Syncing stock…','শিট যুক্ত নেই':'No sheet connected','অ্যাডমিন থেকে গুগল শিট যোগ করলে রিয়েল স্টক আসবে':'Connect a Google Sheet in Admin to load live stock','ট্যাব সমস্যা':'Tab issues','AI ছবি-সার্চ চালু':'AI image search is ready','AI শুরু হচ্ছে…':'Starting AI…','স্টক দেখার permission বন্ধ আছে':'Stock viewing is disabled for this account','এই সার্চ permission বন্ধ আছে':'Search is disabled for this account','এই ফিচার ব্যবহার করতে ইউজার দিয়ে লগইন করুন':'Sign in to use this feature','কোনো courier tracking record নেই':'No courier tracking records','এই কোডের প্রোডাক্ট ক্যাটালগে পাওয়া যায়নি':'This product code was not found in the catalog','আপনার রিপোর্ট permission নেই':'You do not have permission to report this item','রিপোর্ট Admin-এর কাছে গেছে':'Report sent to the admin','রিপোর্ট করা যায়নি':'Could not submit the report'
});
const UI_BN={};
const HAS_BENGALI=s=>/[\u0980-\u09FF]/u.test(String(s??''));
for(const [k,v] of Object.entries(UI_EN)){
 if(HAS_BENGALI(k)&&!HAS_BENGALI(v))UI_BN[v]=k;
 else if(!HAS_BENGALI(k)&&HAS_BENGALI(v))UI_BN[k]=v;
}
Object.assign(UI_BN,{
 'Consignment ID / Customer Number':'ConsID / Phone',
 'Search orders':'অর্ডার খুঁজুন','Order search':'অর্ডার সার্চ','Search':'সার্চ',
 'POS Sales (F9)':'POS সেলস (F9)','Notifications':'নোটিফিকেশন',
 'Dashboard, filter, Excel report and memo PDF —':'ড্যাশবোর্ড, ফিল্টার, Excel রিপোর্ট ও মেমো PDF —',
 'Type a product code or name to search.':'প্রোডাক্ট কোড বা নাম লিখে খুঁজুন।',
 'Products could not be loaded':'প্রোডাক্ট লোড করা যায়নি',
 'No products found':'কোনো প্রোডাক্ট পাওয়া যায়নি',
 'No products are available right now.':'এই মুহূর্তে কোনো প্রোডাক্ট নেই।',
 'No products match your search.':'আপনার সার্চের সঙ্গে কোনো প্রোডাক্ট মেলেনি।',
 'Try another keyword or clear the current filters.':'অন্য কীওয়ার্ড দিন অথবা বর্তমান ফিল্টার পরিষ্কার করুন।',
 'Clear search and filters':'সার্চ ও ফিল্টার পরিষ্কার করুন',
 'Loading products…':'প্রোডাক্ট লোড হচ্ছে…',
 'POS is still loading. Please try again.':'POS এখনও লোড হচ্ছে। আবার চেষ্টা করুন।',
 'Consignment ID':'ConsID','Customer Number':'Customer Number',
 'Enter Consignment ID':'ConsID লিখুন','Enter phone number':'Phone Numbar',
 'Order History Record':'অর্ডার হিস্ট্রি রেকর্ড','Collectible Amount':'আদায়যোগ্য টাকা',
 'ERP Management System':'ERP Management System','Alien Tees ERP Management System':'Alien Tees ERP Management System','ERP MANAGEMENT SYSTEM':' Smart ERP Management System','ALIEN TEES ERP · SALES':'ALIEN TEES ERP · বিক্রয়','Welcome.':'স্বাগতম।',
 'Incorrect username or password':'ইউজারনেম অথবা পাসওয়ার্ড সঠিক নয়',
 'Everything you need, in one place.':'আপনার প্রয়োজনীয় সবকিছু এক জায়গায়।',
 'Stock saved':'স্টক সেভ হয়েছে','Status':'স্ট্যাটাস','Date':'তারিখ',
 'Close':'বন্ধ করুন','Reset':'রিসেট','Today':'আজ',
 'All Products':'সব প্রোডাক্ট','New Arrival':'নতুন পণ্য'
});
Object.assign(UI_BN,{
 'ALIEN · ADMIN':'ALIEN · অ্যাডমিন','My company':'আমার প্রতিষ্ঠান','Name *':'প্রতিষ্ঠানের নাম *','Email':'ইমেইল','Website':'ওয়েবসাইট',
 'Tax / TIN number':'ট্যাক্স / টিআইএন নম্বর','Registration number':'রেজিস্ট্রেশন নম্বর','Delivery Charges (BDT)':'ডেলিভারি চার্জ (টাকা)',
 'Set your default delivery rate for each destination. Existing order and stock data will not be changed.':'প্রতিটি এলাকার ডেলিভারি চার্জ নির্ধারণ করুন। আগের অর্ডার ও স্টকের তথ্য পরিবর্তন হবে না।',
 'Inside Dhaka':'ঢাকার ভেতরে','Dhaka Suburbs (Savar / Gazipur / Narayanganj)':'ঢাকার আশপাশ (সাভার / গাজীপুর / নারায়ণগঞ্জ)','Outside Dhaka':'ঢাকার বাইরে',
 'Logo':'লোগো','PNG, JPG, or WebP up to 2MB.':'সর্বোচ্চ ২ এমবি PNG, JPG বা WebP ছবি।','Save changes':'পরিবর্তন সংরক্ষণ',
 'Purchase No.':'ক্রয় নম্বর','Purchase Date':'ক্রয়ের তারিখ','Supplier':'সরবরাহকারী','Supplier invoice':'সরবরাহকারীর ইনভয়েস',
 'Supplier invoice, received item quantities and payment summary':'সরবরাহকারীর ইনভয়েস, পণ্য গ্রহণের পরিমাণ ও পেমেন্টের হিসাব',
 'Purchase Management':'পণ্য ক্রয় ব্যবস্থাপনা','Purchase History':'ক্রয়ের ইতিহাস','Supplier List':'সরবরাহকারীর তালিকা','Supplier Management':'সরবরাহকারী ব্যবস্থাপনা',
 'Add Supplier':'সরবরাহকারী যোগ করুন','Supplier / Company Name *':'সরবরাহকারী / প্রতিষ্ঠানের নাম *','Opening Due (BDT)':'প্রারম্ভিক বকেয়া (টাকা)',
 'Supplier contacts, opening balance, purchase history and outstanding amounts.':'সরবরাহকারীর যোগাযোগ, প্রারম্ভিক বকেয়া, ক্রয়ের ইতিহাস ও বর্তমান বকেয়ার তথ্য।',
 'Active Suppliers':'সক্রিয় সরবরাহকারী','Total Purchase Value':'মোট ক্রয়মূল্য','Total Paid':'মোট পরিশোধ','Current Due':'বর্তমান বকেয়া',
 'Search purchase / supplier / invoice':'ক্রয় / সরবরাহকারী / ইনভয়েস খুঁজুন','Search':'খুঁজুন','Reset':'রিসেট','Date':'তারিখ',
 'Invoice Ref.':'ইনভয়েস রেফারেন্স','Items / Qty':'পণ্য / পরিমাণ','Total':'মোট','Paid':'পরিশোধ','Due':'বকেয়া','Stock':'স্টক','Action':'কাজ',
 'No purchase records yet.':'এখনো কোনো ক্রয়ের তথ্য নেই।','No purchase history recorded for this supplier yet.':'এই সরবরাহকারীর কোনো ক্রয় ইতিহাস নেই।',
 'PDF Memo':'PDF মেমো','Download Excel':'এক্সেল ডাউনলোড','Close':'বন্ধ করুন',
 'PREPARED BY':'প্রস্তুতকারীর স্বাক্ষর','AUTHORIZED BY':'অনুমোদনকারীর স্বাক্ষর','RECEIVED BY':'গ্রহণকারীর স্বাক্ষর',
 'PURCHASE MEMO':'ক্রয় মেমো','PURCHASED ITEMS':'ক্রয়কৃত পণ্য','PRODUCT CODE':'পণ্যের কোড','PRODUCT NAME':'পণ্যের নাম',
 'SIZE / VARIANT':'সাইজ / ভ্যারিয়েন্ট','QTY':'পরিমাণ','UNIT COST':'প্রতি পণ্যের দাম','LINE TOTAL':'সর্বমোট','INVOICE TOTAL':'ইনভয়েসের মোট',
 'BALANCE DUE':'বকেয়া টাকা','Paid to Date':'এ পর্যন্ত পরিশোধ','PAYMENT RECORD':'পেমেন্টের বিবরণ','NOTES':'নোট',
 'Expense Management':'খরচ ব্যবস্থাপনা','Expense History':'খরচের ইতিহাস','New Expense':'নতুন খরচ','Expense Title *':'খরচের বিবরণ *',
 'Amount (BDT) *':'পরিমাণ (টাকা) *','Expense Date *':'খরচের তারিখ *','Payment Method':'পেমেন্টের মাধ্যম',
 'Paid To / Vendor':'যাকে পরিশোধ করা হয়েছে / বিক্রেতা','Reference / Receipt No.':'রেফারেন্স / রসিদ নম্বর',
 'New Purchase':'নতুন ক্রয়','Stock Sheet *':'স্টক শিট *','Product Code *':'পণ্যের কোড *','Size / Variant *':'সাইজ / ভ্যারিয়েন্ট *',
 'Qty *':'পরিমাণ *','Unit Cost (BDT) *':'প্রতি পণ্যের দাম (টাকা) *','Add Item':'পণ্য যোগ করুন','Invoice Total':'ইনভয়েসের মোট',
 'Due After Initial Payment':'প্রাথমিক পেমেন্টের পর বকেয়া','Save Purchase':'ক্রয় সংরক্ষণ','Supplier Report':'সরবরাহকারীর রিপোর্ট',
 'Purchase Excel':'ক্রয়ের এক্সেল','Expense Excel':'খরচের এক্সেল','Apply Filter':'ফিল্টার প্রয়োগ','All Dates':'সব তারিখ','Purchase Total':'মোট ক্রয়',
 'Supplier Payments':'সরবরাহকারীকে পরিশোধ','Supplier Due':'সরবরাহকারীর বকেয়া','Approved Expenses':'অনুমোদিত খরচ','Pending Expenses':'অপেক্ষমাণ খরচ',
 'Net Cash Outflow':'মোট নগদ ব্যয়','Expenses by Category':'ধরন অনুযায়ী খরচ','Supplier Outstanding':'সরবরাহকারীর বকেয়া',
 'Purchase Records · Pay Due / Delete':'ক্রয়ের তথ্য · বকেয়া পরিশোধ / মুছুন','Expense Records · Delete':'খরচের তথ্য · মুছুন',
 'Product Reports':'পণ্যের রিপোর্ট','Price Change History':'দাম পরিবর্তনের ইতিহাস','Current Password':'বর্তমান পাসওয়ার্ড',
 'POS Orders & Memo':'POS অর্ডার ও মেমো','Courier Tracking':'কুরিয়ার ট্র্যাকিং','Dashboard':'ড্যাশবোর্ড','Stock Management':'স্টক ব্যবস্থাপনা',
 'Catalog':'ক্যাটালগ','Price Editor':'দাম সম্পাদনা','User Management':'ইউজার ব্যবস্থাপনা','Activity Logs':'কার্যক্রমের ইতিহাস',
 'Reports':'রিপোর্ট','Notices':'নোটিশ','Design Settings':'ডিজাইন সেটিংস','Administration & Settings':'প্রশাসন ও সেটিংস',
 'Sales & Operations':'বিক্রয় ও কার্যক্রম','Stock & Catalog':'স্টক ও ক্যাটালগ','View':'দেখুন','PDF Memo':'PDF মেমো'
});
Object.assign(UI_BN,{
 'View':'দেখুন','Receive Stock':'স্টক গ্রহণ করুন','Pay Due':'বকেয়া পরিশোধ','Save Purchase':'ক্রয় সংরক্ষণ','Save Supplier':'সরবরাহকারী সংরক্ষণ',
 'Save Expense':'খরচ সংরক্ষণ','Record Payment':'পেমেন্ট রেকর্ড করুন','Record Payment':'পেমেন্ট সংরক্ষণ','Activate':'সক্রিয় করুন','Deactivate':'নিষ্ক্রিয় করুন',
 'Active':'সক্রিয়','Inactive':'নিষ্ক্রিয়','Received':'স্টক গ্রহণ করা হয়েছে','Not received':'স্টক গ্রহণ করা হয়নি','Saving…':'সংরক্ষণ হচ্ছে…',
 'Select a supplier first.':'আগে সরবরাহকারী নির্বাচন করুন।','Choose the purchase date.':'ক্রয়ের তারিখ নির্বাচন করুন।',
 'Enter product code, valid quantity and unit cost on every line.':'প্রতিটি সারিতে পণ্যের কোড, সঠিক পরিমাণ ও একক মূল্য দিন।',
 'For automatic stock update, every purchase line needs Stock Sheet, Product Code and Size.':'স্টক স্বয়ংক্রিয়ভাবে আপডেট করতে প্রতিটি সারিতে স্টক শিট, পণ্যের কোড ও সাইজ দিতে হবে।',
 'Excel stock workbook is not available. Set up the stock workbook before recording a purchase.':'Excel স্টক ফাইল পাওয়া যায়নি। ক্রয় এন্ট্রির আগে স্টক ফাইল সেটআপ করুন।',
 'Save this purchase invoice and payment details?':'এই ক্রয় ইনভয়েস ও পেমেন্টের তথ্য সংরক্ষণ করবেন?',
 'Purchase saved and Excel stock updated. A PDF Memo button is available in Purchase History.':'ক্রয় সংরক্ষিত হয়েছে এবং Excel স্টক আপডেট হয়েছে। ক্রয়ের ইতিহাস থেকে PDF মেমো দেখা যাবে।',
 'Purchase saved successfully. A PDF Memo button is available in Purchase History.':'ক্রয় সংরক্ষণ হয়েছে। ক্রয়ের ইতিহাস থেকে PDF মেমো দেখা যাবে।',
 'Purchase could not be saved.':'ক্রয় সংরক্ষণ করা যায়নি।','Select supplier':'সরবরাহকারী নির্বাচন করুন','Supplier Name':'সরবরাহকারীর নাম',
 'Supplier / Company Name *':'সরবরাহকারী / প্রতিষ্ঠানের নাম *','Phone number':'ফোন নম্বর','Email address':'ইমেইল ঠিকানা','Supplier address':'সরবরাহকারীর ঠিকানা',
 'Optional supplier notes':'সরবরাহকারী সম্পর্কে অতিরিক্ত নোট','Opening Due':'প্রারম্ভিক বকেয়া','Purchases':'ক্রয়','Payments':'পেমেন্ট',
 'Purchase invoice':'ক্রয় ইনভয়েস','Invoice No.':'ইনভয়েস নম্বর','Invoice Reference':'ইনভয়েস রেফারেন্স','Purchase Date *':'ক্রয়ের তারিখ *',
 'Supplier *':'সরবরাহকারী *','Initial Payment (BDT)':'প্রাথমিক পরিশোধ (টাকা)','Payment Method':'পেমেন্টের মাধ্যম','Payment Reference':'পেমেন্ট রেফারেন্স',
 'Cash':'নগদ','Bank Transfer':'ব্যাংক ট্রান্সফার','Notes':'নোট','Optional notes':'ঐচ্ছিক নোট','Subtotal':'উপমোট','Discount':'ছাড়',
 'Received quantities will be added to stock automatically':'গ্রহণ করা পরিমাণ স্বয়ংক্রিয়ভাবে স্টকে যোগ হবে','No suppliers added.':'কোনো সরবরাহকারী যোগ করা হয়নি।',
 'No active suppliers found.':'কোনো সক্রিয় সরবরাহকারী পাওয়া যায়নি।','Search supplier / phone':'সরবরাহকারী / ফোন খুঁজুন','Search supplier':'সরবরাহকারী খুঁজুন',
 'Supplier Report':'সরবরাহকারীর রিপোর্ট','Supplier List':'সরবরাহকারীর তালিকা','Purchase History —':'ক্রয়ের ইতিহাস —','Purchase / Invoice':'ক্রয় / ইনভয়েস',
 'Product Details':'পণ্যের বিবরণ','Quantity Purchased':'ক্রয় করা পরিমাণ','Created By':'তৈরি করেছেন','Invoice Due':'ইনভয়েসের বকেয়া',
 'Download Excel':'এক্সেল ডাউনলোড','Supplier-Purchase-History':'সরবরাহকারীর-ক্রয়-ইতিহাস','Could not load supplier purchase history.':'সরবরাহকারীর ক্রয়ের ইতিহাস লোড করা যায়নি।',
 'No item rows recorded':'কোনো পণ্য যোগ করা হয়নি','No records found.':'কোনো তথ্য পাওয়া যায়নি।','No purchase records for this period.':'এই সময়সীমায় কোনো ক্রয় নেই।',
 'Expense Title *':'খরচের বিবরণ *','Category *':'খরচের ধরন *','Amount (BDT) *':'পরিমাণ (টাকা) *','Expense Date *':'খরচের তারিখ *',
 'Paid To / Vendor':'যাকে পরিশোধ করা হয়েছে / বিক্রেতা','Reference / Receipt No.':'রেফারেন্স / রসিদ নম্বর','Status':'অবস্থা',
 'Pending Approval':'অনুমোদনের অপেক্ষায়','Approved / Paid':'অনুমোদিত / পরিশোধিত','Purpose, extra information or receipt details':'উদ্দেশ্য, অতিরিক্ত তথ্য বা রসিদের বিবরণ',
 'Office Rent':'অফিস ভাড়া','Salary & Wages':'বেতন ও মজুরি','Utilities':'ইউটিলিটি বিল','Internet & Software':'ইন্টারনেট ও সফটওয়্যার',
 'Marketing':'মার্কেটিং','Packaging':'প্যাকেজিং','Courier & Delivery':'কুরিয়ার ও ডেলিভারি','Transport':'পরিবহন',
 'Office Supplies':'অফিস সামগ্রী','Purchase Miscellaneous':'অন্যান্য ক্রয় খরচ','Other':'অন্যান্য','Expense No.':'খরচ নম্বর',
 'Expense Records':'খরচের তথ্য','Pending Approval':'অনুমোদনের অপেক্ষায়','Approved':'অনুমোদিত','Total Approved':'মোট অনুমোদিত খরচ',
 'Pending Approval':'অনুমোদনের অপেক্ষায়','Pending Expenses':'অপেক্ষমাণ খরচ','Expense Records':'খরচের তথ্য','This Month':'এই মাস',
 'Search title / category / vendor':'বিবরণ / ধরন / বিক্রেতা খুঁজুন','No expense records yet.':'এখনো কোনো খরচের তথ্য নেই।',
 'No expense data for this period.':'এই সময়সীমায় কোনো খরচের তথ্য নেই।','Supplier Outstanding':'সরবরাহকারীর বকেয়া','No supplier balances.':'সরবরাহকারীর বকেয়া নেই।',
 'Apply Filter':'ফিল্টার প্রয়োগ','All Dates':'সব তারিখ','From':'শুরুর তারিখ','To':'শেষ তারিখ','Net Cash Outflow':'মোট নগদ ব্যয়',
 'Delete':'মুছুন','Edit':'সম্পাদনা','Cancel':'বাতিল','Save':'সংরক্ষণ','Add':'যোগ করুন','Update':'আপডেট করুন','Loading…':'লোড হচ্ছে…',
 'Courier Tracking':'কুরিয়ার ট্র্যাকিং','Courier':'কুরিয়ার','Order Status':'অর্ডারের অবস্থা','Delivery Status':'ডেলিভারির অবস্থা','Consignment ID':'কনসাইনমেন্ট আইডি',
 'Tracking':'ট্র্যাকিং','Track Order':'অর্ডার ট্র্যাক করুন','Search Order':'অর্ডার খুঁজুন','Search by phone':'ফোন দিয়ে খুঁজুন',
 'Order History':'অর্ডার ইতিহাস','Return History':'রিটার্ন ইতিহাস','Customer':'কাস্টমার','Customer Name':'কাস্টমারের নাম','Customer Phone':'কাস্টমারের ফোন',
 'Address':'ঠিকানা','Product Price':'পণ্যের দাম','Delivery Charge':'ডেলিভারি চার্জ','Advance Paid':'অগ্রিম পরিশোধ','Collectible Amount':'আদায়যোগ্য টাকা',
 'Assigned To':'দায়িত্বপ্রাপ্ত','No matching order found in Order Sheet history.':'Order Sheet-এর ইতিহাসে মিলে এমন অর্ডার পাওয়া যায়নি।',
 'Enter Consignment ID or customer number':'কনসাইনমেন্ট আইডি অথবা কাস্টমার নম্বর লিখুন','Searching order details…':'অর্ডারের তথ্য খোঁজা হচ্ছে…',
 'POS Sales':'POS বিক্রি','Sales':'বিক্রয়','Memo':'মেমো','Payment':'পেমেন্ট','Print':'প্রিন্ট','Print Memo':'মেমো PDF / প্রিন্ট করুন',
 'Order Details':'অর্ডারের বিবরণ','Order Date':'অর্ডারের তারিখ','New Order':'নতুন অর্ডার','All Orders':'সব অর্ডার',
 'Product':'পণ্য','Products':'পণ্যসমূহ','Size':'সাইজ','Quantity':'পরিমাণ','Price':'দাম','Amount':'টাকা','Total Amount':'মোট টাকা',
 'No data available':'কোনো তথ্য নেই','No results found':'কোনো ফলাফল পাওয়া যায়নি','No records yet':'এখনো কোনো তথ্য নেই','Refresh':'রিফ্রেশ',
 'Dashboard, filter, Excel report and memo PDF —':'ড্যাশবোর্ড, ফিল্টার, Excel রিপোর্ট ও মেমো PDF —','Excel Report':'এক্সেল রিপোর্ট',
 'Upload':'আপলোড','Upload Excel':'এক্সেল আপলোড','Download':'ডাউনলোড','Search History':'সার্চ ইতিহাস','Filter':'ফিল্টার',
 'Website name':'ওয়েবসাইটের নাম','Tagline':'ট্যাগলাইন','Save settings':'সেটিংস সংরক্ষণ','Save notice':'নোটিশ সংরক্ষণ','Write a notice…':'নোটিশ লিখুন…',
 'Design & Language':'ডিজাইন ও ভাষা','Ready-made themes':'প্রস্তুত থিম','Accent':'অ্যাকসেন্ট','Background':'ব্যাকগ্রাউন্ড','Text':'লেখা','Header':'হেডার',
 'Card':'কার্ড','Change Password':'পাসওয়ার্ড পরিবর্তন','New Password':'নতুন পাসওয়ার্ড','Confirm Password':'পাসওয়ার্ড নিশ্চিত করুন',
 'Activity':'কার্যক্রম','Recent Activity':'সাম্প্রতিক কার্যক্রম','User':'ইউজার','Username':'ইউজারনেম','Role':'ভূমিকা','Permission':'অনুমতি',
 'Guest / Customer Access Control':'গেস্ট / কাস্টমার অ্যাক্সেস নিয়ন্ত্রণ','Guest Access':'গেস্ট অ্যাক্সেস','Backup':'ব্যাকআপ','Restore':'পুনরুদ্ধার',
 'Search product':'পণ্য খুঁজুন','Search by code or name':'কোড বা নাম দিয়ে খুঁজুন','All Products':'সব পণ্য','New Arrival':'নতুন পণ্য','Sort by':'সাজান',
 'Price Low to High':'কম দাম থেকে বেশি','Price High to Low':'বেশি দাম থেকে কম','In Stock':'স্টকে আছে','Out of Stock':'স্টকে নেই',
 'Please try again.':'আবার চেষ্টা করুন।','Could not save':'সংরক্ষণ করা যায়নি','Saved successfully':'সফলভাবে সংরক্ষণ হয়েছে','Failed':'ব্যর্থ','Success':'সফল',
 'Close panel':'প্যানেল বন্ধ করুন','Navigation':'মেনু','Purchase Memo Template':'ক্রয় মেমো টেমপ্লেট','Memo Title':'মেমোর শিরোনাম','Subtitle':'সাবটাইটেল',
 'Memo Footer':'মেমোর নিচের লেখা','Show company logo on PDF memo':'PDF মেমোতে কোম্পানির লোগো দেখান'
});
const uiOriginalText=new WeakMap(),uiOriginalAttr=new WeakMap();let uiLanguage='bn',uiTranslating=false;
function translateUi(root=document.body){if(!root||uiTranslating)return;uiTranslating=true;try{const texts=[];if(root.nodeType===3)texts.push(root);else{const walk=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);let n;while(n=walk.nextNode())texts.push(n)}for(const n of texts){const p=n.parentElement;if(!p||p.closest?.('#dr')||/^(SCRIPT|STYLE|TEXTAREA|INPUT)$/.test(p.tagName))continue;if(!uiOriginalText.has(n))uiOriginalText.set(n,n.nodeValue);const original=uiOriginalText.get(n),trim=original.trim(),label=trim.replace(/^[^\p{L}\p{N}]+/u,'').trim(),prefix=trim.slice(0,trim.length-label.length),match=label.match(/^মিলের হার:\s*(.+)$/),digits=s=>s.replace(/[০-৯]/g,d=>'০১২৩৪৫৬৭৮৯'.indexOf(d)),wordEn=UI_EN[trim]||UI_EN[label],wordBn=UI_BN[trim]||UI_BN[label],en=wordEn?(prefix+wordEn):(match?`Similarity: ${match[1].replace(/[০-৯]/g,d=>'০১২৩৪৫৬৭৮৯'.indexOf(d))}`:null),bnText=wordBn?digits(prefix+wordBn):digits(trim),target=uiLanguage==='en'?(en?original.replace(trim,digits(en)):digits(original)):(trim?original.replace(trim,bnText):original);if(n.nodeValue!==target)n.nodeValue=target}const els=root.nodeType===1?[root,...root.querySelectorAll('*')]:[];for(const el of els){if(el.closest?.('#dr'))continue;for(const a of ['placeholder','title','aria-label']){if(!el.hasAttribute?.(a))continue;let saved=uiOriginalAttr.get(el);if(!saved){saved={};uiOriginalAttr.set(el,saved)}if(!(a in saved))saved[a]=el.getAttribute(a);const orig=saved[a],tr=uiLanguage==='en'?UI_EN[orig?.trim()]:UI_BN[orig?.trim()],digits=s=>s.replace(/[০-৯]/g,d=>'০১২৩৪৫৬৭৮৯'.indexOf(d)),target=orig!=null?digits(tr||orig):null;if(target!=null&&el.getAttribute(a)!==target)el.setAttribute(a,target)}}}finally{uiTranslating=false}}
function setUiLanguage(lang){uiLanguage='bn';document.documentElement.lang='bn';translateUi();}
const PRESETS={'গোল্ড ডার্ক':{acc:'#c9a45c',bg:'#f4f2ee',ink:'#1a1512',head:'#1b1512',card:'#ffffff'},'নাইট':{acc:'#8b7bff',bg:'#0f1020',ink:'#ecebff',head:'#07071a',card:'#1a1b33'},'রোজ':{acc:'#e0457b',bg:'#fff5f8',ink:'#2a1020',head:'#3a0f24',card:'#ffffff'},'ফরেস্ট':{acc:'#9bd17a',bg:'#eef3ec',ink:'#12200f',head:'#10230f',card:'#ffffff'}};
function applyTheme(t){const r=document.documentElement.style;['acc','bg','ink','head','card'].forEach(k=>t[k]&&r.setProperty('--'+k,t[k]));if(t.font)r.setProperty('--f',`'${t.font}','Hind Siliguri',sans-serif`);setUiLanguage(t.language||'en');
 if(t.title)$('ttl').textContent=t.title;if(t.tag)$('tag').textContent=t.tag;$('ttl').hidden=false}
async function loadTheme(){applyTheme(await jf('/api/theme'))}

// ---- ডাটা ও প্রোগ্রেস ----
function saveLocalProductCache(){try{const lean=ALL.map(p=>({key:p.key,id:p.id,title:p.title,store:p.store,price:p.price,old:p.old,img:p.img,link:p.link,extra:p.extra,size_chart_image:p.size_chart_image,size_chart_text:p.size_chart_text,display_id:p.display_id}));localStorage.setItem('stockRoomProducts',JSON.stringify(lean))}catch(e){}}
function loadLocalProductCache(){try{const a=JSON.parse(localStorage.getItem('stockRoomProducts')||'[]');if(Array.isArray(a)&&a.length){ALL=a;MAP={};ALL.forEach(p=>MAP[p.key]=p);metaD={products:ALL.length,ready:true,stage:'ক্যাশ থেকে প্রোডাক্ট প্রস্তুত'};drawTabs();drawSizes();render(1);return true}}catch(e){}return false}
async function loadProducts(){if(!me)return;const j=await jf('/api/products');if((j.products||[]).length){ALL=j.products||[];saveLocalProductCache()}else if(!ALL.length){loadLocalProductCache()}MAP={};ALL.forEach(p=>MAP[p.key]=p);metaD=j.meta||metaD;lastOrderVersion=Number(metaD.order_version||0);lastSig='';drawTabs();drawSizes();render(1)}
const SZO=s=>isNaN(+s)?100+(['XS','S','M','L','XL','XXL','2XL','XXXL','3XL'].indexOf(String(s).toUpperCase())+1||99):+s;
function drawTabs(){const stores=[...new Set(ALL.map(p=>p.store))].filter(s=>String(s).trim().toLowerCase()!=='stock excel'),tabs=[...new Set([...(metaD.tabs||[]),...ALL.flatMap(p=>p.tabs||[])])];
 const b=(t,v,on)=>`<button data-t="${t}" data-v="${esc(v)}" class="${on?'on':''} ${t=='t'?'t2':''}">${esc(v)}</button>`;
 const newCount=ALL.filter(p=>p.is_new).length;
 $('tabs').innerHTML=`<button data-t="all" class="${!store&&!tabf&&!newOnly?'on':''}">সব</button>`+(newCount?`<button data-t="new" class="${newOnly?'on':''}">✨ New Arrival <small>${bn(newCount)}</small></button>`:'')+stores.map(s=>b('s',s,store==s)).join('')+tabs.map(s=>b('t',s,tabf==s)).join('');
 $('ft').innerHTML='<option value="">সব ধরন</option>'+tabs.map(s=>`<option ${tabf==s?'selected':''}>${esc(s)}</option>`).join('')}
function toggleSizeFilters(){if(!canDo('size_filter'))return;sizeFiltersExpanded=!sizeFiltersExpanded;const box=$('dsz'),head=document.querySelector('.size-filter-head'),chev=$('sizeFilterChevron');if(box){box.hidden=!sizeFiltersExpanded;box.style.display=sizeFiltersExpanded?'flex':'none'}if(head)head.setAttribute('aria-expanded',String(sizeFiltersExpanded));if(chev)chev.textContent=sizeFiltersExpanded?'−':'＋'}
function drawSizes(){const sizeOn=canDo('size_filter');if(!sizeOn){szf='';sizeFiltersExpanded=false;$('dsz').innerHTML='';$('dsz').hidden=true;$('dsz').style.display='none';$('szh').textContent='';document.querySelectorAll('[data-perm=size-filter-label]').forEach(e=>e.style.display='none');const h=document.querySelector('.size-filter-head');if(h)h.setAttribute('aria-expanded','false');if($('sizeFilterChevron'))$('sizeFilterChevron').textContent='＋';return}const c={};ALL.filter(p=>(!store||p.store==store)&&(!tabf||(p.tabs||[]).includes(tabf))).forEach(p=>Object.keys(p.sizes||{}).forEach(s=>c[s]=(c[s]||0)+1));const L=Object.keys(c).sort((a,b)=>SZO(a)-SZO(b)||a.localeCompare(b));if(szf&&!c[szf])szf='';$('dsz').innerHTML=L.length?L.map(s=>`<button data-v="${esc(s)}" class="${szf==s?'on':''}">${esc(s)}<small>${bn(c[s])}</small></button>`).join(''):'<span style="color:#999;font-size:13px">কোনো সাইজ নেই (শিট যুক্ত আছে কিনা দেখুন)</span>';$('dsz').hidden=!sizeFiltersExpanded;$('dsz').style.display=sizeFiltersExpanded?'flex':'none';const head=document.querySelector('.size-filter-head');if(head)head.setAttribute('aria-expanded',String(sizeFiltersExpanded));if($('sizeFilterChevron'))$('sizeFilterChevron').textContent=sizeFiltersExpanded?'−':'＋';$('szh').textContent=tabf?'— '+tabf+' ট্যাবের':''}
async function logUiAction(action,query='',details=''){try{await post('/api/activity/ui',{action,query,details})}catch(e){}}
$('tabs').onclick=e=>{const b=e.target.closest('button');if(!b)return;const t=b.dataset.t,v=b.dataset.v;if(t=='new'){newOnly=!newOnly;store='';tabf=''}else if(t=='all'){store='';tabf='';newOnly=false;logUiAction('filter_brand','সব','all')}else if(t=='s'){newOnly=false;store=store==v?'':v;logUiAction('filter_brand',store||'সব',store||'clear')}else {newOnly=false;tabf=tabf==v?'':v;logUiAction('filter_sheet',tabf||'সব',tabf||'clear')}drawTabs();drawSizes();render(1)};
$('dsz').onclick=e=>{const b=e.target.closest('button');if(!b||!canDo('size_filter'))return;szf=szf==b.dataset.v?'':b.dataset.v;logUiAction('filter_size',szf||'সব',szf||'clear');drawSizes();render(1);dr(0)};
async function poll(){const j=await jf('/api/stock'),m=j.meta||{};metaD=m;
 if(m.products!=ALL.length||Number(m.order_version||0)!==lastOrderVersion)await loadProducts();
 else if(j.items){const sg=JSON.stringify(j.items)+JSON.stringify(m.tabs);if(sg!=lastSig){lastSig=sg;ALL.forEach(p=>Object.assign(p,j.items[p.key]||{}));drawTabs();drawSizes();render(2)}}
 $('prog').style.display=m.ready?'none':'block';$('pb').style.width=m.pct+'%';
 $('pt').innerHTML=`${esc(m.stage||'')} — ${bn(m.pct||0)}% লোড হয়েছে · ${bn(m.products||0)}টা প্রোডাক্ট${m.ai_ready?' · AI প্রস্তুত':' · AI ব্যাকগ্রাউন্ডে'}`;
 $('sync').innerHTML=m.synced?(m.has_sheet?'🟢 Live stock · Last sync ':'No sheet connected · ')+bt(bnTime(new Date(m.synced*1000)))+(m.errors&&m.errors.length?'<br>⚠ Tab issues: '+esc(m.errors.join(', ')):''):(m.has_sheet?'Syncing stock…':'Connect a Google Sheet in Admin to load live stock');
 $('sync').innerHTML+='<br>'+(m.ai?'🤖 AI image search ready':(m.ai_busy?'⏳ Preparing AI in the background':(m.ai_err?'⚠ AI is unavailable; its model must be set up on the server':'⏳ Starting AI…')));
 aiOn=!!m.ai;dzTxt();
 setTimeout(poll,m.ready&&m.ai_ready?20000:(m.ready?2500:500))}

// ---- দেখানো ----
function render(reset){const q=$('q').value.trim().toLowerCase(),fk=$('fk').value,so=$('so').value; if(!canDo('view_stock')) $('fk').value='';
 let l;if(res){const sm=new Set(res.map(r=>r.key));l=res.map(r=>MAP[r.key]&&Object.assign({},MAP[r.key],{sc:r.score})).filter(Boolean).sort((a,b)=>Number(b.sc||0)-Number(a.sc||0)).concat(ALL.filter(p=>!sm.has(p.key)))}else l=ALL.slice();
 if(newOnly)l=l.filter(p=>p.is_new);if(store)l=l.filter(p=>p.store==store);if(tabf)l=l.filter(p=>(p.tabs||[]).includes(tabf));
 if(q&&!res)l=l.filter(p=>(p.title+' '+p.id+' '+(p.display_id||'')).toLowerCase().includes(q));
 if(szf)l=l.filter(p=>p.sizes&&p.sizes[szf]);if(fk)l=l.filter(p=>fk=='1'?p.instock:p.instock===false);
 if(so&&!res)l.sort((a,b)=>so=='a'?a.price-b.price:b.price-a.price);
 if(res&&l[0]&&l[0].sc!=null)l[0].top=1;
 list=l;const keep=reset==2?shown:0;if(reset){shown=0;$('grid').innerHTML=''}
 $('info').innerHTML=(res?'<button class="pill" onclick="res=null;dzReset();render(1)">✕ Clear image search</button> ':'')+(tabf?`<button class="pill" onclick="tabf='';drawTabs();drawSizes();render(1)">${esc(tabf)} ✕</button> `:'')+(newOnly?`<button class="pill" onclick="newOnly=false;drawTabs();render(1)">New Arrivals ✕</button> `:'')+(szf?`<button class="pill" onclick="szf='';drawSizes();render(1)">Size ${esc(szf)} ✕</button> `:'')+`${bp(l.length)} products`+(sel.size?` · ${bp(sel.size)} selected`:'');
 if(!ALL.length&&!metaD.ready)$('grid').innerHTML='<p class="catalog-empty">Loading products…</p>';
 else if(!ALL.length)$('grid').innerHTML='<p class="catalog-empty">No products are available right now.</p>';
 else if(!list.length)$('grid').innerHTML=`<div class="catalog-empty" style="grid-column:1/-1;text-align:center;padding:36px 18px;color:#526071;background:#fff;border:1px solid #dce6e0;border-radius:16px"><b>No products match your search.</b><p>Try another keyword or clear the current filters.</p><button class="pill" onclick="clearCatalogFilters()">Clear search and filters</button></div>`;
 else if(keep){const n=list.slice(0,keep);shown=n.length;$('grid').insertAdjacentHTML('beforeend',n.map(card).join(''));[...$('grid').children].slice(-n.length).forEach(translateUi)}else more();bar();translateUi($('info'))}
function clearCatalogFilters(){$('q').value='';$('fk').value='';$('so').value='';res=null;store='';tabf='';szf='';newOnly=false;drawTabs();drawSizes();render(1)}
const chips=sz=>sz==null?'':Object.keys(sz).length?Object.entries(sz).map(([k,v])=>{v=Number(v)||0;const m=String(k).trim().match(/^([A-Za-z0-9]+)\s+(\d{2})$/),lab=m?`<b class="sl">${esc(m[1])}</b><em class="sc2">${esc(m[2])}</em>`:`<b class="sl">${esc(k)}</b>`;return `<span class="zc ${v<=0?'no':v<=2?'lo':'ok'}">${lab}<span class="q">${v}</span></span>`}).join(''):'<span class="zc no">স্টক নেই</span>';
const LGO=s=>/alien/i.test(s)?'/logo/a.png':/body/i.test(s)?'/logo/b.png':'';
function card(p){
 const showStock=canDo('view_stock');
 const st=showStock?chips(p.sizes):'';
 const ss=showStock?stockState(p):null;
 const status=ss?(ss.state==='in'?'<span class="stockstate">✅ স্টকে</span>':ss.state==='low'?'<span class="stockstate low">⚠ কম স্টক</span>':'<span class="stockstate out">❌ স্টক নেই</span>'):'';
 const lg=LGO(p.store);
 const fav='',_fav0=(me&&(allowed('favorite')||meRole==='admin'))?`<button class="fav ${favorites.has(p.key)?'on':''}" onclick="event.stopPropagation();toggleFav('${esc(p.key)}')">${favorites.has(p.key)?'★':'☆'}</button>`:'';
 const scoreNumber=(p.sc!==undefined&&p.sc!==null&&String(p.sc).trim()!=='')?Number(p.sc):NaN;const scoreText=Number.isFinite(scoreNumber)?Math.max(0,Math.min(100,scoreNumber)).toFixed(1).replace(/\.0$/,''):'';const lab=scoreText!==''?`<div class="sc" title="Image similarity score" aria-label="মিলের হার ${scoreText} শতাংশ">${p.top?'★ ':''}মিলের হার: <strong class="match-percent">${scoreText}%</strong></div>`:'',tg=(p.tabs||[]).length?`<div class="tg">${esc(p.tabs.join(' · '))}</div>`:'';
 const id=productId(p);
 const idBtn=id?(canDo('id_copy')?`<button class="copycode" title="এক click copies the Product ID" onclick="event.stopPropagation();copyCode('${esc(p.key)}')">📋 ID: ${esc(id)}</button>`:`<small class="product-id">ID: ${esc(id)}</small>`):'';
 const galleryCount=[p.img,...(p.extra||[])].filter(Boolean).length;
 const ck=me&&meRole!=='cs'?`<label class="ck"><input type="checkbox" ${sel.has(p.key)?'checked':''}></label>`:'';
 return `<div class="c ${sel.has(p.key)?'sel':''} ${p.key==hitK?'hit':''}" data-k="${esc(p.key)}">${ck}<div class="iw">${p.img?`<img loading="lazy" src="${im(p.img)}" alt="">`:`<div class="excel-product-placeholder"><b>${esc(id||p.title)}</b><small>Excel Stock</small></div>`}${galleryCount>1?`<span class="photo-count">📸 ${bn(galleryCount)}</span>`:''}${lg?`<img class="lg" src="${lg}" alt="">`:''}${tg}${lab}${fav}${p.is_new?'<span class="newtag">নতুন</span>':''}</div><div class="m"><b>${esc(p.title)}</b><span>${idBtn} · ${esc(p.store)}</span><div class="pr">৳${bn(p.price)}${p.old?`<s>৳${bn(p.old)}</s>`:''}</div>${status}${st?`<div class="st">${st}</div>`:''}</div></div>`
}
function more(){if(shown>=list.length)return;const n=list.slice(shown,shown+60);shown+=n.length;$('grid').insertAdjacentHTML('beforeend',n.map(card).join(''));[...$('grid').children].slice(-n.length).forEach(translateUi)}
new IntersectionObserver(e=>{if(e[0].isIntersecting)more()},{rootMargin:'600px'}).observe($('sent'));
let searchTimer=null,lastLoggedSearch='',textSearchLoaderDelay=null,textSearchRenderDelay=null,searchLoaderHideTimer=null,searchLoaderSeq=0,catalogSearchRequestSeq=0;
function startCatalogLoader(kind='text',force=false){clearTimeout(searchLoaderHideTimer);const token=++searchLoaderSeq;__siteLoaderState.manual=1;__siteLoaderState.kind=kind;const layer=$('searchLoader');if(layer){layer.dataset.token=String(token);layer.dataset.kind=kind}__siteLoaderRefresh(true);return token}
function finishCatalogLoader(token,delay=420){clearTimeout(searchLoaderHideTimer);searchLoaderHideTimer=setTimeout(()=>{if(Number(token)!==searchLoaderSeq)return;__siteLoaderState.manual=0;__siteLoaderRefresh(false)},delay)}
function cancelCatalogLoader(){catalogSearchRequestSeq++;searchLoaderSeq++;clearTimeout(searchLoaderHideTimer);clearTimeout(textSearchLoaderDelay);clearTimeout(textSearchRenderDelay);__siteLoaderState.manual=0;__siteLoaderRefresh(false)}
$('q').oninput=()=>{if(!canDo('search')){toast('এই সার্চ permission বন্ধ আছে',true);$('q').value='';cancelCatalogLoader();return}res=null;clearTimeout(searchTimer);clearTimeout(textSearchLoaderDelay);clearTimeout(textSearchRenderDelay);clearTimeout(searchLoaderHideTimer);const qNow=$('q').value.trim(),requestId=++catalogSearchRequestSeq;if(!qNow){cancelCatalogLoader();render(1)}else{textSearchLoaderDelay=setTimeout(()=>{if(requestId===catalogSearchRequestSeq)startCatalogLoader('text')},90);textSearchRenderDelay=setTimeout(()=>{if(requestId!==catalogSearchRequestSeq)return;const token=startCatalogLoader('text');render(1);finishCatalogLoader(token,460)},230)}searchTimer=setTimeout(()=>{const q=$('q').value.trim();if(q&&q!==lastLoggedSearch){lastLoggedSearch=q;jf('/api/activity/search',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({query:q,results:list.length})})}},650)};
$('fk').onchange=()=>{if(!canDo('view_stock')){toast('স্টক দেখার permission বন্ধ আছে',true);$('fk').value='';return}logUiAction('filter_stock',$('fk').value||'সব',$('fk').value||'clear');render(1)};
$('so').onchange=()=>{logUiAction('sort_price',$('so').value||'ডিফল্ট',$('so').value||'clear');render(1)};
$('grid').onclick=e=>{const c=e.target.closest('.c');if(!c)return;const k=c.dataset.k;if(e.target.closest('.fav')||e.target.closest('.copycode'))return;
 if(e.target.closest('.ck')){e.target.type=='checkbox'&&(e.target.checked?sel.set(k,MAP[k]):sel.delete(k),c.classList.toggle('sel',e.target.checked),bar(),logUiAction('product_select',productId(MAP[k])||k,e.target.checked?'select':'deselect'));return}
 document.querySelectorAll('.c.hit').forEach(x=>x.classList.remove('hit'));hitK=k;c.classList.remove('pk');void c.offsetWidth;c.classList.add('pk','hit');setTimeout(()=>openM(k),260)};
function bar(){$('bar').classList.toggle('on',sel.size>0);$('bc').innerHTML=bn(sel.size)+'টা সিলেক্ট'}
function selAll(){if(sel.size)sel.clear();else list.forEach(p=>sel.set(p.key,p));render(1)}

// ---- মডাল ----
function sizeChartMarkup(p){
 const img=p?.size_chart_image||'',txt=p?.size_chart_text||'';
 if(img)return `<div class="sizechart"><h4>📏 Size Chart</h4><img src="${im(img)}" alt="Size Chart"></div>`;
 if(txt)return `<div class="sizechart"><h4>📏 Size Chart</h4><p>${esc(txt)}</p></div>`;
 return '';
}
function openM(k){const p=MAP[k];if(!p)return;if(!canDo('view_product')){toast('Permission বন্ধ আছে,অনুগ্রহ করে লগইন করুন',true);return}if(me)post('/api/recent',{key:k});else logUiAction('product_view',productId(p)||k,'guest');const showStock=canDo('view_stock'),imgs=[p.img,...(p.extra||[])].filter(Boolean),stk=showStock?chips(p.sizes):'',id=productId(p),status=showStock?(stockState(p).state==='in'?'<span class="stockstate">✅ স্টকে</span>':stockState(p).state==='low'?'<span class="stockstate low">⚠ কম স্টক</span>':'<span class="stockstate out">❌ স্টক নেই</span>'):'';
 $('mb').innerHTML=`<button class="x" onclick="closeM()">✕</button><div class="product-detail-grid"><div class="product-media"><img id="mi" src="${im(imgs[0]||p.img)}"><div class="gallery-label">📸 ${bn(imgs.length)}টি ছবি · প্রতিটি ছবি আলাদাভাবে কপি করা যাবে</div><div class="th">${imgs.map((u,i)=>`<div class="thumb-wrap"><img src="${im(u)}" onclick="mi.src=this.src;mi.dataset.u='${esc(u)}'" title="ছবি ${i+1}"><button type="button" class="thumb-copy" data-perm="copy" onclick="event.stopPropagation();copyImageUrl('${esc(u)}')">📋</button></div>`).join('')}</div>${sizeChartMarkup(p)}</div><div class="product-summary"><h2>${esc(p.title)}</h2><div class="product-meta">${id?`<button class="copycode" onclick="copyCode('${esc(k)}')">📋 ID: ${esc(id)}</button>`:''}<span>· ${esc(p.store)}</span></div><div class="product-modal-price">৳${bn(p.price)}${p.old?` <s>৳${bn(p.old)}</s>`:''}</div>${status}${showStock&&stk?`<div class="product-modal-sizes"><div class="st">${stk}</div></div>`:''}${p.tabs&&p.tabs.length?`<div class="product-modal-tabs">শিট ট্যাব: ${esc(p.tabs.join(', '))}</div>`:''}<div class="acts"><button data-perm="id_copy" onclick="copyCode('${esc(k)}')">📋 ID কপি</button><button data-perm="name_copy" onclick="copyProductName('${esc(k)}')">📝 নাম কপি</button><button data-perm="full_info_copy" onclick="copyFullInfo('${esc(k)}')">📄 Full Info</button><button data-perm="share_product" onclick="shareProduct('${esc(k)}')">📤 Share</button><button data-perm="download" onclick="dl1('${esc(k)}')">⬇ HD ডাউনলোড</button><button data-perm="copy" onclick="copyImg('${esc(k)}')">⧉ এই ছবি কপি</button><button data-perm="copy" onclick="copyProductImages('${esc(k)}')">📸 সব ছবি এক এক করে কপি</button><button data-perm="link_copy" onclick="copyText('${esc(p.link||'')}').then(ok=>{if(ok){logUiAction('link_copy',productId(MAP['${esc(k)}'])||'', 'single product')}toast(ok?'লিংক কপি হয়েছে':'কপি হয়নি',!ok)})">🔗 লিংক কপি</button><button data-perm="report" class="report" onclick="reportProduct('${esc(k)}')"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="flex:none"><path d="M5 21V4M5 4h11l-2 4 2 4H5"/></svg><span>রিপোর্ট</span></button>${p.link?`<a data-perm="website_view" href="${esc(p.link)}" target="_blank">🌐 ওয়েবসাইটে দেখুন</a>`:''}</div><p class="product-description">${esc(p.desc)}</p></div></div>${me?'<div id="stockComments" class="stock-comments">লোড হচ্ছে…</div>':''}<section class="similar-section"><h3>সিমিলার প্রোডাক্ট</h3><div class="sg" id="sim"></div></section>`;translateUi($('mb'));$('mo').classList.add('on');applyPermUI();if(me)loadStockComments(p);jf('/api/similar?key='+encodeURIComponent(k)).then(j=>{const e=$('sim');if(!e)return;const matches=(j.matches||[]).slice().sort((a,b)=>Number(b.score||0)-Number(a.score||0));e.innerHTML=matches.map(m=>Object.assign({sc:m.score},MAP[m.key])).filter(x=>x.key).map(card).join('')||'<span style="color:#888">এখনো পাওয়া যায়নি</span>';[...e.children].forEach(translateUi);e.onclick=ev=>{const c=ev.target.closest('.c');if(c){logUiAction('similar_click',productId(MAP[c.dataset.k])||c.dataset.k,'similar');openM(c.dataset.k)}}})}
let activeStockCommentCode='';
let activeStockTargets=[];
async function loadStockComments(product){const box=$('stockComments');if(!box)return;const candidates=typeof product==='string'?[product]:[productId(product),product?.id,product?.display_id].filter(Boolean).map(String).filter((v,i,a)=>a.indexOf(v)===i);if(!candidates.length){box.innerHTML='<small>স্টক কমেন্টের জন্য Product Code নেই</small>';return}let code=candidates[0],j={targets:[],comments:[]},savedComments=null;for(const candidate of candidates){const result=await jf('/api/excel-stock/comments?code='+encodeURIComponent(candidate));if((result.targets||[]).length){j=result;code=candidate;break}if((result.comments||[]).length)savedComments=result}if(!(j.targets||[]).length&&savedComments)j=savedComments;activeStockCommentCode=code;activeStockTargets=j.targets||[];const sizes=[...new Set(activeStockTargets.map(x=>x.size))];const commentList=(j.comments||[]).map(x=>`<div class="stock-comment-row"><b>${esc(x.sheet)} · ${esc(x.size)} · ${bn(x.sold_qty)} পিস বিক্রি</b><small>${esc(x.username)} · ${esc(new Date(x.created_at*1000).toLocaleString('en-US'))}</small>${x.message?`<p>${esc(x.message)}</p>`:''}${x.can_delete?`<button class="pill" onclick="deleteStockComment(${Number(x.id)||0})">মুছুন</button>`:''}</div>`).join('');const form=me&&(allowed('stock_comment')||isAdminLevelJS())?`<form class="stock-comment-form" onsubmit="submitStockComment(event)"><label>বিক্রি হওয়া সাইজ<input id="commentSize" list="commentSizes" autocomplete="off" placeholder="যেমন M 38 / XL 42" required><datalist id="commentSizes">${sizes.map(x=>`<option value="${esc(x)}">`).join('')}</datalist></label><label>কত পিস বিক্রি হয়েছে?<input id="commentQty" type="number" min="1" step="1" value="1" required></label><label>কমেন্ট (ঐচ্ছিক)<input id="commentText" maxlength="500" placeholder="যেমন: One piece sold"></label><button class="pill save-excel">কমেন্ট সেভ</button></form>`:'';box.innerHTML=`<h4>🧾 স্টক বিক্রির কমেন্ট</h4>${form}<div class="stock-comment-list">${commentList||'<small>এখনো কোনো বিক্রির কমেন্ট নেই</small>'}</div>`}

async function submitStockComment(e){e.preventDefault();const size=$('commentSize')?.value.trim(),matches=activeStockTargets.filter(x=>String(x.size).trim().replace(/\s+/g,' ').toLowerCase()===String(size||'').replace(/\s+/g,' ').toLowerCase()),selected=matches[0]||{sheet:'স্টক তালিকায় নেই',code:activeStockCommentCode,size};if(!size)return toast('সাইজ লিখুন',true);const j=await post('/api/excel-stock/comments',{sheet:selected.sheet,product_code:selected.code,size:selected.size,sold_qty:Number($('commentQty').value),message:$('commentText').value});if(!j._ok)return toast(j.error||'কমেন্ট সেভ হয়নি',true);toast('বিক্রির কমেন্ট সেভ হয়েছে');await loadStockComments(activeStockCommentCode)}

function isSuperAdminJS(){return meRole==='super_admin'}
async function deleteStockComment(id){if(!isAdminLevelJS())return toast('শুধু অ্যাডমিন কমেন্ট মুছতে পারবেন',true);if(!confirm('এই কমেন্টটি মুছে ফেলবেন?'))return;const j=await post('/api/admin/excel-stock/comments/'+id,{},'DELETE');if(!j._ok)return toast(j.error||'মুছতে পারিনি',true);toast('কমেন্ট মুছে ফেলা হয়েছে');await loadStockComments(activeStockCommentCode);loadCmt();if($('cmtSummary'))renderCommentSummary()}
async function clearStockHistory(){if(!isSuperAdminJS())return toast('শুধু Super Admin মুছতে পারবেন',true);if(!confirm('সব স্টক হিস্ট্রি, বিক্রির কমেন্ট ও import/backup-এর তালিকা মুছে ফেলবেন? মূল Excel workbook অপরিবর্তিত থাকবে।'))return;const j=await post('/api/admin/excel-stock/history',{},'DELETE');if(!j._ok)return toast(j.error||'মুছতে পারিনি',true);toast('স্টক হিস্ট্রি ও কমেন্ট মুছে ফেলা হয়েছে');await excelStockHistory()}
function closeM(){$('mo').classList.remove('on')}
document.onkeydown=e=>{if(e.key=='Escape'){closeM();$('ao').classList.remove('on');$('ao').classList.remove('admin-full');dr(0)}};

// ---- ডাউনলোড / কপি ----
const dl1=k=>{if(!canDo('download'))return toast('আপনার HD ডাউনলোড permission নেই',true);const p=MAP[k],a=document.createElement('a');a.href=im(p.img)+'&dl=1&name='+encodeURIComponent(productId(p)||'product')+'&code='+encodeURIComponent(productId(p)||'')+'&price='+encodeURIComponent(p.price??'');a.click()};
async function dlSel(){if(!sel.size)return toast('আগে প্রোডাক্ট সিলেক্ট করুন');let i=0;for(const k of sel.keys()){dl1(k);if(++i<sel.size)await new Promise(r=>setTimeout(r,500))}toast('ডাউনলোড শুরু হয়েছে')}
async function copyImg(k){if(!canDo('copy'))return toast('আপনার HD কপি permission নেই',true);const p=MAP[k];if(!p)return;try{
 const b=await (await fetch(im(p.img)+'&copy=1',{cache:'no-store'})).blob();
 if(!navigator.clipboard||!window.ClipboardItem)throw new Error('clipboard');
 const type=b.type||'image/png';await navigator.clipboard.write([new ClipboardItem({[type]:b})]);
 toast(sel.size>1?'প্রথম ছবিটা কপি হয়েছে (একবারে একটাই কপি হয়)':'HD ছবি কপি হয়েছে');
 }catch(e){try{
 const b=await (await fetch(im(p.img)+'&copy=1',{cache:'no-store'})).blob(),bm=await createImageBitmap(b),cv=document.createElement('canvas');cv.width=bm.width;cv.height=bm.height;cv.getContext('2d').drawImage(bm,0,0);
 const png=await new Promise(r=>cv.toBlob(r,'image/png'));await navigator.clipboard.write([new ClipboardItem({'image/png':png})]);toast(sel.size>1?'প্রথম ছবিটা কপি হয়েছে (একবারে একটাই কপি হয়)':'HD ছবি কপি হয়েছে')}catch(e2){toast('কপি হয়নি — localhost বা https এ খুলুন')}}}
async function copyImageUrl(url, silent=false){
 if(!canDo('copy')){if(!silent)toast('আপনার HD কপি permission নেই',true);return false}
 try{
  const b=await (await fetch(im(url)+'&copy=1',{cache:'no-store'})).blob();
  if(!navigator.clipboard||!window.ClipboardItem)throw new Error('clipboard');
  const type=b.type||'image/png';
  await navigator.clipboard.write([new ClipboardItem({[type]:b})]);
  return true;
 }catch(e){
  try{
   const b=await (await fetch(im(url)+'&copy=1',{cache:'no-store'})).blob(),bm=await createImageBitmap(b),cv=document.createElement('canvas');
   cv.width=bm.width;cv.height=bm.height;cv.getContext('2d').drawImage(bm,0,0);
   const png=await new Promise(r=>cv.toBlob(r,'image/png'));
   await navigator.clipboard.write([new ClipboardItem({'image/png':png})]);
   return true;
  }catch(e2){if(!silent)toast('কপি হয়নি — Chrome-এ clipboard permission দিন বা HTTPS/localhost ব্যবহার করুন',true);return false}
 }
}
async function copyImg(k){
 if(!canDo('copy'))return toast('আপনার HD কপি permission নেই',true);
 const p=MAP[k];if(!p)return;
 const ok=await copyImageUrl(p.img,true);
 if(ok){logUiAction('copy',productId(p)||k,'HD image 1/1');toast('HD ছবি কপি হয়েছে')}
 else toast('কপি হয়নি — Chrome-এ clipboard permission দিন বা HTTPS/localhost ব্যবহার করুন',true);
}
async function copyProductImages(k){
 if(!canDo('copy'))return toast('আপনার HD কপি permission নেই',true);
 const p=MAP[k];if(!p)return;
 const imgs=[p.img,...(p.extra||[])].filter(Boolean);
 let okCount=0;
 for(let i=0;i<imgs.length;i++){
  const ok=await copyImageUrl(imgs[i],true);
  if(ok){okCount++;logUiAction('copy',productId(p)||k,`HD image ${i+1}/${imgs.length}`)}
  if(i<imgs.length-1)await new Promise(r=>setTimeout(r,500));
 }
 toast(okCount===imgs.length?`সব ${bn(imgs.length)}টি ছবি এক এক করে কপি হয়েছে`:`${bn(okCount)}/${bn(imgs.length)}টি ছবি কপি হয়েছে`,okCount!==imgs.length);
}
async function copySelectedSeparately(){
 if(!canDo('copy'))return toast('আপনার HD কপি permission নেই',true);
 const ks=[...sel.keys()];
 if(!ks.length)return toast('আগে প্রোডাক্ট সিলেক্ট করুন');
 dr(0);
 let done=0,total=0;
 for(const k of ks){const p=MAP[k];if(!p)continue;const imgs=[p.img,...(p.extra||[])].filter(Boolean);total+=imgs.length;
  for(let i=0;i<imgs.length;i++){const ok=await copyImageUrl(imgs[i],true);if(ok){done++;logUiAction('copy',productId(p)||k,`HD image ${i+1}/${imgs.length}`)}if(done<total)await new Promise(r=>setTimeout(r,500));}
 }
 toast(`আলাদা আলাদা করে ${bn(done)}টি ছবি কপি করা হয়েছে`+(done<total?` · ${bn(total-done)}টি ব্যর্থ`:''),done<total);
}
function copySel(){return copySelectedSeparately()}
function copyLinks(){if(!canDo('link_copy'))return toast('আপনার লিংক কপি permission নেই',true);navigator.clipboard.writeText([...sel.values()].map(p=>p.link).join('\n')).then(()=>{logUiAction('link_copy','',String(sel.size)+' product');toast('লিংক কপি হয়েছে')})}
// ---- ছবি দিয়ে খোঁজা ----
async function imgSearch(f){if(!f)return;if(!canDo('image_search'))return toast('ইমেজ সার্চ permission বন্ধ আছে',true);showPrev(f);clearTimeout(textSearchLoaderDelay);clearTimeout(textSearchRenderDelay);const requestId=++catalogSearchRequestSeq;$('info').textContent='ছবি মিলিয়ে দেখছি…';const loaderToken=startCatalogLoader('image',true);const fd=new FormData();fd.append('file',f);try{const j=await jf('/api/search-image',{method:'POST',body:fd});if(requestId!==catalogSearchRequestSeq)return;if(!j._ok){toast(j.message||'সমস্যা হয়েছে');render(0);return}res=j.matches;store='';tabf='';szf='';$('q').value='';drawTabs();drawSizes();render(1);if($('grid'))$('grid').scrollIntoView({behavior:'smooth',block:'start'});if(j.ai===false)toast('AI এখনও প্রস্তুত হয়নি — সাধারণ মোডে খোঁজা হয়েছে',true)}catch(e){if(requestId===catalogSearchRequestSeq){toast('ইমেজ সার্চ সম্পন্ন করা যায়নি — আবার চেষ্টা করুন',true);render(0)}}finally{if(requestId===catalogSearchRequestSeq)finishCatalogLoader(loaderToken,560);if(fi)fi.value=''}}
fi.onchange=()=>imgSearch(fi.files[0]);
// The visual drop zone is removed; allow image drag-and-drop anywhere in the main page.
function __droppedImage(e){const dt=e.dataTransfer;if(!dt)return null;const f=[...(dt.files||[])].find(x=>String(x.type||'').startsWith('image/'));if(f)return f;const it=[...(dt.items||[])].find(x=>x.kind==='file'&&String(x.type||'').startsWith('image/'));return it?it.getAsFile():null}
document.addEventListener('dragover',e=>{const dt=e.dataTransfer;if(dt&&([...(dt.types||[])].includes('Files')||__droppedImage(e)))e.preventDefault()});
document.addEventListener('drop',e=>{const f=__droppedImage(e);if(f){e.preventDefault();imgSearch(f)}});
addEventListener('paste',e=>{const cd=e.clipboardData;let f=[...(cd?.files||[])].find(x=>String(x.type||'').startsWith('image/'));if(!f&&cd?.items){const it=[...cd.items].find(x=>String(x.type||'').startsWith('image/'));if(it)f=it.getAsFile()}if(f){e.preventDefault();imgSearch(f)}});

// ---- অ্যাডমিন ----
function scheduleSessionExpiry(expiresAt){if(window.sessionExpiryTimer)clearTimeout(window.sessionExpiryTimer);const left=Number(expiresAt)*1000-Date.now();if(left<=0){logout();return}window.sessionExpiryTimer=setTimeout(()=>{toast('৯ ঘণ্টার সেশন শেষ হয়েছে—আবার লগইন করুন');logout()},left)}
async function checkMe(){const __restorePanel=(!window.__adminPanelRestoreChecked?sessionStorage.getItem('activeAdminPanel'):'');window.__adminPanelRestoreChecked=true;const j=await jf('/api/me');me=j.user;meRole=j.role||'';if(me&&j.expires_at)scheduleSessionExpiry(j.expires_at);perms=j.permissions||{};guestPerms=j.guest_permissions||{};document.querySelectorAll('.direct-admin-link').forEach(b=>b.classList.remove('direct-visible'));document.querySelectorAll('.dr-menu-group').forEach(g=>g.hidden=true);if(!me){try{sessionStorage.removeItem('activeAdminPanel')}catch(_){}ALL=[];MAP={};grid.innerHTML='';prog.hidden=true;openAdminFrame(false);loginForm();return}prog.hidden=false;$('lo').hidden=!me;$('stockdash').hidden=!canUseStockDashboard();$('adbtn').textContent=me?(isAdminLevelJS()?'⚙ Admin Panel':'👤 My Account'):'🔐 Sign In';const directAllowed={posorders:isAdminLevelJS()||allowed('pos_orders')||allowed('pos_delete'),couriers:isAdminLevelJS()||allowed('courier_view')||allowed('courier_update'),purchases:isAdminLevelJS()||allowed('purchase_view')||allowed('purchase_manage'),suppliers:isAdminLevelJS()||allowed('supplier_manage')||allowed('purchase_manage')||allowed('purchase_view'),expenses:isAdminLevelJS()||allowed('expense_view')||allowed('expense_manage'),erpReports:isAdminLevelJS()||allowed('purchase_view')||allowed('purchase_manage')||allowed('expense_view')||allowed('expense_manage'),company:isAdminLevelJS(),hr:isAdminLevelJS()||allowed('hr_view')||allowed('hr_manage')||allowed('hr_payroll')||allowed('hr_leave_approve'),excelstock:canUseStockDashboard(),stockentry:canUseStockDashboard(),dashboard:isAdminLevelJS(),cat:isAdminLevelJS(),price:isAdminLevelJS()||allowed('price_edit'),notice:isAdminLevelJS(),photosettings:isAdminLevelJS(),orders:isAdminLevelJS()||allowed('order_view'),returns:isAdminLevelJS()||allowed('return_view'),design:isAdminLevelJS(),users:isAdminLevelJS(),backup:isAdminLevelJS(),guest:isAdminLevelJS(),logs:isAdminLevelJS(),reports:isAdminLevelJS()||allowed('report'),password:!!me};document.querySelectorAll('.direct-admin-link').forEach(b=>{if(directAllowed[b.dataset.direct])b.classList.add('direct-visible')});document.querySelectorAll('.dr-menu-group').forEach(g=>{g.hidden=!g.querySelector('.direct-admin-link.direct-visible')});applyPermUI();const welcome=document.getElementById('userWelcome');if(welcome)welcome.textContent='Welcome, '+String(me.username||me.name||'User')+'!';await loadPersonal();if(ALL.length)render(1);if(__restorePanel&&me&&canOpenAdminTabJS(__restorePanel)){window.__adminPanelRestored=true;openAdminFrame(true);setTimeout(()=>panel(__restorePanel),0)}else if(__restorePanel){try{sessionStorage.removeItem('activeAdminPanel')}catch(_){}}}
function isAdminLevelJS(){return meRole==='admin'||meRole==='super_admin'}
function canUseStockDashboard(){return isAdminLevelJS()||allowed('stock_dashboard')||allowed('stock_manage')||allowed('stock_history_view')}
function canSeeExcelStock(){return isAdminLevelJS()||allowed('stock_view')||allowed('stock_manage')}
function canSeeStockValue(){return isAdminLevelJS()||allowed('stock_value_view')}
function canSeeStockHistory(){return isAdminLevelJS()||allowed('stock_history_view')}
function canEditExcelStock(){return isAdminLevelJS()||allowed('stock_manage')||allowed('stock_set')||(canSeeExcelStock()&&(allowed('stock_add')||allowed('stock_reduce')))}
function roleLabel(r){return r==='super_admin'?'Super Admin':r==='admin'?'Admin':r==='cs'?'CS':'Staff'}
function allowed(k){return !!perms[k]}
function guestAllowed(k){return !!guestPerms[k]}
function canDo(k){return me?allowed(k):guestAllowed(k)}
function applyPermUI(){
 {const pb=$('posBtn');if(pb)pb.hidden=!(me&&(allowed('pos_use')||isAdminLevelJS()))}
 const sizeOn=canDo('size_filter');if(!sizeOn)szf='';
 $('dsz').style.display=sizeOn?'flex':'none';$('szh').style.display=sizeOn?'':'none';
 document.querySelectorAll('[data-perm=size-filter-label]').forEach(e=>e.style.display=sizeOn?'':'none');
 document.querySelectorAll('[data-perm]').forEach(e=>{const k=e.dataset.perm;if(k&&k!=='size-filter-label')e.style.display=canDo(k)?'':'none'});
 const admin=isAdminLevelJS(),cs=meRole==='cs',guest=!me;
 $('staffprice').hidden=!(me&&allowed('price_edit')&&!admin);
 $('useradmin').hidden=!admin;$('logadmin').hidden=!admin;$('dashadmin').hidden=!admin;$('reportadmin').hidden=!admin;$('guestadmin').hidden=!admin;$('passadmin').hidden=!admin;
 if(cs||guest){sel.clear();document.querySelectorAll('.ck,.fav').forEach(e=>e.style.display='none');$('bar').classList.remove('on');['fk','so','ft'].forEach(id=>{if($(id))$(id).style.display='none'})}
 else{document.querySelectorAll('.ck').forEach(e=>e.style.display='');if($('fk'))$('fk').style.display='';if($('so'))$('so').style.display='';if($('ft'))$('ft').style.display='';if($('quickbtn'))$('quickbtn').style.display=''}
}
async function logout(){sessionStorage.removeItem('activeAdminPanel');await post('/api/logout');location.replace('/')}
function openAdminFrame(full){$('ao').classList.toggle('admin-full',!!full);$('ao').classList.add('on')}
function closeAdminPanel(){sessionStorage.removeItem('activeAdminPanel');$('ao').classList.remove('on','admin-full')}
function adminBackToSlide(){closeAdminPanel();dr(1)}
let hrEmployeePhotoData='',hrAttendancePhotoData='',hrCameraStream=null;
function hrDuration(minutes){minutes=Math.max(0,Math.round(Number(minutes)||0));return Math.floor(minutes/60)+'h '+String(minutes%60).padStart(2,'0')+'m'}
async function hrSaveShiftHours(){const el=document.getElementById('hrShiftHours'),hours=Number(el?.value);if(!Number.isFinite(hours)||hours<1||hours>24)return toast('Standard shift hours must be between 1 and 24.',true);const j=await post('/api/hr/settings/shift-hours',{hours});if(!j._ok)return toast(j.error||'Overtime settings could not be saved.',true);toast('Standard shift and overtime settings saved.');await panel('hr');hrShowSection('attendance',document.querySelectorAll('.hr-subtabs button')[1])}
async function hrSaveLateRules(){const allowed_days=Number(document.getElementById('hrLateAllowedDays')?.value),deduction_per_day=Number(document.getElementById('hrLateDeductionPerDay')?.value);if(!Number.isInteger(allowed_days)||allowed_days<0||allowed_days>31)return toast('Allowed late days must be between 0 and 31.',true);if(!Number.isFinite(deduction_per_day)||deduction_per_day<0||deduction_per_day>1000000000)return toast('Deduction per late day must be a valid non-negative amount.',true);const j=await post('/api/hr/settings/late-rules',{allowed_days,deduction_per_day});if(!j._ok)return toast(j.error||'Late-entry rules could not be saved.',true);toast('Late-entry salary deduction rules saved.');await panel('hr');hrShowSection('attendance',document.querySelectorAll('.hr-subtabs button')[1])}
function hrShowSection(name,button){document.querySelectorAll('.hr-pane').forEach(x=>x.hidden=x.id!=='hr-pane-'+name);document.querySelectorAll('.hr-subtabs button').forEach(x=>x.classList.toggle('active',x===button));}
function hrAttendanceRowsHtml(rows){return (rows||[]).map(a=>`<tr><td>${esc(a.attendance_date)}</td><td>${esc(a.employee_code)} · ${esc(a.full_name)}</td><td>${esc(a.check_in||'—')}</td><td>${esc(a.check_out||'—')}</td><td>${hrDuration(Number(a.worked_minutes||Math.round(Number(a.work_hours||0)*60)))}</td><td><b class="${Number(a.overtime_minutes||0)>0?'hr-ot':'hr-zero'}">${hrDuration(Number(a.overtime_minutes||0))}</b></td><td><span class="hr-badge ${a.status==='Absent'?'rejected':''}">${esc(a.status)}</span></td><td>${a.photo_file?`<a href="/api/hr/photo/${encodeURIComponent(a.photo_file)}" target="_blank" rel="noopener">View Photo</a>`:'—'}</td></tr>`).join('')||'<tr><td colspan="8" class="hr-empty">No attendance records found for this employee and month.</td></tr>'}
async function hrLoadAttendanceReport(){const employeeId=document.getElementById('hrAttendanceEmployee')?.value||'',month=document.getElementById('hrAttendanceMonth')?.value||new Date().toISOString().slice(0,7),tbody=document.getElementById('hrAttendanceRows'),summary=document.getElementById('hrAttendanceSummary'),link=document.getElementById('hrAttendanceExport');if(!/^\d{4}-\d{2}$/.test(month))return toast('Select a valid month.',true);const qs=new URLSearchParams({month});if(employeeId)qs.set('employee_id',employeeId);if(link)link.href='/api/hr/attendance-export.xlsx?'+qs.toString();if(tbody)tbody.innerHTML='<tr><td colspan="8" class="hr-empty">Loading attendance…</td></tr>';const j=await jf('/api/hr/attendance-report?'+qs.toString());if(!j._ok){if(tbody)tbody.innerHTML=`<tr><td colspan="8" class="hr-empty">${esc(j.error||'Attendance report could not be loaded.')}</td></tr>`;return toast(j.error||'Attendance report could not be loaded.',true)}if(tbody)tbody.innerHTML=hrAttendanceRowsHtml(j.rows||[]);if(summary)summary.textContent=`${(j.rows||[]).length} record(s) · Worked ${hrDuration(j.total_worked_minutes||0)} · Overtime ${hrDuration(j.total_overtime_minutes||0)}`;}
function hrFilterTable(inputId,tableId){const q=(document.getElementById(inputId)?.value||'').toLowerCase();document.querySelectorAll('#'+tableId+' tbody tr[data-search]').forEach(tr=>tr.style.display=(tr.dataset.search||'').includes(q)?'':'none')}
function hrReadEmployeePhoto(input){const f=input.files?.[0];if(!f){hrEmployeePhotoData='';return}if(!f.type.startsWith('image/')){toast('Choose an image file.',true);input.value='';return}if(f.size>5*1024*1024){toast('Profile photo must be under 5 MB.',true);input.value='';return}const reader=new FileReader();reader.onload=()=>{hrEmployeePhotoData=String(reader.result||'');const im=document.getElementById('hrEmployeePhotoPreview');if(im){im.src=hrEmployeePhotoData;im.hidden=false}};reader.onerror=()=>toast('Could not read profile photo.',true);reader.readAsDataURL(f)}
function hrResetEmployeeForm(){const f=document.getElementById('hrEmployeeForm');if(f)f.reset();if(document.getElementById('hrEmployeeId'))document.getElementById('hrEmployeeId').value='';hrEmployeePhotoData='';const p=document.getElementById('hrEmployeePhotoPreview');if(p){p.hidden=true;p.removeAttribute('src')}const id=document.getElementById('hrEmployeeId');if(id)id.value=''}
async function hrEditEmployee(id){const j=await jf('/api/hr/employees/'+id);if(!j.ok)return toast(j.error||'Employee could not be loaded.',true);const e=j.employee;const vals={hrEmployeeId:e.id,hrEmployeeCode:e.employee_code,hrEmployeeName:e.full_name,hrEmployeePhone:e.phone,hrEmployeeEmail:e.email,hrEmployeeDepartment:e.department,hrEmployeeDesignation:e.designation,hrEmployeeType:e.employment_type,hrEmployeeJoinDate:e.join_date,hrEmployeeSalary:e.salary,hrEmployeeOvertimeRate:e.overtime_rate||0,hrEmployeeShiftHours:e.shift_hours||0,hrEmployeeEmergency:e.emergency_contact,hrEmployeeStatus:e.status,hrEmployeeAddress:e.address,hrEmployeeNotes:e.notes};Object.entries(vals).forEach(([k,v])=>{const el=document.getElementById(k);if(el)el.value=v??''});hrEmployeePhotoData='';const p=document.getElementById('hrEmployeePhotoPreview');if(p){if(e.photo_file){p.src='/api/hr/photo/'+encodeURIComponent(e.photo_file);p.hidden=false}else{p.hidden=true;p.removeAttribute('src')}}document.getElementById('hr-pane-employees')?.scrollIntoView({behavior:'smooth',block:'start'});toast('Employee loaded for editing.')}
async function hrSaveEmployee(){const id=document.getElementById('hrEmployeeId')?.value;const payload={id:id?Number(id):null,employee_code:document.getElementById('hrEmployeeCode')?.value||'',full_name:document.getElementById('hrEmployeeName')?.value||'',phone:document.getElementById('hrEmployeePhone')?.value||'',email:document.getElementById('hrEmployeeEmail')?.value||'',department:document.getElementById('hrEmployeeDepartment')?.value||'',designation:document.getElementById('hrEmployeeDesignation')?.value||'',employment_type:document.getElementById('hrEmployeeType')?.value||'Full-time',join_date:document.getElementById('hrEmployeeJoinDate')?.value||'',salary:Number(document.getElementById('hrEmployeeSalary')?.value||0),overtime_rate:Number(document.getElementById('hrEmployeeOvertimeRate')?.value||0),shift_hours:Number(document.getElementById('hrEmployeeShiftHours')?.value||0),emergency_contact:document.getElementById('hrEmployeeEmergency')?.value||'',status:document.getElementById('hrEmployeeStatus')?.value||'Active',address:document.getElementById('hrEmployeeAddress')?.value||'',notes:document.getElementById('hrEmployeeNotes')?.value||'',photo_data:hrEmployeePhotoData};if(!payload.full_name.trim())return toast('Enter the employee name.',true);const j=await post('/api/hr/employees',payload);if(!j._ok)return toast(j.error||'Employee could not be saved.',true);toast(id?'Employee updated.':'Employee added.');hrResetEmployeeForm();await panel('hr')}
async function hrSetEmployeeStatus(id,status){if(!confirm(`Set employee status to ${status}?`))return;const j=await post('/api/hr/employees/'+id+'/status',{status});if(!j._ok)return toast(j.error||'Employee status could not be updated.',true);toast('Employee status updated.');await panel('hr')}
async function hrStartCamera(){const v=document.getElementById('hrCameraVideo'),msg=document.getElementById('hrCameraStatus');try{if(!navigator.mediaDevices?.getUserMedia)throw Error('Webcam requires HTTPS or localhost and a supported browser.');hrCameraStream=await navigator.mediaDevices.getUserMedia({video:{facingMode:'user',width:{ideal:640},height:{ideal:480}},audio:false});v.srcObject=hrCameraStream;v.hidden=false;await v.play();if(msg)msg.textContent='Camera is ready. Capture a snapshot for the attendance record.'}catch(e){if(msg)msg.textContent=e.message||'Camera access was denied.';toast(e.message||'Could not open webcam.',true)}}
function hrCaptureCamera(){const v=document.getElementById('hrCameraVideo'),c=document.getElementById('hrCameraCanvas'),p=document.getElementById('hrAttendancePhotoPreview'),msg=document.getElementById('hrCameraStatus');if(!v||!v.videoWidth)return toast('Open the webcam and wait until the video is ready.',true);c.width=v.videoWidth;c.height=v.videoHeight;c.getContext('2d').drawImage(v,0,0,c.width,c.height);hrAttendancePhotoData=c.toDataURL('image/jpeg',.82);p.src=hrAttendancePhotoData;p.hidden=false;if(msg)msg.textContent='Snapshot captured. It will be attached to the next attendance save.';toast('Attendance snapshot captured.')}
function hrStopCamera(){if(hrCameraStream)hrCameraStream.getTracks().forEach(t=>t.stop());hrCameraStream=null;const v=document.getElementById('hrCameraVideo');if(v){v.pause();v.srcObject=null;v.hidden=true}const m=document.getElementById('hrCameraStatus');if(m)m.textContent='Camera is closed.'}
async function hrMarkAttendance(action){const id=Number(document.getElementById('hrAttEmployee')?.value||0);if(!id)return toast('Add an employee first.',true);const payload={employee_id:id,attendance_date:document.getElementById('hrAttDate')?.value||'',status:document.getElementById('hrAttStatus')?.value||'Present',check_in:document.getElementById('hrAttIn')?.value||'',check_out:document.getElementById('hrAttOut')?.value||'',notes:document.getElementById('hrAttNotes')?.value||'',photo_data:hrAttendancePhotoData,action};const j=await post('/api/hr/attendance',payload);if(!j._ok)return toast(j.error||'Attendance could not be saved.',true);if(j.attendance){if(document.getElementById('hrAttIn'))document.getElementById('hrAttIn').value=j.attendance.check_in||'';if(document.getElementById('hrAttOut'))document.getElementById('hrAttOut').value=j.attendance.check_out||''}hrAttendancePhotoData='';const p=document.getElementById('hrAttendancePhotoPreview');if(p){p.hidden=true;p.removeAttribute('src')}toast(action==='clock_in'?'Clock-in saved.':action==='clock_out'?'Clock-out saved.':'Attendance saved.');hrStopCamera();await panel('hr')}
async function hrSaveLeave(){const payload={employee_id:Number(document.getElementById('hrLeaveEmployee')?.value||0),leave_type:document.getElementById('hrLeaveType')?.value||'Casual',start_date:document.getElementById('hrLeaveStart')?.value||'',end_date:document.getElementById('hrLeaveEnd')?.value||'',reason:document.getElementById('hrLeaveReason')?.value||''};if(!payload.employee_id||!payload.start_date||!payload.end_date||!payload.reason.trim())return toast('Complete the employee, dates and reason.',true);const j=await post('/api/hr/leaves',payload);if(!j._ok)return toast(j.error||'Leave request could not be saved.',true);toast('Leave request submitted.');await panel('hr');hrShowSection('leaves',document.querySelectorAll('.hr-subtabs button')[2])}
async function hrReviewLeave(id,status){if(!['Approved','Rejected'].includes(status))return;if(!confirm(`${status} this leave request?`))return;const j=await post('/api/hr/leaves/'+id+'/review',{status});if(!j._ok)return toast(j.error||'Leave request could not be reviewed.',true);toast('Leave request '+status.toLowerCase()+'.');await panel('hr');hrShowSection('leaves',document.querySelectorAll('.hr-subtabs button')[2])}
let hrPayrollPreviewRequest=0;
async function hrPayrollPreview(){
 const employeeId=Number(document.getElementById('hrPayEmployee')?.value||0),month=document.getElementById('hrPayMonth')?.value||'',summary=document.getElementById('hrPayOtSummary');
 if(!employeeId||!/^\d{4}-\d{2}$/.test(month))return;
 const seq=++hrPayrollPreviewRequest,option=document.getElementById('hrPayEmployee')?.selectedOptions?.[0];
 if(summary)summary.textContent='Calculating monthly overtime…';
 const j=await jf('/api/hr/payroll/preview?'+new URLSearchParams({employee_id:String(employeeId),pay_month:month}).toString());
 if(seq!==hrPayrollPreviewRequest)return;
 if(!j._ok){const basic=document.getElementById('hrPayBasic');if(basic&&option)basic.value=option.dataset.salary||'0';if(summary)summary.textContent='Overtime preview unavailable. Save uses the employee profile and recorded attendance.';return}
 const v=j.preview||{};
 const set=(id,value)=>{const el=document.getElementById(id);if(el&&value!==undefined&&value!==null)el.value=String(value)};
 set('hrPayBasic',v.basic_salary??(option?.dataset.salary||0));set('hrPayAllowances',v.allowances??0);set('hrPayBonus',v.bonus??0);set('hrPayDeductions',v.deductions??0);set('hrPayStatus',v.status||'Draft');
 set('hrPayPaidDate',v.paid_date||new Date().toISOString().slice(0,10));set('hrPayMethod',v.payment_method||'Bank Transfer');set('hrPayReference',v.reference||'');set('hrPayNotes',v.notes||'');
 if(summary){const hours=Number(v.overtime_hours||0),rate=Number(v.overtime_rate||0),pay=Number(v.overtime_pay||0),lateDays=Number(v.late_days||0),allowedLate=Number(v.late_allowed_days||0),excessLate=Number(v.late_excess_days||0),lateDeduction=Number(v.late_deduction||0);summary.innerHTML=`<b>Monthly Overtime:</b> ${hours.toFixed(2)} hours × ৳${rate.toLocaleString('en-BD',{maximumFractionDigits:2})}/hour = <b>৳${pay.toLocaleString('en-BD',{maximumFractionDigits:2})}</b><div style="margin-top:7px"><b>Late entries:</b> ${lateDays} day(s) · Grace allowed: ${allowedLate} day(s) · Deductible: ${excessLate} day(s) · <b>Late deduction: ৳${lateDeduction.toLocaleString('en-BD',{maximumFractionDigits:2})}</b></div><small style="display:block;margin-top:4px">Late-entry settings are global for each month. Payroll automatically deducts the configured amount only for late days beyond the allowed count. ${v.existing?' Existing payroll found; saved values are loaded.':'Basic salary came from the employee profile.'}</small>`;}
}
async function hrSavePayroll(){const payload={employee_id:Number(document.getElementById('hrPayEmployee')?.value||0),pay_month:document.getElementById('hrPayMonth')?.value||'',basic_salary:Number(document.getElementById('hrPayBasic')?.value||0),allowances:Number(document.getElementById('hrPayAllowances')?.value||0),bonus:Number(document.getElementById('hrPayBonus')?.value||0),deductions:Number(document.getElementById('hrPayDeductions')?.value||0),status:document.getElementById('hrPayStatus')?.value||'Draft',paid_date:document.getElementById('hrPayPaidDate')?.value||'',payment_method:document.getElementById('hrPayMethod')?.value||'Bank Transfer',reference:document.getElementById('hrPayReference')?.value||'',notes:document.getElementById('hrPayNotes')?.value||''};if(!payload.employee_id||!/\d{4}-(0[1-9]|1[0-2])/.test(payload.pay_month))return toast('Select an employee and valid payroll month.',true);if([payload.basic_salary,payload.allowances,payload.bonus,payload.deductions].some(x=>!Number.isFinite(x)||x<0))return toast('Amounts must be zero or more.',true);const j=await post('/api/hr/payroll',payload);if(!j._ok)return toast(j.error||'Payroll could not be saved.',true);toast('Payroll saved, including attendance-based overtime.');await panel('hr');hrShowSection('payroll',document.querySelectorAll('.hr-subtabs button')[3])}

async function cashClosingReload(){window.cashClosingDate=$('cashCloseDate')?.value||window.cashClosingDate||'';await panel('cashclosing')}
function cashClosingPreview(){const opening=Math.max(0,Number($('cashCloseOpening')?.value||0)),actual=Math.max(0,Number($('cashCloseActual')?.value||0));const sales=Number($('cashCloseSales')?.dataset.value||0),expenses=Number($('cashCloseExpenses')?.dataset.value||0),supplier=Number($('cashCloseSupplier')?.dataset.value||0);const expected=Math.round((opening+sales-expenses-supplier)*100)/100;const el=$('cashCloseExpected'),diff=$('cashCloseDifference');if(el)el.textContent='৳ '+expected.toLocaleString('en-BD',{maximumFractionDigits:2});if(diff){const d=Math.round((actual-expected)*100)/100;diff.textContent=(d<0?'Short: ':'Over: ')+'৳ '+Math.abs(d).toLocaleString('en-BD',{maximumFractionDigits:2});diff.style.color=Math.abs(d)<0.01?'#176745':(d<0?'#b42318':'#9a6700')}}
async function saveCashClosing(){const payload={date:$('cashCloseDate')?.value||'',opening_cash:Number($('cashCloseOpening')?.value||0),actual_cash:Number($('cashCloseActual')?.value||0),notes:$('cashCloseNotes')?.value||''};const b=$('cashCloseSave');if(b){b.disabled=true;b.textContent='সেভ হচ্ছে…'}try{const j=await post('/api/admin/cash-closing',payload);if(!j._ok)throw Error(j.error||'Cash closing save হয়নি');toast('Daily cash closing সেভ হয়েছে');await panel('cashclosing')}catch(e){toast(e.message||'Cash closing save হয়নি',true)}finally{if(b){b.disabled=false;b.textContent='💾 Close Day & Save'}}}
function canOpenAdminTabJS(t){
 if(isAdminLevelJS())return true;
 if(!me)return false;
 if(t==='excelstock'||t==='stockentry')return canUseStockDashboard();
 if(t==='password')return !!me;
 const permsByTab={stockrecon:['stock_dashboard','stock_manage','stock_set','stock_add','stock_reduce','stock_history_view'],price:['price_edit'],orders:['order_view'],returns:['return_view'],couriers:['courier_view','courier_update'],posorders:['pos_orders','pos_delete'],purchases:['purchase_view','purchase_manage'],suppliers:['supplier_manage','purchase_manage','purchase_view'],expenses:['expense_view','expense_manage'],erpReports:['purchase_view','purchase_manage','expense_view','expense_manage'],hr:['hr_view','hr_manage','hr_leave_approve','hr_payroll']};
 return (permsByTab[t]||[]).some(k=>allowed(k));
}
function adminTab(t){if(!canOpenAdminTabJS(t))return;dr(0);openAdminFrame(true);panel(t)}
function staffPriceOpen(){if(!(me&&allowed('price_edit')))return toast('আপনার দাম এডিট permission নেই');dr(0);openAdminFrame(true);panel('price')}
function adminOpen(){dr(0);if(me){if(isAdminLevelJS()){openAdminFrame(true);panel('dashboard');return}const first=canUseStockDashboard()?'excelstock':allowed('order_view')?'orders':allowed('returns')?'returns':(allowed('courier_view')||allowed('courier_update'))?'couriers':(allowed('pos_orders')||allowed('pos_delete'))?'posorders':allowed('price_edit')?'price':(allowed('hr_view')||allowed('hr_manage')||allowed('hr_payroll')||allowed('hr_leave_approve'))?'hr':'';if(first){openAdminFrame(true);panel(first)}else{openAdminFrame(false);permissionInfo()}}else{openAdminFrame(false);loginForm()}}
const PERMS_LABELS={order_view:'🧾 অর্ডার দেখা',return_view:'↩ রিটার্ন দেখা',courier_view:'🚚 কুরিয়ার ট্র্যাকিং দেখা',courier_update:'🚚 কুরিয়ার আপডেট',pos_use:'🛒 POS বিক্রি',pos_orders:'🧾 POS অর্ডার ও মেমো',pos_delete:'🗑 POS বিক্রি সিলেক্ট করে / সব মুছুন','purchase_view':'🛒 Purchase Management দেখা','purchase_manage':'🛒 Purchase Entry / Manage','supplier_manage':'🏭 Supplier ও Payment Manage','expense_view':'💸 Expense Management দেখা','expense_manage':'💸 Expense Entry / Approve',hr_view:'👥 HR: View',hr_manage:'👥 HR: Manage Employees & Attendance',hr_leave_approve:'👥 HR: Approve Leaves',hr_payroll:'👥 HR: Payroll & Payslips',order_sheet_upload:'📤 Order Sheet আপলোড',return_sheet_upload:'↩ Return Sheet আপলোড',stock_dashboard:'স্টক ড্যাশবোর্ড',stock_view:'স্টকের পরিমাণ দেখা',stock_value_view:'দাম ও স্টক মূল্য দেখা',stock_history_view:'স্টক হিস্ট্রি দেখা',stock_manage:'পূর্ণ স্টক ব্যবস্থাপনা',stock_add:'স্টক যোগ',stock_reduce:'স্টক কমানো',stock_set:'পরিমাণ সরাসরি সেট',stock_comment:'বিক্রির কমেন্ট',stock_comment_delete:'বিক্রির কমেন্ট মুছতে পারবে',price_edit:'দাম এডিট',search:'টেক্সট সার্চ',image_search:'ইমেজ সার্চ',view_product:'প্রোডাক্ট দেখা',size_filter:'সাইজ ফিল্টার',download:'HD ডাউনলোড',copy:'HD কপি',link_copy:'লিংক কপি',id_copy:'ID কপি',name_copy:'নাম কপি',full_info_copy:'Full Info',share_product:'Share',website_view:'ওয়েবসাইটে দেখা',favorite:'ফেভারিট',report:'রিপোর্ট'};
function permissionInfo(){$('ab').innerHTML=`<button class="x" onclick="ao.classList.remove('on')">✕</button><h3>${esc(me)} — ${esc(roleLabel(meRole))}</h3><p>আপনার অনুমতি:</p><div class="perms">${Object.entries(PERMS_LABELS).map(([k,v])=>`<label>${v}: ${allowed(k)?'✅':'❌'}</label>`).join('')}</div><button class="pill" onclick="logout()">লগআউট</button>`;translateUi($('ab'))}
function loginForm(){$('ab').innerHTML=`<button class="x" onclick="ao.classList.remove('on')">✕</button><h3>স্টাফ / অ্যাডমিন লগইন</h3><div class="row"><input id="lu" placeholder="ইউজারনেম"><input id="lp" type="password" placeholder="পাসওয়ার্ড" onkeydown="event.key=='Enter'&&doLogin()"><button class="pill" onclick="doLogin()">লগইন</button></div><div style="font-size:12px;color:#888;text-align:center;margin-top:8px">ইউজারনেম ও পাসওয়ার্ড পেতে IT এডমিনের সাথে যোগাযোগ করুন।</div>`;translateUi($('ab'))}
async function doLogin(){const j=await post('/api/login',{username:$('lu').value,password:$('lp').value});if(j._ok){await checkMe();isAdminLevelJS()?panel('dashboard'):permissionInfo()}else toast('ইউজারনেম বা পাসওয়ার্ড ভুল')}
async function savePhotoDownloadSettings(){const j=await post('/api/admin/photo-download-settings',{show_code:!!$('photoShowCode')?.checked,show_price:!!$('photoShowPrice')?.checked,access:$('photoAccess')?.value||'all'});if(!j._ok)return toast(j.error||'সেটিংস সেভ হয়নি',true);toast('Photo Download সেটিংস সেভ হয়েছে')}
async function courierPanelMain(){const c=isAdminLevelJS()?await jf('/api/admin/couriers'):{couriers:[]};const opt=await jf('/api/admin/courier-options');const courierName=window.activeCourier||'Steadfast';const courierQ=(typeof window.courierQuery==='string'?window.courierQuery:'');window.courierQuery=courierQ;const tq=new URLSearchParams({q:courierQ,courier:courierName,from:window.courierFrom||'',to:window.courierTo||'',status:window.courierStatus||'',assigned_to:window.courierAssigned||''});const tr=await jf('/api/admin/courier-tracking?'+tq.toString());if(tr._s===401){location.reload();return ''}window.courierLoadErr=tr._ok?'':('Courier তালিকা লোড হয়নি ('+(tr.error||tr._s||'network')+')');const fh=await jf('/api/admin/courier-fraud-history?limit=8');if(!tr._ok){return `<div class="courier-dashboard"><section class="courier-card"><h3>Courier dashboard load হয়নি</h3><p>${esc(tr.error||('HTTP '+tr._s))}</p><button class="pill" onclick="panel('couriers')">আবার Refresh করুন</button></section></div>`}const names=(c.couriers||[]).map(x=>x.name);(opt.couriers||[]).forEach(x=>{if(!names.includes(x))names.push(x)});const selected=(c.couriers||[]).find(x=>x.name===courierName)||(c.couriers||[])[0]||{name:courierName};const rows=tr.rows||[],counts=tr.totals||tr.counts||{delivered:0,partial_delivered:0,returned:0,cancelled:0,pending:0,inway:0},money=tr.money||{},max=Math.max(1,...Object.values(counts));if(isAdminLevelJS()||allowed('courier_update'))setTimeout(()=>startCourierAutoSync(rows),0);const report='/api/admin/courier-report.xlsx?'+new URLSearchParams({from:window.courierFrom||'',to:window.courierTo||'',courier:courierName,status:window.courierStatus||'',assigned_to:window.courierAssigned||'',q:courierQ});return `<div class="courier-dashboard"><div class="courier-heading"><div><span class="eyebrow">ALIEN TEES ERP · DELIVERY</span><h2>Courier & Delivery</h2><p>Courier status, fraud check ও Excel report</p></div><div class="courier-head-actions">${isAdminLevelJS()?`<button class="pill" id="courierAutoSyncBtn" onclick="toggleCourierAutoSync()">${localStorage.getItem('courierAutoSync')==='off'?'Auto sync চালু করুন':'Auto sync · ১০ মিনিট'}</button>`:""}<a class="pill courier-report-link" href="${report}">⬇ Excel রিপোর্ট</a><span id="courierSyncProgress" role="status">Order Sheet Cons ID auto sync</span></div></div><div class="courier-provider-tabs">${names.map(n=>`<button class="${n===courierName?'on':''} courier-provider-${esc(n.toLowerCase().replace(/[^a-z0-9]+/g,'-'))}" onclick="selectCourier('${esc(n)}')">${esc(n)}</button>`).join('')}</div><div class="courier-stats">${[['ডেলিভারি সম্পন্ন',counts.delivered,money.delivered,'#17865b','মোট আদায়'],['আংশিক ডেলিভারি',counts.partial_delivered,0,'#4d83cf',''],['রিটার্ন',counts.returned,money.returned,'#d55454','রিটার্নের মোট টাকা'],['বাতিল',counts.cancelled,money.cancelled,'#d55454','বাতিল অর্ডারের মোট টাকা'],['পেন্ডিং',counts.pending,money.pending,'#d89a2b','মোট পেন্ডিং টাকা']].map(x=>`<article><small>${x[0]}</small><b>${bp(x[1])}</b>${x[4]?`<span class="courier-money">৳ ${bp(x[2])} · ${x[4]}</span>`:''}<div class="courier-meter"><i style="width:${Math.round(x[1]/max*100)}%;background:${x[3]}"></i></div></article>`).join('')}</div>${isAdminLevelJS()?`<section class="courier-card courier-config"><button class="courier-config-toggle" type="button" aria-expanded="${window.courierConfigOpen?'true':'false'}" onclick="toggleCourierConfig()"><span><b>কুরিয়ার API কনফিগারেশন</b><small>API কী, সিক্রেট ও স্ট্যাটাস পাথ</small></span><span class="courier-chevron">${window.courierConfigOpen?'−':'＋'}</span></button><div class="courier-config-body" ${window.courierConfigOpen?'':'hidden'}><p class="admin-note">API Key ও Secret Key courier account থেকে নিন। Saved দেখালে credential সংরক্ষিত আছে।</p><div class="courier-tabs">${(c.couriers||[]).map(x=>`<button class="${selected.name===x.name?'on':''}" onclick="selectCourier('${esc(x.name)}')">${esc(x.name)} <i>${x.enabled?'ON':'OFF'}</i></button>`).join('')}<button onclick="addCourierOption()">＋ Add Courier</button></div><div class="courier-fields"><label>কুরিয়ারের নাম<input id="courierName" value="${esc(selected.name||'Steadfast')}"></label><label>API বেস URL<input id="courierBase" value="${esc(selected.base_url||'')}" placeholder="https://..." autocomplete="url"></label><label>API কী<input id="courierKey" type="password" value="" placeholder="${selected.has_api_key?'Saved · blank রাখলে আগেরটাই থাকবে':'API Key'}" autocomplete="new-password"></label><label>API authentication<select id="courierAuthType"><option value="steadfast" ${selected.auth_type==='steadfast'?'selected':''}>Steadfast key + secret</option><option value="bearer" ${selected.auth_type==='bearer'?'selected':''}>Bearer token</option><option value="api-key" ${selected.auth_type==='api-key'?'selected':''}>API-Key header</option><option value="none" ${selected.auth_type==='none'?'selected':''}>No auth</option></select></label><label>Status field name<input id="courierStatusField" value="${esc(selected.status_field||'delivery_status')}"></label><label>Optional Fraud API path<input id="courierFraudPath" value="${esc(selected.fraud_path||'')}" placeholder="/fraud-check/{phone}"></label><label>সিক্রেট কী<input id="courierSecret" type="password" value="" placeholder="${selected.has_secret_key?'Saved · blank রাখলে আগেরটাই থাকবে':'Secret Key'}" autocomplete="new-password"></label><label class="courier-toggle"><input id="courierEnabled" type="checkbox" ${selected.enabled?'checked':''}> এই courier চালু</label><label>Status API path<input id="courierStatusPath" value="${esc(selected.mapping?.status_path||'/status_by_cid/{cons_id}')}"></label></div><div class="courier-actions"><button class="pill" onclick="saveCourier()">সেটিংস সেভ</button><button class="pill" type="button" onclick="testCourierConnection()">Steadfast connection test</button><span id="courierTestResult" role="status"></span></div></div></section>`:''}<section class="courier-card"><h3>ফ্রড চেক / কনসাইনমেন্ট আপডেট</h3><div class="courier-fields"><label>ফোন নম্বর<input id="fraudPhone" inputmode="tel" placeholder="01XXXXXXXXX"></label><label>কনসাইনমেন্ট আইডি<input id="trackCons" placeholder="কুরিয়ার কনসাইনমেন্ট নম্বর"></label><label>কুরিয়ার<select id="trackCourier">${names.map(x=>`<option ${x===selected.name?'selected':''}>${esc(x)}</option>`).join('')}</select></label><label>ডেলিভারি স্ট্যাটাস<select id="trackStatus"><option value="Pending">পেন্ডিং</option><option value="Pickup Pending">পিকআপ পেন্ডিং</option><option value="In Way">পথে আছে</option><option value="Delivered">ডেলিভারি সম্পন্ন</option><option value="Partial Delivered">আংশিক ডেলিভারি</option><option value="Cancelled">বাতিল</option><option value="Returned">রিটার্ন</option></select></label><label>কুরিয়ার রেটিং<input id="trackRating" type="number" min="0" max="5" step="0.1" placeholder="০–৫"></label><label>কুরিয়ার নোট<input id="trackDetails" placeholder="ডেলিভারি / রিটার্ন নোট"></label></div><div class="courier-actions"><button class="pill" onclick="runFraudCheck()">🛡 ফ্রড চেক</button>${isAdminLevelJS()||allowed('courier_update')?'<button class="pill" onclick="syncCourierStatus()">API থেকে স্ট্যাটাস আনুন</button><button class="pill" onclick="saveTracking()">স্ট্যাটাস আপডেট ও অর্ডার হিস্ট্রি সিঙ্ক</button>':''}<span id="fraudResult">${esc(window.courierFraudText||sessionStorage.getItem('courierFraudLast')||'')}</span></div>${fraudHistoryTable(fh)}</section><section class="courier-card"><div class="courier-heading"><div><h3>স্ট্যাটাস ড্যাশবোর্ড</h3><p>API থেকে পাওয়া স্ট্যাটাস অর্ডার হিস্ট্রি ও Excel রিপোর্টে সংরক্ষিত হবে।</p></div><div class="excel-actions"><label class="courier-date">শুরুর তারিখ (ঐচ্ছিক) <input type="date" id="courierFrom" value="${esc(window.courierFrom||'')}" onchange="window.courierFrom=this.value;panel('couriers')"></label><label class="courier-date">শেষের তারিখ (ঐচ্ছিক) <input type="date" id="courierTo" value="${esc(window.courierTo||'')}" onchange="window.courierTo=this.value;panel('couriers')"></label><select id="courierStatusFilter" onchange="window.courierStatus=this.value;panel('couriers')"><option value="">সব স্ট্যাটাস</option>${[['Pending','পেন্ডিং'],['Pickup Pending','পিকআপ পেন্ডিং'],['In Way','পথে আছে'],['Delivered','ডেলিভারি সম্পন্ন'],['Partial Delivered','আংশিক ডেলিভারি'],['Cancelled','বাতিল'],['Returned','রিটার্ন']].map(([v,l])=>`<option value="${v}" ${window.courierStatus===v?'selected':''}>${l}</option>`).join('')}</select><select id="courierAssignedFilter" onchange="window.courierAssigned=this.value;panel('couriers')"><option value="">সব দায়িত্বপ্রাপ্ত</option>${(tr.assigned||[]).map(a=>`<option value="${esc(a)}" ${window.courierAssigned===a?'selected':''}>${esc(a)}</option>`).join('')}</select><input id="courierQuery" placeholder="কনসাইনমেন্ট / কুরিয়ার / স্ট্যাটাস খুঁজুন" value="${esc(window.courierQuery||'')}" onkeydown="if(event.key==='Enter'){event.preventDefault();window.courierQuery=String(event.currentTarget.value||'').trim();panel('couriers')}"><button class="pill" onclick="window.courierQuery=String(document.getElementById('courierQuery')?.value||'').trim();panel('couriers')">খুঁজুন</button><button class="pill" onclick="resetCourierFilters()">রিসেট</button></div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>বুকিংয়ের তারিখ</th><th>কনসাইনমেন্ট আইডি</th><th>কুরিয়ার</th><th>স্ট্যাটাস</th><th>দায়িত্বপ্রাপ্ত</th><th>গ্রাহক</th><th>ফোন</th><th>প্রোডাক্ট</th><th>প্রোডাক্টের দাম</th><th>কুরিয়ার চার্জ</th><th>অগ্রিম</th><th>আদায়যোগ্য / API টাকা</th><th>রেটিং</th><th>বিস্তারিত</th><th>অর্ডার মিল</th><th>স্ট্যাটাস আপডেট</th><th>API</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.order_date||'—')}</td><td>${esc(x.cons_id)}</td><td>${esc(x.courier)}</td><td><b>${esc(({'Pending':'পেন্ডিং','Pickup Pending':'পিকআপ পেন্ডিং','In Way':'পথে আছে','Delivered':'ডেলিভারি সম্পন্ন','Partial Delivered':'আংশিক ডেলিভারি','Cancelled':'বাতিল','Canceled':'বাতিল','Returned':'রিটার্ন'})[x.status]||x.status)}</b></td><td>${esc(x.assigned_to||'—')}</td><td>${esc(x.customer_name||'—')}</td><td>${esc(x.phone||'—')}</td><td>${esc(x.product_code||'—')} ${esc(x.size||'')}</td><td>৳ ${bp(x.product_price||0)}</td><td>৳ ${bp(x.delivery_charge||0)}</td><td>৳ ${bp(x.advance_paid||0)}</td><td title="${x.amount_is_estimate?'Order Sheet estimate':'Courier API amount'}">${String(x.status||'').toLowerCase()=='partial delivered'?'—':`৳ ${bp(x.collectible_amount||0)} <small>${x.amount_is_estimate?'Order Sheet':'API'}</small>${x.api_amount!=null?`<small>API: ৳ ${bp(x.api_amount)}</small>`:''}`}</td><td>${esc(x.rating)}</td><td>${esc(x.details)}</td><td>${x.order_id?'মিলেছে #'+bp(x.order_id):'মেলেনি'}</td><td>${new Date(x.updated_at*1000).toLocaleString('bn-BD')}</td><td><button class="pill" onclick="syncCourierStatus('${esc(x.courier)}','${esc(x.cons_id)}')">Refresh</button></td></tr>`).join('')||'<tr><td colspan="17">কোনো কুরিয়ার ট্র্যাকিং রেকর্ড নেই</td></tr>'}</tbody></table></div></section></div>`}async function selectCourier(name){window.activeCourier=name;await panel('couriers')}
function addCourierOption(){const n=prompt('কুরিয়ার প্রতিষ্ঠানের নাম দিন');if(!n)return;window.activeCourier=n.trim();panel('couriers')}
async function saveCourier(){if(!isAdminLevelJS())return toast('শুধু Admin courier API settings বদলাতে পারবে',true);const name=$('courierName').value.trim(),mapping={status_path:$('courierStatusPath').value};const j=await post('/api/admin/couriers',{name,base_url:$('courierBase').value,api_key:$('courierKey').value,secret_key:$('courierSecret').value,enabled:!!$('courierEnabled').checked,auth_type:$('courierAuthType').value,status_field:$('courierStatusField').value,fraud_path:$('courierFraudPath').value,mapping});if(!j._ok)return toast(j.error||'Courier settings সেভ হয়নি',true);window.activeCourier=name;toast('Courier API settings সেভ হয়েছে');panel('couriers')}
async function testCourierConnection(){if(!isAdminLevelJS())return toast('শুধু Admin API connection test করতে পারবে',true);const el=$('courierTestResult');if(el)el.textContent='Steadfast API পরীক্ষা হচ্ছে…';const j=await post('/api/admin/courier-test',{name:$('courierName').value.trim()||'Steadfast'});if(el)el.textContent=j._ok?'✅ Steadfast API connected':('⚠ '+(j.error||'Connection test failed'));}
async function saveTracking(){if(!allowed('courier_update')&&!isAdminLevelJS())return toast('Courier status update permission নেই',true);const j=await post('/api/admin/courier-tracking',{cons_id:$('trackCons').value,courier:$('trackCourier').value,status:$('trackStatus').value,rating:$('trackRating').value,details:$('trackDetails').value,phone:$('fraudPhone').value});if(!j._ok)return toast(j.error||'স্ট্যাটাস সেভ হয়নি',true);window.courierKeep={};try{sessionStorage.removeItem('courierKeep')}catch(_){}toast(j.order_matched?'Status Excel report ও Order History-তে সেভ হয়েছে':'Status report-এ সেভ হয়েছে · Order match নেই');await panel('couriers');refreshNotificationCount()}
async function syncCourierStatus(courier=$('trackCourier')?.value,cons=$('trackCons')?.value){if(!allowed('courier_update')&&!isAdminLevelJS())return toast('Courier status update permission নেই',true);if(!cons)return toast('Consignment ID দিন',true);const j=await post('/api/admin/courier-sync',{courier,cons_id:cons});if(!j._ok)return toast(j.error||'Courier API থেকে status আসেনি',true);toast('Courier API status: '+j.status);await panel('couriers');refreshNotificationCount()}
function toggleCourierConfig(){window.courierConfigOpen=!window.courierConfigOpen;const body=document.querySelector('.courier-config-body'),button=document.querySelector('.courier-config-toggle');if(body)body.hidden=!window.courierConfigOpen;if(button)button.setAttribute('aria-expanded',window.courierConfigOpen?'true':'false');const mark=document.querySelector('.courier-chevron');if(mark)mark.textContent=window.courierConfigOpen?'−':'＋'}
function toggleCourierAutoSync(){const enabled=localStorage.getItem('courierAutoSync')==='off';localStorage.setItem('courierAutoSync',enabled?'on':'off');toast(enabled?'Auto sync চালু · প্রতি ১০ মিনিটে':'Auto sync বন্ধ');panel('couriers')}
async function openNotificationTarget(id,ref){try{await post('/api/notifications/read',{id});closeAdminPanel();const raw=String(ref||'').trim(),r=raw.toLowerCase();if(r.includes('return'))return panel('returns');if(r.includes('order'))return panel('orders');if(r.includes('courier')||r.includes('cons')){const cons=raw.replace(/^(courier|consignment|cons|order)[:#-]?/i,'').trim();if(cons)window.courierQuery=cons;return panel('couriers')}if(r.includes('stock'))return panel('excelstock');if(r.includes('product')||r.includes('catalog')){const q=raw.replace(/^(product|catalog)[:#-]?/i,'').trim();if(q){$('q').value=q;render(1)}return}if(r.includes('user'))return panel('users');if(r.includes('report'))return panel('reports');if(r.includes('notice'))return panel('notice');if(raw){window.courierQuery=raw.replace(/^(consignment|cons|courier)[:#-]?/i,'').trim();return panel('couriers')}await refreshNotificationCount()}catch(e){toast('Notifications খোলা যায়নি',true)}}
async function openNotifications(){const j=await jf('/api/notifications');const rows=j.items||[];$('notifyCount').textContent=bp(j.unread||0);openAdminFrame(false);$('ao').classList.remove('admin-full');$('ab').innerHTML=`<div class="notification-panel"><button class="x" onclick="closeAdminPanel()">✕</button><h3>🔔 Notifications · ${bp(j.unread||0)}টি অপঠিত</h3><div class="notification-list">${rows.map(x=>`<article role="button" tabindex="0" class="notification-item ${x.read_at?'read':'unread'}" style="cursor:pointer" onclick="openNotificationTarget(${Number(x.id)},'${esc(x.ref||'').replace(/'/g, '&#39;')}')" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openNotificationTarget(${Number(x.id)},'${esc(x.ref||'').replace(/'/g, '&#39;')}')}"><div><b>${esc(x.title)}</b><small>${new Date(x.created_at*1000).toLocaleString('bn-BD')}</small><p>${esc(x.message)}</p></div>${x.read_at?'<span class="read-label">দেখা হয়েছে</span>':`<button class="pill" onclick="event.stopPropagation();markNotificationRead(${Number(x.id)})">দেখেছি</button>`}</article>`).join('')||'<div class="empty-sheet">কোনো Notifications নেই</div>'}</div><button class="pill" onclick="markAllNotificationsRead()">সবগুলো দেখেছি</button></div>`;translateUi($('ab'))} async function markNotificationRead(id){await post('/api/notifications/read',{id});await openNotifications()} async function markAllNotificationsRead(){await post('/api/notifications/read',{});await openNotifications()}
async function refreshNotificationCount(){if(!me)return;const j=await jf('/api/notifications');if(j._ok)$('notifyCount').textContent=bp(j.unread||0)}
function orderLookupAmount(value){const raw=String(value??'').trim();if(!raw)return '—';const numeric=Number(raw.replace(/,/g,'').replace(/[^0-9.\-]/g,''));return Number.isFinite(numeric)?'৳ '+bp(numeric):esc(raw)}
async function runOrderLookup(){const input=$('orderQuickSearch'),q=input?.value?.trim()||'',box=$('quickOrderResults');if(!q)return toast('Enter Consignment ID or customer number',true);if(box){box.hidden=false;box.innerHTML='<div class="quick-order-empty">Searching order details…</div>'}let found=[],errors=[];for(const type of ['consignment','number']){const j=await jf('/api/order-lookup?type='+encodeURIComponent(type)+'&q='+encodeURIComponent(q));if(j._ok&&!j.error)found.push(...(j.rows||[]));else if(j.error)errors.push(j.error)}const seen=new Set();const rows=found.filter(r=>{const k=String(r.id||'')+'|'+String(r.cons_id||'')+'|'+String(r.product_code||'')+'|'+String(r.size||'');if(seen.has(k))return false;seen.add(k);return true});if(box){box.hidden=false;box.innerHTML=`<div class="quick-order-summary"><span>Search: ${esc(q)} · ${bp(rows.length)} matching order item(s)</span><button type="button" class="quick-order-close" onclick="closeOrderLookup()" aria-label="Close search results" title="Close">×</button></div>${rows.length?`<div style="overflow:auto"><table class="quick-order-table"><thead><tr><th>Order Date</th><th>Order ID</th><th>Consignment ID</th><th>Customer</th><th>Customer Number</th><th>Product Code</th><th>Size</th><th>Product Price</th><th>Delivery Charge</th><th>Advance Paid</th><th>Collectible Amount (Sheet)</th><th>Assigned To</th><th>Order Status</th><th>Delivery Status</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.order_date||'—')}</td><td>${bp(r.id)}</td><td>${esc(r.cons_id||'—')}</td><td>${esc(r.customer_name||'—')}</td><td>${esc(r.phone||'—')}</td><td>${esc(r.product_code||'—')}</td><td>${esc(r.size||'—')}</td><td>৳ ${bp(r.product_price||0)}</td><td>৳ ${bp(r.delivery_charge||0)}</td><td>৳ ${bp(r.advance_paid||0)}</td><td title="${r.collectible_source==='calculated'?'Calculated because no collectible amount was found in the uploaded sheet':'Collectible amount from uploaded sheet'}"><b>${orderLookupAmount(r.collectible_amount)}</b></td><td>${esc(r.assigned_to||'Unassigned')}</td><td>${esc(r.order_status||'—')}</td><td>${esc(r.delivery_status||'—')}</td></tr>`).join('')}</tbody></table></div>`:'<div class="quick-order-empty">No matching order found in Order Sheet history.</div>'}`}}

function closeOrderLookup(){const box=document.getElementById('quickOrderResults');if(box){box.hidden=true;box.innerHTML=''}const input=document.getElementById('orderQuickSearch');if(input)input.blur()}
document.addEventListener('click',function(e){const section=document.getElementById('quickOrderLookup');const box=document.getElementById('quickOrderResults');if(section&&box&&!section.contains(e.target))box.hidden=true;});
document.addEventListener('keydown',function(e){if(e.key==='Escape'){const box=document.getElementById('quickOrderResults');if(box&&!box.hidden)closeOrderLookup()}});
async function saveNotice(){const duration_seconds=Math.max(1,Math.min(300,Number($('noticeDuration')?.value||6)));const j=await post('/api/admin/notice',{text:$('noticeText')?.value||'',enabled:!!$('noticeEnabled')?.checked,duration_seconds});if(!j._ok)return toast(j.error||'নোটিশ সেভ হয়নি',true);toast('নোটিশ সেভ হয়েছে')}
window.stockReconQuery=window.stockReconQuery||'';window.stockReconSheet=window.stockReconSheet||'';window.stockReconOffset=window.stockReconOffset||0;
async function stockReconSearch(){window.stockReconQuery=String(document.getElementById('stockReconQuery')?.value||'').trim();window.stockReconSheet=document.getElementById('stockReconSheet')?.value||'';window.stockReconOffset=0;await panel('stockrecon')}
async function stockReconPage(delta){window.stockReconOffset=Math.max(0,Number(window.stockReconOffset||0)+delta);await panel('stockrecon')}
async function stockReconSave(btn,action){const tr=btn?.closest('tr[data-recon]');if(!tr)return;const payload={sheet:tr.dataset.sheet,row_number:Number(tr.dataset.row),code:tr.dataset.code,size:tr.dataset.size,counted_qty:Number(tr.querySelector('.stock-recon-count')?.value),note:tr.querySelector('.stock-recon-note')?.value||'',action};if(!Number.isInteger(payload.counted_qty)||payload.counted_qty<0)return toast('গণনা করা stock ০ বা তার বেশি পূর্ণ সংখ্যা হতে হবে।',true);if(action==='apply'&&!confirm('এই counted quantity Excel stock-এ লিখে বর্তমান quantity পরিবর্তন করবেন? আগে backup নেওয়া হবে।'))return;btn.disabled=true;try{const j=await post('/api/admin/stock-reconciliation',payload);if(!j._ok)throw Error(j.error||'Reconciliation save হয়নি');toast(action==='apply'?(j.stock_updated?'Stock count apply হয়েছে।':'Stock একই ছিল।'):'Physical count রেকর্ড হয়েছে; system stock বদলায়নি।');await panel('stockrecon')}catch(e){toast(e.message||'Stock reconciliation failed',true)}finally{btn.disabled=false}}
async function panel(t){if(!canOpenAdminTabJS(t))return;try{sessionStorage.setItem('activeAdminPanel',t)}catch(_){}
 let content='';
 if(t=='stockentry'&&(isAdminLevelJS()||canUseStockDashboard())){const d=await jf('/api/excel-stock/data?offset=0&limit=250');const sheets=d.sheets||[];content=`<div class="admin-section"><h3>➕ Stock Entry</h3><div class="admin-note">স্টক যোগের জন্য workbook থেকে sheet, product ও size বেছে নিন।</div>${d.has_file?`<div class="visual-card stock-entry-grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;align-items:end"><label>Sheet<select id="entrySheet" onchange="loadEntryRows()">${sheets.map(x=>`<option>${esc(x.name)}</option>`).join('')}</select></label><label>Product Code<input id="entryCode" list="entryCodes" placeholder="Type product code" autocomplete="off" oninput="previewEntryCatalogProduct()" onchange="loadEntrySizes()" onkeydown="if(event.key==='Enter'){event.preventDefault();loadEntrySizes()}"><datalist id="entryCodes"></datalist></label><button class="pill" type="button" onclick="loadEntrySizes()">🔎 Find Code</button><label>Size<select id="entrySize"></select></label><label>Quantity<input id="entryQty" type="number" min="1" value="1"></label><button class="pill" onclick="submitStockEntry()">➕ Add Stock</button><div id="entryCatalogPreview" class="entry-catalog-preview" style="grid-column:1/-1"></div></div>`:`<div class="admin-note">আগে Stock Management-এ Excel workbook যোগ করুন।`}</div>`;setTimeout(loadEntryRows,0)} if(t=='excelstock'){
  const d=await jf('/api/excel-stock/data?offset='+excelOffset+'&limit=250'+(excelSheet?'&sheet='+encodeURIComponent(excelSheet):'')+(excelQuery?'&q='+encodeURIComponent(excelQuery):'')).catch(()=>({has_file:false}));
  if(!d.has_file){content=`<div class="excel-hero"><div class="stepper"><span class="active">01 · ফাইল</span><i></i><span>02 · ট্যাব</span><i></i><span>03 · স্টক</span><i></i><span>04 · ইতিহাস</span></div><div class="welcome"><div class="eyebrow">ALIEN STOCK WORKSPACE</div><h2>আপনার স্টক, এক জায়গায়</h2><p>Excel workbook আপলোড করুন। প্রতিটি tab ও product code ওয়েবে sheet-এর মতো দেখাবে, এখান থেকে পরিমাণ বদলালেই মূল ফাইলে সেভ হবে।</p></div>${isAdminLevelJS()?`<label class="upload-card"><div class="upload-icon">↑</div><b>Excel ফাইল বেছে নিন</b><span>শুধু .xlsx · সর্বোচ্চ 25 MB · প্রথম সারিতে product code এবং size header</span><input id="excelUpload" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onchange="uploadExcelStock(this)"></label>`:`<div class="upload-card">এখনো Excel ফাইল যোগ করা হয়নি। Admin-কে workbook আপলোড করতে বলুন।</div>`}<div class="excel-tip"><b>ফাইলের গঠন</b><span>প্রথম কলামে Product Code, পরের কলামগুলোতে size (S, M, L, XL...)। আলাদা worksheet-গুলো dashboard-এ আলাদা tab হবে।</span></div></div>`}
  else {const today=d.today||{},sheets=d.sheets||[],sel=d.selected||{headers:[],rows:[]};excelSheet=sel.name||'__all__';const allTabs=[{name:'__all__',products:sheets.reduce((n,x)=>n+x.products,0),total:sheets.reduce((n,x)=>n+x.total,0),value:sheets.reduce((n,x)=>n+(x.value||0),0)},...sheets],maxTabStock=Math.max(1,...sheets.map(x=>x.total));let donutCursor=0;const donutColors=['#255a42','#c8a45f','#6f97b1','#bd7770','#8d83b8','#8ca66b'],donutGradient=sheets.map((x,i)=>{const start=donutCursor;donutCursor+=Math.max(0,x.total)/Math.max(1,allTabs[0].total)*100;return `${donutColors[i%donutColors.length]} ${start}% ${donutCursor}%`}).join(',')||'#dce8d9 0 100%';content=`<div class="excel-hero"><div class="stepper"><span class="done">01 · ফাইল</span><i></i><span class="done">02 · ট্যাব</span><i></i><span class="active">03 · স্টক</span><i></i><span>04 · ইতিহাস</span></div><div class="welcome"><div class="eyebrow">LIVE EXCEL WORKSPACE</div><h2>Alien Tees Stock Center</h2><p>স্টক পরিবর্তনগুলো Save Changes দিলে Excel-এ সেভ হবে।</p></div><div class="excel-stats"><div><small>মোট স্টক</small><b>${canSeeExcelStock()?bn(allTabs[0].total):'<span class="eye-mask" title="স্টক দেখার অনুমতি নেই">👁️</span>'}</b><em>ইউনিট</em></div><div><small>মোট প্রোডাক্ট</small><b>${bn(allTabs[0].products)}</b><em>সারি</em></div><div><small>স্টক ট্যাব</small><b>${bn(sheets.length)}</b><em>worksheets</em></div><div><small>মোট স্টক মূল্য</small><b>${canSeeStockValue()?'৳'+bp(allTabs[0].value):'<span class="eye-mask" title="স্টক মূল্য দেখার অনুমতি নেই">👁️</span>'}</b><em>পণ্যমূল্য × স্টক</em></div><div><small>অপেক্ষমাণ পরিবর্তন</small><b id="pendingCount">${bn(Object.keys(excelPending).length)}</b><em>সেভ না হওয়া</em></div></div><div style="display:${today?'block':'none'}"><div class="today-title">আজকের স্টক কার্যক্রম</div><div class="today-grid"><div><small>আজকের সেভ</small><b>${bn(today.saves||0)}</b></div><div><small>পরিবর্তিত স্টক ঘর</small><b>${bn(today.changes||0)}</b></div><div class="plus"><small>স্টক যোগ হয়েছে</small><b>+${bn(today.plus||0)}</b></div><div class="minus"><small>স্টক কমেছে</small><b>−${bn(today.minus||0)}</b></div></div></div><div class="stock-visuals" style="display:${canSeeExcelStock()?'grid':'none'}"><div class="visual-card"><div class="visual-title">ট্যাব অনুযায়ী স্টক</div>${sheets.map(s=>`<div class="visual-row"><span>${esc(s.name)}</span><div class="visual-track"><i style="width:${Math.max(1,Math.round(s.total/maxTabStock*100))}%"></i></div><b>${bn(s.total)}</b></div>`).join('')||'<small>কোনো ট্যাব নেই</small>'}</div><div class="visual-card donut-card"><div class="visual-title">স্টকের বণ্টন</div><div class="donut" style="--donut:${donutGradient}"><div><b>${bn(allTabs[0].total)}</b><small>মোট ইউনিট</small></div></div><div class="donut-caption">${sheets.slice(0,6).map((s,i)=>`<span><i style="background:${donutColors[i%donutColors.length]}"></i>${esc(s.name)} <b>${bn(s.total)}</b></span>`).join('')}</div></div></div><div class="excel-tools"><div class="sheet-tabs">${allTabs.map(s=>`<button class="${s.name===excelSheet?'on':''}" onclick="openExcelSheet('${esc(s.name)}')">${s.name==='__all__'?'সব ট্যাব':esc(s.name)} <small>${bn(s.products)}</small></button>`).join('')}</div><div class="excel-actions"><input id="excelQuery" value="${esc(excelQuery)}" placeholder="কোড / সাইজ খুঁজুন" onkeydown="if(event.key==='Enter')searchExcelStock()"><button class="pill" onclick="searchExcelStock()">খুঁজুন</button><button id="saveExcelBtn" class="pill save-excel" onclick="saveExcelChanges()" ${Object.keys(excelPending).length?'':'disabled'}>💾 Save Changes · ${bp(Object.keys(excelPending).length)}</button>${canSeeStockHistory()?'<button class="pill" onclick="excelStockHistory()">🕘 History</button>':''}${isSuperAdminJS()?'<button class="pill" onclick="restoreLatestExcelBackup()">↩ Restore</button>':''}${isAdminLevelJS()?'<a class="pill" href="/api/admin/excel-stock/download">⬇ Excel ডাউনলোড</a>':''}${isSuperAdminJS()?'<label class="pill upload-replace">↻ নতুন ফাইল<input type="file" accept=".xlsx" onchange="uploadExcelStock(this)"></label>':''}</div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th class="rowhead">#</th>${sel.name==='__all__'?'<th>ট্যাব</th>':''}<th>ছবি</th><th>PRODUCT CODE</th><th>PRODUCT NAME</th><th>💰 দাম</th><th>📦 স্টক মূল্য</th>${sel.headers.slice(1).map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${sel.rows.map(r=>`<tr><th class="rowhead">${r.row}</th>${sel.name==='__all__'?`<td>${esc(r.sheet)}</td>`:''}<td>${stockThumb(r.values[0])}</td><td><b class="code-cell">${esc(r.values[0])}</b></td><td class="stock-product-name">${esc(r.product?.title||findStockCatalogProduct(r.values[0])?.title||'—')}</td><td class="stock-price">${canSeeStockValue()?'৳'+bp(r.unit_price||0):'<span class="eye-mask" title="দাম দেখার অনুমতি নেই">👁️</span>'}</td><td class="stock-value">${canSeeStockValue()?'৳'+bp(r.stock_value||0):'<span class="eye-mask" title="স্টক মূল্য দেখার অনুমতি নেই">👁️</span>'}</td>${r.values.slice(1).map((v,j)=>{const i=j+1;const h=sel.headers[i]||'',editable=i>0&&h.toLowerCase()&&!['total','মোট','price','rate','name','product','title','sum','grand total','amount','cost'].includes(h.toLowerCase());return `<td>${editable&&canEditExcelStock()?`<input class="excel-cell" type="number" min="0" value="${canSeeExcelStock()?esc(v):''}" placeholder="${canSeeExcelStock()?'':'নতুন সংখ্যা'}" title="${canSeeExcelStock()?'':'বর্তমান স্টক গোপন'}" data-original="${canSeeExcelStock()?esc(v):''}" data-sheet="${esc(r.sheet||sel.name)}" data-row="${r.row}" data-code="${esc(r.values[0])}" data-size="${esc(h)}" onchange="queueExcelCell(this)">`:!canSeeExcelStock()&&editable?'<span class="eye-mask" title="স্টক দেখার অনুমতি নেই">👁️</span>':i===0?`<b class="code-cell">${esc(v)}</b>`:esc(v)}</td>`}).join('')}</tr>`).join('')||`<tr><td colspan="${sel.headers.length+(sel.name==='__all__'?3:2)}" class="empty-sheet">এই ট্যাবে কোনো সারি নেই</td></tr>`}</tbody></table></div><div class="table-foot">${bn(excelOffset+1)}–${bn(excelOffset+sel.rows.length)} / ${bn(sel.total_matches||0)} টি সারি <button class="pill" onclick="excelPage(-1)" ${excelOffset===0?"disabled":""}>← আগের</button> <button class="pill" onclick="excelPage(1)" ${excelOffset+sel.rows.length>=(sel.total_matches||0)?"disabled":""}>পরের →</button></div><div class="backup-note">🛡️ পূর্ণ ব্যাকআপে অ্যাপের ইউজার/অনুমতি/সেটিংস/হিস্ট্রি, বর্তমান Excel এবং আগের Excel snapshot এক ZIP ফাইলে থাকবে।</div></div>`}
 }
 if(t=='cat'||t=='sheet'){
  const j=await jf('/api/admin/sources'),L=j[t=='cat'?'catalog':'sheet'],m=j.meta;
  content+=(t=='cat'?`<div class="admin-note">লোড: <b>${bn(m.pct)}%</b> · ${esc(m.stage)} · মোট <b>${bn(m.products)}</b>টা প্রোডাক্ট ${m.ai_ready?'·  AI Ready':''} ${m.cat_err.length?'· ⚠ ফেইল: '+esc(m.cat_err.join(', ')):''}</div>`:`<div class="admin-note">শিট "Anyone with the link can view" করা থাকতে হবে। প্রথম কলাম = কোড, বাকি কলাম = সাইজ। সব ট্যাব নিজে থেকে আসবে।</div>`)+
  `<div class="admin-table-wrap"><table>${L.map(s=>`<tr><td><b>${esc(s.name)}</b><br><small style="word-break:break-all;color:#888">${esc(s.url)}</small></td><td>${isSuperAdminJS()?`<button class="pill" onclick="rmSrc(${s.id},'${t}')">রিমুভ</button>`:'শুধু Super Admin'}</td></tr>`).join('')||'<tr><td style="color:#888">এখনো কিছু যোগ করা হয়নি</td></tr>'}</table></div>
  <div class="row"><input id="sn" placeholder="নাম (যেমন Alien Tees)"><input id="su" placeholder="লিংক পেস্ট করুন" style="flex:1;min-width:200px"><button class="pill" onclick="addSrc('${t=='cat'?'catalog':'sheet'}','${t}')">+ যোগ করুন</button></div><button class="pill" onclick="post('/api/admin/refresh').then(()=>toast('রিফ্রেশ শুরু'))">↻ এখনই রিফ্রেশ</button>`;
 }
 if(t==='stockrecon'&&(isAdminLevelJS()||canUseStockDashboard()||allowed('stock_manage')||allowed('stock_set')||allowed('stock_add')||allowed('stock_reduce')||allowed('stock_history_view'))){
  const rq=new URLSearchParams({q:window.stockReconQuery||'',sheet:window.stockReconSheet||'',offset:String(window.stockReconOffset||0),limit:'150'});const sj=await jf('/api/admin/stock-reconciliation?'+rq.toString());
  if(!sj._ok){content+=`<div class="admin-section erp-module"><div class="dash-hero"><div><h3>📦 Stock Reconciliation</h3><small>Excel stock workbook unavailable or permission missing.</small></div><button class="pill" onclick="panel('dashboard')">← Dashboard</button></div><p>${esc(sj.error||'Stock reconciliation could not load.')}</p></div>`}
  else {const moneyQty=n=>Number(n||0).toLocaleString('en-US');const canApply=isAdminLevelJS()||allowed('stock_manage')||allowed('stock_set');content+=`<div class="admin-section erp-module stock-recon-module"><div class="dash-hero"><div><h3>📦 Stock Reconciliation</h3><small>Excel system quantity-এর সঙ্গে হাতে গোনা physical stock মিলিয়ে দেখুন। Record Count শুধু audit করে; Apply Count workbook-এ quantity পরিবর্তন করে এবং আগে backup নেয়।</small></div><div class="dash-actions"><button class="pill" onclick="panel('dashboard')">← Dashboard</button><button class="pill" onclick="panel('stockrecon')">↻ Refresh</button></div></div><div class="excel-stats"><div><small>Matched stock rows</small><b>${bn(sj.total||0)}</b><em>Code + size variants</em></div><div><small>Unapplied variances</small><b>${bn(sj.unresolved_variances||0)}</b><em>Counted, not applied</em></div><div><small>Showing</small><b>${bn((sj.rows||[]).length)}</b><em>এই পৃষ্ঠায়</em></div></div><div class="excel-tools"><div class="excel-actions"><input id="stockReconQuery" placeholder="Product name / code / size / sheet খুঁজুন" value="${esc(window.stockReconQuery||'')}" onkeydown="if(event.key==='Enter'){event.preventDefault();stockReconSearch()}"><select id="stockReconSheet" onchange="stockReconSearch()"><option value="">সব Sheet</option>${(sj.sheets||[]).map(x=>`<option value="${esc(x)}" ${window.stockReconSheet===x?'selected':''}>${esc(x)}</option>`).join('')}</select><button class="pill" onclick="stockReconSearch()">🔎 Search</button><button class="pill" onclick="window.stockReconQuery='';window.stockReconSheet='';window.stockReconOffset=0;panel('stockrecon')">Reset</button></div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Sheet</th><th>Product</th><th>Code</th><th>Size</th><th>System Qty</th><th>Physical Count</th><th>Note</th><th>Action</th></tr></thead><tbody>${(sj.rows||[]).map(r=>`<tr data-recon data-sheet="${esc(r.sheet)}" data-row="${Number(r.row_number)}" data-code="${esc(r.code)}" data-size="${esc(r.size)}"><td>${esc(r.sheet)} <small>#${bn(r.row_number)}</small></td><td><b>${esc(r.title)}</b></td><td>${esc(r.code)}</td><td>${esc(r.size)}</td><td>${bn(r.system_qty)}</td><td><input class="stock-recon-count" type="number" min="0" step="1" inputmode="numeric" value="${Number(r.system_qty)||0}" style="width:100px;min-width:75px"></td><td><input class="stock-recon-note" maxlength="500" placeholder="Optional note" style="min-width:140px"></td><td class="stock-recon-actions"><button class="pill" onclick="stockReconSave(this,'counted')">Record</button>${canApply?`<button class="pill" onclick="stockReconSave(this,'apply')">Apply</button>`:''}</td></tr>`).join('')||'<tr><td colspan="8">কোনো stock row পাওয়া যায়নি।</td></tr>'}</tbody></table></div><div class="table-foot">${bn(Number(sj.offset||0)+1)}–${bn(Number(sj.offset||0)+(sj.rows||[]).length)} / ${bn(sj.total||0)} <button class="pill" ${Number(sj.offset||0)<=0?'disabled':''} onclick="stockReconPage(-150)">← Previous</button> <button class="pill" ${!sj.has_more?'disabled':''} onclick="stockReconPage(150)">Next →</button></div><div class="chart-card" style="margin-top:16px"><h4>🧾 Recent Reconciliation Log</h4><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Date</th><th>Sheet / Code</th><th>Size</th><th>System</th><th>Counted</th><th>Variance</th><th>Action</th><th>User</th><th>Note</th></tr></thead><tbody>${(sj.history||[]).map(x=>`<tr><td>${new Date(Number(x.created_at||0)*1000).toLocaleString()}</td><td>${esc(x.sheet)} / ${esc(x.product_code)} #${bn(x.row_number)}</td><td>${esc(x.size)}</td><td>${bn(x.system_qty)}</td><td>${bn(x.counted_qty)}</td><td style="font-weight:800;color:${Number(x.variance)===0?'#177245':Number(x.variance)<0?'#b42318':'#9a6700'}">${Number(x.variance)>0?'+':''}${bn(x.variance)}</td><td>${x.action==='apply'?'Applied':'Count recorded'}</td><td>${esc(x.username)}</td><td>${esc(x.note||'')}</td></tr>`).join('')||'<tr><td colspan="9">No reconciliation records yet.</td></tr>'}</tbody></table></div></div></div>`}
 }
 if(t=='cashclosing'&&isAdminLevelJS()){
  const d=await jf('/api/admin/cash-closing?date='+encodeURIComponent(window.cashClosingDate||''));
  if(!d._ok){content+=`<div class="admin-section"><h3>💵 Daily Cash Closing</h3><p>${esc(d.error||'Cash close data could not be loaded.')}</p><button class="pill" onclick="panel('cashclosing')">Retry</button></div>`}
  else {window.cashClosingDate=d.date;const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});content+=`<div class="admin-section erp-module"><div class="dash-hero"><div><h3>💵 Daily Cash Closing</h3><small>দিনের নগদ বিক্রি, অনুমোদিত নগদ খরচ ও supplier cash payment মিলিয়ে closing balance দেখুন।</small></div><button class="pill" onclick="panel('dashboard')">← Dashboard</button></div><div class="visual-card" style="padding:16px;margin-top:12px"><div class="co-grid"><label>Closing Date<input id="cashCloseDate" type="date" value="${esc(d.date)}" onchange="cashClosingReload()"></label><label>Opening Cash (৳)<input id="cashCloseOpening" type="number" min="0" step="0.01" value="${Number(d.opening_cash||0)}" oninput="cashClosingPreview()"></label><label>Actual Counted Cash (৳)<input id="cashCloseActual" type="number" min="0" step="0.01" value="${Number(d.actual_cash||0)}" oninput="cashClosingPreview()"></label><label>Expected Cash (auto)<div id="cashCloseExpected" style="padding:12px;border:1px solid #d6e8dc;border-radius:9px;background:#f5fbf7;font-weight:850">${money(d.expected_cash)}</div></label><label>Difference<div id="cashCloseDifference" style="padding:12px;border:1px solid #d6e8dc;border-radius:9px;background:#f5fbf7;font-weight:800">${Number(d.difference||0)<0?'Short: ':'Over: '}${money(Math.abs(Number(d.difference||0)))}</div></label><label class="wide">Notes / mismatch reason<textarea id="cashCloseNotes" maxlength="2000" placeholder="ক্যাশ কম/বেশি হলে কারণ লিখুন">${esc(d.notes||'')}</textarea></label></div><div class="order-summary-grid"><div class="visual-card" style="padding:12px"><small>Cash POS Sales · ${bn(d.cash_sale_count||0)} transactions</small><h3 id="cashCloseSales" data-value="${Number(d.cash_sales||0)}">${money(d.cash_sales)}</h3></div><div class="visual-card" style="padding:12px"><small>Approved Cash Expenses · ${bn(d.cash_expense_count||0)} entries</small><h3 id="cashCloseExpenses" data-value="${Number(d.cash_expenses||0)}">− ${money(d.cash_expenses)}</h3></div><div class="visual-card" style="padding:12px"><small>Supplier Cash Payments · ${bn(d.cash_supplier_payment_count||0)} payments</small><h3 id="cashCloseSupplier" data-value="${Number(d.cash_supplier_payments||0)}">− ${money(d.cash_supplier_payments)}</h3></div><div class="visual-card" style="padding:12px"><small>Previous Closing</small><h3>${esc(d.previous_close_date||'প্রথম closing')}</h3><small>আগের দিনের Actual Cash opening হিসেবে এসেছে</small></div></div><p class="admin-note">Expected = Opening + Cash POS collection − Approved Cash Expenses − Supplier Cash Payments. Card/Mobile Banking বিক্রি cash drawer-এর হিসাবে যোগ করা হয়নি। এন্ট্রি দিনে পরে পরিবর্তিত হলে re-close করুন।</p><button id="cashCloseSave" class="pill" onclick="saveCashClosing()">💾 Close Day & Save</button> <button class="pill" onclick="cashClosingReload()">↻ Reload</button></div></div>`;setTimeout(cashClosingPreview,0)}
 }
 if(t=='dashboard'){
  const fin=await jf('/api/admin/business-dashboard').catch(()=>({ok:false}));
  const j=await jf('/api/admin/dashboard');const ls=j.low_stock||[],os=j.out_of_stock||[],acts=j.action_stats||{},users=j.byuser||[],days=j.daily||[],tops=j.top_queries||[];
  const maxAct=Math.max(1,...Object.values(acts).map(Number));
  const maxUser=Math.max(1,...users.map(x=>Number(x.n||0)));
  const maxDay=Math.max(1,...days.map(x=>Number(x.n||0)));
  const orderMini=await jf('/api/admin/order-history/summary');const returnMini=await jf('/api/admin/return-history'); const labels={text_search:'টেক্সট সার্চ',image_search:'ইমেজ সার্চ',download:'HD ডাউনলোড',copy:'HD কপি',favorite_add:'ফেভারিট',product_report:'রিপোর্ট'};
  content=`<div class="admin-section"><div class="dash-hero"><div><h3> আজকের Dashboard</h3><small>লাইভ স্টেটাস · মোট ${bn(j.total_products||0)}টা product · ${j.ai_ready?'🤖 AI Ready':'⏳ AI প্রস্তুত হচ্ছে'} · ${esc(j.stage||'')}</small></div></div>
  <div class="dashgrid"><div class="stat"><small>সার্চ করা অর্ডার</small><b>${bn(orderMini.total_orders||0)}</b></div><div class="stat"><small>মোট রিটার্ন</small><b>${bn(returnMini.total||0)}</b></div><div class="stat"><small>আজকের Activity</small><b>${bn(j.today||0)}</b></div><div class="stat"><small>Search</small><b>${bn(j.searches||0)}</b></div><div class="stat"><small>HD Download</small><b>${bn(j.downloads||0)}</b></div><div class="stat"><small>HD Copy</small><b>${bn(j.copies||0)}</b></div><div class="stat"><small>Open Report</small><b>${bn(j.reports||0)}</b></div><div class="stat"><small>Users / Staff / CS</small><b>${bn(j.users||0)}</b></div><div class="stat"><small>কম স্টক</small><b>${bn(ls.length)}</b></div><div class="stat"><small>স্টক নেই</small><b>${bn(os.length)}</b></div></div>
  <div class="chart-grid"><div class="chart-card"><h4>আজকের কাজের পরিমাণ</h4>${Object.entries(labels).map(([k,v])=>`<div class="barrow"><span>${v}</span><div class="bartrack"><div class="barfill" style="width:${Math.round((Number(acts[k]||0)/maxAct)*100)}%"></div></div><b>${bn(acts[k]||0)}</b></div>`).join('')}</div>
  <div class="chart-card"><h4> গত ২৪ ঘণ্টায় User Activity</h4>${users.slice(0,10).map(x=>`<div class="barrow"><span>${esc(x.username)}</span><div class="bartrack"><div class="barfill" style="width:${Math.round((Number(x.n||0)/maxUser)*100)}%"></div></div><b>${bn(x.n||0)}</b></div>`).join('')||'<small style="color:#999">কোনো activity নেই</small>'}</div>
  <div class="chart-card"><h4> গত ৭ দিনের Activity</h4><div class="mini-line">${days.map(x=>`<div class="daycol"><b>${bn(x.n||0)}</b><div class="daybar" title="${esc(x.label||'')} · ${bp(x.n||0)} activity" style="height:${Math.max(4,Math.round((Number(x.n||0)/maxDay)*105))}px"></div><small>${esc(x.label||'')}</small></div>`).join('')}</div></div>
  <div class="chart-card"><h4> বেশি খোঁজা Query</h4><div class="toplist">${tops.map(x=>`<div class="item"><span>${esc(x.query)}</span><b>${bn(x.n||0)}</b></div>`).join('')||'<small style="color:#999">কোনো search data নেই</small>'}</div></div></div>
  <div class="stock-columns"><div class="userbox"><b>⚠ কম স্টক (${bn(ls.length)})</b><div class="row"><input id="stockth" type="number" min="1" max="99" value="${j.low_stock_threshold||2}" style="width:120px"><button class="pill" onclick="saveStockThreshold()">Threshold সেভ</button></div><small style="color:#777">এই সংখ্যার সমান বা কম stock হলে “কম স্টক” দেখাবে।</small><div class="toplist">${ls.slice(0,25).map(x=>`<div class="item"><span><b>${esc(x.id||'—')}</b> — ${esc(x.title)}</span><b>${bn(x.total||0)}</b></div>`).join('')||'<small style="color:#999">কম স্টকের product নেই</small>'}</div></div>
  <div class="userbox"><b> স্টক নেই (${bn(os.length)})</b><div class="toplist">${os.slice(0,25).map(x=>`<div class="item"><span><b>${esc(x.id||'—')}</b> — ${esc(x.title)}</span><span>Out</span></div>`).join('')||'<small style="color:#999">Out of Stock product নেই</small>'}</div></div></div></div>`;
  if(fin.ok){const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});const companies=fin.by_company||[];const due=fin.top_customer_due||[];const top=fin.top_products||[];const coverage=Number(fin.cost_coverage||0);const profitText=fin.estimated_gross_profit===null?'ক্রয়মূল্য মেলেনি':money(fin.estimated_gross_profit);content=`<div class="admin-section business-dashboard"><div class="dash-hero"><div><h3>📊 Business & Finance Overview</h3><small>${esc(fin.as_of||'')} · current month: ${esc(fin.month_start||'')} থেকে আজ পর্যন্ত</small></div><div style="display:flex;gap:8px;flex-wrap:wrap"><button class="pill" onclick="panel('cashclosing')">💵 Daily Cash Closing</button><button class="pill" onclick="panel('stockrecon')">📦 Stock Reconciliation</button><button class="pill" onclick="panel('dashboard')">↻ Refresh</button></div></div><div class="bizdash-grid"><div class="bizdash-tile primary"><small>আজকের বিক্রি</small><b>${money(fin.today_sales)}</b><em>${bn(fin.today_orders||0)} orders · আদায় ${money(fin.today_collected)}</em></div><div class="bizdash-tile"><small>এই মাসের বিক্রি</small><b>${money(fin.month_sales)}</b><em>${bn(fin.month_orders||0)} orders</em></div><div class="bizdash-tile"><small>Customer Due (POS)</small><b>${money(fin.customer_due)}</b><em>সকল বকেয়া বিক্রির যোগফল</em></div><div class="bizdash-tile"><small>Supplier Due</small><b>${money(fin.supplier_due)}</b><em>কেনাকাটা − নথিভুক্ত পেমেন্ট</em></div><div class="bizdash-tile"><small>এই মাসের Purchase</small><b>${money(fin.month_purchases)}</b><em>${bn(fin.purchase_count||0)} invoices</em></div><div class="bizdash-tile"><small>Approved Expenses</small><b>${money(fin.approved_expenses)}</b><em>Pending ${money(fin.pending_expenses)} (${bn(fin.pending_expense_count||0)}টি)</em></div><div class="bizdash-tile"><small>Matched-item Gross Profit Estimate</small><b>${profitText}</b><em>Cost coverage: ${coverage}% units · ${Number(fin.cost_coverage_sales||0)}% sales value</em></div><div class="bizdash-tile"><small>Estimated COGS</small><b>${fin.estimated_cogs===null?'—':money(fin.estimated_cogs)}</b><em>Matching product + size cost</em></div><div class="bizdash-tile"><small>Net after approved expenses · Estimate</small><b>${fin.estimated_net_after_approved_expenses===null?'—':money(fin.estimated_net_after_approved_expenses)}</b><em>${fin.estimated_net_after_approved_expenses===null?'Hidden until cost coverage is ≥99.9%':'Cost coverage '+coverage+'% · matched POS items'}</em></div></div><div class="bizdash-columns"><div class="chart-card"><h4>🏷 কোম্পানি অনুযায়ী বিক্রি · এই মাস</h4>${companies.map(x=>`<div class="barrow"><span>${esc(x.company_key==='bodyfitters'?'Bodyfitters':'Alien Tees')}</span><div class="bartrack"><div class="barfill" style="width:${Math.max(0,Math.min(100,Number(x.sales||0)/Math.max(1,...companies.map(y=>Number(y.sales||0)))*100))}%"></div></div><b>${money(x.sales)}</b></div>`).join('')||'<small>এই মাসে POS sale নেই</small>'}<small class="bizdash-note">${esc(fin.profit_note||'')}</small></div><div class="chart-card"><h4>🔥 বেশি বিক্রি হওয়া পণ্য</h4>${top.map(x=>`<div class="item"><span><b>${esc(x.name)}</b><small style="display:block;color:#78867e">${esc(x.code||'')} · ${bn(x.units)} pcs · stock ${bn(x.stock_on_hand||0)} · ≈ ${x.estimated_days_left===null?'—':bn(Math.round(Number(x.estimated_days_left)))} days left</small></span><b>${money(x.sales)}</b></div>`).join('')||'<small>এই মাসে কোনো sale item নেই</small>'}</div><div class="chart-card"><h4>💳 সবচেয়ে বেশি Customer Due</h4>${due.map(x=>`<div class="item"><span><b>${esc(x.name||'Customer')}</b><small style="display:block;color:#78867e">${esc(x.phone||'')} · ${bn(x.orders||0)} orders</small></span><b>${money(x.due)}</b></div>`).join('')||'<small>POS-এ বকেয়া নেই</small>'}</div><div class="chart-card"><h4>⚠ Smart Alerts</h4>${(fin.smart_alerts||[]).map(a=>`<button class="biz-alert-row ${esc(a.level||'medium')}" style="display:flex;width:100%;align-items:flex-start;justify-content:space-between;gap:12px;text-align:left;border:0;border-bottom:1px solid #e8eee9;background:transparent;padding:10px 2px;cursor:pointer;color:inherit" onclick="panel('${esc(a.panel||'dashboard')}')"><span><b>${esc(a.title||'Alert')}</b><small style="display:block;color:#718076;margin-top:3px;white-space:normal">${esc(a.detail||'')}</small></span><span aria-hidden="true">›</span></button>`).join('')||'<div class="item"><span>✅ No urgent alerts</span><b>0</b></div>'}<small class="bizdash-note">Profit and inventory warnings use recorded data only. Fix missing costs and unapplied stock counts before relying on estimates.</small></div></div></div>`+content;}
 }
 if(t=='reports'){const j=await jf('/api/admin/reports');content+=`<div class="admin-section"><h3> Product Reports</h3><div class="admin-table-wrap"><table><tr><th>সময়</th><th>ইউজার</th><th>Product</th><th>কারণ</th><th>Status</th></tr>${(j.reports||[]).map(x=>{const d=new Date((x.created_at||0)*1000);return `<tr><td>${esc(d.toLocaleString('en-US'))}</td><td>${esc(x.username)}</td><td>${esc(x.product_key)}</td><td>${esc(x.reason)}</td><td>${esc(x.status)}</td></tr>`}).join('')||'<tr><td colspan="5">কোনো রিপোর্ট নেই</td></tr>'}</table></div></div>`}
 if(t=='price'){const h=await jf('/api/admin/price-history?limit=60');content+=`<div class="admin-section"><h3> দাম এডিট</h3><div class="admin-note">যে product-এর দাম বদলাবেন, তার <b>photo + ID + পুরনো/নতুন দাম + কে বদলেছে + সময়</b> নিচের History-তে থাকবে।</div><input id="pf" placeholder="কোড বা নাম খুঁজুন" oninput="priceRows()" style="width:100%"><div id="pr" style="max-height:52vh;overflow:auto"></div><h4 style="margin-top:18px"> Price Change History</h4><div class="price-history">${(h.history||[]).map(x=>{const d=new Date((x.created_at||0)*1000);return `<div class="ph-row"><img src="${im(x.img)}" loading="lazy"><div class="ph-info"><b>${esc(x.display_id||'ID নেই')} — ${esc(x.title||'')}</b><small>${esc(x.store||'')} · ${esc(x.username||'')} · ${esc(d.toLocaleString('en-US'))}</small><div><s>৳${bp(x.old_price||0)}</s> → <strong>৳${bp(x.new_price||0)}</strong> ${x.action==='restore'?'<em>আসল দাম ফিরেছে</em>':''}</div></div></div>`}).join('')||'<small style="color:#999">এখনো কোনো price change history নেই</small>'}</div></div>`}
 if(t=='photosettings'){const d=await jf('/api/admin/photo-download-settings');content+=`<div class="admin-section"><h3>📸 Photo Download Control</h3><p class="admin-note">ছবিতে Product Code/Price দেখানো এবং কারা ছবি ডাউনলোড করতে পারবে—এখান থেকে নিয়ন্ত্রণ করুন। সার্ভারেও অনুমতি যাচাই হবে।</p><label style="display:flex;align-items:center;gap:9px;margin:14px 0"><input id="photoShowCode" type="checkbox" ${d.show_code?'checked':''}> ডাউনলোড ছবিতে CODE দেখাবে</label><label style="display:flex;align-items:center;gap:9px;margin:14px 0"><input id="photoShowPrice" type="checkbox" ${d.show_price?'checked':''}> ডাউনলোড ছবিতে PRICE দেখাবে</label><label style="display:grid;gap:7px;margin:14px 0">কারা ছবি ডাউনলোড করতে পারবে?<select id="photoAccess" style="max-width:360px;padding:10px;border:1px solid #ddd;border-radius:9px;font:inherit"><option value="all" ${d.access==='all'?'selected':''}>সব ইউজার (বর্তমান permission সাপেক্ষে)</option><option value="logged" ${d.access==='logged'?'selected':''}>শুধু লগইন করা ইউজার</option><option value="admin" ${d.access==='admin'?'selected':''}>শুধু Admin / Super Admin</option></select></label><button class="pill" onclick="savePhotoDownloadSettings()">💾 সেটিংস সেভ করুন</button></div>`}
 if(t=='couriers'&&(isAdminLevelJS()||allowed('courier_view')||allowed('courier_update'))){content+=await courierPanel()} if(t=='posorders'&&(isAdminLevelJS()||allowed('pos_orders')||allowed('pos_delete'))){openAdminFrame(true);content+=await posOrdersPanel()} if(t=='company'&&isAdminLevelJS()){content+=await companyPanel()} if(t=='notice'){const n=await jf('/api/notice');content+=`<div class="admin-section"><h3>📢 সাইট নোটিশ</h3><p class="admin-note">নোটিশ ও প্রদর্শনের সময় ঠিক করুন।</p><textarea id="noticeText" rows="5" maxlength="2000" style="width:100%;padding:12px;border:1px solid #ddd;border-radius:10px;font:inherit;box-sizing:border-box" placeholder="নোটিশ লিখুন...">${esc(n.text||'')}</textarea><label style="display:flex;gap:8px;align-items:center;margin:12px 0"><input id="noticeEnabled" type="checkbox" ${n.enabled?'checked':''}> নোটিশ চালু থাকবে</label><label style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:12px 0">⏱️ নোটিশ কত সেকেন্ড থাকবে <input id="noticeDuration" type="number" min="1" max="300" value="${Number(n.duration_seconds||6)}" style="width:100px;padding:9px;border:1px solid #ddd;border-radius:8px;font:inherit"> <small>১–৩০০ সেকেন্ড</small></label><button class="pill" onclick="saveNotice()">নোটিশ সেভ করুন</button></div>`}
 if(t=='design'){const T=await jf('/api/theme'),v=k=>T[k]||getComputedStyle(document.documentElement).getPropertyValue('--'+k).trim()||'#000000';content+=`<div class="admin-section"><h3> ডিজাইন ও ভাষা</h3><p>এখান থেকে পুরো সাইটের ভাষা নিয়ন্ত্রণ করুন।</p><h4>রেডিমেড টেমপ্লেট</h4><div class="tb">${Object.keys(PRESETS).map(k=>`<button onclick='setPre("${k}")'>${k}</button>`).join('')}</div><div class="row">${['acc','bg','ink','head','card'].map(k=>`<label>${{acc:'অ্যাকসেন্ট',bg:'ব্যাকগ্রাউন্ড',ink:'লেখা',head:'হেডার',card:'কার্ড'}[k]} <input type="color" id="c_${k}" value="${v(k).length==7?v(k):'#000000'}"></label>`).join('')}</div><div class="row"><label>ওয়েবসাইটের ভাষা<select id="c_language"><option value="bn" ${(T.language||'en')==='bn'?'selected':''}>বাংলা</option><option value="en" ${(T.language||'en')==='en'?'selected':''}>Professional English</option></select></label><select id="c_font">${['Hind Siliguri','Noto Sans Bengali','Tiro Bangla','Baloo Da 2'].map(f=>`<option ${T.font==f?'selected':''}>${f}</option>`).join('')}</select><input id="c_title" placeholder="সাইটের নাম" value="${esc(T.title||'')}"><input id="c_tag" placeholder="ট্যাগলাইন" value="${esc(T.tag||'')}"></div><button class="pill" onclick="saveTheme()">সেভ করুন</button> <button class="pill" onclick="location.reload()">রিসেট প্রিভিউ</button></div>`}
 if(t=='guest'){const j=await jf('/api/admin/guest-perms');const gp=j.permissions||guestPerms||{};content+=`<div class="admin-section"><h3> Guest / Customer Access Control</h3><div class="admin-note">লগইন ছাড়া Customer কী কী করতে পারবে এখান থেকে নিয়ন্ত্রণ করুন। checkbox বদলে সেভ করুন।</div><div class="perms guest-perms">${Object.entries(GUEST_PERM_LABELS).map(([k,v])=>`<label><input type="checkbox" id="guest_${k}" ${gp[k]?'checked':''}> ${v}</label>`).join('')}</div><button class="pill" onclick="saveGuestPerms()"> Guest Access সেভ করুন</button><div class="admin-section" style="margin-top:18px"><h3>👨‍💼 CS Access</h3><div class="admin-note"><b>CS role-এর permission User Management থেকে checkbox দিয়ে নিয়ন্ত্রণ করুন।</div></div></div>`}
 if(t=='orders'&&(isAdminLevelJS()||allowed('order_view'))){
 const oq=new URLSearchParams({q:window.orderSearch||'',from:window.orderFrom||'',to:window.orderTo||''}); const [j,summary]=await Promise.all([jf('/api/admin/order-history?'+oq.toString()),jf('/api/admin/order-history/summary?'+oq.toString())]); const rows=j.rows||[]; const money=n=>'৳'+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});
 content+=`<div class="admin-section order-history"><div class="dash-hero"><div><h3>🧾 Order History Record</h3><small>প্রতিদিনের Excel রিপোর্ট এখানে যোগ হবে; আগের রেকর্ড মুছে বা replace হবে না।</small></div>${(meRole==='admin'||meRole==='super_admin'||perms.order_sheet_upload)?`<label class="pill upload-replace order-upload-pill">＋ আজকের Excel আপলোড<input type="file" accept=".xlsx,.xlsm" onchange="uploadOrderHistory(this)"></label>`:'<span class="pill upload-locked">🔒 Order Upload permission দরকার</span>'}</div><div class="admin-note">Admin/Super Admin এবং Super Admin থেকে Order Sheet Upload permission পাওয়া ইউজাররা রিপোর্ট আপলোড করতে পারবেন। সিলেক্ট করা বা সব রেকর্ড ডিলিট করার সুবিধা শুধু Super Admin-এর জন্য। Product Code দিয়ে ক্যাটালগে ছবি মিললে ছবিও দেখাবে। Excel-এ Customer Name, Phone, Product Code কলাম থাকা প্রয়োজন।</div><div class="excel-stats"><div><small>মোট অর্ডার</small><b>${bn(summary.total_orders||j.total||0)}</b><em>বর্তমান সার্চ/তারিখ অনুযায়ী</em></div><div><small>মোট Collectible</small><b>${money(summary.collectible_total||0)}</b><em>সব অর্ডারের যোগফল</em></div><div><small>মোট বিল (দাম + ডেলিভারি)</small><b>${money(summary.total_bill||0)}</b><em>Product Price ও Delivery Charge-এর যোগফল</em></div><div><small>Assigned To</small><b>${bn((summary.assigned||[]).length)}</b><em>অ্যাসাইনি সংখ্যা</em></div><div><small>Product Code</small><b>${bn((summary.codes||[]).length)}</b><em>আলাদা কোড</em></div></div><div class="order-summary-grid"><div class="visual-card"><h3>👤 Assigned To অনুযায়ী অর্ডার</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Assigned To</th><th>অর্ডার</th><th>Collectible মোট</th></tr></thead><tbody>${(summary.assigned||[]).map(x=>`<tr><td><b>${esc(x.name)}</b></td><td>${bn(x.orders)}</td><td>${money(x.collectible)}</td></tr>`).join('')||'<tr><td colspan="3">ডাটা নেই</td></tr>'}</tbody></table></div></div><div class="visual-card"><h3>🏷️ Product Code অনুযায়ী অর্ডার</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Product Code</th><th>অর্ডার</th><th>Collectible মোট</th></tr></thead><tbody>${(summary.codes||[]).map(x=>`<tr><td><b>${esc(x.code)}</b></td><td>${bn(x.orders)}</td><td>${money(x.collectible)}</td></tr>`).join('')||'<tr><td colspan="3">ডাটা নেই</td></tr>'}</tbody></table></div></div></div><div class="order-lookup"><label>📦 Consignment Search<input id="orderConsQuery" placeholder="Enter Consignment ID" onkeydown="if(event.key==='Enter')searchOrderLookup()"></label><label>📱 Number Search<input id="orderPhoneQuery" inputmode="tel" placeholder="Enter phone number" onkeydown="if(event.key==='Enter')searchOrderLookup()"></label><button class="pill" onclick="searchOrderLookup()">🔎 Search</button><span>Matching orders and total order count will appear below.</span></div><div class="excel-tools"><div class="excel-actions"><input id="orderQuery" placeholder="Name / phone / code / consignment" value="${esc(window.orderSearch||'')}"><button class="pill" onclick="searchOrderHistory(1)">🔎 Search</button><button class="pill" onclick="window.orderSearch='';window.orderFrom='';window.orderTo='';panel('orders')">Reset</button><input id="orderFrom" type="date" value="${esc(window.orderFrom||'')}"><input id="orderTo" type="date" value="${esc(window.orderTo||'')}"><button class="pill" onclick="searchOrderHistory(1)">Date filter</button><a class="pill" href="/api/admin/order-history/export?q=${encodeURIComponent(window.orderSearch||'')}&from=${encodeURIComponent(window.orderFrom||'')}&to=${encodeURIComponent(window.orderTo||'')}">⬇ Excel রিপোর্ট</a></div></div><div class="excel-actions order-bulk-actions"><label class="pill"><input type="checkbox" id="orderSelectAll" onchange="toggleOrderSelectAll(this.checked)"> সব সিলেক্ট</label><button class="pill" onclick="deleteSelectedOrders()">🗑️ সিলেক্ট করা ডিলিট</button><button class="pill" onclick="deleteAllOrders()">সব অর্ডার ডিলিট</button><small>ডিলিট করার অনুমতি শুধু Super Admin-এর</small></div><div class="excel-table-wrap"><table class="excel-grid order-grid"><thead><tr><th>সিলেক্ট</th><th>ক্রমিক</th><th>তারিখ</th><th>উৎস</th><th>দায়িত্বপ্রাপ্ত</th><th>পেমেন্টের অবস্থা</th><th>অর্ডারের অবস্থা</th><th>ডেলিভারির অবস্থা</th><th>কনসাইনমেন্ট আইডি</th><th>ছবি</th><th>কাস্টমারের নাম</th><th>ফোন</th><th>প্রোডাক্ট কোড</th><th>সাইজ</th><th>পণ্যের দাম</th><th>ডেলিভারি</th><th>পরিশোধ</th><th>বাকি/Collectible</th><th>Rating</th><th>ঠিকানা</th><th>নোট</th><th>আপলোড</th></tr></thead><tbody>${rows.map(r=>`<tr><td><input type="checkbox" class="order-row-check" value="${Number(r.id)}" aria-label="অর্ডার ${Number(r.id)} সিলেক্ট"></td><td>${bn(r.id)}</td><td>${esc(r.order_date)}</td><td>${esc(r.source)}</td><td>${esc(r.assigned_to)}</td><td>${esc(r.payment_status)}</td><td>${esc(r.order_status)}</td><td>${esc(r.delivery_status)}</td><td>${esc(r.cons_id)}</td><td>${r.image_url?`<img class="order-thumb" src="${esc(r.image_url)}" alt="product">`:'—'}</td><td><b>${esc(r.customer_name)}</b></td><td>${esc(r.phone)}</td><td><button class="code-cell" style="border:0;background:transparent;color:#2563eb;text-decoration:underline;cursor:pointer;font-weight:800" onclick="selectOrderProduct(this,${Number(r.id)},'${esc(r.customer_name||'').replace(/'/g, "&#39;")}','${esc(r.product_code||'').replace(/'/g, "&#39;")}')">${esc(r.product_code)}</button></td><td>${esc(r.size)}</td><td>${esc(r.product_price)}</td><td>${esc(r.delivery_charge)}</td><td>${esc(r.advance_paid)}</td><td><b>${esc((r.collectible_amount||'').trim() || (Number(String(r.product_price||'0').replace(/[^0-9.\-]/g,''))+Number(String(r.delivery_charge||'0').replace(/[^0-9.\-]/g,''))-Number(String(r.advance_paid||'0').replace(/[^0-9.\-]/g,''))).toLocaleString('en-BD'))}</b></td><td>${esc(r.rating)}</td><td class="order-address">${esc(r.address)}</td><td>${esc(r.note)}</td><td>${esc(r.source_file)}<small class="order-uploader">${esc(r.uploaded_by)} · ${new Date(r.uploaded_at*1000).toLocaleDateString()}</small></td></tr>`).join('')||'<tr><td colspan="22" class="empty-sheet">কোনো অর্ডার রেকর্ড নেই। প্রথম Excel রিপোর্ট আপলোড করুন।</td></tr>'}</tbody></table></div><div class="table-foot">${bn(rows.length)}টি দেখানো · মোট ${bn(j.total||0)}টি রেকর্ড · নতুন অর্ডার আগে</div></div>`;
 }
 if(t=='returns'&&(isAdminLevelJS()||allowed('return_view'))){
 const rj=await jf('/api/admin/return-history?q='+encodeURIComponent(window.returnSearch||'')); const rm=n=>'৳'+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});
 content+=`<div class="admin-section order-history"><div class="dash-hero"><div><h3>↩ রিটার্ন হিস্ট্রি</h3><small>রিটার্নের Excel আপলোড করুন; মূল অর্ডার হিস্ট্রির সঙ্গে কনসাইনমেন্ট আইডি ও প্রোডাক্ট কোড মিলিয়ে দেখুন।</small></div>${(meRole==='admin'||meRole==='super_admin'||perms.return_sheet_upload)?`<label class="pill upload-replace return-upload-pill">＋ Return Excel আপলোড<input type="file" accept=".xlsx,.xlsm" onchange="uploadReturnHistory(this)"></label>`:'<span class="pill upload-locked">🔒 আপলোডের অনুমতি নেই</span>'}</div><div class="excel-stats"><div><small>মোট অর্ডার</small><b>${bn(rj.order_total||0)}</b><em>Main Order Sheet</em></div><div><small>মোট রিটার্ন</small><b>${bn(rj.total||0)}</b><em>Return Sheet rows</em></div><div><small>Assigned To</small><b>${bn((rj.assigned||[]).length)}</b><em>রিটার্ন পাওয়া অ্যাসাইনি</em></div><div><small>Return Reason</small><b>${bn((rj.reasons||[]).length)}</b><em>আলাদা কারণ</em></div></div><div class="order-summary-grid"><div class="visual-card"><h3>👤 কার কত রিটার্ন</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Assigned To</th><th>Return Count</th></tr></thead><tbody>${(rj.assigned||[]).map(x=>`<tr><td>${esc(x.name)}</td><td>${bn(x.returns)}</td></tr>`).join('')||'<tr><td colspan="2">ডাটা নেই</td></tr>'}</tbody></table></div></div><div class="visual-card"><h3>📦 রিটার্নের কারণ</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>কারণ</th><th>সংখ্যা</th></tr></thead><tbody>${(rj.reasons||[]).map(x=>`<tr><td>${esc(x.name)}</td><td>${bn(x.count)}</td></tr>`).join('')||'<tr><td colspan="2">ডাটা নেই</td></tr>'}</tbody></table></div></div></div><div class="excel-tools"><div class="excel-actions"><button class="pill" onclick="showManualReturnEntry()">＋ Return Entry</button><button class="pill" onclick="syncOrderSheetReturns()">🔄 Order Sheet থেকে API Return Sync</button><button class="pill" onclick="deleteSelectedReturns()">🗑️ সিলেক্টেড ডিলিট</button><button class="pill" onclick="deleteAllReturns()">🗑️ সব রিটার্ন ডিলিট</button><input id="returnQuery" placeholder="নাম / ফোন / প্রোডাক্ট কোড / কনসাইনমেন্ট আইডি / কারণ" value="${esc(window.returnSearch||'')}"><button class="pill" onclick="searchReturnHistory()">🔎 খুঁজুন</button><button class="pill" onclick="window.returnSearch='';panel('returns')">Reset</button><a class="pill" href="/api/admin/return-history/export.xlsx">Excel রিপোর্ট</a></div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th><input type="checkbox" onchange="toggleReturnSelectAll(this.checked)" title="সব সিলেক্ট"></th><th>ক্রমিক</th><th>দায়িত্বপ্রাপ্ত</th><th>কনসাইনমেন্ট আইডি</th><th>ঠিকানা</th><th>ফোন</th><th>গ্রাহকের নাম</th><th>প্রোডাক্ট কোড</th><th>সাইজ</th><th>আদায়কৃত টাকা</th><th>নোট</th><th>পাঠাও বিলের অবস্থা</th><th>রিটার্নের কারণ</th><th>মূল অর্ডার</th><th>অ্যাকশন</th></tr></thead><tbody>${(rj.rows||[]).map(r=>`<tr><td><input class="return-row-check" type="checkbox" value="${bn(r.id)}"></td><td>${bn(r.id)}</td><td>${esc(r.assigned_to)}</td><td>${esc(r.cons_id)}</td><td class="order-address">${esc(r.address)}</td><td>${esc(r.phone)}</td><td><b>${esc(r.customer_name)}</b></td><td><b>${esc(r.product_code)}</b></td><td>${esc(r.size)}</td><td>${esc(r.collected_amount)}</td><td>${esc(r.note)}</td><td>${esc(r.pathao_bill_status)}</td><td>${esc(r.return_reason)}</td><td><button class="pill" onclick="matchReturnOrder('${esc(r.cons_id||'').replace(/'/g, '&#39;')}','${esc(r.product_code||'').replace(/'/g, '&#39;')}')">অর্ডার দেখুন</button></td><td><button class="pill" onclick="deleteOneReturn(${bn(r.id)})">🗑️ Delete</button></td></tr>`).join('')||'<tr><td colspan="15" class="empty-sheet">কোনো Return record নেই। Return Entry বা Return Excel ব্যবহার করুন।</td></tr>'}</tbody></table></div><div class="table-foot">মোট ${bn(rj.total||0)}টি রিটার্ন রেকর্ড</div></div>`;
 }
 if(t=='users'){const j=await jf('/api/admin/users');const canMakeAdmin=meRole==='super_admin';content+=`<div class="admin-section"><h3> ইউজার ম্যানেজমেন্ট</h3><div class="admin-note">প্রতিটি ইউজারের জন্য প্রয়োজনীয় কাজের অনুমতি দিন।</div><div class="row"><input id="un" placeholder="নতুন ইউজারনেম"><input id="up" type="password" placeholder="পাসওয়ার্ড"><select id="ur"><option value="staff">Staff</option><option value="cs">CS</option>${canMakeAdmin?'<option value="admin">Admin</option>':''}</select><button class="pill" onclick="addUser()">+ ইউজার দিন</button></div><div id="ul">${(j.users||[]).map(userRow).join('')}</div></div>`}
 if(t=='backup'){content+=`<div class="admin-section"><div class="dash-hero"><div><h3>🛡 ব্যাকআপ সেন্টার</h3><small>আপনার ডেটার আলাদা আলাদা নিরাপদ কপি ডাউনলোড করুন</small></div></div><div class="admin-note">পূর্ণ ব্যাকআপে সব অ্যাপ ডেটা, ইউজার/অনুমতি, সেটিংস, হিস্ট্রি, বর্তমান workbook, আগের Excel snapshot এবং app source থাকবে। ZIP ফাইলটি অন্য ড্রাইভ বা ক্লাউডে রেখে দিন।</div><div class="backup-grid"><a class="backup-card primary" href="/api/admin/full-backup?scope=all"><span>✦</span><b>সম্পূর্ণ ব্যাকআপ</b><small>সবকিছু এক ZIP ফাইলে</small><em>ডাউনলোড ZIP →</em></a><a class="backup-card" href="/api/admin/full-backup?scope=database"><span>🗄️</span><b>অ্যাপ ডেটাবেস</b><small>ইউজার, অনুমতি, দাম, সেটিংস ও সব অ্যাপ ডেটা</small><em>ডাউনলোড DB ZIP →</em></a><a class="backup-card" href="/api/admin/full-backup?scope=stock"><span>📦</span><b>স্টক ফাইল</b><small>বর্তমান Excel workbook এবং আগের snapshot</small><em>ডাউনলোড স্টক ZIP →</em></a><a class="backup-card" href="/api/admin/full-backup?scope=history"><span>🧾</span><b>হিস্ট্রি ও কমেন্ট</b><small>স্টক পরিবর্তন, বিক্রির কমেন্ট, দাম ও activity history</small><em>ডাউনলোড হিস্ট্রি →</em></a><a class="backup-card" href="/api/admin/full-backup?scope=users-settings"><span>👥</span><b>ইউজার ও সেটিংস</b><small>ইউজার, access permissions, থিম ও সংযুক্ত উৎস</small><em>ডাউনলোড সেটিংস →</em></a></div><div class="admin-note">এই ব্যাকআপগুলো ডাউনলোড কপি। শুধু সার্ভারে রেখে দিলে কম্পিউটার বা সার্ভার নষ্ট হলে কপি উদ্ধার করা যাবে না—ডাউনলোড করে অন্য জায়গায় রাখুন।</div></div>`}
 if(t=='logs'){const [j,us]=await Promise.all([jf('/api/admin/logs?limit=1000'+(logUserFilter?'&username='+encodeURIComponent(logUserFilter):'')+(logActionFilter?'&action='+encodeURIComponent(logActionFilter):'')),jf('/api/admin/users')]);content+=`<div class="admin-section"><h3>📋 সব Activity</h3><p style="color:#777;font-size:13px">কে কখন কী ব্যবহার করেছে—search, product view, filter, copy, download, price edit, favorite, report, login এবং অন্যান্য activity এখানে থাকবে।</p><div class="row"><select id="logUser" onchange="logUserFilter=this.value;panel('logs')"><option value="">সব User</option>${(us.users||[]).map(u=>`<option value="${esc(u.username)}" ${logUserFilter===u.username?'selected':''}>${esc(u.username)} (${esc(u.role)})</option>`).join('')}</select><select id="logAction" onchange="logActionFilter=this.value;panel('logs')"><option value="">সব Activity</option>${Object.entries(ACTIVITY_LABELS).map(([k,v])=>`<option value="${k}" ${logActionFilter===k?'selected':''}>${v}</option>`).join('')}</select><button class="pill" onclick="logUserFilter='';logActionFilter='';panel('logs')">Reset</button></div><div class="admin-table-wrap"><table><tr><th>সময়</th><th>ইউজার</th><th>কাজ</th><th>Search / ID</th><th>বিস্তারিত</th></tr>${(j.logs||[]).map(logRow).join('')||'<tr><td colspan="5">কোনো activity নেই</td></tr>'}</table></div></div>`}
 if(t=='purchases'){
  const [sj,pj,meta,catalog]=await Promise.all([jf('/api/erp/suppliers'),jf('/api/erp/purchases?limit=200&q='+encodeURIComponent(window.erpPurchaseQuery||'')),jf('/api/erp/meta'),jf('/api/erp/catalog-options')]);
  window.erpSheets=meta.sheets||[];window.erpCatalogProducts=catalog.products||[];const supp=sj.suppliers||[],rows=pj.rows||[],sm=pj.summary||{};const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});
  const sheetOptions=`<option value="">No stock sheet</option>${(window.erpSheets||[]).map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('')}`;
  content+=`<div class="admin-section erp-module"><div class="dash-hero"><div><h3>🛒 Purchase Management</h3><small>Supplier invoices, item-level cost, payments, due balance and automatic Excel stock receiving.</small></div><div class="dash-actions"><button class="pill" onclick="panel('suppliers')">🏭 Manage Suppliers</button><a class="pill" href="#" onclick="erpDownloadExcel('/api/erp/purchases/export.xlsx','Alien-Tees-Purchase-Report.xlsx');return false;">⬇ Excel Report</a></div></div>
  <div class="excel-stats"><div><small>Total Purchases</small><b>${money(sm.total||0)}</b><em>All recorded invoices</em></div><div><small>Paid to Suppliers</small><b>${money(sm.paid||0)}</b><em>Recorded payments</em></div><div><small>Supplier Due</small><b>${money(sm.due||0)}</b><em>Outstanding balance</em></div><div><small>Purchase Invoices</small><b>${Number(sm.count||0).toLocaleString('en-US')}</b><em>Total records</em></div></div>
  <div class="visual-card" style="padding:18px;margin:16px 0"><h3 style="margin-top:0">＋ New Purchase Entry</h3>${supp.length?`<div class="co-grid">
  <label>Supplier *<select id="erpPurchaseSupplier" required><option value="">Select supplier</option>${supp.filter(x=>x.active).map(x=>`<option value="${x.id}">${esc(x.name)}${x.phone?' · '+esc(x.phone):''}</option>`).join('')}</select></label>
  <label>Purchase Date *<input id="erpPurchaseDate" type="date" value="${new Date().toISOString().slice(0,10)}"></label>
  <label>Supplier Invoice No.<input id="erpPurchaseInvoice" maxlength="100" placeholder="Supplier invoice / challan number"></label>
  <label>Initial Payment (BDT)<input id="erpPurchasePaid" type="number" min="0" step="0.01" value="0" oninput="erpCalcPurchaseTotals()"></label>
  <label>Payment Method<select id="erpPurchaseMethod"><option>Cash</option><option>Bank Transfer</option><option>bKash</option><option>Nagad</option><option>Card</option><option>Other</option></select></label>
  <label>Discount (BDT)<input id="erpPurchaseDiscount" type="number" min="0" step="0.01" value="0" oninput="erpCalcPurchaseTotals()"></label>
  <label class="wide">Purchase Notes<textarea id="erpPurchaseNotes" placeholder="Optional notes"></textarea></label></div>
  <div class="excel-table-wrap" style="margin-top:14px"><table class="excel-grid" id="erpPurchaseItemsTable"><thead><tr><th>Product Name *</th><th>Size / Variant *</th><th>Qty *</th><th>Unit Cost (BDT) *</th><th>Line Total</th><th></th></tr></thead><tbody id="erpPurchaseItems"></tbody></table><datalist id="erpPurchaseProductNames">${(window.erpCatalogProducts||[]).map(p=>`<option value="${esc(p.title||'')}"></option>`).join('')}</datalist></div>
  <div class="row" style="justify-content:space-between;margin-top:10px"><button type="button" class="pill" onclick="erpAddPurchaseItem()">＋ Add Item</button><span style="font-size:13px;color:#0D7552;font-weight:600">✓ Received quantities will be added to stock automatically</span></div>
  <div class="excel-stats" style="margin-top:12px"><div><small>Subtotal</small><b id="erpPurchaseSubtotal">৳ 0</b></div><div><small>Discount</small><b id="erpPurchaseDiscountLabel">৳ 0</b></div><div><small>Invoice Total</small><b id="erpPurchaseTotal">৳ 0</b></div><div><small>Due After Initial Payment</small><b id="erpPurchaseDue">৳ 0</b></div></div>
  <div class="row" style="justify-content:flex-end;margin-top:14px"><button type="button" class="pill" onclick="erpSavePurchase()">💾 Save Purchase</button></div>`:`<div class="admin-note">Add a supplier before recording the first purchase.</div><button class="pill" onclick="panel('suppliers')">＋ Add Supplier</button>`}</div>
  <div class="visual-card" style="padding:18px;margin:16px 0"><div class="row" style="justify-content:space-between"><h3 style="margin:0">Purchase History</h3><div class="row"><input id="erpPurchaseQuery" placeholder="Search purchase / supplier / invoice" onkeydown="if(event.key==='Enter')erpFilterPurchases()"><button class="pill" onclick="erpFilterPurchases()">Search</button><button class="pill" onclick="window.erpPurchaseQuery='';panel('purchases')">Reset</button></div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Date</th><th>Purchase No.</th><th>Supplier</th><th>Invoice Ref.</th><th>Items / Qty</th><th>Total</th><th>Paid</th><th>Due</th><th>Stock</th><th>Action</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.purchase_date)}</td><td><b>${esc(r.purchase_no)}</b></td><td>${esc(r.supplier_name)}</td><td>${esc(r.invoice_no||'—')}</td><td>${r.item_count} items / ${r.quantity} pcs</td><td>${money(r.total)}</td><td>${money(r.paid)}</td><td><b>${money(r.due)}</b></td><td>${r.stock_received?'<span class="stockstate">Received</span>':'<span class="stockstate low">Not received</span>'}</td><td><button class="pill" onclick="erpViewPurchase(${r.id})">View</button> <button class="pill" onclick="erpOpenPurchaseMemo(${r.id})">📄 মেমো দেখুন</button> <button class="pill" onclick="erpDownloadPurchaseMemo(${r.id})">⬇ ডাউনলোড</button>${!r.stock_received&&r.stock_receivable&&window.erpSheets.length?` <button class="pill" onclick="erpReceivePurchaseStock(${r.id})">Receive Stock</button>`:''}${r.due>0?` <button class="pill" onclick="erpRecordPurchasePayment(${r.id},${r.due})">Pay Due</button>`:''}</td></tr>`).join('')||'<tr><td colspan="10">No purchase records yet.</td></tr>'}</tbody></table></div></div></div>`;
  if(!window.erpSheets)window.erpSheets=[];setTimeout(()=>{if($('erpPurchaseItems')&&$('erpPurchaseItems').children.length===0)erpAddPurchaseItem();},0);
 }
 if(t=='suppliers'){
  const j=await jf('/api/erp/suppliers?include_inactive=1'),rows=j.suppliers||[],sm=j.summary||{};const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});
  content+=`<div class="admin-section erp-module"><div class="dash-hero"><div><h3>🏭 Supplier Management</h3><small>Supplier contacts, opening balance, purchase history and outstanding amounts.</small></div><div class="dash-actions"><button class="pill" onclick="panel('purchases')">🛒 Purchase Management</button><a class="pill" href="#" onclick="erpDownloadExcel('/api/erp/suppliers/export.xlsx','Alien-Tees-Suppliers.xlsx');return false;">⬇ Supplier Report</a></div></div><div class="excel-stats"><div><small>Active Suppliers</small><b>${sm.active_count||0}</b></div><div><small>Total Purchase Value</small><b>${money(sm.purchase_total||0)}</b></div><div><small>Total Paid</small><b>${money(sm.paid_total||0)}</b></div><div><small>Current Due</small><b>${money(sm.due_total||0)}</b></div></div>
  <div class="visual-card" style="padding:18px;margin:16px 0"><h3 style="margin-top:0">＋ Add Supplier</h3><div class="co-grid"><label>Supplier / Company Name *<input id="erpSupplierName" maxlength="160" placeholder="Supplier name"></label><label>Phone<input id="erpSupplierPhone" maxlength="60" placeholder="Phone number"></label><label>Email<input id="erpSupplierEmail" type="email" maxlength="160" placeholder="Email address"></label><label>Opening Due (BDT)<input id="erpSupplierOpening" type="number" min="0" step="0.01" value="0"></label><label class="wide">Address<input id="erpSupplierAddress" maxlength="500" placeholder="Supplier address"></label><label class="wide">Notes<textarea id="erpSupplierNotes" placeholder="Optional supplier notes"></textarea></label></div><div class="row" style="justify-content:flex-end;margin-top:12px"><button class="pill" onclick="erpSaveSupplier()">💾 Save Supplier</button></div></div>
  <div class="visual-card" style="padding:18px;margin:16px 0"><div class="row" style="justify-content:space-between"><h3 style="margin:0">Supplier List</h3><input id="erpSupplierQuery" placeholder="Search supplier / phone" oninput="erpFilterSupplierRows()"></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Supplier</th><th>Phone / Email</th><th>Opening Due</th><th>Purchases</th><th>Paid</th><th>Current Due</th><th>Status</th><th>Action</th></tr></thead><tbody id="erpSupplierRows">${rows.map(r=>`<tr data-search="${esc((r.name+' '+r.phone+' '+r.email).toLowerCase())}"><td><b>${esc(r.name)}</b><small style="display:block">${esc(r.address||'')}</small></td><td>${esc(r.phone||'—')}<small style="display:block">${esc(r.email||'')}</small></td><td>${money(r.opening_balance)}</td><td>${money(r.purchases)}</td><td>${money(r.paid)}</td><td><b>${money(r.due)}</b></td><td>${r.active?'<span class="stockstate">Active</span>':'<span class="stockstate out">Inactive</span>'}</td><td><button class="pill" onclick="erpShowSupplierHistory(${r.id})">📋 ক্রয়ের ইতিহাস</button> <button class="pill" onclick="erpDownloadExcel('/api/erp/suppliers/${r.id}/purchases/export.xlsx','Supplier-Purchase-History-${r.id}.xlsx')">⬇ Excel</button>${r.due>0?` <button class="pill" onclick="erpRecordSupplierPayment(${r.id},${r.due})">Record Payment</button>`:''} <button class="pill" onclick="erpToggleSupplier(${r.id},${r.active?0:1})">${r.active?'Deactivate':'Activate'}</button></td></tr>`).join('')||'<tr><td colspan="8">No suppliers added.</td></tr>'}</tbody></table></div></div></div>`;
 }
 if(t=='expenses'){
  const q=window.erpExpenseQuery||'',j=await jf('/api/erp/expenses?limit=300&q='+encodeURIComponent(q)),rows=j.rows||[],sm=j.summary||{};const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});
  content+=`<div class="admin-section erp-module"><div class="dash-hero"><div><h3>💸 Expense Management</h3><small>Track operating expenses, receipts/reference numbers, payment methods and approvals.</small></div><div class="dash-actions"><a class="pill" href="/api/erp/expenses/export.xlsx">⬇ Excel Report</a><button class="pill" onclick="panel('erpReports')">📊 Reports</button></div></div><div class="excel-stats"><div><small>Total Approved</small><b>${money(sm.approved||0)}</b></div><div><small>Pending Approval</small><b>${money(sm.pending||0)}</b></div><div><small>Expense Records</small><b>${sm.count||0}</b></div><div><small>This Month</small><b>${money(sm.month_total||0)}</b></div></div>
  <div class="visual-card" style="padding:18px;margin:16px 0"><h3 style="margin-top:0">＋ New Expense</h3><div class="co-grid"><label>Expense Title *<input id="erpExpenseTitle" maxlength="180" placeholder="e.g. Office Rent"></label><label>Category *<select id="erpExpenseCategory"><option>Office Rent</option><option>Salary & Wages</option><option>Utilities</option><option>Internet & Software</option><option>Marketing</option><option>Packaging</option><option>Courier & Delivery</option><option>Transport</option><option>Office Supplies</option><option>Purchase Miscellaneous</option><option>Other</option></select></label><label>Amount (BDT) *<input id="erpExpenseAmount" type="number" min="0.01" step="0.01" placeholder="0.00"></label><label>Expense Date *<input id="erpExpenseDate" type="date" value="${new Date().toISOString().slice(0,10)}"></label><label>Payment Method<select id="erpExpenseMethod"><option>Cash</option><option>Bank Transfer</option><option>bKash</option><option>Nagad</option><option>Card</option><option>Other</option></select></label><label>Paid To / Vendor<input id="erpExpensePaidTo" maxlength="160" placeholder="Person or vendor"></label><label>Reference / Receipt No.<input id="erpExpenseRef" maxlength="120" placeholder="Receipt number / transaction ID"></label><label>Status<select id="erpExpenseStatus"><option value="Pending">Pending Approval</option><option value="Approved">Approved / Paid</option></select></label><label class="wide">Notes<textarea id="erpExpenseNotes" placeholder="Purpose, extra information or receipt details"></textarea></label></div><div class="row" style="justify-content:flex-end;margin-top:12px"><button class="pill" onclick="erpSaveExpense()">💾 Save Expense</button></div></div>
  <div class="visual-card" style="padding:18px;margin:16px 0"><div class="row" style="justify-content:space-between"><h3 style="margin:0">Expense History</h3><div class="row"><input id="erpExpenseQuery" value="${esc(q)}" placeholder="Search title / category / vendor" onkeydown="if(event.key==='Enter')erpFilterExpenses()"><button class="pill" onclick="erpFilterExpenses()">Search</button><button class="pill" onclick="window.erpExpenseQuery='';panel('expenses')">Reset</button></div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Date</th><th>Expense No.</th><th>Title</th><th>Category</th><th>Paid To</th><th>Method</th><th>Amount</th><th>Status</th><th>Action</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.expense_date)}</td><td>${esc(r.expense_no)}</td><td><b>${esc(r.title)}</b><small style="display:block">${esc(r.notes||'')}</small></td><td>${esc(r.category)}</td><td>${esc(r.paid_to||'—')}</td><td>${esc(r.payment_method)}</td><td><b>${money(r.amount)}</b></td><td>${r.status==='Approved'?'<span class="stockstate">Approved</span>':r.status==='Rejected'?'<span class="stockstate out">Rejected</span>':'<span class="stockstate low">Pending</span>'}</td><td>${r.status==='Pending'?`<button class="pill" onclick="erpSetExpenseStatus(${r.id},'Approved')">Approve</button> <button class="pill" onclick="erpSetExpenseStatus(${r.id},'Rejected')">Reject</button>`:''}</td></tr>`).join('')||'<tr><td colspan="9">No expense records found.</td></tr>'}</tbody></table></div></div></div>`;
 }
 if(t=='erpReports'){
  const from=window.erpReportFrom||'',to=window.erpReportTo||'';
  const qs=new URLSearchParams({from,to,limit:'1000'});
  const [j,pj,ej]=await Promise.all([jf('/api/erp/summary?'+new URLSearchParams({from,to}).toString()),jf('/api/erp/purchases?'+qs.toString()),jf('/api/erp/expenses?'+qs.toString())]);
  if(!j._ok)return `<div class="admin-section erp-module"><h3>Purchase & Expense Reports</h3><p>${esc(j.error||'Could not load financial report.')}</p><button class="pill" onclick="panel('erpReports')">Retry</button></div>`;
  const s=j.summary||{},cats=j.expense_categories||[],sup=j.supplier_dues||[],purchases=pj.rows||[],expenses=ej.rows||[];
  const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});
  const canPay=isAdminLevelJS()||allowed('purchase_manage')||allowed('supplier_manage');
  const canDelete=isAdminLevelJS();
  content+=`<div class="admin-section erp-module"><div class="dash-hero"><div><h3>📊 Purchase & Expense Reports</h3><small>Financial snapshot, supplier outstanding, invoice due payments and record controls.</small></div><div class="dash-actions"><a class="pill" href="/api/erp/purchases/export.xlsx?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}">⬇ Purchase Excel</a><a class="pill" href="/api/erp/expenses/export.xlsx?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}">⬇ Expense Excel</a></div></div>
  <div class="visual-card" style="padding:14px;margin:14px 0"><div class="row"><label>From <input id="erpReportFrom" type="date" value="${esc(from)}"></label><label>To <input id="erpReportTo" type="date" value="${esc(to)}"></label><button class="pill" onclick="erpApplyReports()">Apply Filter</button><button class="pill" onclick="window.erpReportFrom='';window.erpReportTo='';panel('erpReports')">All Dates</button></div></div>
  <div class="excel-stats"><div><small>Purchase Total</small><b>${money(s.purchase_total||0)}</b></div><div><small>Supplier Payments</small><b>${money(s.supplier_paid||0)}</b></div><div><small>Supplier Due</small><b>${money(s.supplier_due||0)}</b></div><div><small>Approved Expenses</small><b>${money(s.expenses_approved||0)}</b></div><div><small>Pending Expenses</small><b>${money(s.expenses_pending||0)}</b></div><div><small>Net Cash Outflow</small><b>${money((s.supplier_paid||0)+(s.expenses_approved||0))}</b><em>Supplier payments + approved expenses</em></div></div>
  <div class="order-summary-grid"><div class="visual-card"><h3>Expenses by Category</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Category</th><th>Records</th><th>Approved</th><th>Pending</th></tr></thead><tbody>${cats.map(x=>`<tr><td>${esc(x.category)}</td><td>${x.count}</td><td>${money(x.approved)}</td><td>${money(x.pending)}</td></tr>`).join('')||'<tr><td colspan="4">No expense data for this period.</td></tr>'}</tbody></table></div></div>
  <div class="visual-card"><h3>Supplier Outstanding</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Supplier</th><th>Purchase Total</th><th>Paid</th><th>Due</th><th>Action</th></tr></thead><tbody>${sup.slice(0,100).map(x=>`<tr><td>${esc(x.name)}</td><td>${money(x.purchases)}</td><td>${money(x.paid)}</td><td><b>${money(x.due)}</b></td><td>${canPay&&x.due>0?`<button class="pill" onclick="erpRecordSupplierPayment(${Number(x.id)},${Number(x.due)},'erpReports')">Pay Due</button>`:'—'}</td></tr>`).join('')||'<tr><td colspan="5">No supplier balances.</td></tr>'}</tbody></table></div></div></div>
  <div class="visual-card" style="padding:16px;margin-top:14px"><div style="display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap"><h3 style="margin:0">Purchase Records · Pay Due / Delete</h3><span style="color:#657a6d;font-size:12px">${purchases.length} records in current filter</span></div><div class="excel-table-wrap" style="margin-top:12px"><table class="excel-grid"><thead><tr><th>Date</th><th>Purchase No.</th><th>Supplier</th><th>Invoice</th><th>Total</th><th>Paid</th><th>Due</th><th>Stock</th><th>Actions</th></tr></thead><tbody>${purchases.map(x=>`<tr><td>${esc(x.purchase_date)}</td><td><b>${esc(x.purchase_no)}</b></td><td>${esc(x.supplier_name)}</td><td>${esc(x.invoice_no||'—')}</td><td>${money(x.total)}</td><td>${money(x.paid)}</td><td><b>${money(x.due)}</b></td><td>${x.stock_received?'<span class="stockstate">Received</span>':'<span class="stockstate low">Not received</span>'}</td><td style="white-space:nowrap">${canPay&&Number(x.due)>0?`<button class="pill" onclick="erpRecordPurchasePayment(${Number(x.id)},${Number(x.due)},'erpReports')">Pay Due</button> `:''}${canDelete?`<button class="pill" style="border-color:#d95b5b;color:#ad2626" onclick="erpDeletePurchase(${Number(x.id)},'${esc(x.purchase_no).replace(/'/g,'&#39;')}',${!!x.stock_received})">Delete</button>`:'—'}</td></tr>`).join('')||'<tr><td colspan="9">No purchase records for this period.</td></tr>'}</tbody></table></div><small style="display:block;margin-top:8px;color:#6a7c71">Deleting a received purchase first attempts to subtract its received quantities from Excel stock. If quantities cannot be safely reversed, the app will stop the deletion.</small></div>
  <div class="visual-card" style="padding:16px;margin-top:14px"><div style="display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap"><h3 style="margin:0">Expense Records · Delete</h3><span style="color:#657a6d;font-size:12px">${expenses.length} records in current filter</span></div><div class="excel-table-wrap" style="margin-top:12px"><table class="excel-grid"><thead><tr><th>Date</th><th>Expense No.</th><th>Title</th><th>Category</th><th>Paid To</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead><tbody>${expenses.map(x=>`<tr><td>${esc(x.expense_date)}</td><td><b>${esc(x.expense_no)}</b></td><td>${esc(x.title)}</td><td>${esc(x.category)}</td><td>${esc(x.paid_to||'—')}</td><td>${money(x.amount)}</td><td>${esc(x.status)}</td><td>${canDelete?`<button class="pill" style="border-color:#d95b5b;color:#ad2626" onclick="erpDeleteExpense(${Number(x.id)},'${esc(x.expense_no).replace(/'/g,'&#39;')}')">Delete</button>`:'—'}</td></tr>`).join('')||'<tr><td colspan="8">No expense records for this period.</td></tr>'}</tbody></table></div></div>
  </div>`;
 }
 if(t=='hr'){
  const hr=await jf('/api/hr/overview').catch(()=>({ok:false,error:'Could not load HR data'}));
  if(!hr.ok){content+=`<div class="admin-section"><h3>👥 HR Management</h3><div class="admin-note">${esc(hr.error||'HR data could not be loaded. Check your permissions.')}</div><button class="pill" onclick="panel('hr')">↻ Retry</button></div>`}
  else {
   const employees=hr.employees||[],attendance=hr.attendance||[],leaves=hr.leaves||[],payroll=hr.payroll||[],money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2}),eopts=employees.map(e=>`<option value="${Number(e.id)}" data-salary="${Number(e.salary||0)}" data-overtime-rate="${Number(e.overtime_rate||0)}">${esc(e.employee_code)} — ${esc(e.full_name)}</option>`).join(''),attendanceEopts='<option value="">All employees</option>'+eopts;
   content+=`<style>
   .hr-shell{display:grid;gap:18px;color:#173d30}.hr-hero{border-radius:20px;padding:24px;color:#fff;background:radial-gradient(circle at 95% 0%,#3f9972 0,#1c654c 34%,#123d30 75%);display:flex;align-items:center;justify-content:space-between;gap:18px;flex-wrap:wrap;box-shadow:0 14px 36px #123d3020}.hr-hero h2{margin:0;font-size:26px}.hr-hero p{margin:7px 0 0;color:#d4eadc;font-size:13px;max-width:670px}.hr-hero .pill{background:#f5d990;color:#1b3f31;border-color:#f5d990}.hr-stats{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}.hr-stat{padding:15px 16px;background:#fff;border:1px solid #e0ebe3;border-radius:14px;box-shadow:0 5px 16px #193f3008}.hr-stat small{display:block;color:#718078;font-size:11px}.hr-stat b{display:block;color:#143f30;font-size:23px;margin-top:5px}.hr-section{background:#fff;border:1px solid #e0e8e2;border-radius:16px;padding:18px;box-shadow:0 6px 20px #193f3007;min-width:0}.hr-section h3{margin:0 0 5px;font-size:17px;color:#143f30}.hr-muted{color:#78877d;font-size:12px;margin:0 0 14px}.hr-form-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:11px;align-items:end}.hr-form-grid label{display:grid;gap:5px;color:#64756b;font-size:11px;font-weight:700;min-width:0}.hr-form-grid input,.hr-form-grid select,.hr-form-grid textarea{width:100%;box-sizing:border-box;border:1px solid #d8e4db;border-radius:9px;background:#fbfdfb;padding:10px 11px;color:#172e23;font:13px system-ui,sans-serif;min-width:0}.hr-form-grid textarea{min-height:38px;resize:vertical}.hr-form-grid .wide{grid-column:span 2}.hr-tools{display:flex;gap:8px;align-items:center;justify-content:space-between;flex-wrap:wrap;margin:14px 0 10px}.hr-tools .pill{white-space:nowrap}.hr-table-wrap{max-width:100%;overflow:auto;border:1px solid #e5ece6;border-radius:11px}.hr-table{width:100%;border-collapse:collapse;min-width:750px;font-size:12px}.hr-table th,.hr-table td{padding:10px 11px;border-bottom:1px solid #edf1ed;text-align:left;vertical-align:middle}.hr-table th{background:#f2f7f3;color:#52685b;font-size:10px;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}.hr-table tr:last-child td{border-bottom:0}.hr-employee{display:flex;align-items:center;gap:9px;min-width:175px}.hr-avatar{width:36px;height:36px;border-radius:10px;object-fit:cover;background:#eaf1eb;border:1px solid #dae6dc;display:grid;place-items:center;font-weight:800;color:#47725a;flex-shrink:0}.hr-badge{display:inline-flex;padding:4px 8px;border-radius:99px;background:#eaf5ed;color:#246340;font-weight:750;font-size:10px}.hr-badge.pending{background:#fff3d7;color:#805d0f}.hr-badge.rejected,.hr-badge.inactive{background:#fdebe9;color:#a0342a}.hr-actions{display:flex;gap:5px;flex-wrap:wrap}.hr-actions .pill{padding:6px 8px;font-size:10px}.hr-camera-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:10px 0}.hr-camera-row video{width:min(300px,100%);aspect-ratio:4/3;object-fit:cover;border-radius:12px;background:#14241c}.hr-photo-preview{width:72px;height:72px;object-fit:cover;border-radius:12px;border:1px solid #dbe6dc;background:#eef4ef}.hr-empty{padding:18px;text-align:center;color:#849087}.hr-ot{color:#0d7552;font-weight:850;white-space:nowrap}.hr-zero{color:#7c897f;font-weight:650;white-space:nowrap}.hr-subtabs{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:14px}.hr-subtabs button{background:#f3f7f3;color:#325641;border:1px solid #dce7de;border-radius:10px;padding:9px 12px;font-weight:700;cursor:pointer}.hr-subtabs button.active{background:#0d7552;color:#fff;border-color:#0d7552}.hr-status-note{background:#f4f8f4;border:1px dashed #d8e5d9;border-radius:10px;padding:10px 12px;color:#65786b;font-size:11px;margin:12px 0 0}
   .hr-table-wrap{max-width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.hr-table{min-width:760px}.hr-actions .pill,.hr-shell .pill{min-height:38px;touch-action:manipulation}.hr-form-grid input,.hr-form-grid select,.hr-form-grid textarea{min-height:42px;font-size:14px}
   @media(max-width:1050px){.hr-stats{grid-template-columns:repeat(3,minmax(0,1fr))}.hr-form-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:620px){.hr-shell{gap:11px}.hr-hero{padding:17px 14px;border-radius:14px}.hr-stats{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.hr-stat{padding:11px}.hr-stat b{font-size:20px}.hr-section{padding:13px;border-radius:13px}.hr-form-grid{grid-template-columns:1fr}.hr-form-grid .wide{grid-column:auto}.hr-hero h2{font-size:22px}.hr-tools{align-items:stretch;flex-direction:column}.hr-tools input{width:100%;min-width:0!important}.hr-subtabs{display:flex;overflow-x:auto;gap:6px;padding-bottom:4px;-webkit-overflow-scrolling:touch}.hr-subtabs button{flex:0 0 auto;min-height:40px}.hr-camera-row{align-items:stretch}.hr-camera-row video{max-width:100%;height:auto}}
   </style><div class="admin-section hr-shell"><div class="hr-hero"><div><div style="font-size:10px;letter-spacing:.18em;font-weight:800;color:#b9e5cb;margin-bottom:8px">ALIEN TEES · PEOPLE OPERATIONS</div><h2>HR Management</h2><p>Employee records, daily attendance, leave approvals, payroll and printable payslips in one secure workspace.</p></div><div style="display:flex;gap:8px;flex-wrap:wrap"><a class="pill" href="/api/hr/export.xlsx">⬇ Export HR Reports</a><button class="pill" onclick="panel('hr')">↻ Refresh Data</button></div></div>
   <div class="hr-stats"><div class="hr-stat"><small>Total Employees</small><b>${Number(hr.summary?.employees||0).toLocaleString('en-US')}</b></div><div class="hr-stat"><small>Present Today</small><b>${Number(hr.summary?.present_today||0).toLocaleString('en-US')}</b></div><div class="hr-stat"><small>Pending Leave Requests</small><b>${Number(hr.summary?.pending_leaves||0).toLocaleString('en-US')}</b></div><div class="hr-stat"><small>Current Month Payroll</small><b style="font-size:18px">${money(hr.summary?.payroll_month_total||0)}</b></div><div class="hr-stat"><small>Active Employees</small><b>${Number(hr.summary?.active_employees||0).toLocaleString('en-US')}</b></div></div>
   <div class="hr-subtabs"><button class="active" onclick="hrShowSection('employees',this)">👥 Employees</button><button onclick="hrShowSection('attendance',this)">🕒 Attendance</button><button onclick="hrShowSection('leaves',this)">🌿 Leave Management</button><button onclick="hrShowSection('payroll',this)">💳 Payroll & Payslips</button><button onclick="hrShowSection('reports',this)">📊 Reports</button></div>
   <div class="hr-section hr-pane" id="hr-pane-employees"><h3>Employee Directory</h3><p class="hr-muted">Create employee profiles and maintain contact, role, department and compensation details.</p><form id="hrEmployeeForm" class="hr-form-grid" onsubmit="event.preventDefault();hrSaveEmployee()"><input type="hidden" id="hrEmployeeId"><label>Employee Code<input id="hrEmployeeCode" placeholder="Auto-generated if blank" maxlength="40"></label><label>Full Name *<input id="hrEmployeeName" required maxlength="120" placeholder="Employee full name"></label><label>Phone<input id="hrEmployeePhone" maxlength="40"></label><label>Email<input id="hrEmployeeEmail" type="email" maxlength="180"></label><label>Department<input id="hrEmployeeDepartment" maxlength="100" placeholder="e.g. Sales"></label><label>Designation<input id="hrEmployeeDesignation" maxlength="100" placeholder="e.g. Sales Executive"></label><label>Employment Type<select id="hrEmployeeType"><option>Full-time</option><option>Part-time</option><option>Contract</option><option>Intern</option></select></label><label>Joining Date<input id="hrEmployeeJoinDate" type="date"></label><label>Monthly Salary (BDT)<input id="hrEmployeeSalary" type="number" min="0" step="0.01" value="0"></label><label>Overtime Pay / Hour (BDT)<input id="hrEmployeeOvertimeRate" type="number" min="0" step="0.01" value="0" placeholder="e.g. 100"></label><label>Daily Shift / OT Limit (hours)<input id="hrEmployeeShiftHours" type="number" min="0" max="24" step="0.25" value="0"><small class="hr-muted">0 = use global default (${Number(hr.summary?.shift_hours||8)} hours)</small></label><label>Emergency Contact<input id="hrEmployeeEmergency" maxlength="100"></label><label>Status<select id="hrEmployeeStatus"><option>Active</option><option>Inactive</option></select></label><label>Profile Photo<input id="hrEmployeePhoto" type="file" accept="image/*" onchange="hrReadEmployeePhoto(this)"></label><label class="wide">Address<textarea id="hrEmployeeAddress" maxlength="500"></textarea></label><label class="wide">Notes<textarea id="hrEmployeeNotes" maxlength="800"></textarea></label><div style="display:flex;gap:8px;flex-wrap:wrap"><button class="pill" type="submit" style="background:#0d7552;color:#fff;border-color:#0d7552">💾 Save Employee</button><button class="pill" type="button" onclick="hrResetEmployeeForm()">Clear Form</button></div><img id="hrEmployeePhotoPreview" class="hr-photo-preview" alt="Employee photo preview" hidden></form><div class="hr-tools"><b>Employee List · ${employees.length} displayed</b><input id="hrEmployeeQuery" placeholder="Search employee, code or department" oninput="hrFilterTable('hrEmployeeQuery','hrEmployeeTable')" style="min-width:220px;max-width:100%;padding:10px;border:1px solid #d8e4db;border-radius:9px"></div><div class="hr-table-wrap"><table id="hrEmployeeTable" class="hr-table"><thead><tr><th>Employee</th><th>Department / Role</th><th>Contact</th><th>Join Date</th><th>Monthly Salary</th><th>OT Rate / Hour</th><th>Daily Shift / OT</th><th>Status</th><th>Actions</th></tr></thead><tbody>${employees.map(e=>`<tr data-search="${esc([e.employee_code,e.full_name,e.department,e.designation,e.phone,e.email].join(' ').toLowerCase())}"><td><div class="hr-employee">${e.photo_file?`<img class="hr-avatar" src="/api/hr/photo/${encodeURIComponent(e.photo_file)}" alt="">`:`<span class="hr-avatar">${esc((e.full_name||'?').slice(0,1).toUpperCase())}</span>`}<span><b>${esc(e.full_name)}</b><small style="display:block;color:#7c897f;margin-top:3px">${esc(e.employee_code)}</small></span></div></td><td>${esc(e.department||'—')}<small style="display:block;color:#78877d;margin-top:3px">${esc(e.designation||'—')}</small></td><td>${esc(e.phone||'—')}<small style="display:block;color:#78877d;margin-top:3px">${esc(e.email||'')}</small></td><td>${esc(e.join_date||'—')}</td><td>${money(e.salary)}</td><td>${money(e.overtime_rate||0)} / hr</td><td>${Number(e.shift_minutes||0)>0?hrDuration(Number(e.shift_minutes)):('Global · '+hrDuration(Number(hr.summary?.shift_hours||8)*60))}</td><td><span class="hr-badge ${e.status==='Active'?'':'inactive'}">${esc(e.status)}</span></td><td><div class="hr-actions"><button class="pill" onclick="hrEditEmployee(${Number(e.id)})">Edit</button><button class="pill" onclick="hrSetEmployeeStatus(${Number(e.id)},'${e.status==='Active'?'Inactive':'Active'}')">${e.status==='Active'?'Deactivate':'Activate'}</button></div></td></tr>`).join('')||'<tr><td colspan="9" class="hr-empty">No employees yet. Add the first employee above.</td></tr>'}</tbody></table></div></div>
   <div class="hr-section hr-pane" id="hr-pane-attendance" hidden><h3>Daily Attendance & Overtime</h3><p class="hr-muted">Record exact worked minutes. Overtime is the time beyond the standard daily shift; the default is 8 hours.</p><div class="hr-form-grid" style="margin:10px 0 12px;padding:12px;border:1px solid #dce9de;border-radius:12px;background:#f6faf6"><label>Standard Shift Hours<input id="hrShiftHours" type="number" min="1" max="24" step="0.25" value="${Number(hr.summary?.shift_hours||8)}"></label><div><button class="pill" onclick="hrSaveShiftHours()">💾 Save OT Settings</button><small class="hr-muted" style="display:block;margin-top:5px">Overtime starts after this daily duration for future attendance saves.</small></div></div><div class="hr-form-grid" style="margin:0 0 16px;padding:12px;border:1px solid #e8dfc5;border-radius:12px;background:#fffaf0"><label>Monthly Allowed Late Days<input id="hrLateAllowedDays" type="number" min="0" max="31" step="1" value="${Number(hr.summary?.late_allowed_days??3)}"></label><label>Deduction per Extra Late Day (BDT)<input id="hrLateDeductionPerDay" type="number" min="0" step="0.01" value="${Number(hr.summary?.late_deduction_per_day??100)}"></label><div><button class="pill" onclick="hrSaveLateRules()">💾 Save Late Rules</button><small class="hr-muted" style="display:block;margin-top:5px">Every attendance marked “Late” counts once per date. Deduction applies only after the monthly grace days.</small></div></div><div class="hr-form-grid"><label>Employee<select id="hrAttEmployee">${eopts}</select></label><label>Attendance Date<input id="hrAttDate" type="date" value="${new Date().toISOString().slice(0,10)}"></label><label>Status<select id="hrAttStatus"><option>Present</option><option>Late</option><option>Half Day</option><option>Absent</option><option>Leave</option><option>Holiday</option></select></label><label>Check In<input id="hrAttIn" type="time"></label><label>Check Out<input id="hrAttOut" type="time"></label><label class="wide">Notes<input id="hrAttNotes" maxlength="500" placeholder="Late reason, shift note, etc."></label><div style="display:flex;gap:8px;flex-wrap:wrap;align-items:end"><button class="pill" onclick="hrMarkAttendance('clock_in')">🟢 Clock In</button><button class="pill" onclick="hrMarkAttendance('clock_out')">🔴 Clock Out</button><button class="pill" onclick="hrMarkAttendance('manual')">Save Attendance</button></div></div><div class="hr-camera-row"><button class="pill" onclick="hrStartCamera()">📷 Open Webcam</button><button class="pill" onclick="hrCaptureCamera()">Capture Snapshot</button><button class="pill" onclick="hrStopCamera()">Close Camera</button><video id="hrCameraVideo" autoplay playsinline muted hidden></video><canvas id="hrCameraCanvas" hidden></canvas><img id="hrAttendancePhotoPreview" class="hr-photo-preview" alt="Attendance snapshot preview" hidden><small id="hrCameraStatus" class="hr-muted" style="margin:0">Camera is optional. Allow browser camera access to capture a photo.</small></div><h3 style="margin-top:20px">Monthly Attendance Report</h3><p class="hr-muted">Select one employee and a month to view every attendance record and download the same filtered report in Excel.</p><div class="hr-form-grid" style="margin:10px 0 12px"><label>Employee<select id="hrAttendanceEmployee" onchange="hrLoadAttendanceReport()">${attendanceEopts}</select></label><label>Month<input id="hrAttendanceMonth" type="month" value="${new Date().toISOString().slice(0,7)}" onchange="hrLoadAttendanceReport()"></label><div style="display:flex;gap:8px;flex-wrap:wrap;align-items:end"><button class="pill" type="button" onclick="hrLoadAttendanceReport()">🔎 Show Month</button><a class="pill" id="hrAttendanceExport" href="/api/hr/attendance-export.xlsx?month=${new Date().toISOString().slice(0,7)}">⬇ Download Excel</a></div></div><div class="hr-tools"><b id="hrAttendanceSummary">Loading monthly attendance…</b></div><div class="hr-table-wrap"><table id="hrAttendanceTable" class="hr-table"><thead><tr><th>Date</th><th>Employee</th><th>Check In</th><th>Check Out</th><th>Work Duration</th><th>Overtime (H:MM)</th><th>Status</th><th>Photo</th></tr></thead><tbody id="hrAttendanceRows">${attendance.map(a=>`<tr><td>${esc(a.attendance_date)}</td><td>${esc(a.employee_code)} · ${esc(a.full_name)}</td><td>${esc(a.check_in||'—')}</td><td>${esc(a.check_out||'—')}</td><td>${hrDuration(Number(a.worked_minutes||Math.round(Number(a.work_hours||0)*60)))}</td><td><b class="${Number(a.overtime_minutes||0)>0?'hr-ot':'hr-zero'}">${hrDuration(Number(a.overtime_minutes||0))}</b></td><td><span class="hr-badge ${a.status==='Absent'?'rejected':''}">${esc(a.status)}</span></td><td>${a.photo_file?`<a href="/api/hr/photo/${encodeURIComponent(a.photo_file)}" target="_blank" rel="noopener">View Photo</a>`:'—'}</td></tr>`).join('')||'<tr><td colspan="8" class="hr-empty">No attendance records yet.</td></tr>'}</tbody></table></div></div>
   <div class="hr-section hr-pane" id="hr-pane-leaves" hidden><h3>Leave Management</h3><p class="hr-muted">Submit leave requests and review pending requests. Approved and rejected decisions are stored with reviewer details.</p><form class="hr-form-grid" onsubmit="event.preventDefault();hrSaveLeave()"><label>Employee<select id="hrLeaveEmployee" required>${eopts}</select></label><label>Leave Type<select id="hrLeaveType"><option>Casual</option><option>Sick</option><option>Annual</option><option>Emergency</option><option>Maternity / Paternity</option><option>Unpaid</option></select></label><label>Start Date<input id="hrLeaveStart" type="date" required></label><label>End Date<input id="hrLeaveEnd" type="date" required></label><label class="wide">Reason<textarea id="hrLeaveReason" maxlength="1000" required placeholder="Reason for leave"></textarea></label><div><button class="pill" type="submit" style="background:#0d7552;color:white;border-color:#0d7552">＋ Submit Leave</button></div></form><div class="hr-tools"><b>Leave Requests</b><span class="hr-muted" style="margin:0">Pending: ${leaves.filter(x=>x.status==='Pending').length}</span></div><div class="hr-table-wrap"><table class="hr-table"><thead><tr><th>Employee</th><th>Type</th><th>Dates</th><th>Days</th><th>Reason</th><th>Status</th><th>Review</th></tr></thead><tbody>${leaves.map(l=>`<tr><td>${esc(l.employee_code)} · ${esc(l.full_name)}</td><td>${esc(l.leave_type)}</td><td>${esc(l.start_date)} – ${esc(l.end_date)}</td><td>${Number(l.days||0)}</td><td>${esc(l.reason||'—')}</td><td><span class="hr-badge ${l.status==='Pending'?'pending':l.status==='Rejected'?'rejected':''}">${esc(l.status)}</span><small style="display:block;color:#78877d;margin-top:3px">${l.reviewed_by?'By '+esc(l.reviewed_by):''}</small></td><td>${l.status==='Pending'?`<div class="hr-actions"><button class="pill" onclick="hrReviewLeave(${Number(l.id)},'Approved')">Approve</button><button class="pill" onclick="hrReviewLeave(${Number(l.id)},'Rejected')">Reject</button></div>`:'—'}</td></tr>`).join('')||'<tr><td colspan="7" class="hr-empty">No leave requests yet.</td></tr>'}</tbody></table></div></div>
   <div class="hr-section hr-pane" id="hr-pane-payroll" hidden><h3>Payroll & Payslips</h3><p class="hr-muted">Prepare monthly salary, allowances, bonus and deductions. Net pay is calculated server-side and can be marked as paid with payment reference.</p><form class="hr-form-grid" onsubmit="event.preventDefault();hrSavePayroll()"><label>Employee<select id="hrPayEmployee" required onchange="hrPayrollPreview()">${eopts}</select></label><label>Payroll Month<input id="hrPayMonth" type="month" value="${new Date().toISOString().slice(0,7)}" onchange="hrPayrollPreview()" required></label><label>Basic Salary (BDT)<input id="hrPayBasic" type="number" min="0" step="0.01" value="0" required></label><label>Allowances (BDT)<input id="hrPayAllowances" type="number" min="0" step="0.01" value="0"></label><label>Bonus (BDT)<input id="hrPayBonus" type="number" min="0" step="0.01" value="0"></label><label>Deductions (BDT)<input id="hrPayDeductions" type="number" min="0" step="0.01" value="0"></label><label>Status<select id="hrPayStatus"><option>Draft</option><option>Paid</option></select></label><label>Paid Date<input id="hrPayPaidDate" type="date" value="${new Date().toISOString().slice(0,10)}"></label><label>Payment Method<select id="hrPayMethod"><option>Bank Transfer</option><option>Cash</option><option>bKash</option><option>Nagad</option><option>Other</option></select></label><label>Reference / Transaction ID<input id="hrPayReference" maxlength="120"></label><label class="wide">Notes<textarea id="hrPayNotes" maxlength="500"></textarea></label><div class="wide" id="hrPayOtSummary" style="padding:12px 14px;border:1px solid #d8e8dc;border-radius:10px;background:#f4faf5;color:#1b503a">Loading monthly overtime calculation…</div><div><button class="pill" type="submit" style="background:#0d7552;color:white;border-color:#0d7552">💾 Save Payroll</button></div></form><div class="hr-table-wrap" style="margin-top:14px"><table class="hr-table"><thead><tr><th>Employee</th><th>Month</th><th>Basic</th><th>Allowance</th><th>Bonus</th><th>OT Pay</th><th>Late Deduction</th><th>Other Deductions</th><th>Net Pay</th><th>Status</th><th>Payslip</th></tr></thead><tbody>${payroll.map(p=>`<tr><td>${esc(p.employee_code)} · ${esc(p.full_name)}</td><td>${esc(p.pay_month)}</td><td>${money(p.basic_salary)}</td><td>${money(p.allowances)}</td><td>${money(p.bonus)}</td><td>${money(p.overtime_pay||0)}<small style="display:block;color:#78877d">${Number(p.overtime_hours||0).toFixed(2)} hrs</small></td><td>${money(Math.max(0,Number(p.late_deduction||0))) }<small style="display:block;color:#78877d">${Math.max(0,Number(p.late_days||0))} late day(s)</small></td><td>${money(p.deductions)}</td><td><b>${money(p.net_pay)}</b></td><td><span class="hr-badge ${p.status==='Paid'?'':'pending'}">${esc(p.status)}</span></td><td><a class="pill" target="_blank" rel="noopener" href="/api/hr/payslip/${Number(p.id)}">Print / View</a></td></tr>`).join('')||'<tr><td colspan="11" class="hr-empty">No payroll records yet.</td></tr>'}</tbody></table></div></div>
   <div class="hr-section hr-pane" id="hr-pane-reports" hidden><h3>HR Reports</h3><p class="hr-muted">Download a multi-sheet Excel workbook with employee directory, attendance, leave decisions and payroll.</p><div class="hr-form-grid"><div><a class="pill" href="/api/hr/export.xlsx">⬇ Download HR Excel Report</a></div><div><button class="pill" onclick="window.print()">🖨 Print Current View</button></div></div><div class="hr-status-note">Security: profile and attendance photos are served only to authenticated accounts with HR view permission. Payroll endpoints require payroll permission. Webcam photos are evidence snapshots and do not confirm a person's identity biometrically.</div></div>
   </div>`;
  }
 }
 if(t=='password'){content+=`<div class="admin-section"><h3> Admin / Super Admin পাসওয়ার্ড পরিবর্তন</h3><div class="admin-note">শুধু নিজের Admin account-এর password এখানে বদলাবেন। অন্য User-এর password User Management থেকে বদলানো যাবে।</div><div class="row" style="flex-direction:column;align-items:stretch;max-width:520px"><input id="oldpw" type="password" placeholder="বর্তমান পাসওয়ার্ড"><input id="newpw" type="password" placeholder="নতুন পাসওয়ার্ড (কমপক্ষে ৬ অক্ষর)"><input id="newpw2" type="password" placeholder="নতুন পাসওয়ার্ড আবার লিখুন"><button class="pill" onclick="changeAdminPassword()">পাসওয়ার্ড পরিবর্তন করুন</button></div></div>`}
 $('ab').innerHTML=`<div class="admin-head"><div><h2>ALIEN · ADMIN</h2><p id="userWelcome">${esc(roleLabel(meRole))} access · ${esc(me||'')}</p></div><div class="admin-head-actions"><button class="pill admin-menu-return" onclick="adminBackToSlide()">☰ মেনু</button><button class="x" onclick="closeAdminPanel()" aria-label="Close admin panel">✕</button></div></div><div class="admin-body admin-body-no-tabs"><section class="admin-content" id="adminContent">${content}</section></div>`;translateUi($('ab'));
 if(t=='price')priceRows();if(t=='excelstock')setTimeout(()=>{markStockCommentRows();renderCommentSummary()},0);if(t=='hr')setTimeout(()=>{hrLoadAttendanceReport();hrPayrollPreview()},0);
}
async function markStockCommentRows(){const table=document.querySelector('#adminContent .excel-grid');if(!table)return;const sheet=excelSheet||'__all__',url='/api/excel-stock/data?sheet='+encodeURIComponent(sheet)+'&offset='+excelOffset+'&limit=250';const d=await jf(url);const rows=d.selected?.rows||[];const map=new Map(rows.filter(r=>r.comment_count).map(r=>[String(r.sheet||sheet).toLowerCase()+'|'+String((r.values||[])[0]||'').trim().toLowerCase(),r]));table.querySelectorAll('tbody tr').forEach(tr=>{const codeCell=tr.querySelector('.code-cell')?.closest('td');if(!codeCell)return;const code=tr.querySelector('.code-cell')?.textContent.trim();const rowSheet=sheet==='__all__'?tr.cells[1]?.textContent.trim():sheet;const item=map.get(String(rowSheet||'').toLowerCase()+'|'+String(code||'').toLowerCase());if(!item)return;tr.classList.add('has-stock-comments');tr.querySelector('.code-cell')?.classList.add('commented-code');if(!codeCell.querySelector('.stock-comment-jump')){const b=document.createElement('button');b.className='stock-comment-jump';b.textContent='কমেন্ট '+bn(item.comment_count)+' · খুলুন';b.onclick=()=>openStockProductFromRow(code);codeCell.append(' ',b)}})}
let excelSheet='__all__',excelQuery='',excelOffset=0,excelPending={};
async function loadEntryRows(){const sh=$("entrySheet")?.value;if(!sh)return;const d=await jf("/api/excel-stock/data?sheet="+encodeURIComponent(sh)+"&offset=0&limit=250"),list=$("entryCodes");if(list){const codes=new Map();for(const r of (d.selected?.rows||[])){const code=String(r.values?.[0]||'').trim();if(code)codes.set(code,'')}for(const p of (ALL||[])){for(const code of [p.id,p.display_id,productId(p)]){const key=String(code||'').trim();if(key)codes.set(key,String(p.title||''))}}list.innerHTML=[...codes].map(([code,title])=>`<option value="${esc(code)}" label="${esc(title)}"></option>`).join('')}const code=$("entryCode")?.value?.trim();if(code)await loadEntrySizes();else{window.entryRow=null;window.entryMatch=null;previewEntryCatalogProduct()}}
function previewEntryCatalogProduct(){const code=$("entryCode")?.value?.trim()||"",box=$("entryCatalogPreview");if(!box)return;const p=findStockCatalogProduct(code),m=window.entryMatch; if(!code){box.innerHTML='<small>Enter a product code to see the product and current size quantities.</small>';return}const sizes=m?.sizes||[];const sizesHtml=sizes.length?`<div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px">${sizes.map(x=>`<span style="display:inline-flex;gap:5px;align-items:center;padding:5px 8px;border:1px solid #dbe7dd;border-radius:8px;background:#f5faf6;font-size:12px"><b>${esc(x.size)}</b><span>Stock: ${Number(x.qty||0)}</span></span>`).join('')}</div>`:'<small style="display:block;margin-top:4px;color:#8a6a3a">Choose / find a product to load size variants from its worksheet.</small>';const image=(m?.product_img||p?.img)?`<img src="${im(m?.product_img||p?.img)}" alt="" style="width:76px;height:92px;object-fit:cover;border-radius:10px;background:#f2f4f2" onerror="this.style.display='none'">`:'';box.innerHTML=`<div style="display:flex;align-items:flex-start;gap:12px;padding:10px;border:1px solid #e1e8e2;border-radius:12px;background:#fff">${image}<div style="min-width:0;flex:1"><b style="display:block;color:#173d31">${esc(m?.code||code)}</b><span>${esc(m?.product_name||p?.title||'Catalog product')}</span><small style="display:block;color:#758179;margin-top:4px">${m?.sheet?`Matched stock sheet: ${esc(m.sheet)}`:(p?'Catalog match · workbook row not located':'No matching catalog product')}</small>${sizesHtml}</div></div>`}
async function loadEntrySizes(){const sh=$("entrySheet")?.value,code=$("entryCode")?.value?.trim(),sel=$("entrySize");if(!code)return toast("Type a product code first",true);if(sel){sel.disabled=true;sel.innerHTML='<option value="">Looking up sizes…</option>'}const seq=(window.entryLookupSeq||0)+1;window.entryLookupSeq=seq;const d=await jf("/api/erp/stock-item-lookup?q="+encodeURIComponent(code));if(seq!==window.entryLookupSeq)return;const matches=d.matches||[],match=matches.find(x=>x.sheet===sh)||matches[0];if(!d._ok||!match){window.entryRow=null;window.entryMatch=null;if(sel){sel.disabled=true;sel.innerHTML='<option value="">No size variants found</option>'}previewEntryCatalogProduct();return toast(d.error||"এই Product Code কোনো Stock Sheet-এ পাওয়া যায়নি",true)}window.entryMatch=match;window.entryRow=match.row;if(match.sheet&&$('entrySheet')&&$('entrySheet').value!==match.sheet)$('entrySheet').value=match.sheet;if(match.code&&$('entryCode'))$('entryCode').value=match.code;if(sel){sel.innerHTML='<option value="">Select size / variant</option>'+(match.sizes||[]).map(x=>`<option value="${esc(x.size)}">${esc(x.size)} · current stock ${Number(x.qty||0)}</option>`).join('');sel.disabled=!(match.sizes||[]).length}previewEntryCatalogProduct();toast("Product পাওয়া গেছে — current size এবং stock count দেখানো হয়েছে")}
async function submitStockEntry(){const sheet=$("entrySheet")?.value,code=$("entryCode")?.value?.trim(),size=$("entrySize")?.value,qty=Math.floor(Number($("entryQty")?.value||0)),m=window.entryMatch;if(!code||!size||!qty||!m)return toast("Product code খুঁজে size নির্বাচন করুন এবং quantity দিন",true);if(m.sheet!==sheet){return toast("এই product-এর সঠিক stock sheet নির্বাচন করুন: "+m.sheet,true)}const sizeInfo=(m.sizes||[]).find(x=>x.size===size);if(!sizeInfo)return toast("Selected size is not available in the matched worksheet.",true);const j=await post("/api/excel-stock/update",{sheet:m.sheet,row:m.row,code:m.code,size,quantity:Number(sizeInfo.qty||0)+qty,note:"Stock entry"});if(!j._ok)return toast(j.error||"Stock could not be added",true);toast("Added "+bp(qty)+" units to "+size+" · previous stock "+bp(sizeInfo.qty||0));window.entryMatch=null;window.entryRow=null;await panel("stockentry")}function queueExcelCell(el){const quantity=el.value.trim(),original=el.dataset.original,oldQty=Number(original||0),newQty=Number(quantity);if(!isAdminLevelJS()&&!allowed("stock_manage")&&!allowed("stock_set")&&((newQty>oldQty&&!allowed("stock_add"))||(newQty<oldQty&&!allowed("stock_reduce")))){el.value=original;return toast("এই পরিবর্তনের অনুমতি আপনার নেই",true)}const key=[el.dataset.sheet,el.dataset.row,el.dataset.size].join('|');if(quantity===''||!Number.isInteger(Number(quantity))||Number(quantity)<0){el.value=original;return toast('০ বা তার বেশি পূর্ণসংখ্যা দিন',true)}if(quantity===original){delete excelPending[key];el.classList.remove('pending')}else{excelPending[key]={sheet:el.dataset.sheet,row:Number(el.dataset.row),code:el.dataset.code,size:el.dataset.size,quantity:Number(quantity)};el.classList.add('pending')}const count=Object.keys(excelPending).length;$('pendingCount').textContent=bp(count);const b=$('saveExcelBtn');if(b){b.disabled=!count;b.textContent='💾 Save Changes · '+bp(count)}}
async function saveExcelChanges(){const changes=Object.values(excelPending);if(!changes.length)return toast('সেভ করার মতো পরিবর্তন নেই');const b=$('saveExcelBtn');if(b){b.disabled=true;b.textContent='⏳ সেভ হচ্ছে…'}try{const j=await post('/api/excel-stock/update-batch',{changes});if(!j._ok)throw Error(j.error||'সেভ হয়নি');excelPending={};toast(`${bp(j.saved||0)}টি পরিবর্তন Excel-এ সেভ হয়েছে`);await panel('excelstock');setTimeout(poll,0)}catch(e){toast(e.message||'সেভ হয়নি',true);if(b){b.disabled=false;b.textContent='💾 Save Changes · '+bp(changes.length)}}}
function discardPendingForNavigation(){if(!Object.keys(excelPending).length)return true;if(!confirm('অসেভ করা পরিবর্তন বাদ দিয়ে এগোবেন? সেভ করতে Cancel চাপুন, তারপর Save Changes দিন।'))return false;excelPending={};return true}
async function openExcelSheet(name){if(!discardPendingForNavigation())return;excelSheet=name;excelOffset=0;await panel('excelstock')}
async function searchExcelStock(){if(!discardPendingForNavigation())return;excelQuery=$('excelQuery')?.value||'';excelOffset=0;await panel('excelstock')}
async function excelPage(delta){if(!discardPendingForNavigation())return;excelOffset=Math.max(0,excelOffset+delta*250);await panel('excelstock')}
// Purchase & Expense Management interactions (additive; separate from POS / sheet logic)
(function(){if(document.getElementById('erp-module-style'))return;const st=document.createElement('style');st.id='erp-module-style';st.textContent='.erp-module{max-width:1600px;margin:0 auto}.erp-module .visual-card{border:1px solid #dce9e1;border-radius:16px;background:linear-gradient(145deg,#fff,#f7fbf8);box-shadow:0 8px 26px #173d2a0c}.erp-module .dash-hero{background:linear-gradient(120deg,#0d7552,#164a3a);color:#fff;border-radius:18px;padding:22px;display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap}.erp-module .dash-hero h3{margin:0;color:#fff;font-size:22px}.erp-module .dash-hero small{display:block;margin-top:6px;color:#e1f6ea}.erp-module .dash-actions{display:flex;gap:8px;flex-wrap:wrap}.erp-module .dash-actions .pill{background:#ffffff18;color:#fff;border-color:#ffffff70}.erp-module .excel-stats>div{border:1px solid #dce9e1;border-radius:13px;background:#fff;box-shadow:0 4px 15px #0d75520a}.erp-module .excel-stats small{color:#597265}.erp-module .excel-stats b{color:#0d7552}.erp-module .co-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.erp-module .co-grid label{display:grid;gap:6px;font-size:13px;font-weight:650;color:#344a40;min-width:0}.erp-module .co-grid .wide{grid-column:1/-1}.erp-module .co-grid input,.erp-module .co-grid select,.erp-module .co-grid textarea,.erp-module .row input,.erp-module .visual-card>input{box-sizing:border-box;width:100%;min-width:0;padding:11px 12px;border:1px solid #cadbd1;border-radius:9px;background:#fff;color:#18392b;font:inherit}.erp-module .co-grid textarea{min-height:76px;resize:vertical}.erp-module .pill{white-space:nowrap}.erp-module .excel-grid th{background:#eaf5ee;color:#174532;position:sticky;top:0}.erp-module .excel-grid td{vertical-align:middle}.erp-module .excel-grid input,.erp-module .excel-grid select{padding:8px;border:1px solid #cfded4;border-radius:8px;min-width:80px;box-sizing:border-box}.erp-module .excel-table-wrap{max-width:100%;overflow:auto;border-radius:12px}.erp-module .row input{width:auto}.erp-module .order-summary-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-top:14px}.erp-module .stockstate{white-space:nowrap}@media(max-width:760px){.erp-module .co-grid,.erp-module .order-summary-grid{grid-template-columns:1fr}.erp-module .co-grid .wide{grid-column:auto}.erp-module .dash-hero{padding:16px}.erp-module .dash-hero h3{font-size:19px}.erp-module .excel-stats{grid-template-columns:repeat(2,minmax(0,1fr))}.erp-module .row{flex-wrap:wrap}}';document.head.appendChild(st)})();
window.erpSheets=window.erpSheets||[];window.erpPurchaseQuery=window.erpPurchaseQuery||'';window.erpExpenseQuery=window.erpExpenseQuery||'';window.erpReportFrom=window.erpReportFrom||'';window.erpReportTo=window.erpReportTo||'';
function erpMoney(n){return '৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2})}
function erpAddPurchaseItem(){const body=$('erpPurchaseItems');if(!body)return;const tr=document.createElement('tr');tr.className='erp-purchase-item-row';tr.innerHTML=`<td><input class="erp-pi-name" list="erpPurchaseProductNames" autocomplete="off" placeholder="Search / choose product name" aria-label="Product Name" required onchange="erpPurchaseNameChanged(this)" onkeydown="if(event.key==='Enter'){event.preventDefault();erpPurchaseNameChanged(this)}"><input type="hidden" class="erp-pi-code"><input type="hidden" class="erp-pi-sheet"><small class="erp-pi-stock-hint">Choose a product name to load available sizes and current quantities.</small></td><td><select class="erp-pi-size" aria-label="Size / Variant" required disabled><option value="">Select product first</option></select></td><td><input class="erp-pi-qty" type="number" min="1" step="1" value="1" aria-label="Quantity" oninput="erpCalcPurchaseTotals()" style="width:80px"></td><td><input class="erp-pi-cost" type="number" min="0" step="0.01" value="0" aria-label="Unit Cost" oninput="erpCalcPurchaseTotals()" style="width:120px"></td><td class="erp-pi-total">৳ 0</td><td><button type="button" class="pill" aria-label="Remove item" onclick="this.closest('tr').remove();if(!$('erpPurchaseItems')?.children.length)erpAddPurchaseItem();erpCalcPurchaseTotals()">×</button></td>`;body.appendChild(tr);erpCalcPurchaseTotals()}
async function erpPurchaseNameChanged(input){const tr=input?.closest('.erp-purchase-item-row');if(!tr)return;const name=String(input.value||'').trim(),seq=Number(tr.dataset.lookupSeq||0)+1;tr.dataset.lookupSeq=String(seq);const codeEl=tr.querySelector('.erp-pi-code'),sheetEl=tr.querySelector('.erp-pi-sheet'),sizeEl=tr.querySelector('.erp-pi-size'),hint=tr.querySelector('.erp-pi-stock-hint');if(codeEl)codeEl.value='';if(sheetEl)sheetEl.value='';if(sizeEl){sizeEl.innerHTML='<option value="">Loading sizes…</option>';sizeEl.disabled=true}if(!name){if(hint)hint.textContent='Choose a product name to load current stock sizes.';return}if(hint)hint.textContent='Finding product and stock sizes…';const j=await jf('/api/erp/stock-item-lookup?q='+encodeURIComponent(name));if(Number(tr.dataset.lookupSeq)!==seq)return;const match=(j.matches||[])[0];if(!j._ok||!match){if(hint)hint.textContent=(j.product?.title||name)+' — '+(j.error||'No matching product row / size columns found in the stock workbook.');if(sizeEl)sizeEl.innerHTML='<option value="">No stock sizes found</option>';return}if(codeEl)codeEl.value=match.code||'';if(sheetEl)sheetEl.value=match.sheet||'';if(input)input.value=j.product?.title||match.product_name||name;if(sizeEl){sizeEl.innerHTML='<option value="">Select size / variant</option>'+(match.sizes||[]).map(x=>`<option value="${esc(x.size)}">${esc(x.size)} · current ${Number(x.qty||0).toLocaleString('en-BD')} pcs</option>`).join('');sizeEl.disabled=!(match.sizes||[]).length}if(hint){const photo=(j.product?.img||match.product_img)?`<img src="${im(j.product?.img||match.product_img)}" alt="" style="width:32px;height:42px;object-fit:cover;border-radius:5px;vertical-align:middle;margin-right:6px" onerror="this.style.display='none'">`:'';hint.innerHTML=photo+'Current stock · '+(match.sizes||[]).map(x=>`${esc(x.size)}: ${Number(x.qty||0)} pcs`).join(' · ')} }
function erpPurchaseItemData(){return [...document.querySelectorAll('.erp-purchase-item-row')].map(tr=>({sheet:tr.querySelector('.erp-pi-sheet')?.value||'',code:tr.querySelector('.erp-pi-code')?.value.trim()||'',name:tr.querySelector('.erp-pi-name')?.value.trim()||'',size:tr.querySelector('.erp-pi-size')?.value.trim()||'',qty:Number(tr.querySelector('.erp-pi-qty')?.value||0),unit_cost:Number(tr.querySelector('.erp-pi-cost')?.value||0)}))}
function erpCalcPurchaseTotals(){let subtotal=0;document.querySelectorAll('.erp-purchase-item-row').forEach(tr=>{const qty=Math.max(0,Number(tr.querySelector('.erp-pi-qty')?.value||0)),cost=Math.max(0,Number(tr.querySelector('.erp-pi-cost')?.value||0)),total=qty*cost;subtotal+=total;const el=tr.querySelector('.erp-pi-total');if(el)el.textContent=erpMoney(total)});const discount=Math.max(0,Number($('erpPurchaseDiscount')?.value||0)),total=Math.max(0,subtotal-discount),paid=Math.max(0,Number($('erpPurchasePaid')?.value||0));if($('erpPurchaseSubtotal'))$('erpPurchaseSubtotal').textContent=erpMoney(subtotal);if($('erpPurchaseDiscountLabel'))$('erpPurchaseDiscountLabel').textContent=erpMoney(discount);if($('erpPurchaseTotal'))$('erpPurchaseTotal').textContent=erpMoney(total);if($('erpPurchaseDue'))$('erpPurchaseDue').textContent=erpMoney(Math.max(0,total-paid))}
async function erpSavePurchase(){const supplier_id=Number($('erpPurchaseSupplier')?.value||0),purchase_date=$('erpPurchaseDate')?.value||'',items=erpPurchaseItemData(),payload={supplier_id,purchase_date,invoice_no:$('erpPurchaseInvoice')?.value||'',discount:Number($('erpPurchaseDiscount')?.value||0),paid_amount:Number($('erpPurchasePaid')?.value||0),payment_method:$('erpPurchaseMethod')?.value||'Cash',notes:$('erpPurchaseNotes')?.value||'',items,receive_to_stock:true};if(!supplier_id)return toast('Select a supplier first.',true);if(!purchase_date)return toast('Choose the purchase date.',true);if(!items.length||items.some(x=>!x.name||!x.code||!Number.isInteger(x.qty)||x.qty<1||!Number.isFinite(x.unit_cost)||x.unit_cost<0))return toast('Choose a product name, valid quantity and unit cost on every line.',true);if(items.some(x=>!x.sheet||!x.size))return toast('Choose a product name first, then select a size / variant from its current stock.',true);if(!window.erpSheets.length)return toast('Excel stock workbook is not available. Set up the stock workbook before recording a purchase.',true);if(!confirm('Save this purchase invoice and payment details?'))return;const b=document.querySelector('[onclick="erpSavePurchase()"]');if(b){b.disabled=true;b.textContent='Saving…'}try{const j=await post('/api/erp/purchases',payload);if(!j._ok)throw Error(j.error||'Purchase could not be saved.');toast(j.stock_error?`Purchase saved, but stock was not updated: ${j.stock_error}`:j.stock_received?'Purchase saved and Excel stock updated. A PDF Memo button is available in Purchase History.': 'Purchase saved successfully. A PDF Memo button is available in Purchase History.');await panel('purchases')}catch(e){toast(e.message||'Purchase could not be saved.',true)}finally{if(b){b.disabled=false;b.textContent='💾 Save Purchase'}}}
async function erpFilterPurchases(){window.erpPurchaseQuery=$('erpPurchaseQuery')?.value||'';const j=await jf('/api/erp/purchases?limit=200&q='+encodeURIComponent(window.erpPurchaseQuery));if(!j._ok)return toast(j.error||'Could not filter purchases.',true);await panel('purchases')}
async function erpViewPurchase(id){const j=await jf('/api/erp/purchases/'+id);if(!j._ok)return toast(j.error||'Could not load purchase.',true);const p=j.purchase,lines=(j.items||[]).map(x=>`${x.item_name} (${x.item_code||'No code'} / ${x.size||'—'}) · ${x.quantity} × ${erpMoney(x.unit_cost)} = ${erpMoney(x.line_total)}`).join('\n');alert(`Purchase: ${p.purchase_no}\nSupplier: ${p.supplier_name}\nDate: ${p.purchase_date}\nInvoice: ${p.invoice_no||'—'}\n\n${lines}\n\nSubtotal: ${erpMoney(p.subtotal)}\nDiscount: ${erpMoney(p.discount)}\nTotal: ${erpMoney(p.total)}\nPaid on this invoice: ${erpMoney(p.paid)}\nInvoice due: ${erpMoney(p.due)}\nStock status: ${p.stock_received?'Received':'Not received'}`)}
function erpOpenPurchaseMemo(id){const a=document.createElement('a');a.href='/api/erp/purchases/'+encodeURIComponent(id)+'/memo.pdf';a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();a.remove()}
function erpDownloadPurchaseMemo(id){const a=document.createElement('a');a.href='/api/erp/purchases/'+encodeURIComponent(id)+'/memo.pdf?download=1';a.download='Purchase-Memo-'+id+'.pdf';document.body.appendChild(a);a.click();a.remove()}
function erpOpenPaymentForm(kind,id,due,title,sourcePanel){
 if(!(isAdminLevelJS()||allowed('purchase_manage')||allowed('supplier_manage')))return toast('You do not have permission to record supplier payments.',true);
 due=Math.max(0,Number(due)||0);if(due<=0)return toast('There is no outstanding due to pay.',true);
 erpClosePaymentModal();window.erpPaymentContext={kind:String(kind),id:Number(id),due:Math.round(due*100)/100,sourcePanel:sourcePanel||'erpReports'};
 const overlay=document.createElement('div');overlay.id='erpPaymentModal';overlay.style.cssText='position:fixed;inset:0;z-index:100050;background:rgba(8,24,17,.62);display:flex;align-items:center;justify-content:center;padding:16px;backdrop-filter:blur(3px)';
 overlay.innerHTML=`<form onsubmit="erpSubmitPayment(event)" style="width:min(520px,100%);max-height:92vh;overflow:auto;background:#fff;color:#173b2b;border-radius:18px;padding:22px;box-shadow:0 24px 70px #0004;border:1px solid #dce9e1" aria-label="Supplier payment form"><div style="display:flex;align-items:center;justify-content:space-between;gap:12px"><div><small style="color:#0d7552;font-weight:800;letter-spacing:.08em">ALIEN TEES ERP · PAYMENT</small><h3 style="margin:5px 0 0;font-size:21px">${esc(title||'Record Due Payment')}</h3></div><button type="button" class="pill" onclick="erpClosePaymentModal()" aria-label="Close payment form">✕</button></div><p style="margin:12px 0 16px;color:#53685d">Outstanding amount: <b style="color:#0d7552">${erpMoney(due)}</b></p><label style="display:grid;gap:6px;margin-bottom:12px;font-size:13px;font-weight:700">Payment Amount (BDT) *<input id="erpPayAmount" type="number" min="0.01" max="${due.toFixed(2)}" step="0.01" value="${due.toFixed(2)}" required style="box-sizing:border-box;width:100%;padding:12px;border:1px solid #cadbd1;border-radius:10px;font:inherit"></label><div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px"><label style="display:grid;gap:6px;font-size:13px;font-weight:700">Payment Date *<input id="erpPayDate" type="date" value="${new Date().toISOString().slice(0,10)}" required style="box-sizing:border-box;width:100%;padding:11px;border:1px solid #cadbd1;border-radius:10px;font:inherit"></label><label style="display:grid;gap:6px;font-size:13px;font-weight:700">Payment Method *<select id="erpPayMethod" required style="box-sizing:border-box;width:100%;padding:11px;border:1px solid #cadbd1;border-radius:10px;font:inherit"><option>Cash</option><option>Bank Transfer</option><option>bKash</option><option>Nagad</option><option>Card</option><option>Other</option></select></label></div><label style="display:grid;gap:6px;margin:12px 0;font-size:13px;font-weight:700">Reference / Transaction ID<input id="erpPayReference" maxlength="120" placeholder="Optional reference" style="box-sizing:border-box;width:100%;padding:11px;border:1px solid #cadbd1;border-radius:10px;font:inherit"></label><label style="display:grid;gap:6px;margin-bottom:16px;font-size:13px;font-weight:700">Notes<textarea id="erpPayNotes" maxlength="800" placeholder="Optional payment notes" style="box-sizing:border-box;width:100%;padding:11px;border:1px solid #cadbd1;border-radius:10px;font:inherit;min-height:64px"></textarea></label><div style="display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap"><button type="button" class="pill" onclick="erpClosePaymentModal()">Cancel</button><button id="erpPaySaveBtn" type="submit" class="pill" style="background:#0d7552;color:#fff;border-color:#0d7552">💾 Save Payment</button></div></form>`;
 overlay.addEventListener('click',e=>{if(e.target===overlay)erpClosePaymentModal()});document.body.appendChild(overlay);setTimeout(()=>document.getElementById('erpPayAmount')?.focus(),30);
}
function erpClosePaymentModal(){document.getElementById('erpPaymentModal')?.remove();delete window.erpPaymentContext}
async function erpSubmitPayment(event){
 event.preventDefault();const ctx=window.erpPaymentContext;if(!ctx)return;
 const amount=Number(document.getElementById('erpPayAmount')?.value||0),due=Number(ctx.due||0);
 if(!Number.isFinite(amount)||amount<=0||amount>due+0.009)return toast(`Enter an amount greater than zero and no more than ${erpMoney(due)}.`,true);
 const payment_date=document.getElementById('erpPayDate')?.value||'',payment_method=document.getElementById('erpPayMethod')?.value||'Cash',reference=document.getElementById('erpPayReference')?.value||'',notes=document.getElementById('erpPayNotes')?.value||'';
 if(!payment_date)return toast('Choose a payment date.',true);
 const button=document.getElementById('erpPaySaveBtn');if(button){button.disabled=true;button.textContent='Saving…'}
 const endpoint=ctx.kind==='purchase'?`/api/erp/purchases/${ctx.id}/payments`:`/api/erp/suppliers/${ctx.id}/payments`;
 try{const j=await post(endpoint,{amount:Math.round(amount*100)/100,payment_date,payment_method,reference,notes});if(!j._ok)throw Error(j.error||`Payment could not be saved (${j._s||'network error'}).`);const target=ctx.sourcePanel||'erpReports';erpClosePaymentModal();toast(`Payment saved successfully · ${erpMoney(amount)}`);await panel(target)}catch(e){if(button){button.disabled=false;button.textContent='💾 Save Payment'}toast(e.message||'Payment could not be saved.',true)}
}
async function erpRecordPurchasePayment(id,due,sourcePanel='purchases'){return erpOpenPaymentForm('purchase',id,due,'Pay Purchase Invoice',sourcePanel)}
async function erpReceivePurchaseStock(id){if(!confirm('Receive this purchase into the linked Excel stock workbook now? Stock quantities will increase.'))return;const j=await post(`/api/erp/purchases/${id}/receive-stock`,{});if(!j._ok)return toast(j.error||'Stock could not be updated.',true);toast('Purchase quantities added to Excel stock.');await panel('purchases')}
async function erpSaveSupplier(){const payload={name:$('erpSupplierName')?.value||'',phone:$('erpSupplierPhone')?.value||'',email:$('erpSupplierEmail')?.value||'',address:$('erpSupplierAddress')?.value||'',opening_balance:Number($('erpSupplierOpening')?.value||0),notes:$('erpSupplierNotes')?.value||''};if(!payload.name.trim())return toast('Supplier name is required.',true);const j=await post('/api/erp/suppliers',payload);if(!j._ok)return toast(j.error||'Supplier could not be saved.',true);toast('Supplier saved successfully.');await panel('suppliers')}
async function erpDownloadExcel(url,filename){try{const resp=await fetch(url,{credentials:'same-origin',headers:{'Accept':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}});const ct=(resp.headers.get('content-type')||'').toLowerCase();if(!resp.ok||!ct.includes('spreadsheetml')){let msg='Excel download failed (HTTP '+resp.status+').';try{const j=await resp.json();msg=j.error||j.message||msg;if(msg==='login_required')msg='Please log in again before downloading.';if(msg==='permission_denied')msg='You do not have permission to export this report.'}catch(_){}throw Error(msg)}const blob=await resp.blob();if(!blob.size)throw Error('The Excel report is empty.');const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=filename||'report.xlsx';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),1500);toast('Excel report downloaded.')}catch(e){toast(e.message||'Excel download failed.',true)}}
async function erpShowSupplierHistory(id){const j=await jf('/api/erp/suppliers/'+id+'/purchases?limit=1000');if(!j._ok)return toast(j.error||'Could not load supplier purchase history.',true);const s=j.supplier||{},rows=j.purchases||[];let total=0,paid=0,due=0,qty=0;rows.forEach(p=>{total+=Number(p.total||0);paid+=Number(p.paid||0);due+=Number(p.due||0);(p.items||[]).forEach(i=>qty+=Number(i.quantity||0))});const money=n=>'৳ '+Number(n||0).toLocaleString('en-BD',{maximumFractionDigits:2});const body=rows.map(p=>`<tr><td>${esc(p.purchase_date||'')}</td><td><b>${esc(p.purchase_no||'')}</b><small style="display:block">Invoice: ${esc(p.invoice_no||'—')}</small><button class="pill" style="margin-top:5px" onclick="erpOpenPurchaseMemo(${p.id})">📄 মেমো দেখুন</button> <button class="pill" onclick="erpDownloadPurchaseMemo(${p.id})">⬇ ডাউনলোড</button></td><td>${(p.items||[]).map(i=>`<div style="padding:5px 0;border-bottom:1px solid #eee"><b>${esc(i.item_name||'')}</b><small style="display:block">Code: ${esc(i.item_code||'—')} · Size: ${esc(i.size||'—')} · Qty: ${Number(i.quantity||0)} × ${money(i.unit_cost)}</small></div>`).join('')||'—'}</td><td>${money(p.total)}</td><td>${money(p.paid)}</td><td><b>${money(p.due)}</b></td><td>${p.stock_received?'Received':'Not received'}</td><td>${esc(p.created_by||'')}</td></tr>`).join('')||'<tr><td colspan="8" style="text-align:center;padding:22px">No purchase history recorded for this supplier yet.</td></tr>';const old=document.getElementById('erpSupplierHistoryModal');if(old)old.remove();document.body.insertAdjacentHTML('beforeend',`<div id="erpSupplierHistoryModal" style="position:fixed;inset:0;z-index:99999;background:rgba(8,18,15,.68);display:flex;align-items:center;justify-content:center;padding:16px" onclick="if(event.target===this)this.remove()"><div style="background:#fff;color:#20252a;width:min(1500px,98vw);max-height:92vh;overflow:auto;border-radius:16px;padding:20px;box-shadow:0 25px 80px #0005"><div style="display:flex;gap:12px;align-items:center;justify-content:space-between;flex-wrap:wrap"><div><h2 style="margin:0 0 5px">Purchase History — ${esc(s.name||'Supplier')}</h2><div style="color:#667085">${esc(s.phone||'')} · ${esc(s.email||'')}</div></div><div style="display:flex;gap:8px;flex-wrap:wrap"><button class="pill" onclick="erpDownloadExcel('/api/erp/suppliers/${id}/purchases/export.xlsx','Supplier-Purchase-History-${id}.xlsx')">⬇ Download Excel</button><button class="pill" onclick="document.getElementById('erpSupplierHistoryModal').remove()">✕ Close</button></div></div><div class="excel-stats" style="margin:16px 0"><div><small>Total Purchase</small><b>${money(total)}</b></div><div><small>Paid</small><b>${money(paid)}</b></div><div><small>Invoice Due</small><b>${money(due)}</b></div><div><small>Quantity Purchased</small><b>${qty.toLocaleString('en-US')}</b></div></div><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>Date</th><th>Purchase / Invoice</th><th>Product Details</th><th>Total</th><th>Paid</th><th>Due</th><th>Stock</th><th>Created By</th></tr></thead><tbody>${body}</tbody></table></div></div></div>`)}

function erpFilterSupplierRows(){const q=($('erpSupplierQuery')?.value||'').toLowerCase();document.querySelectorAll('#erpSupplierRows tr[data-search]').forEach(tr=>tr.style.display=tr.dataset.search.includes(q)?'':'none')}
async function erpToggleSupplier(id,active){const j=await post(`/api/erp/suppliers/${id}`,{active:!!active});if(!j._ok)return toast(j.error||'Supplier status was not updated.',true);toast(active?'Supplier activated.':'Supplier deactivated.');await panel('suppliers')}
async function erpRecordSupplierPayment(id,due,sourcePanel='suppliers'){return erpOpenPaymentForm('supplier',id,due,'Pay Supplier Due',sourcePanel)}
async function erpSaveExpense(){const payload={title:$('erpExpenseTitle')?.value||'',category:$('erpExpenseCategory')?.value||'Other',amount:Number($('erpExpenseAmount')?.value||0),expense_date:$('erpExpenseDate')?.value||'',payment_method:$('erpExpenseMethod')?.value||'Cash',paid_to:$('erpExpensePaidTo')?.value||'',reference:$('erpExpenseRef')?.value||'',notes:$('erpExpenseNotes')?.value||'',status:$('erpExpenseStatus')?.value||'Pending'};if(!payload.title.trim()||!payload.category||!payload.expense_date||!Number.isFinite(payload.amount)||payload.amount<=0)return toast('Enter title, category, valid amount and date.',true);const j=await post('/api/erp/expenses',payload);if(!j._ok)return toast(j.error||'Expense could not be saved.',true);toast('Expense saved successfully.');await panel('expenses')}
async function erpFilterExpenses(){window.erpExpenseQuery=$('erpExpenseQuery')?.value||'';await panel('expenses')}
async function erpSetExpenseStatus(id,status){if(!['Approved','Rejected'].includes(status))return;if(!confirm(`Mark this expense as ${status}?`))return;const j=await post(`/api/erp/expenses/${id}/status`,{status});if(!j._ok)return toast(j.error||'Expense status was not updated.',true);toast('Expense status updated.');await panel('expenses')}
async function erpApplyReports(){window.erpReportFrom=$('erpReportFrom')?.value||'';window.erpReportTo=$('erpReportTo')?.value||'';if(window.erpReportFrom&&window.erpReportTo&&window.erpReportFrom>window.erpReportTo)return toast('From date must not be after To date.',true);await panel('erpReports')}
async function erpDeletePurchase(id,purchaseNo,stockReceived){
 if(!isAdminLevelJS())return toast('Only Admin / Super Admin can delete purchase records.',true);
 const extra=stockReceived?'\n\nThis purchase was received into stock. The app will try to reverse those received quantities first; deletion will stop if it would make stock negative.':'';
 if(!confirm(`Delete purchase ${purchaseNo||'#'+id} and its linked invoice payments?${extra}\n\nThis action cannot be undone. Make a backup first.`))return;
 const j=await post(`/api/erp/purchases/${id}`,{},'DELETE');if(!j._ok)return toast(j.error||'Purchase could not be deleted.',true);
 toast(j.stock_reversed?'Purchase deleted and received stock quantities reversed.':'Purchase record deleted.');await panel('erpReports')
}
async function erpDeleteExpense(id,expenseNo){
 if(!isAdminLevelJS())return toast('Only Admin / Super Admin can delete expense records.',true);
 if(!confirm(`Delete expense ${expenseNo||'#'+id}? This action cannot be undone. Make a backup first.`))return;
 const j=await post(`/api/erp/expenses/${id}`,{},'DELETE');if(!j._ok)return toast(j.error||'Expense could not be deleted.',true);
 toast('Expense record deleted.');await panel('erpReports')
}
window.orderSearch=window.orderSearch||'';window.orderFrom=window.orderFrom||'';window.orderTo=window.orderTo||'';
function searchOrderLookup(){const c=$("orderConsQuery")?.value.trim()||"",p=$("orderPhoneQuery")?.value.trim()||"";window.orderSearch=c||p;$("orderQuery").value=window.orderSearch;searchOrderHistory(1)}function applyOrderDatePreset(preset){window.orderDatePreset=preset;if(preset==='custom')return;const end=new Date();end.setHours(0,0,0,0);const start=new Date(end);if(preset==='yesterday'){start.setDate(start.getDate()-1);end.setDate(end.getDate()-1)}else if(preset==='last7')start.setDate(start.getDate()-6);else if(preset==='last15')start.setDate(start.getDate()-14);else if(preset==='last30')start.setDate(start.getDate()-29);else if(preset==='month')start.setDate(1);else if(preset==='lastmonth'){start.setDate(1);start.setMonth(start.getMonth()-1);end.setDate(0)}else if(preset==='last3months'){start.setDate(1);start.setMonth(start.getMonth()-2)}else if(preset==='year'){start.setMonth(0,1)}else if(preset==='lastyear'){start.setFullYear(start.getFullYear()-1,0,1);end.setFullYear(end.getFullYear()-1,11,31)}else if(preset==='lifetime'){window.orderFrom='';window.orderTo='';return searchOrderHistory(1)}const fmt=d=>[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');window.orderFrom=fmt(start);window.orderTo=fmt(end);searchOrderHistory(1)}async function searchOrderHistory(page=1){window.orderSearch=$('orderQuery')?.value||'';window.orderFrom=$('orderFrom')?.value||'';window.orderTo=$('orderTo')?.value||'';const qs=new URLSearchParams({q:window.orderSearch,from:window.orderFrom,to:window.orderTo,page:String(page)});const j=await jf('/api/admin/order-history?'+qs.toString());if(!j.ok)return toast('অর্ডার খুঁজতে সমস্যা হয়েছে',true);await panel('orders');}
function toggleOrderSelectAll(checked){document.querySelectorAll('.order-row-check').forEach(x=>x.checked=checked)}
async function deleteSelectedOrders(){if(!isSuperAdminJS())return toast('শুধু Super Admin অর্ডার ডিলিট করতে পারবেন',true);const ids=[...document.querySelectorAll('.order-row-check:checked')].map(x=>Number(x.value)).filter(Number.isFinite);if(!ids.length)return toast('আগে অর্ডার সিলেক্ট করো',true);if(!confirm(`সিলেক্ট করা ${ids.length}টি অর্ডার স্থায়ীভাবে ডিলিট করবেন?`))return;const j=await post('/api/admin/order-history',{ids},'DELETE');if(!j._ok)return toast(j.error||'ডিলিট হয়নি',true);toast(`${bp(j.deleted||0)}টি অর্ডার ডিলিট হয়েছে`);await panel('orders')}
async function deleteAllOrders(){if(!isSuperAdminJS())return toast('শুধু Super Admin অর্ডার ডিলিট করতে পারবেন',true);if(!confirm('সতর্কতা: সব Order History স্থায়ীভাবে ডিলিট হবে। আগে Excel ব্যাকআপ নিয়েছেন তো?'))return;if(!confirm('শেষবার নিশ্চিত করুন—সব অর্ডার রেকর্ড মুছে ফেলবেন?'))return;const j=await post('/api/admin/order-history',{all:true},'DELETE');if(!j._ok)return toast(j.error||'ডিলিট হয়নি',true);toast(`${bp(j.deleted||0)}টি অর্ডার ডিলিট হয়েছে`);await panel('orders')}
async function uploadOrderHistory(input){const f=input.files?.[0];if(!f)return;if(!confirm('এই Excel-এর সব অর্ডার স্থায়ী Order History-তে যোগ হবে। আগের রেকর্ড মুছবে না। চালিয়ে যাবেন?')){input.value='';return}const fd=new FormData();fd.append('file',f);toast('অর্ডার রিপোর্ট যোগ হচ্ছে…');try{const r=await fetch('/api/admin/order-history/upload',{method:'POST',body:fd}),j=await r.json();if(!r.ok||!j.ok)throw Error(j.error||'আপলোড হয়নি');toast(`${bp(j.inserted||0)}টি অর্ডার স্থায়ী রেকর্ডে যোগ হয়েছে`);window.orderSearch='';window.orderFrom='';window.orderTo='';await panel('orders')}catch(e){toast(e.message||'আপলোড হয়নি',true)}input.value=''}
window.returnSearch=window.returnSearch||'';
async function syncOrderSheetReturns(){
 if(!isAdminLevelJS()&&!allowed('courier_update'))return toast('Courier API sync permission নেই',true);
 if(!confirm('Order Sheet-এর সর্বশেষ 100টি Consignment ID configured Courier API-তে যাচাই হবে। Returned status পাওয়া গেলে Return History-তে স্বয়ংক্রিয় এন্ট্রি হবে। চালাবেন?'))return;
 toast('Order Sheet থেকে Courier API status যাচাই হচ্ছে…');
 const j=await post('/api/admin/return-history/sync-courier',{courier:'all'});
 if(!j._ok)return toast(j.error||'API sync হয়নি',true);
 const err=(j.errors||[]).length?` · API error ${bp(j.errors.length)}`:'';
 toast(`API sync সম্পন্ন · যাচাই ${bp(j.checked||0)} · নতুন Return ${bp(j.returned||0)}${err}`);
 window.returnSearch='';await panel('returns');refreshNotificationCount();
}
async function uploadReturnHistory(input){const f=input.files?.[0];if(!f)return;if(!confirm('Return Excel-এর রেকর্ড Return History-তে যোগ হবে। চালিয়ে যাবেন?')){input.value='';return}const fd=new FormData();fd.append('file',f);try{const r=await fetch('/api/admin/return-history/upload',{method:'POST',body:fd}),j=await r.json();if(!r.ok||!j.ok)throw Error(j.error||'আপলোড হয়নি');toast(`${bp(j.inserted||0)}টি রিটার্ন যোগ হয়েছে`);window.returnSearch='';await panel('returns')}catch(e){toast(e.message||'আপলোড হয়নি',true)}input.value=''}
function showManualReturnEntry(){if(!isAdminLevelJS()&&!allowed('return_sheet_upload'))return toast('Return entry permission নেই',true);const html=`<div class="return-manual-modal" style="position:fixed;inset:0;background:#0f172a99;z-index:99999;display:flex;align-items:center;justify-content:center;padding:16px" onclick="if(event.target===this)this.remove()"><div style="background:#fff;color:#172033;width:min(680px,100%);max-height:90vh;overflow:auto;border-radius:18px;padding:20px"><div style="display:flex;justify-content:space-between;align-items:center"><h3>＋ নতুন Return Entry</h3><button class="pill" onclick="this.closest('.return-manual-modal').remove()">✕</button></div><div class="excel-actions"><input id="manualReturnCons" placeholder="Consignment ID *" required><input id="manualReturnCode" placeholder="Product Code *" required><input id="manualReturnAssigned" placeholder="Assigned To"><input id="manualReturnName" placeholder="Customer Name"><input id="manualReturnPhone" placeholder="Phone"><input id="manualReturnSize" placeholder="Size"><input id="manualReturnAmount" placeholder="Collected Amount"><input id="manualReturnStatus" placeholder="Courier Status" value="Returned"><input id="manualReturnReason" placeholder="Return Reason"><input id="manualReturnNote" placeholder="Note"></div><button class="pill" style="margin-top:12px" onclick="saveManualReturnEntry()">Save Return</button></div></div>`;document.body.insertAdjacentHTML('beforeend',html)}
async function saveManualReturnEntry(){const cons=$('manualReturnCons')?.value.trim(),code=$('manualReturnCode')?.value.trim();if(!cons||!code)return toast('Enter Consignment ID and Product Code.',true);const body={cons_id:cons,product_code:code,assigned_to:$('manualReturnAssigned')?.value,customer_name:$('manualReturnName')?.value,phone:$('manualReturnPhone')?.value,size:$('manualReturnSize')?.value,collected_amount:$('manualReturnAmount')?.value,pathao_bill_status:$('manualReturnStatus')?.value,return_reason:$('manualReturnReason')?.value,note:$('manualReturnNote')?.value};const btn=document.querySelector('.return-manual-modal button[onclick="saveManualReturnEntry()"]');if(btn){btn.disabled=true;btn.textContent='Saving…'}try{const j=await post('/api/admin/return-history/manual',body);if(!j._ok)throw Error(j.error||'Return entry could not be saved.');document.querySelector('.return-manual-modal')?.remove();toast('Return entry saved');await panel('returns')}catch(e){toast(e.message||'Return entry could not be saved.',true)}finally{if(btn){btn.disabled=false;btn.textContent='Save Return'}}}
function toggleReturnSelectAll(checked){document.querySelectorAll('.return-row-check').forEach(x=>x.checked=checked)}
async function deleteOneReturn(id){if(!isSuperAdminJS())return toast('শুধু Super Admin রিটার্ন ডিলিট করতে পারবেন',true);if(!confirm('এই Return record স্থায়ীভাবে ডিলিট করবেন?'))return;const j=await post('/api/admin/return-history',{ids:[id]},'DELETE');if(!j._ok)return toast(j.error||'ডিলিট হয়নি',true);toast('Return record ডিলিট হয়েছে');await panel('returns')}
async function deleteSelectedReturns(){if(!isSuperAdminJS())return toast('শুধু Super Admin রিটার্ন ডিলিট করতে পারবেন',true);const ids=[...document.querySelectorAll('.return-row-check:checked')].map(x=>Number(x.value)).filter(Number.isFinite);if(!ids.length)return toast('আগে রিটার্ন সিলেক্ট করুন',true);if(!confirm(`${ids.length}টি Return record ডিলিট করবেন?`))return;const j=await post('/api/admin/return-history',{ids},'DELETE');if(!j._ok)return toast(j.error||'ডিলিট হয়নি',true);toast(`${bp(j.deleted||0)}টি রিটার্ন ডিলিট হয়েছে`);await panel('returns')}
async function deleteAllReturns(){if(!isSuperAdminJS())return toast("শুধু Super Admin রিটার্ন ডিলিট করতে পারবেন",true);if(!confirm("সব Return History স্থায়ীভাবে মুছবেন? আগে Excel backup নিয়েছেন তো?"))return;const j=await post("/api/admin/return-history",{all:true},"DELETE");if(!j._ok)return toast(j.error||"ডিলিট হয়নি",true);toast(`${bp(j.deleted||0)}টি রিটার্ন ডিলিট হয়েছে`);await panel("returns")}async function searchReturnHistory(){window.returnSearch=$('returnQuery')?.value||'';await panel('returns')}
async function matchReturnOrder(cons,code){window.orderSearch=cons||code||'';await panel('orders')}
function selectOrderProduct(el,id,customer,code){document.querySelectorAll('.code-cell.is-selected').forEach(x=>x.classList.remove('is-selected'));el.classList.add('is-selected');el.closest('tr')?.classList.add('order-row-selected');document.querySelectorAll('.order-row-selected').forEach(x=>{if(x!==el.closest('tr'))x.classList.remove('order-row-selected')});openOrderComments(id,customer,code)}
async function openOrderComments(id,customer,code){const j=await jf('/api/admin/order-history/'+id+'/comments');if(!j.ok)return toast(j.error||'কমেন্ট লোড হয়নি',true);const html=`<div class="comment-modal" style="position:fixed;inset:0;background:#11182788;z-index:99999;display:flex;align-items:flex-start;justify-content:center;padding:20px 12px;overflow-y:auto;overscroll-behavior:contain;-webkit-overflow-scrolling:touch" onclick="if(event.target===this)this.remove()"><div style="background:#fff;color:#1f2937;width:min(600px,100%);max-height:calc(100dvh - 40px);overflow-y:auto;overscroll-behavior:contain;-webkit-overflow-scrolling:touch;border-radius:18px;padding:20px;margin:auto 0;flex-shrink:0;box-shadow:0 20px 70px #0004"><div style="display:flex;justify-content:space-between;gap:10px;align-items:center"><h3>💬 Order Comment · ${esc(customer)} / ${esc(code)}</h3><button class="pill" onclick="this.closest('.comment-modal').remove()">✕</button></div><div>${(j.comments||[]).map(c=>`<div style="padding:10px;border-bottom:1px solid #eee"><b>${esc(c.username)}</b> <small>${esc(new Date(c.created_at*1000).toLocaleString())}</small><p style="white-space:pre-wrap">${esc(c.message)}</p></div>`).join('')||'<p>এখনো কোনো কমেন্ট নেই।</p>'}</div><textarea id="orderCommentText" maxlength="2000" placeholder="এই অর্ডার সম্পর্কে কমেন্ট লিখুন…" style="width:100%;min-height:90px;padding:12px;border:1px solid #dce2ea;border-radius:10px;margin-top:12px;box-sizing:border-box"></textarea><button class="pill" style="margin-top:10px" onclick="saveOrderComment(${id})">কমেন্ট সেভ করুন</button></div></div>`;document.body.insertAdjacentHTML('beforeend',html)}
async function saveOrderComment(id){const el=$('orderCommentText'),message=el?.value.trim();if(!message)return toast('আগে কমেন্ট লিখুন',true);const j=await post('/api/admin/order-history/'+id+'/comments',{message});if(!j._ok)return toast(j.error||'কমেন্ট সেভ হয়নি',true);document.querySelector('.comment-modal')?.remove();toast('কমেন্ট সেভ হয়েছে');await openOrderComments(id,'অর্ডার','')}
async function uploadExcelStock(input){const f=input.files?.[0];if(!f)return;if(!discardPendingForNavigation()){input.value='';return}const fd=new FormData();fd.append('file',f);toast('Excel আপলোড হচ্ছে…');try{const r=await fetch('/api/admin/excel-stock/upload',{method:'POST',body:fd}),j=await r.json();if(!r.ok||!j.ok)throw Error(j.error||'আপলোড হয়নি');excelSheet=j.data?.selected?.name||'';toast('Excel আপলোড হয়েছে');await panel('excelstock');setTimeout(poll,0)}catch(e){toast(e.message||'আপলোড হয়নি',true)}input.value=''}
async function saveExcelCell(el){const old=el.dataset.old??el.defaultValue;el.dataset.old=el.value;el.disabled=true;try{const j=await post('/api/excel-stock/update',{sheet:el.dataset.sheet,row:Number(el.dataset.row),code:el.dataset.code,size:el.dataset.size,quantity:el.value});if(!j._ok)throw Error(j.error||'সেভ হয়নি');el.defaultValue=el.value;toast(`স্টক আপডেট: ${el.dataset.code} · ${el.dataset.size}`);setTimeout(poll,0)}catch(e){el.value=old;toast(e.message||'সেভ হয়নি',true)}finally{el.disabled=false}}
async function restoreLatestExcelBackup(){if(!discardPendingForNavigation())return;if(!confirm('সর্বশেষ ব্যাকআপ দিয়ে বর্তমান workbook ফিরিয়ে আনবেন?'))return;const j=await post('/api/admin/excel-stock/restore-latest',{});if(!j._ok)return toast(j.error||'ব্যাকআপ ফেরানো যায়নি',true);excelPending={};excelSheet='__all__';excelOffset=0;toast('সর্বশেষ ব্যাকআপ ফিরিয়ে আনা হয়েছে');await panel('excelstock')}
async function excelStockHistory(){if(!canSeeStockHistory())return toast('স্টক হিস্ট্রি দেখার অনুমতি নেই',true);if(!discardPendingForNavigation())return;const j=await jf('/api/excel-stock/history');const body=(j.edits||[]).map(x=>`<tr><td>${esc(new Date(x.created_at*1000).toLocaleString('en-US'))}</td><td>${esc(x.username)}</td><td>${esc(x.sheet)}</td><td>${esc(x.product_code)}</td><td>${esc(x.size)}</td><td>${bn(x.old_qty)} → ${bn(x.new_qty)}</td><td>${esc(x.note||'')}</td></tr>`).join('');const c=$('adminContent');if(c)c.innerHTML=`<div class="excel-hero"><div class="stepper"><span class="done">01 · ফাইল</span><i></i><span class="done">02 · ট্যাব</span><i></i><span class="done">03 · স্টক</span><i></i><span class="active">04 · ইতিহাস</span></div><div class="welcome"><div class="eyebrow">AUDIT TRAIL</div><h2>স্টক পরিবর্তনের ইতিহাস</h2><p>কে, কখন, কোন শিটের কোন পণ্যের পরিমাণ বদলেছেন—সব এখানে রেকর্ড থাকে।</p></div><button class="pill" onclick="panel('excelstock')">← ড্যাশবোর্ডে ফিরুন</button>${isSuperAdminJS()?'<button class="pill" onclick="clearStockHistory()">🗑 হিস্ট্রি ও কমেন্ট মুছুন</button>':''}<div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>সময়</th><th>ইউজার</th><th>ট্যাব</th><th>কোড</th><th>সাইজ</th><th>পরিবর্তন</th><th>নোট</th></tr></thead><tbody>${body||'<tr><td colspan="7" class="empty-sheet">এখনো কোনো পরিবর্তন নেই</td></tr>'}</tbody></table></div><h3>🧾 বিক্রির কমেন্ট</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>সময়</th><th>ইউজার</th><th>ট্যাব</th><th>প্রোডাক্ট কোড</th><th>সাইজ</th><th>বিক্রি</th><th>কমেন্ট</th></tr></thead><tbody>${(j.comments||[]).map(x=>`<tr><td>${esc(new Date(x.created_at*1000).toLocaleString('en-US'))}</td><td>${esc(x.username)}</td><td>${esc(x.sheet)}</td><td>${esc(x.product_code)}</td><td>${esc(x.size)}</td><td>${bn(x.sold_qty)} পিস</td><td>${esc(x.message||'—')}</td></tr>`).join('')||'<tr><td colspan="7">কোনো বিক্রির কমেন্ট নেই</td></tr>'}</tbody></table></div><h3>ফাইল আপলোডের ইতিহাস</h3><div class="excel-table-wrap"><table class="excel-grid"><thead><tr><th>সময়</th><th>ফাইল</th><th>আপলোড করেছেন</th><th>আকার</th></tr></thead><tbody>${(j.imports||[]).map(x=>`<tr><td>${esc(new Date(x.created_at*1000).toLocaleString('en-US'))}</td><td>${esc(x.filename)}</td><td>${esc(x.username)}</td><td>${bn(Math.round(x.bytes/1024))} KB</td></tr>`).join('')||'<tr><td colspan="4">কোনো আপলোড নেই</td></tr>'}</tbody></table></div></div>`}
async function addSrc(kind,t){const j=await post('/api/admin/sources',{kind,name:$('sn').value,url:$('su').value});if(j._ok){toast('যোগ হয়েছে');panel(t);setTimeout(poll,0)}else toast(j.error||'সমস্যা')}
async function rmSrc(id,t){if(!confirm('রিমুভ করবেন?'))return;await post('/api/admin/sources/'+id,{},'DELETE');panel(t)}
function userRow(u){
 const p=u.permissions||{},isSA=u.role==='super_admin',isAdmin=u.role==='admin',isCS=u.role==='cs',isLegacy=u.role==='user',isStaff=u.role==='staff'||isLegacy;
 const lockRole=isSA||(isAdmin&&meRole!=='super_admin');
 const roleOpts=isSA
   ? `<option value="super_admin" selected>Super Admin</option>`
   : `<option value="staff" ${isStaff?'selected':''}>Staff</option><option value="cs" ${isCS?'selected':''}>CS</option>${meRole==='super_admin'?`<option value="admin" ${isAdmin?'selected':''}>Admin</option>`:''}`;
 const stockChecks=Object.entries(PERMS_LABELS).filter(([k])=>k.startsWith('stock_')).map(([k,v])=>`<label><input type="checkbox" data-u="${u.id}" data-p="${k}" ${p[k]?'checked':''}> ${v}</label>`).join('');
 const checkboxes=!isSA&&!isAdmin?`<div class="row" style="margin:10px 0"><button class="pill" type="button" onclick="setUserPerms(${u.id},true)">সব Permission চালু</button><button class="pill" type="button" onclick="setUserPerms(${u.id},false)">সব Permission বন্ধ</button></div><div class="perms">${Object.entries(PERMS_LABELS).map(([k,v])=>`<label><input type="checkbox" data-u="${u.id}" data-p="${k}" ${p[k]?'checked':''}> ${v}</label>`).join('')}</div>`
   : isSA?`<small style="color:#777"> Super Admin — সব permission চালু। এই account UI থেকে demote/delete করা যাবে না।</small>`
   : `<small style="color:#777">Admin হলে সব permission স্বয়ংক্রিয়ভাবে চালু থাকবে। Admin manage করতে Super Admin দরকার।</small>`;
 return `<div class="userbox"><div class="row" style="justify-content:space-between"><b>${esc(u.username)}</b><span style="color:#888">${esc(roleLabel(u.role))}</span></div><div class="row"><select id="role_${u.id}" ${lockRole?'disabled':''}>${roleOpts}</select><input id="pw_${u.id}" type="password" placeholder="${isSA?'Super Admin password পরিবর্তন করুন':'নতুন পাসওয়ার্ড (ঐচ্ছিক)'}" style="flex:1" ${isSA?'disabled':''}><button class="pill" onclick="saveUser(${u.id})" ${isSA?'disabled':''}>সেভ</button>${meRole==='super_admin'?`<button class="pill" onclick="deleteUser(${u.id})">মুছুন</button>`:''}</div>${checkboxes}</div>`;
}
let logUserFilter='',logActionFilter='';
const ACTIVITY_LABELS={text_search:'টেক্সট সার্চ',image_search:'ইমেজ সার্চ',product_view:'প্রোডাক্ট দেখা',filter_size:'সাইজ ফিল্টার',filter_stock:'স্টক ফিল্টার',filter_brand:'Brand ফিল্টার',filter_sheet:'শিট/ক্যাটাগরি ফিল্টার',sort_price:'দাম সাজানো',product_select:'প্রোডাক্ট সিলেক্ট',quick_mode:'দ্রুত মোড',link_copy:'লিংক কপি',selection_clear:'সিলেকশন ক্লিয়ার',id_copy:'ID কপি',name_copy:'নাম কপি',full_info_copy:'Full Info কপি',share_product:'Share',similar_click:'Similar click',login:'লগইন',login_failed:'লগইন ব্যর্থ',user_add:'ইউজার যোগ',user_edit:'ইউজার এডিট',user_delete:'ইউজার মুছে ফেলা',download:'HD ডাউনলোড',copy:'HD কপি',price_edit:'দাম এডিট',favorite_add:'ফেভারিট যোগ',favorite_remove:'ফেভারিট বাদ',product_report:'প্রোডাক্ট রিপোর্ট',admin_password_change:'Admin password change',stock_threshold:'Stock threshold',stock_sold_comment:'বিক্রির কমেন্ট',stock_sold_comment_delete:'কমেন্ট মুছেছেন',stock_history_clear:'স্টক হিস্ট্রি মুছেছেন',guest_perms_update:'Guest permission update'};
function replaySearch(q){$('ao').classList.remove('on');$('q').value=q||'';res=null;render(1);$('q').focus()}
function logRow(x){const d=new Date((x.created_at||0)*1000);const tm=d.toLocaleString('en-US',{dateStyle:'short',timeStyle:'medium'});const act=ACTIVITY_LABELS[x.action]||x.action;return `<tr><td>${esc(tm)}</td><td><b>${esc(x.username)}</b></td><td>${esc(act)}</td><td style="word-break:break-word">${x.action==='text_search'&&x.query?`<button class="copycode" onclick="replaySearch('${esc(x.query).replace(/\\/g,'\\\\').replace(/'/g,"\\'")}')">${esc(x.query)}</button>`:esc(x.query||'')}</td><td style="word-break:break-word">${esc(x.details||'')}</td></tr>`}
function setUserPerms(id,on){document.querySelectorAll(`[data-u="${id}"][data-p]`).forEach(e=>e.checked=!!on)}
/* ===================== v-fix: Courier state, My Company, POS অর্ডার ও মেমো ===================== */
(function(){const st=document.createElement('style');st.textContent=`
.co-page,.poso{max-width:1450px;margin:0 auto;color:#202a36}
.co-card,.poso-card{background:#fff!important;border:1px solid #e1e6ed!important;border-radius:14px!important;box-shadow:0 5px 18px #1720330c!important;padding:18px;margin:14px 0;backdrop-filter:none!important}
.co-card h3,.poso-card h3{margin:0 0 12px}
.co-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}
.co-grid label{display:grid;gap:6px;font-size:13px;font-weight:650;color:#344054;min-width:0}
.co-grid label.wide{grid-column:1/-1}
.co-grid input,.co-grid textarea{width:100%;box-sizing:border-box;padding:11px 12px;border:1px solid #d5dce5!important;border-radius:8px!important;background:#fff!important;font:inherit;color:#202a36;box-shadow:none!important}
.co-grid textarea{min-height:92px;resize:vertical}
.co-logo{display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.co-logo .lg{width:86px;height:86px;border-radius:10px;border:1px solid #d5dce5;background:#111 center/contain no-repeat;display:grid;place-items:center;color:#9aa;font-size:11px;overflow:hidden}
.co-logo .lg img{width:100%;height:100%;object-fit:contain}
.co-foot{display:flex;justify-content:flex-end;margin-top:18px}
.co-save{background:#1d4ed8;color:#fff;border:0;border-radius:8px;padding:11px 22px;font-weight:700;cursor:pointer}
.poso-head{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap}
.poso-head h2{margin:2px 0;font-size:26px}
.poso-tabs,.poso-chips{display:flex;gap:8px;flex-wrap:wrap}
.poso-tabs button,.poso-chips button{border:1px solid #d9e0e8;background:#fff;color:#344054;border-radius:999px;padding:8px 14px;cursor:pointer;font-weight:650}
.poso-tabs button.on,.poso-chips button.on{background:#164a3a;color:#fff;border-color:#164a3a}
.poso-filters{display:flex;gap:8px;flex-wrap:wrap;align-items:end;margin-top:12px}
.poso-filters label{display:grid;gap:4px;font-size:11px;font-weight:650;color:#68766f}
.poso-filters input,.poso-filters select{padding:9px 10px;border:1px solid #d5dce5!important;border-radius:8px!important;background:#fff!important;color:#202a36;min-width:130px}
.poso-kpi{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:12px;margin:14px 0}
.poso-kpi article{background:#fff;border:1px solid #e1e6ed;border-radius:14px;padding:14px;box-shadow:0 5px 18px #1720330c}
.poso-kpi small{display:block;color:#667085}.poso-kpi b{font-size:22px;color:#172b3b}
.poso-charts{display:grid;grid-template-columns:2fr 1fr 1fr;gap:14px}
.poso-bars{display:flex;align-items:flex-end;gap:4px;height:120px}
.poso-bars i{flex:1;min-width:4px;background:linear-gradient(#27a56f,#176b4c);border-radius:4px 4px 0 0;position:relative}
.poso-hbar{display:grid;grid-template-columns:90px 1fr auto;gap:8px;align-items:center;font-size:12px;margin:6px 0}
.poso-hbar span.t{height:9px;background:#edf1f5;border-radius:6px;overflow:hidden}.poso-hbar span.t i{display:block;height:100%;background:#27a56f}
.poso-act{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:10px 0}
.poso-act .sel{font-weight:700;color:#164a3a}
.poso-table{width:100%;border-collapse:collapse;font-size:13px;min-width:1000px}
.poso-table th{background:#eaf5ef;text-align:left;padding:9px 8px;position:sticky;top:0}
.poso-table td{padding:8px;border-bottom:1px solid #eef1f4;vertical-align:middle}
.poso-table img{width:40px;height:50px;object-fit:cover;border-radius:6px;background:#eee}
.poso-table .ph{width:40px;height:50px;border-radius:6px;background:#e4efe9;display:grid;place-items:center;font-size:9px;color:#5a7;text-align:center}
.poso-pill{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11px;font-weight:700;background:#e4efe9;color:#164a3a}
.poso-pill.due{background:#fdecec;color:#b42318}
@media(max-width:1000px){.poso-kpi{grid-template-columns:repeat(3,minmax(0,1fr))}.poso-charts{grid-template-columns:1fr}.co-grid{grid-template-columns:1fr}}
`;document.head.appendChild(st)})();
const fmtN=n=>Number(n||0).toLocaleString('en-US',{maximumFractionDigits:2}),taka=n=>'৳ '+fmtN(n);

/* ---------- Courier: refresh হলেও search/filter/ফলাফল থাকবে ---------- */
const COURIER_STATE_KEYS=['courierQuery','courierFrom','courierTo','courierStatus','courierAssigned','activeCourier'];
function loadCourierState(){if(window.courierStateLoaded)return;window.courierStateLoaded=true;try{const s=JSON.parse(sessionStorage.getItem('courierState')||'{}');COURIER_STATE_KEYS.forEach(k=>{if(window[k]===undefined&&s[k]!==undefined)window[k]=s[k]})}catch(e){}}
function saveCourierState(){try{const s={};COURIER_STATE_KEYS.forEach(k=>{if(window[k]!==undefined)s[k]=window[k]});sessionStorage.setItem('courierState',JSON.stringify(s))}catch(e){}}
function courierKeepLoad(){if(!window.courierKeep){try{window.courierKeep=JSON.parse(sessionStorage.getItem('courierKeep')||'{}')}catch(e){window.courierKeep={}}}return window.courierKeep}
document.addEventListener('input',e=>{const el=e.target;if(!el||!el.id||!el.closest||!el.closest('.courier-dashboard'))return;if(['fraudPhone','trackCons','trackRating','trackDetails'].includes(el.id)){const k=courierKeepLoad();k[el.id]=el.value;try{sessionStorage.setItem('courierKeep',JSON.stringify(k))}catch(_){}}});
function restoreCourierInputs(){const k=courierKeepLoad();Object.keys(k).forEach(id=>{const el=document.getElementById(id);if(el&&!el.value)el.value=k[id]})}
async function courierPanel(){
 loadCourierState();saveCourierState();
 try{
  const html=await courierPanelMain();
  setTimeout(restoreCourierInputs,30);
  const err=window.courierLoadErr?`<div class="admin-note" style="border-color:#e0a0a0!important">⚠ ${esc(window.courierLoadErr)} <button class="pill" onclick="panel('couriers')">আবার চেষ্টা</button></div>`:'';
  return err+html;
 }catch(e){console.error(e);return `<div class="courier-dashboard"><section class="courier-card"><h3>Courier ডাটা লোড হয়নি</h3><p>${esc(String(e&&e.message||e))}</p><button class="pill" onclick="panel('couriers')">আবার চেষ্টা করুন</button></section></div>`}
}
function resetCourierFilters(){window.courierQuery='';window.courierFrom='';window.courierTo='';window.courierStatus='';window.courierAssigned='';panel('couriers')}
function fraudBrief(r){if(r==null)return '';if(typeof r!=='object')return String(r);if(r.error)return 'API error: '+r.error;const o=(r.data&&typeof r.data==='object')?r.data:r;const parts=Object.entries(o).filter(([k,v])=>v!==null&&typeof v!=='object').slice(0,8).map(([k,v])=>k+': '+v);return parts.join(', ')||JSON.stringify(r).slice(0,180)}
function fraudHistoryTable(fh){const rows=(fh&&fh.checks)||[];if(!rows.length)return '';return `<div class="excel-table-wrap" style="margin-top:12px"><table class="excel-grid"><thead><tr><th>সময়</th><th>Phone</th><th>Courier</th><th>Orders</th><th>Delivered</th><th>Returned</th><th>Courier API</th><th>By</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${new Date(x.checked_at*1000).toLocaleString('en-US')}</td><td>${esc(x.phone)}</td><td>${esc(x.courier||'Local')}</td><td>${x.orders}</td><td>${x.delivered}</td><td>${x.returned}</td><td>${esc(x.provider_brief||'')}</td><td>${esc(x.checked_by||'')}</td></tr>`).join('')}</tbody></table></div>`}
async function runFraudCheck(){const phone=$('fraudPhone').value.trim();if(!phone)return toast('Phone number দিন',true);const el=$('fraudResult');el.textContent='সব courier-এ history খোঁজা হচ্ছে…';const j=await jf('/api/admin/courier-fraud-check?courier=all&phone='+encodeURIComponent(phone));if(!j._ok){el.textContent=j.error||'Fraud check হয়নি';return}const parts=(j.providers||[]).map(p=>p.courier+': '+fraudBrief(p.result));const text=`${phone} · অর্ডার ${bp(j.orders||0)} · Delivered ${bp(j.delivered||0)} · Returned ${bp(j.returned||0)} · Return rate ${Number(j.return_rate||0).toFixed(0)}%`+(parts.length?' · '+parts.join(' | '):' · কোনো courier-এ Fraud API path সেট নেই (শুধু নিজের Order/Return history)');window.courierFraudText=text;try{sessionStorage.setItem('courierFraudLast',text)}catch(e){}el.textContent=text;panel('couriers')}
function courierSoftRefresh(){if(sessionStorage.getItem('activeAdminPanel')!=='couriers')return;const a=document.activeElement;if(a&&a.closest&&a.closest('.courier-dashboard')&&/^(INPUT|SELECT|TEXTAREA)$/.test(a.tagName)){if(!window.courierPendingRefresh){window.courierPendingRefresh=true;a.addEventListener('blur',()=>{window.courierPendingRefresh=false;setTimeout(courierSoftRefresh,400)},{once:true})}return}panel('couriers')}
function startCourierAutoSync(rows){
 if(window.courierAutoTimer){clearInterval(window.courierAutoTimer);window.courierAutoTimer=null}
 if(localStorage.getItem('courierAutoSync')==='off')return;
 window.courierAutoRows=rows||[];
 const run=async()=>{
  if(sessionStorage.getItem('activeAdminPanel')!=='couriers'){clearInterval(window.courierAutoTimer);window.courierAutoTimer=null;return}
  if(window.courierAutoBusy)return;window.courierAutoBusy=true;let changed=false;
  try{
   const active=(window.courierAutoRows||[]).filter(x=>x.cons_id&&!/return|cancel/i.test(String(x.status||''))).sort((a,b)=>Number(a.updated_at||0)-Number(b.updated_at||0));
   if(active.length){
    let index=window.courierAutoIndex||0;const cycle=[...Array(Math.min(5,active.length))].map((_,i)=>active[(index+i)%active.length]);window.courierAutoIndex=(index+cycle.length)%active.length;let done=0;
    for(const x of cycle){
     try{const prior=x.status,result=await post('/api/admin/courier-sync',{courier:x.courier||window.activeCourier||'Steadfast',cons_id:x.cons_id});if(result._ok){x.status=result.status;if(prior!==result.status)changed=true}}catch(e){}
     done++;const label=document.getElementById('courierSyncProgress');if(label)label.textContent='Order Sheet status sync: '+done+' / '+cycle.length+' · '+x.cons_id}
   }
   try{sessionStorage.setItem('courierAutoLast',String(Date.now()))}catch(e){}
  }finally{window.courierAutoBusy=false}
  if(changed){refreshNotificationCount();courierSoftRefresh()}
 };
 window.courierAutoTimer=setInterval(run,600000);
 const last=Number(sessionStorage.getItem('courierAutoLast')||0);
 if(Date.now()-last>120000)run();
}
function canOpenSavedPanel(t){if(isAdminLevelJS())return true;if(t==='excelstock')return canUseStockDashboard();if(t==='couriers')return allowed('courier_view')||allowed('courier_update');if(t==='orders')return allowed('order_view');if(t==='returns')return allowed('return_view');if(t==='price')return allowed('price_edit');if(t==='posorders')return allowed('pos_orders')||allowed('pos_delete');return false}

/* ---------- My Company ---------- */
async function companyPanel(){
 window.companyEditKey=window.companyEditKey||'alien_tees';
 const key=window.companyEditKey==='bodyfitters'?'bodyfitters':'alien_tees';
 const j=await jf('/api/company?company='+encodeURIComponent(key));const c=j.company||{},mt=j.sales_memo_template||j.purchase_memo_template||{};
 const f=(id,label,val,ph,type)=>`<label>${label}<input id="${id}" type="${type||'text'}" value="${esc(val||'')}" placeholder="${esc(ph||'')}" oninput="coMemoLivePreview()"></label>`;
 const moneyField=(id,label,val)=>`<label>${label}<input id="${id}" type="number" min="0" step="0.01" inputmode="decimal" value="${esc(val==null?'0':val)}" placeholder="0.00"></label>`;
 const col=/^#[0-9a-fA-F]{6}$/.test(mt.color||'')?mt.color:'#0D7552';
 const title=mt.title||(key==='bodyfitters'?'BODYFITTERS SALES MEMO':'SALES MEMO'),subtitle=mt.subtitle||'Customer invoice, itemized sale and payment summary',footer=mt.footer||c.invoice_note||c.name||(key==='bodyfitters'?'Bodyfitters':'Alien Tees');
 return `<div class="co-page"><span class="eyebrow">ALIEN TEES ERP · সেটিংস</span><h2 style="margin:4px 0">আমার প্রতিষ্ঠান</h2><p class="admin-note">প্রতিষ্ঠানের তথ্য ও লোগো এখানে সেট করুন। POS বিক্রির PDF মেমোর ডিজাইন ও QR এই কোম্পানি অনুযায়ী নিয়ন্ত্রণ হবে।</p>
 <div class="co-card"><div class="co-grid">
 <label class="wide" style="grid-column:1/-1;background:#eff8f1;border:1px solid #d5e8da;border-radius:12px;padding:12px">কোন কোম্পানির তথ্য ও মেমো edit করবেন?<select id="coCompanySelect" style="margin-top:6px" onchange="companySwitchEdit(this.value)"><option value="alien_tees" ${key==='alien_tees'?'selected':''}>Alien Tees</option><option value="bodyfitters" ${key==='bodyfitters'?'selected':''}>Bodyfitters</option></select><small style="display:block;color:#667085;margin-top:5px">দুই কোম্পানির নাম, ফোন, ঠিকানা, লোগো ও মেমো টেমপ্লেট আলাদা করে সংরক্ষণ হবে।</small></label>
 ${f('coName','প্রতিষ্ঠানের নাম *',c.name,key==='bodyfitters'?'Bodyfitters':'Alien Tees')}${f('coEmail','ইমেইল',c.email,'info@example.com','email')}
 ${f('coPhone','ফোন / হটলাইন (মেমোতে দেখাবে)',c.phone,'01XXXXXXXXX')}${f('coWebsite','ওয়েবসাইট',c.website,'https://')}
 ${f('coTin','ট্যাক্স / টিআইএন নম্বর',c.tin,'')}${f('coReg','রেজিস্ট্রেশন নম্বর',c.reg_no,'')}
 <div class="wide" style="grid-column:1/-1;border-top:1px solid #e6eaf0;padding-top:14px;margin-top:4px"><h3 style="margin:0 0 4px;font-size:16px">ডেলিভারি চার্জ (টাকা)</h3><p style="margin:0;color:#667085;font-size:12px">প্রতিটি এলাকার ডেলিভারি চার্জ নির্ধারণ করুন। আগের অর্ডার ও স্টকের তথ্য পরিবর্তন হবে না।</p></div>
 ${moneyField('coDeliveryInside','ঢাকার ভেতরে',c.delivery_inside_dhaka)}
 ${moneyField('coDeliverySuburbs','ঢাকার আশপাশ (সাভার / গাজীপুর / নারায়ণগঞ্জ)',c.delivery_dhaka_suburbs)}
 ${moneyField('coDeliveryOutside','ঢাকার বাইরে',c.delivery_outside_dhaka)}
 <label class="wide">ঠিকানা (মেমোর ডানে দেখাবে)<textarea id="coAddress">${esc(c.address||'')}</textarea></label>
 <label class="wide">প্রতিষ্ঠানের নোট (মেমোর নিচে — ডেলিভারি/রিটার্ন নিয়ম)<textarea id="coNote">${esc(c.company_note||'')}</textarea></label>
 ${f('coInv','ইনভয়েসের উপরের লাইন',c.invoice_note,'কম্পিউটারে তৈরি ইনভয়েস')}
 <div class="wide" style="grid-column:1/-1;border-top:1px solid #e6eaf0;padding-top:15px;margin-top:8px"><h3 style="margin:0 0 5px;font-size:16px">${key==='bodyfitters'?'Bodyfitters':'Alien Tees'} মেমো টেমপ্লেট</h3><p style="margin:0 0 12px;color:#667085;font-size:12px">এই কোম্পানির POS Sales Memo-র শিরোনাম, সাবটাইটেল, নিচের লেখা, রং ও লোগো আলাদাভাবে বদলাতে পারবেন। QR স্ক্যান করলে সরাসরি Memo Text দেখাবেন, নাকি Memo PDF খুলবেন—নিচের সেটিং থেকে বেছে নিন। ক্রয় মেমোর আগের ডিজাইন সেটিংস আলাদা থাকবে।</p>
 <div class="co-grid">${f('coMemoTitle','মেমোর শিরোনাম',title,'SALES MEMO')}${f('coMemoSubtitle','শিরোনামের নিচের লেখা',subtitle,'')}${f('coMemoFooter','মেমোর নিচের লেখা',footer,'')}
 <label>মেমোর প্রধান রং<input id="coMemoColor" type="color" value="${col}" oninput="coMemoLivePreview()"></label>
 <label style="display:flex;align-items:center;gap:8px"><input id="coMemoShowLogo" type="checkbox" ${mt.show_logo!==false?'checked':''} onchange="coMemoLivePreview()"> PDF মেমোতে লোগো দেখান</label><label>QR স্ক্যান করলে কী দেখাবে?<select id="coQrMode"><option value="text" ${(c.qr_mode||'text')==='text'?'selected':''}>সরাসরি Memo Text</option><option value="pdf" ${c.qr_mode==='pdf'?'selected':''}>Memo PDF খুলবে</option></select><small style="display:block;color:#667085">এই কোম্পানির নতুন Sales/Order/Purchase Memo QR-এর জন্য প্রযোজ্য।</small></label></div>
 <div id="coMemoPreview" style="margin-top:14px;border:1px solid #d7e4dc;border-radius:14px;overflow:hidden;background:#fff"><div id="coMemoPreviewHead" style="background:${col};color:white;padding:14px;display:flex;align-items:center;gap:12px"><span id="coMemoPreviewLogo" style="display:${mt.show_logo!==false&&j.has_logo?'block':'none'};background:#fff;border-radius:8px;width:54px;height:54px;overflow:hidden;flex:0 0 54px">${j.has_logo?`<img src="/api/company/logo?company=${key}&v=${j.logo_v||0}" style="width:100%;height:100%;object-fit:contain">`:''}</span><span><b id="coMemoPreviewTitle">${esc(title)}</b><small id="coMemoPreviewSubtitle" style="display:block;margin-top:4px">${esc(subtitle)}</small><small style="display:block;margin-top:5px">${esc(c.name||'Alien Tees')} · ${esc(c.phone||'')}</small></span></div><div style="padding:16px;color:#25352d"><div style="font-size:12px;color:#68766e">ইনভয়েস নম্বর · কাস্টমার · তারিখ</div><table style="width:100%;border-collapse:collapse;margin-top:10px;font-size:12px"><tr style="background:#edf5ef"><th style="text-align:left;padding:8px">পণ্যের কোড</th><th style="text-align:left;padding:8px">পণ্যের নাম</th><th style="text-align:left;padding:8px">সাইজ</th><th style="text-align:right;padding:8px">পরিমাণ</th><th style="text-align:right;padding:8px">মোট</th></tr><tr><td style="padding:8px;border-bottom:1px solid #eee">AT-001</td><td style="padding:8px;border-bottom:1px solid #eee">নমুনা পণ্য</td><td style="padding:8px;border-bottom:1px solid #eee">L</td><td style="padding:8px;text-align:right;border-bottom:1px solid #eee">2</td><td style="padding:8px;text-align:right;border-bottom:1px solid #eee">৳ 1,000</td></tr></table><div style="display:flex;justify-content:space-between;gap:10px;margin-top:25px;text-align:center;font-size:11px"><span>________________<br>প্রস্তুতকারীর স্বাক্ষর</span><span>________________<br>অনুমোদনকারীর স্বাক্ষর</span><span>________________<br>গ্রহণকারীর স্বাক্ষর</span></div><small id="coMemoPreviewFooter" style="display:block;border-top:1px solid #eee;margin-top:16px;padding-top:9px;color:#68766e">${esc(footer)}</small></div></div></div>
 <div class="wide" style="grid-column:1/-1;border-top:1px solid #e6eaf0;padding-top:15px;margin-top:8px"><b style="font-size:13px">প্রতিষ্ঠানের লোগো</b><div class="co-logo" style="margin-top:8px"><div class="lg" id="coLogoPrev">${j.has_logo?`<img src="/api/company/logo?company=${key}&v=${j.logo_v||0}">`:'লোগো যোগ করা হয়নি'}</div><div><input id="coLogo" type="file" accept="image/png,image/jpeg,image/webp" onchange="coLogoPreview(this)"><div style="font-size:12px;color:#667085;margin-top:6px">সর্বোচ্চ ২ এমবি PNG, JPG বা WebP ছবি।</div>${j.has_logo?`<label style="display:flex;gap:6px;align-items:center;margin-top:8px;font-weight:500"><input id="coRmLogo" type="checkbox"> লোগো মুছে ফেলুন</label>`:''}</div></div></div>
 </div><div class="co-foot"><button class="co-save" onclick="saveCompany()">পরিবর্তন সংরক্ষণ করুন</button></div></div></div>`;
}
function companySwitchEdit(key){window.companyEditKey=key==='bodyfitters'?'bodyfitters':'alien_tees';panel('company')}
function coLogoPreview(inp){const f=inp.files[0];if(!f)return;const p=$('coLogoPrev');if(p)p.innerHTML='<img src="'+URL.createObjectURL(f)+'">';coMemoLivePreview()}
function coMemoLivePreview(){const title=$('coMemoTitle'),sub=$('coMemoSubtitle'),foot=$('coMemoFooter'),color=$('coMemoColor'),show=$('coMemoShowLogo'),head=$('coMemoPreviewHead'),t=$('coMemoPreviewTitle'),st=$('coMemoPreviewSubtitle'),ft=$('coMemoPreviewFooter'),logo=$('coMemoPreviewLogo'),holder=$('coLogoPrev');if(t&&title)t.textContent=title.value||'ক্রয় মেমো';if(st&&sub)st.textContent=sub.value;if(ft&&foot)ft.textContent=foot.value;if(head&&color)head.style.background=color.value;if(logo&&show){const img=holder?.querySelector('img');logo.style.display=show.checked&&img?'block':'none';if(img)logo.innerHTML='<img src="'+img.src+'" style="width:100%;height:100%;object-fit:contain">'}}
async function saveCompany(){const fd=new FormData();fd.append('company',window.companyEditKey==='bodyfitters'?'bodyfitters':'alien_tees');fd.append('name',$('coName').value);fd.append('email',$('coEmail').value);fd.append('phone',$('coPhone').value);fd.append('website',$('coWebsite').value);fd.append('tin',$('coTin').value);fd.append('reg_no',$('coReg').value);fd.append('delivery_inside_dhaka',$('coDeliveryInside').value);fd.append('delivery_dhaka_suburbs',$('coDeliverySuburbs').value);fd.append('delivery_outside_dhaka',$('coDeliveryOutside').value);fd.append('address',$('coAddress').value);fd.append('company_note',$('coNote').value);fd.append('invoice_note',$('coInv').value);fd.append('qr_mode',$('coQrMode')?.value==='pdf'?'pdf':'text');fd.append('purchase_memo_title',$('coMemoTitle')?.value||'');fd.append('purchase_memo_subtitle',$('coMemoSubtitle')?.value||'');fd.append('purchase_memo_footer',$('coMemoFooter')?.value||'');fd.append('purchase_memo_color',$('coMemoColor')?.value||'#0D7552');fd.append('purchase_memo_show_logo',$('coMemoShowLogo')?.checked?'1':'0');const lf=$('coLogo')?.files?.[0];if(lf)fd.append('logo',lf);if($('coRmLogo')&&$('coRmLogo').checked)fd.append('remove_logo','1');const r=await fetch('/api/admin/company',{method:'POST',body:fd});const j=await r.json().catch(()=>({}));if(!r.ok)return toast(j.error||'সেভ হয়নি',true);toast('প্রতিষ্ঠান ও মেমো টেমপ্লেট সংরক্ষণ হয়েছে');panel('company')}

/* ---------- POS অর্ডার তালিকা + dashboard ---------- */
function ymd(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}
function posoRange(p){const n=new Date(),d=x=>new Date(n.getFullYear(),n.getMonth(),n.getDate()+x);if(p==='today')return[ymd(d(0)),ymd(d(0))];if(p==='yesterday')return[ymd(d(-1)),ymd(d(-1))];if(p==='7d')return[ymd(d(-6)),ymd(d(0))];if(p==='30d')return[ymd(d(-29)),ymd(d(0))];if(p==='month')return[ymd(new Date(n.getFullYear(),n.getMonth(),1)),ymd(d(0))];if(p==='lastmonth')return[ymd(new Date(n.getFullYear(),n.getMonth()-1,1)),ymd(new Date(n.getFullYear(),n.getMonth(),0))];return['','']}
function posoLoad(){if(!window.posoSt){try{window.posoSt=JSON.parse(sessionStorage.getItem('posoState')||'{}')}catch(e){window.posoSt={}}}const s=window.posoSt;s.source=s.source||'pos';s.preset=s.preset||'today';s.q=s.q||'';s.user=s.user||'';s.method=s.method||'';s.status=s.status||'';if(!window.posoSel)window.posoSel=new Set();return s}
function posoSave(){try{sessionStorage.setItem('posoState',JSON.stringify(window.posoSt))}catch(e){}}
function posoQS(extra){const s=window.posoSt;const q=new URLSearchParams({source:s.source,from:s.from||'',to:s.to||'',q:s.q||'',user:s.user||'',method:s.method||'',status:s.status||''});return q.toString()+(extra||'')}
async function posOrdersPanel(){
 const s=posoLoad();if(!(isAdminLevelJS()||allowed('pos_orders')))s.source='pos';if(s.preset!=='custom'){const r=posoRange(s.preset);s.from=r[0];s.to=r[1]}posoSave();
 const j=await jf('/api/admin/pos-orders?'+posoQS());
 if(j._s===401){location.reload();return ''}
 if(!j._ok)return `<div class="poso"><section class="poso-card"><h3>অর্ডার তালিকা লোড হয়নি</h3><p>${esc(j.error||j._s||'network')}</p><button class="pill" onclick="panel('posorders')">আবার চেষ্টা করুন</button></section></div>`;
 window.posoRows=j.rows||[];const sm=j.summary||{},online=s.source==='online',rows=window.posoRows;
 const presets=[['today','আজ'],['yesterday','গতকাল'],['7d','৭ দিন'],['30d','৩০ দিন'],['month','এই মাস'],['lastmonth','গত মাস'],['all','সব সময়'],['custom','Custom']];
 const maxDay=Math.max(1,...(sm.by_day||[]).map(x=>x.total));
 const hbars=(arr,key)=>{const m=Math.max(1,...arr.map(x=>x[key||'total']));return arr.slice(0,6).map(x=>`<div class="poso-hbar"><span>${esc(x.name||x.code)}</span><span class="t"><i style="width:${Math.round((x[key||'total'])/m*100)}%"></i></span><b>${key==='pieces'?fmtN(x.pieces)+' পিস':taka(x.total)}</b></div>`).join('')||'<small>ডাটা নেই</small>'};
 const kpi=[['অর্ডার',fmtN(sm.orders)],['মোট পিস',fmtN(sm.pieces)],['মোট বিক্রি',taka(sm.total)],[online?'Advance':'পরিশোধ',taka(sm.paid)],['গড় অর্ডার',taka(sm.avg)]];
 const canDelete = !online && allowed('pos_delete');
 const canViewOnline = isAdminLevelJS() || allowed('pos_orders');
 return `<div class="poso"><div class="poso-head"><div><span class="eyebrow">ALIEN TEES ERP · SALES</span><h2>POS অর্ডার ও মেমো</h2><p style="margin:2px 0;color:#667085">Dashboard, filter, Excel রিপোর্ট ও মেমো PDF — যেকোনো সময়</p></div>
 <div class="poso-tabs"><button class="${online?'':'on'}" onclick="posoSet('source','pos')">🛒 POS বিক্রি</button>${canViewOnline?`<button class="${online?'on':''}" onclick="posoSet('source','online')">🌐 Online / Courier অর্ডার</button>`:''}</div></div>
 <section class="poso-card"><div class="poso-chips">${presets.map(p=>`<button class="${s.preset===p[0]?'on':''}" onclick="posoPreset('${p[0]}')">${p[1]}</button>`).join('')}</div>
 <div class="poso-filters"><label>শুরু<input type="date" id="posoFrom" value="${esc(s.from||'')}" ${s.preset==='custom'?'':'disabled'}></label><label>শেষ<input type="date" id="posoTo" value="${esc(s.to||'')}" ${s.preset==='custom'?'':'disabled'}></label>
 <label>${online?'Assigned To':'ইউজার'}<select id="posoUser"><option value="">সব</option>${(j.users||[]).map(u=>`<option ${s.user===u?'selected':''}>${esc(u)}</option>`).join('')}</select></label>
 <label>${online?'Payment':'পেমেন্ট মেথড'}<select id="posoMethod"><option value="">সব</option>${(j.methods||[]).map(u=>`<option ${s.method===u?'selected':''}>${esc(u)}</option>`).join('')}</select></label>
 <label>${online?'Delivery Status':'অবস্থা'}<select id="posoStatus"><option value="">সব</option>${(j.statuses||[]).map(u=>`<option ${s.status===u?'selected':''}>${esc(u)}</option>`).join('')}</select></label>
 <label>খুঁজুন<input id="posoQ" value="${esc(s.q||'')}" placeholder="নাম / ফোন / কোড / অর্ডার নং" onkeydown="if(event.key==='Enter')posoApply()"></label>
 <button class="pill" onclick="posoApply()">ফিল্টার করুন</button><button class="pill" onclick="posoReset()">রিসেট</button></div></section>
 <div class="poso-kpi">${kpi.map(k=>`<article><small>${k[0]}</small><b>${k[1]}</b></article>`).join('')}</div>
 <div class="poso-charts"><section class="poso-card"><h3>দিন অনুযায়ী বিক্রি</h3><div class="poso-bars">${(sm.by_day||[]).map(x=>`<i title="${esc(x.day)} · ${taka(x.total)} · ${x.orders}টি" style="height:${Math.max(4,Math.round(x.total/maxDay*100))}%"></i>`).join('')||'<small>ডাটা নেই</small>'}</div></section>
 <section class="poso-card"><h3>${online?'Delivery status':'পেমেন্ট মেথড'}</h3>${hbars(online?(sm.by_status||[]):(sm.by_method||[]))}</section><section class="poso-card"><h3>সেরা প্রোডাক্ট</h3>${hbars((sm.top_products||[]).map(x=>({name:x.code,pieces:x.pieces,total:x.total})),'pieces')}</section></div>
 <section class="poso-card"><div class="poso-act"><b>${fmtN(j.total)}টি অর্ডার</b>${j.total>rows.length?`<small>(প্রথম ${fmtN(rows.length)}টি দেখানো হচ্ছে — Excel-এ সব পাবেন)</small>`:''}<span class="sel" id="posoSelCount">${window.posoSel.size?window.posoSel.size+'টি সিলেক্ট':''}</span>
 <button class="pill" onclick="posoSelectAll(true)">☑ সব সিলেক্ট</button><button class="pill" onclick="posoSelectAll(false)">✕ সিলেক্ট মুছুন</button>
 <button class="pill" onclick="posoMemo('sel')">🧾 সিলেক্ট করাগুলোর মেমো PDF</button><button class="pill" onclick="posoMemo('all')">🧾 ফিল্টার করা সবার মেমো</button>
 <button class="pill" onclick="posoExcel(false)">⬇ Excel রিপোর্ট (ফিল্টার অনুযায়ী)</button><button class="pill" onclick="posoExcel(true)">⬇ সিলেক্ট করা Excel</button>
 ${canDelete?'<button class="pill danger" data-mode="selected" style="border:1px solid #d92d20;color:#b42318" onclick="posoDelete(this.dataset.mode)">🗑 সিলেক্ট করা POS বিক্রি মুছুন</button><button class="pill danger" data-mode="all_filtered" style="border:1px solid #d92d20;color:#b42318" onclick="posoDelete(this.dataset.mode)">🗑 ফিল্টার অনুযায়ী সব মুছুন</button>':''}</div>
 <div class="excel-table-wrap" style="max-height:620px;overflow:auto"><table class="poso-table"><thead><tr><th></th><th>অর্ডার নং</th><th>তারিখ</th><th>ছবি</th><th>প্রোডাক্ট</th><th>কাস্টমার</th><th>ফোন</th><th>পিস</th><th>মোট</th><th>বাকি</th><th>${online?'Payment':'মেথড'}</th><th>${online?'Delivery':'অবস্থা'}</th><th>${online?'Assigned':'ইউজার'}</th><th>মেমো</th></tr></thead><tbody>
 ${rows.map(r=>`<tr><td><input type="checkbox" ${window.posoSel.has(r.id)?'checked':''} onchange="posoToggle(${r.id},this.checked)"></td><td><b>${esc(r.no)}</b></td><td>${esc(r.date)}</td><td>${r.img?`<img loading="lazy" src="/api/img?url=${encodeURIComponent(r.img)}" onerror="this.outerHTML='<div class=ph>No photo</div>'">`:'<div class="ph">No photo</div>'}</td><td>${esc(r.items_text)}</td><td>${esc(r.customer||'—')}</td><td>${esc(r.phone||'—')}</td><td>${fmtN(r.pieces)}</td><td><b>${taka(r.total)}</b></td><td>${r.due>0?`<span class="poso-pill due">${taka(r.due)}</span>`:'—'}</td><td>${esc(r.method||'—')}</td><td><span class="poso-pill ${r.status==='Due'?'due':''}">${esc(r.status||'—')}</span></td><td>${esc(r.user||'—')}</td><td><button class="pill" onclick="openMemos('${s.source}',[${r.id}])">🖨 মেমো PDF / প্রিন্ট</button></td></tr>`).join('')||'<tr><td colspan="14" style="padding:22px;text-align:center">এই ফিল্টারে কোনো অর্ডার নেই</td></tr>'}
 </tbody></table></div></section></div>`;
}
function posoSet(k,v){const s=posoLoad();s[k]=v;s.user='';s.method='';s.status='';window.posoSel=new Set();posoSave();panel('posorders')}
function posoPreset(p){const s=posoLoad();s.preset=p;if(p==='custom'){s.from=$('posoFrom')?.value||s.from||'';s.to=$('posoTo')?.value||s.to||''}window.posoSel=new Set();posoSave();panel('posorders')}
function posoApply(){const s=posoLoad();if(s.preset==='custom'){s.from=$('posoFrom').value;s.to=$('posoTo').value}s.user=$('posoUser').value;s.method=$('posoMethod').value;s.status=$('posoStatus').value;s.q=$('posoQ').value.trim();window.posoSel=new Set();posoSave();panel('posorders')}
function posoReset(){const s=posoLoad();s.preset='today';s.q='';s.user='';s.method='';s.status='';window.posoSel=new Set();posoSave();panel('posorders')}
function posoUpdSel(){const e=$('posoSelCount');if(e)e.textContent=window.posoSel.size?window.posoSel.size+'টি সিলেক্ট':''}
function posoToggle(id,on){if(on)window.posoSel.add(id);else window.posoSel.delete(id);posoUpdSel()}
function posoSelectAll(on){(window.posoRows||[]).forEach(r=>on?window.posoSel.add(r.id):window.posoSel.delete(r.id));document.querySelectorAll('.poso-table tbody input[type=checkbox]').forEach(c=>c.checked=on);posoUpdSel()}
function posoExcel(selOnly){const s=posoLoad();if(selOnly&&!window.posoSel.size)return toast('আগে অর্ডার সিলেক্ট করুন',true);location.href='/api/admin/pos-orders/export.xlsx?'+posoQS(selOnly?'&ids='+[...window.posoSel].join(','):'')}
function posoMemo(mode){const s=posoLoad();let ids=mode==='sel'?[...window.posoSel]:(window.posoRows||[]).map(r=>r.id);if(!ids.length)return toast(mode==='sel'?'আগে অর্ডার সিলেক্ট করুন':'কোনো অর্ডার নেই',true);if(ids.length>100&&!confirm(ids.length+'টি মেমো তৈরি হবে। চালাবেন?'))return;openMemos(s.source,ids)}
async function posoDelete(mode){
 if(!allowed('pos_delete'))return toast('POS বিক্রি মুছার অনুমতি নেই',true);
 const s=posoLoad();if(s.source!=='pos')return toast('এই delete action শুধু POS বিক্রির জন্য',true);
 const ids=mode==='selected'?[...(window.posoSel||[])]:[];
 if(mode==='selected'&&!ids.length)return toast('আগে POS বিক্রি সিলেক্ট করুন',true);
 if(mode==='all_filtered') {const typed=prompt('সতর্কতা: বর্তমান filter-এর সব POS sale মুছে যাবে এবং stock ফেরত যোগ হবে। নিশ্চিত করতে DELETE লিখুন।');if(typed!=='DELETE')return;}
 const reason=prompt('Delete reason লিখুন (যেমন: ভুল/duplicate entry):');if(!reason||reason.trim().length<3)return toast('Delete reason দিতে হবে',true);
 const filters={source:'pos',from:s.from||'',to:s.to||'',q:s.q||'',user:s.user||'',method:s.method||'',status:s.status||''};
 if(mode==='selected'&&!confirm(`${ids.length}টি POS sale delete করে stock restore করবেন?`))return;
 const btns=[...document.querySelectorAll('.poso-act button')];btns.forEach(b=>b.disabled=true);
 try{const j=await post('/api/admin/pos-orders/delete',{mode,ids,filters,reason:reason.trim()});if(!j._ok)throw Error(j.error||'Delete failed');window.posoSel=new Set();toast(j.message||'POS বিক্রি মুছে ফেলা হয়েছে');posoSave();await panel('posorders');}
 catch(e){toast(e.message||'POS বিক্রি মুছা যায়নি',true)}finally{btns.forEach(b=>b.disabled=false)}
}

/* ---------- মেমো (A4 / PDF) ---------- */
function memoHtml(m,co,hasLogo,logoV){
 const mt=m.memo_template||{};const showMemoLogo=mt.show_logo!==false;const memoColor=/^#[0-9a-fA-F]{6}$/.test(mt.color||'')?mt.color:'#0d7552';
 const memoTitle=mt.title||'SALES MEMO';const memoFooter=mt.footer||co.invoice_note||'Computer generated invoice';
 const money=n=>Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
 const img=u=>u?`<img class="memo-product-img" src="${location.origin}/api/img?url=${encodeURIComponent(u)}" onerror="this.style.display='none'">`:'<span class="memo-product-img memo-product-placeholder">AT</span>';
 const rows=(m.items||[]).map((it,i)=>`<tr><td class="memo-index">${String(i+1).padStart(2,'0')}</td><td class="memo-product"><span class="memo-product-media">${img(it.img)}</span><span class="memo-product-name"><b>${esc(it.title||it.code||'Product')}</b><small>${esc(it.code||'')}</small></span></td><td class="memo-center">${esc(it.variation||'—')}</td><td class="memo-center">${Number(it.qty||0)}</td><td class="memo-money">৳ ${money(it.price)}</td><td class="memo-money memo-line-total">৳ ${money(Number(it.price||0)*Number(it.qty||0))}</td></tr>`).join('');
 const notes=`${m.customer_note?`<div class="memo-note-card nc"><b>Customer note</b><p>${esc(m.customer_note)}</p></div>`:''}${m.office_note?`<div class="memo-note-card no"><b>Office note</b><p>${esc(m.office_note)}</p></div>`:''}${co.company_note?`<div class="memo-note-card nm"><b>Company note</b><p>${esc(co.company_note)}</p></div>`:''}`;
 const q=Array.isArray(m.qr_data)?m.qr_data[0]:m.qr_data;const qrMode=String(co.qr_mode||'text').toLowerCase()==='pdf'?'pdf':'text';
 const contact=[co.phone,co.email].filter(Boolean).map(esc).join(' · ');
 return `<article class="memo" style="--memo-green:${memoColor};--memo-dark:${memoColor}">
  <div class="memo-topline"><span class="memo-kicker"><i></i> OFFICIAL ${m.source==='online'?'ORDER':'SALES'} MEMO</span><span class="memo-copy">CUSTOMER COPY · KEEP FOR YOUR RECORDS</span></div>
  <header class="memo-hero"><div class="memo-brand"><div class="co">${esc(co.name||'Alien Tees')}</div>${contact?`<div class="memo-contact">${contact}</div>`:''}${co.address?`<div class="memo-address">${esc(co.address)}</div>`:''}${(co.tin||co.reg_no)?`<div class="memo-registration">${co.tin?'TIN: '+esc(co.tin)+' ':''}${co.reg_no?'REG: '+esc(co.reg_no):''}</div>`:''}</div>
   ${hasLogo&&showMemoLogo?`<div class="memo-logo-wrap"><img class="memo-logo" src="${location.origin}/api/company/logo?company=${encodeURIComponent(m.company_key||'alien_tees')}&v=${logoV}" alt="${esc(co.name||'Company')} logo"></div>`:`<div class="memo-brand-mark">${esc((co.name||'AT').split(/\s+/).map(x=>x[0]||'').join('').slice(0,3).toUpperCase())}</div>`}
  </header>
  <section class="memo-title-row"><div><div class="memo-eyebrow">TRANSACTION SUMMARY</div><h1>${esc(memoTitle)}</h1><p>${esc(mt.subtitle||co.invoice_note||'Itemized invoice and payment summary')}</p></div><div class="memo-invoice-chip"><span>INVOICE / ORDER NO.</span><b>${esc(m.order_no||'—')}</b><small>${esc(m.date||'—')} · ${esc(m.kind||'—')}</small></div></section>
  <section class="memo-customer-grid"><div class="memo-customer-card"><div class="memo-section-label">BILL TO</div><b class="memo-customer-name">${esc(m.name||'Walk-in Customer')}</b><div class="memo-customer-detail">${esc(m.phone||'Phone not provided')}</div><div class="memo-customer-detail">${esc(m.address||'Address not provided')}</div>${m.seller?`<div class="memo-customer-detail memo-seller"><b>Created by:</b> ${esc(m.seller)}</div>`:''}</div><div class="memo-payment-card"><div class="memo-section-label">PAYMENT METHOD</div><b>${esc(m.kind||'—')}</b><span class="memo-status ${Number(m.due||0)>0?'memo-status-due':'memo-status-paid'}">${Number(m.due||0)>0?'PAYMENT DUE':'PAID IN FULL'}</span></div></section>
  <section class="memo-items-section"><div class="memo-section-heading"><div><span class="memo-eyebrow">YOUR PURCHASE</span><h2>Item details</h2></div><span class="memo-items-count">${(m.items||[]).length} line(s) · ${(m.items||[]).reduce((a,it)=>a+Number(it.qty||0),0)} pieces</span></div>
   <table class="items"><thead><tr><th class="memo-index">#</th><th style="text-align:left">PRODUCT</th><th>SIZE</th><th>QTY</th><th class="memo-money">UNIT PRICE</th><th class="memo-money">LINE TOTAL</th></tr></thead><tbody>${rows||'<tr><td colspan="6" class="memo-empty">No item details recorded</td></tr>'}</tbody></table>
  </section>
  <section class="memo-lower-grid"><div class="memo-notes-area">${notes||'<div class="memo-thanks"><span class="memo-thanks-icon">✓</span><div><b>Thank you for your business</b><small>We appreciate your trust and support.</small></div></div>'}<div class="memo-signatures"><div><span></span><small>PREPARED BY</small></div><div><span></span><small>AUTHORIZED BY</small></div><div><span></span><small>RECEIVED BY</small></div></div></div>
   <div class="tot"><div class="memo-tot-heading">PAYMENT SUMMARY</div><div class="pr"><span>Product subtotal</span><span>৳ ${money(m.subtotal)}</span></div>${m.discount?`<div class="pr"><span>Discount</span><span>− ৳ ${money(m.discount)}</span></div>`:''}${m.delivery?`<div class="pr"><span>Delivery${m.delivery_zone_label?' · '+esc(m.delivery_zone_label):''}</span><span>৳ ${money(m.delivery)}</span></div>`:''}<div class="pr memo-total-row"><b>Total payable</b><b>৳ ${money(m.total)}</b></div>${Number(m.advance||0)>0?`<div class="pr"><span>Paid / advance</span><span>৳ ${money(m.advance)}</span></div>`:''}<div class="due"><span>Balance due</span><b>৳ ${money(m.due)}</b></div></div>
   <aside class="qr-area">${q?`<div class="memo-qr-card"><div class="memo-qr-heading">QUICK MEMO</div><img class="memo-qr" src="${q}" alt="Memo QR"><span class="memo-qr-invoice">${esc(m.order_no||'')}</span></div>`:'<div class="memo-qr-card memo-qr-missing">QR unavailable</div>'}</aside>
  </section>
  <footer class="memo-bottom"><span>${esc(memoFooter)}</span><span>${esc(co.name||'Alien Tees')} · ${esc(m.order_no||'')}</span></footer>
 </article>`;
}
async function openMemos(source,ids,nos){
 const w=window.open('','_blank');if(!w)return toast('Pop-up ব্লক করা আছে — অনুমতি দিন',true);
 w.document.write('<meta charset="utf-8"><title>মেমো তৈরি হচ্ছে…</title><body style="font:16px sans-serif;padding:30px">মেমো তৈরি হচ্ছে…</body>');
 const j=await post('/api/admin/memos',{source,ids:ids||[],nos:nos||[]});
 if(!j._ok||!(j.memos||[]).length){w.document.body.textContent=j.error||'মেমো পাওয়া যায়নি';return}
 const per=1;
 const css=`:root{--memo-green:#0d7552;--memo-dark:#16392d;--memo-gold:#d9b96e;--memo-ink:#24372d;--memo-muted:#6b7d72;--memo-line:#e0e9e2;--memo-pale:#f2f8f3}*{box-sizing:border-box}html{background:#edf2ee}body{font:12px Arial,'Noto Sans Bengali','Nirmala UI',sans-serif;margin:0;color:var(--memo-ink);background:#edf2ee;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.bar{position:sticky;top:0;z-index:5;background:#14352a;color:#fff;padding:10px 16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;font-size:12px;box-shadow:0 3px 14px #13291e22}.bar label{display:flex;gap:5px;align-items:center;cursor:pointer}.bar button{background:#e2f0e7;color:#14352a;border:0;border-radius:8px;padding:9px 15px;font-weight:800;cursor:pointer}.bar select{padding:5px;border-radius:6px;border:0}
.memo{position:relative;margin:24px auto 30px;width:calc(100% - 32px);max-width:1000px;padding:28px 34px 20px;background:#fff;border:1px solid #dbe7dd;border-top:7px solid var(--memo-green);border-radius:16px;box-shadow:0 16px 45px #173c2916;break-inside:avoid;overflow:hidden}.memo:before{content:'';position:absolute;right:-70px;top:-85px;width:210px;height:210px;border-radius:50%;border:1px solid #e1eee5;box-shadow:0 0 0 18px #f8fbf8,0 0 0 19px #e4efe6;pointer-events:none}.memo-topline{position:relative;z-index:1;display:flex;justify-content:space-between;align-items:center;gap:10px;border-bottom:1px solid var(--memo-line);padding-bottom:12px;color:#789084;font-size:8px;font-weight:800;letter-spacing:.13em}.memo-kicker{display:flex;align-items:center;gap:7px;color:var(--memo-green)}.memo-kicker i{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--memo-gold);box-shadow:0 0 0 4px #f7efd9}.memo-copy{font-size:7px;letter-spacing:.09em;text-align:right}.memo-hero{position:relative;z-index:1;display:flex;align-items:center;justify-content:space-between;gap:18px;padding:18px 0 17px;border-bottom:1px solid var(--memo-line)}.memo-brand{min-width:0}.co{font-size:27px;font-weight:850;letter-spacing:-.045em;line-height:1.12;color:var(--memo-dark);margin-bottom:6px}.memo-contact{font-size:10px;color:var(--memo-green);font-weight:700;margin-bottom:4px}.memo-address,.memo-registration{font-size:9px;line-height:1.5;color:var(--memo-muted)}.memo-logo-wrap{width:95px;height:70px;display:grid;place-items:center;flex:0 0 auto}.memo-logo{max-width:95px;max-height:70px;object-fit:contain}.memo-brand-mark{flex:0 0 70px;width:70px;height:70px;display:grid;place-items:center;border:1px solid #deebe0;border-radius:14px;background:linear-gradient(135deg,#eff7f1,#fff);color:var(--memo-green);font-weight:900;font-size:20px;letter-spacing:-.05em}.memo-title-row{position:relative;z-index:1;display:flex;align-items:center;justify-content:space-between;gap:18px;padding:20px 0 16px}.memo-eyebrow{display:block;color:var(--memo-green);font-size:8px;letter-spacing:.16em;font-weight:900;text-transform:uppercase;margin-bottom:5px}.memo-title-row h1{font-size:23px;line-height:1.15;letter-spacing:-.025em;margin:0;color:#1c3126}.memo-title-row p{font-size:10px;margin:6px 0 0;color:var(--memo-muted)}.memo-invoice-chip{min-width:175px;max-width:220px;padding:12px 15px;border-radius:12px;background:var(--memo-pale);border:1px solid #dcebe0;border-right:4px solid var(--memo-green)}.memo-invoice-chip span,.memo-invoice-chip small{display:block;color:#718477;font-size:8px;font-weight:800;letter-spacing:.09em}.memo-invoice-chip b{display:block;font-size:15px;overflow-wrap:anywhere;color:var(--memo-dark);margin:5px 0}.memo-invoice-chip small{font-size:9px;letter-spacing:0;font-weight:500}.memo-customer-grid{display:grid;grid-template-columns:minmax(0,1fr) 190px;gap:12px;margin-bottom:19px}.memo-customer-card,.memo-payment-card{padding:13px 15px;background:#fbfdfb;border:1px solid var(--memo-line);border-radius:11px;min-width:0}.memo-section-label{font-size:8px;letter-spacing:.12em;color:#75887c;font-weight:900;margin-bottom:7px}.memo-customer-name{display:block;font-size:13px;color:#1c382a;margin-bottom:4px;overflow-wrap:anywhere}.memo-customer-detail{font-size:10px;color:#53675a;line-height:1.5;overflow-wrap:anywhere}.memo-payment-card> b{display:block;font-size:13px;margin:2px 0 9px;color:var(--memo-dark)}.memo-status{display:inline-block;padding:5px 8px;border-radius:999px;font-size:7px;font-weight:900;letter-spacing:.08em}.memo-status-paid{background:#e4f5e9;color:#167348}.memo-status-due{background:#fff2dc;color:#9a5a08}.memo-section-heading{display:flex;align-items:end;justify-content:space-between;gap:12px;margin:0 0 9px}.memo-section-heading h2{font-size:15px;line-height:1.2;margin:0;color:#23392c;text-transform:capitalize}.memo-items-count{font-size:9px;color:#738579}.items{width:100%;border-collapse:separate;border-spacing:0;border:1px solid #dfe9e1;border-radius:10px;overflow:hidden;margin:0 0 4px}.items th{padding:10px 8px;font-size:8px;letter-spacing:.07em;text-transform:uppercase;text-align:center;background:var(--memo-green);color:#fff;border:0}.items td{padding:8px 7px;vertical-align:middle;border-bottom:1px solid #e6eee8;font-size:10px}.items tbody tr:last-child td{border-bottom:0}.items tbody tr:nth-child(even){background:#f8fbf8}.memo-index{width:28px;text-align:center;color:#7c8e82;font-size:9px!important}.memo-product{display:flex;align-items:center;gap:9px;min-width:180px}.memo-product-media{display:grid;place-items:center;flex:0 0 42px;width:42px;height:47px;background:#edf3ee;border:1px solid #e2eae3;border-radius:6px;overflow:hidden}.memo-product-img{width:42px;height:47px;object-fit:cover;display:block}.memo-product-placeholder{display:grid;place-items:center;font-size:9px;font-weight:900;color:#9aafa0}.memo-product-name{min-width:0}.memo-product-name b{display:block;font-size:10px;line-height:1.3;color:#2a3f32;overflow-wrap:anywhere}.memo-product-name small{display:block;font-size:8px;color:#849186;margin-top:3px}.memo-center{text-align:center;white-space:nowrap}.memo-money{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}.memo-line-total{font-weight:800;color:#213c2c}.memo-empty{text-align:center!important;padding:25px!important;color:#8a978e!important}.memo-lower-grid{display:grid;grid-template-areas:"notes qr summary";grid-template-columns:minmax(0,1fr) 145px minmax(215px,250px);align-items:stretch;gap:12px;margin-top:17px}.memo-notes-area{grid-area:notes;display:flex;flex-direction:column;gap:8px;min-width:0}.memo-note-card{padding:10px 11px;border-left:3px solid #aacbb5;border-radius:0 7px 7px 0;background:#f7faf7;font-size:9px}.memo-note-card b{font-size:8px;color:var(--memo-green);text-transform:uppercase;letter-spacing:.06em}.memo-note-card p{margin:4px 0 0;line-height:1.5;overflow-wrap:anywhere}.memo-thanks{display:flex;align-items:center;gap:10px;padding:12px;border:1px solid #e0ebe2;border-radius:10px;background:#f8fbf8;color:#315f43}.memo-thanks-icon{width:27px;height:27px;display:grid;place-items:center;border-radius:50%;background:#dff0e3;font-weight:900}.memo-thanks b,.memo-thanks small{display:block}.memo-thanks b{font-size:10px}.memo-thanks small{font-size:8px;margin-top:3px;color:#7b8b7f}.memo-signatures{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:auto;padding-top:18px}.memo-signatures div{text-align:center}.memo-signatures span{display:block;border-top:1px solid #aebdb2;margin-bottom:5px}.memo-signatures small{font-size:7px;color:#829084;letter-spacing:.06em}.tot{grid-area:summary;min-width:0;padding:13px 13px 11px;border:1px solid #dce8df;border-radius:12px;background:#f8fbf8;align-self:start}.memo-tot-heading{font-size:8px;font-weight:900;letter-spacing:.13em;color:var(--memo-green);padding-bottom:9px;margin-bottom:4px;border-bottom:1px solid #dce8df}.tot .pr{display:flex;justify-content:space-between;gap:12px;padding:6px 0;font-size:9px}.tot .pr span:last-child,.tot .pr b:last-child{white-space:nowrap;font-variant-numeric:tabular-nums}.memo-total-row{border-top:1px solid #dce8df;margin-top:3px;padding-top:9px!important;font-size:10px!important;color:var(--memo-dark)}.due{display:flex;justify-content:space-between;gap:8px;border-top:1px dashed #c6d8cb;margin-top:5px;padding-top:9px!important;font-weight:850;color:#995c0a;font-size:10px}.due b{white-space:nowrap}.qr-area{grid-area:qr;min-width:0;width:100%;max-width:145px;justify-self:end;text-align:center;align-self:start}.memo-qr-card{padding:11px 7px 9px;border:1px solid #d5e5d9;border-radius:12px;background:linear-gradient(155deg,#f5faf6 0%,#fff 75%);display:flex;flex-direction:column;align-items:center;gap:6px;box-shadow:0 3px 8px #173c2908}.memo-qr-heading{align-self:stretch;color:var(--memo-green);font-size:8px;font-weight:900;letter-spacing:.12em;border-bottom:1px solid #e1ebe3;padding-bottom:7px}.memo-qr{display:block;width:108px;height:108px;object-fit:contain;image-rendering:pixelated;background:#fff;padding:3px;border:1px solid #e0e9e1;border-radius:6px}.memo-qr-card>b{font-size:7px;letter-spacing:.07em;color:#254d36}.memo-qr-card>small{font-size:7px;line-height:1.35;color:#718277}.memo-qr-invoice{display:block;width:100%;overflow-wrap:anywhere;background:#edf5ef;color:#486854;padding:4px;border-radius:5px;font-size:8px;font-weight:800}.memo-qr-missing{min-height:145px;display:grid;place-items:center;color:#97a59a;font-size:9px}.memo-bottom{display:flex;justify-content:space-between;gap:12px;margin-top:17px;padding-top:10px;border-top:1px solid #dfe8e1;color:#809087;font-size:8px}.memo-bottom span:first-child{max-width:65%;overflow-wrap:anywhere}.memo-bottom span:last-child{text-align:right;font-weight:700}
.no-photo .memo-product-media{display:none}.no-price .memo-money{display:none!important}.no-office .memo-note-card.no{display:none}.no-cust .memo-note-card.nc{display:none}.no-comp .memo-note-card.nm{display:none}.no-logo .memo-logo-wrap,.no-logo .memo-brand-mark{display:none}
@media(max-width:760px){.memo{margin:8px;width:calc(100% - 16px);padding:17px 15px 14px;border-radius:11px}.memo-topline{font-size:7px}.memo-copy{display:none}.memo-hero{padding:13px 0}.co{font-size:21px}.memo-logo-wrap{width:70px;height:58px}.memo-logo{max-width:70px;max-height:58px}.memo-brand-mark{width:54px;height:54px;flex-basis:54px;font-size:16px}.memo-title-row{align-items:flex-start;gap:9px;padding:15px 0 13px}.memo-title-row h1{font-size:19px}.memo-title-row p{font-size:9px}.memo-invoice-chip{min-width:125px;padding:9px 10px}.memo-invoice-chip b{font-size:12px}.memo-customer-grid{grid-template-columns:minmax(0,1fr) 135px;gap:8px}.memo-customer-card,.memo-payment-card{padding:10px}.memo-lower-grid{grid-template-areas:"notes notes" "qr summary";grid-template-columns:135px minmax(0,1fr);gap:9px}.memo-notes-area{grid-area:notes;grid-column:1/-1}.qr-area{grid-area:qr;grid-column:1;grid-row:2;max-width:135px;justify-self:start}.tot{grid-area:summary;grid-column:2;grid-row:2}.memo-qr{width:100px;height:100px}.memo-signatures{padding-top:9px}.memo-product{min-width:140px}.items th{font-size:7px;padding:7px 4px}.items td{padding:6px 4px;font-size:9px}.memo-product-media,.memo-product-img{width:31px;height:36px;flex-basis:31px}.memo-product-name b{font-size:9px}.memo-money{font-size:8px}.memo-bottom{font-size:7px}}
@media print{html,body{background:#fff!important}.bar{display:none!important}.memo{width:auto;max-width:none;min-height:0;margin:0;padding:8mm 9mm 6mm;border-radius:0;border:0;border-top:2mm solid var(--memo-green);box-shadow:none;overflow:visible;break-inside:avoid;page-break-inside:avoid}.memo:before{display:none}.memo-topline{padding-bottom:2.5mm}.memo-hero{padding:4mm 0}.co{font-size:20pt}.memo-title-row{padding:4mm 0 3mm}.memo-title-row h1{font-size:17pt}.memo-customer-grid{margin-bottom:4mm}.memo-customer-card,.memo-payment-card{padding:3mm}.items{margin:0}.items th{padding:2.2mm 1.5mm;font-size:7pt}.items td{padding:2mm 1.5mm;font-size:7.5pt}.memo-product-media,.memo-product-img{width:9mm;height:10mm;flex-basis:9mm}.memo-product-name b{font-size:7.5pt}.memo-product-name small{font-size:6.5pt}.memo-lower-grid{grid-template-areas:"notes qr summary";grid-template-columns:minmax(0,1fr) 34mm 53mm;gap:3mm;margin-top:4mm}.memo-notes-area{grid-area:notes}.qr-area{grid-area:qr;max-width:34mm;justify-self:start}.tot{grid-area:summary}.memo-notes-area{gap:2mm}.memo-note-card{padding:2mm}.tot{padding:3mm}.tot .pr{font-size:7.4pt;padding:1.5mm 0}.due{font-size:8pt;padding-top:2mm!important}.memo-qr-card{padding:2mm 1mm;box-shadow:none}.memo-qr{width:27mm;height:27mm;padding:1mm}.memo-qr-heading{padding-bottom:1.5mm}.memo-qr-card>b,.memo-qr-card>small{font-size:5.6pt}.memo-qr-invoice{font-size:6.2pt}.memo-signatures{padding-top:3mm}.memo-bottom{margin-top:4mm;padding-top:2.5mm;font-size:6.5pt} .memo:after{display:none} @page{size:A4;margin:0}}
`;
 const body=j.memos.map(m=>memoHtml(m,m.company||j.company,(m.has_logo!==undefined?m.has_logo:j.has_logo),(m.logo_v!==undefined?m.logo_v:j.logo_v))).join('');
 w.document.open();
 w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Memo · ${j.memos.length}টি</title><style>${css}</style><style id="pp"></style></head><body><div class="bar"><b>🧾 ${j.memos.length}টি মেমো</b>
 <label><input type="checkbox" id="k_photo" checked> ছবি</label><label><input type="checkbox" id="k_price" checked> দাম</label><label><input type="checkbox" id="k_office" checked> Office Note</label><label><input type="checkbox" id="k_cust" checked> Customer Note</label><label><input type="checkbox" id="k_comp" checked> Company Note</label><label><input type="checkbox" id="k_logo" checked> Logo</label>
 <label>প্রতি পাতায় <select id="k_per"><option>1</option><option>2</option><option>3</option><option>4</option></select> টি</label><button id="k_print">⬇ PDF ডাউনলোড / Print</button><span style="opacity:.8;font-size:12px">Print window-এ Destination → "Save as PDF"</span></div><div id="sheet">${body}</div></body></html>`);
 w.document.close();
 const d=w.document,sheet=d.getElementById('sheet');
 const tg=(id,cls)=>d.getElementById(id).addEventListener('change',e=>sheet.classList.toggle(cls,!e.target.checked));
 tg('k_photo','no-photo');tg('k_price','no-price');tg('k_office','no-office');tg('k_cust','no-cust');tg('k_comp','no-comp');tg('k_logo','no-logo');
 const setPer=n=>{d.getElementById('pp').textContent='@media print{#sheet>.memo:nth-child('+n+'n){break-after:page}}'};
 d.getElementById('k_per').value=String(per);setPer(per);d.getElementById('k_per').addEventListener('change',e=>setPer(Number(e.target.value)));
 d.getElementById('k_print').addEventListener('click',async()=>{const imgs=[...d.images].filter(i=>!i.complete);await Promise.race([Promise.all(imgs.map(i=>new Promise(r=>{i.onload=i.onerror=r}))),new Promise(r=>setTimeout(r,6000))]);w.focus();w.print()});
}

async function addUser(){const role=$('ur').value;const perms={order_view:false,return_view:false,courier_view:false,courier_update:false,pos_use:false,pos_orders:false,price_edit:false,view_stock:false,stock_comment:true,search:true,image_search:true,view_product:true,size_filter:true,download:true,copy:true,link_copy:true,id_copy:true,name_copy:true,full_info_copy:true,share_product:true,website_view:true,favorite:true,report:true};const j=await post('/api/admin/users',{username:$('un').value,password:$('up').value,role,permissions:perms});if(j._ok){toast(role==='admin'?'Admin তৈরি হয়েছে':role==='cs'?'CS তৈরি হয়েছে':'Staff তৈরি হয়েছে');panel('users')}else toast(j.error||'ইউজার তৈরি হয়নি',true)}
async function saveUser(id){const r=$('role_'+id),role=r&&!r.disabled?r.value:(r?.value||'staff'),permissions={};document.querySelectorAll(`[data-u="${id}"]`).forEach(e=>permissions[e.dataset.p]=e.checked);const j=await post('/api/admin/users/'+id,{role,permissions,password:$('pw_'+id)?.value||''});if(j._ok){toast('User/Role/Permission সেভ হয়েছে');panel('users')}else toast(j.error||'সেভ হয়নি',true)}
async function deleteUser(id){if(!confirm('এই user account মুছে ফেলবেন?'))return;const j=await post('/api/admin/users/'+id,{},'DELETE');if(j._ok){toast('ইউজার মুছে ফেলা হয়েছে');panel('users')}else toast(j.error||'মুছতে পারিনি',true)}
async function saveGuestPerms(){
 const permissions={};
 Object.keys(GUEST_PERM_LABELS).forEach(k=>{const el=$('guest_'+k);permissions[k]=!!el?.checked});
 const j=await post('/api/admin/guest-perms',{permissions});
 if(j._ok){guestPerms=j.permissions||permissions;applyPermUI();render(1);toast('Guest/Customer Access সেভ হয়েছে')}
 else toast(j.error||'Guest permission সেভ হয়নি',true)
}
function priceRows(){const q=($('pf')?.value||'').toLowerCase();$('pr').innerHTML='<table>'+ALL.filter(p=>!q||(p.title+p.id).toLowerCase().includes(q)).slice(0,60).map(p=>`<tr><td style="display:flex;align-items:center;gap:10px"><img src="${im(p.img)}" loading="lazy" style="width:42px;height:52px;object-fit:cover;border-radius:7px"><div>${esc(productId(p)||'—')}<br><small>${esc(p.title)} · ${esc(p.store)}</small></div></td><td><input size="7" value="${p.edited?p.price:''}" placeholder="${bp(p.edited?p.base:p.price)}" onchange="setPrice('${esc(p.key)}',this.value)"></td></tr>`).join('')+'</table>'}
async function changeAdminPassword(){const a=$('oldpw').value,b=$('newpw').value,c=$('newpw2').value;if(!b||b.length<6)return toast('নতুন পাসওয়ার্ড কমপক্ষে ৬ অক্ষর দিন',true);if(b!==c)return toast('নতুন পাসওয়ার্ড দুইবার একই নয়',true);const j=await post('/api/admin/password',{old_password:a,new_password:b});if(j._ok){toast('Admin পাসওয়ার্ড পরিবর্তন হয়েছে');$('oldpw').value='';$('newpw').value='';$('newpw2').value=''}else toast(j.error||'পাসওয়ার্ড পরিবর্তন হয়নি',true)}
async function saveStockThreshold(){const v=Math.max(1,Math.min(99,Number($('stockth').value||2)));const j=await post('/api/admin/stock-threshold',{threshold:v});if(j._ok){localStorage.setItem('stockThreshold',String(v));toast('Stock Alert threshold সেভ হয়েছে');render(1)}else toast(j.error||'সেভ হয়নি',true)}
async function setPrice(k,v){const j=await post('/api/admin/price',{key:k,price:v});if(!j._ok)return toast(j.error||'দাম সেভ হয়নি',true);await loadProducts();toast(v===''?'আসল দাম ফিরেছে':'দাম সেভ হয়েছে');if(meRole==='admin'||allowed('price_edit'))panel('price')}
function setPre(n){const p=PRESETS[n];for(const k in p)$('c_'+k).value=p[k];applyTheme(p)}
async function saveTheme(){const t={font:$('c_font').value,title:$('c_title').value,tag:$('c_tag').value,language:$('c_language')?.value||'en'};['acc','bg','ink','head','card'].forEach(k=>t[k]=$('c_'+k).value);
 const j=await post('/api/admin/theme',t);if(j._ok){await loadTheme();toast('ডিজাইন সেভ হয়েছে')}else toast('সেভ হয়নি')}

loadTheme();
loadLocalProductCache();
if(localStorage.getItem('quickMode')==='1'){document.body.classList.add('quick');setTimeout(()=>{if($('quickbtn'))$('quickbtn').textContent=' দ্রুত মোড: চালু'},0)}
checkMe().then(()=>{setUiLanguage('bn');translateUi(document.body);if(me){if(!window.__adminPanelRestored)loadProducts().then(poll);refreshNotificationCount()}});
new MutationObserver(()=>{if(uiLanguage!=='bn')setUiLanguage('bn');translateUi(document.body)}).observe(document.body,{childList:true,subtree:true});
// Remove only the redundant ✕ Off badge from the POS Orders & Memos sidebar item.
function cleanPosMemoOffLabel(){const root=document.getElementById('dr');if(!root)return;root.querySelectorAll('button,a,[role=button],li,div').forEach(item=>{const label=(item.innerText||item.textContent||'').replace(/\s+/g,' ').trim();if(/POS Orders & Memos|POS অর্ডার ও মেমো/i.test(label)&&/✕\s*Off|×\s*Off/i.test(label)){item.querySelectorAll('*').forEach(el=>{if(el.children.length===0&&/^(✕\s*Off|×\s*Off)$/i.test((el.textContent||'').trim()))el.remove()});item.querySelectorAll('*').forEach(el=>{if(el.children.length===0&&/^(✕|×)\s*$/i.test((el.textContent||'').trim()))el.remove()})}})}
cleanPosMemoOffLabel();new MutationObserver(cleanPosMemoOffLabel).observe(document.getElementById('dr')||document.body,{childList:true,subtree:true,characterData:true});
</script>
<script>
/* ===== PREMIUM JS (XEROX-S1AM) ===== */
(function(){
 /* সাইডবারে আইকন CSS দিয়ে বসে, তাই বাটনের শুরুর ইমোজি সরানো হয় (শুধু সাইডবারে) */
 const isIcon=cp=>(cp>=0x2190&&cp<=0x21FF)||(cp>=0x2300&&cp<=0x23FF)||(cp>=0x25A0&&cp<=0x25FF)||(cp>=0x2B00&&cp<=0x2BFF)||(cp>=0x1F000&&cp<=0x1FAFF)||(cp>=0x2600&&cp<=0x27BF)||(cp>=0xFE00&&cp<=0xFE0F)||cp===0x200D||(cp>=0x1F1E6&&cp<=0x1F1FF);
 function strip(s){let i=0;while(i<s.length){if(s[i].trim()===''){i++;continue}const cp=s.codePointAt(i);if(!isIcon(cp))break;i+=cp>0xFFFF?2:1}return s.slice(i)}
 function fix(root){root.querySelectorAll('#dr button:not(.dr-x):not(.personal-toggle)').forEach(b=>{b.childNodes.forEach(n=>{if(n.nodeType===3){const t=strip(n.nodeValue);if(t!==n.nodeValue)n.nodeValue=t}})})}
 const dr0=document.getElementById('dr');if(dr0){fix(document);new MutationObserver(()=>fix(document)).observe(dr0,{childList:true,subtree:true,characterData:true})}
 /* ব্যাকড্রপে ক্লিক করলে অ্যাডমিন/ লগইন বন্ধ */
 const ao=document.getElementById('ao');if(ao)ao.addEventListener('mousedown',e=>{if(e.target===ao&&!ao.classList.contains('admin-full'))ao.classList.remove('on')});
})();

let CMT={};
async function loadCmt(){if(!me)return;try{const j=await fetch('/api/excel-stock/comment-latest').then(r=>r.json());CMT=j.map||{};applyCmt()}catch(e){}}
function applyCmt(){
 document.querySelectorAll('#grid .c[data-k]').forEach(el=>{
  const p=MAP[el.dataset.k];if(!p)return;const id=String(productId(p)||'').trim().toLowerCase();const o=id&&CMT[id];
  let n=el.querySelector('.cmt-note');
  if(!o||!o.last){if(n)n.remove();return}
  const l=o.last,txt=`${esc(l.size||'')} · ${bn(l.sold_qty||0)} পিস বিক্রি${l.message?' — '+esc(l.message):''}`;
  const html=`<span class="cmt-ic">💬</span><div><b>${txt}</b><small>${esc(l.username||'')}${o.n>1?` · আরও ${bn(o.n-1)}টি কমেন্ট`:''}</small></div>`;
  if(!n){n=document.createElement('div');n.className='cmt-note';el.querySelector('.m').appendChild(n)}
  if(n.dataset.h!==html){n.innerHTML=html;n.dataset.h=html}
 })
}
new MutationObserver(()=>{clearTimeout(window._cmtT);window._cmtT=setTimeout(applyCmt,40)}).observe(document.getElementById('grid'),{childList:true});
setInterval(()=>{if(typeof me!=='undefined'&&me)loadCmt()},45000);
setTimeout(loadCmt,1500);setTimeout(loadCmt,5000);

/* প্রোডাক্টে ক্লিক করলে ওয়েবসাইটের সব ছবি আনা */
function renderThumbs(k){
 const p=MAP[k];if(!p||!$('mb'))return;const th=document.querySelector('#mb .th');if(!th)return;
 const imgs=[p.img,...(p.extra||[])].filter(Boolean);
 const lab=document.querySelector('#mb .gallery-label');if(lab)lab.innerHTML='📸 '+bn(imgs.length)+'টি ছবি · ছবিতে ক্লিক করলে বড় হবে, দিয়ে আলাদা কপি';
 th.innerHTML=imgs.map((u,i)=>`<div class="thumb-wrap"><img src="${im(u)}" loading="lazy" onclick="pickModalImg(this,'${esc(u)}')" title="ছবি ${i+1}"><button type="button" class="thumb-copy" data-perm="copy" onclick="event.stopPropagation();copyImageUrl('${esc(u)}')">📋</button></div>`).join('');
 applyPermUI();
}
function pickModalImg(el,u){const m=$('mi');if(!m)return;m.src=el.src;m.dataset.u=u;document.querySelectorAll('#mb .thumb-wrap img').forEach(x=>x.classList.toggle('cur',x===el))}
async function loadGallery(k){
 const p=MAP[k];if(!p)return;
 try{const j=await fetch('/api/product-gallery?key='+encodeURIComponent(k)).then(r=>r.json());
  const add=(j.images||[]).filter(u=>u&&u!==p.img);if(!add.length)return;
  const merged=[...new Set([...(p.extra||[]),...add])];
  if(merged.length===(p.extra||[]).length)return;
  p.extra=merged;if(typeof saveLocalProductCache==='function')try{saveLocalProductCache()}catch(e){}
  renderThumbs(k)}catch(e){}
}
(function(){const o=openM;openM=function(k){o(k);loadGallery(k)}})();
/* কমেন্ট করলেই প্রোডাক্টের ওপরে/নিচে সবার জন্য সঙ্গে সঙ্গে দেখানো */
function showTopComments(){
 const rows=[...document.querySelectorAll('#stockComments .stock-comment-row')].slice(0,3);
 const host=document.querySelector('#mb .product-summary');if(!host)return;
 let box=document.getElementById('modalCmtTop');
 if(!rows.length){if(box)box.remove();return}
 if(!box){box=document.createElement('div');box.id='modalCmtTop';const anchor=host.querySelector('.product-modal-sizes')||host.querySelector('.stockstate');anchor?anchor.after(box):host.querySelector('.acts').before(box)}
 box.innerHTML=rows.map(r=>`<div class="cmt-note"><span class="cmt-ic">💬</span><div><b>${esc(r.querySelector('b')?.textContent||'')}</b>${r.querySelector('p')?`<b style="font-weight:500">${esc(r.querySelector('p').textContent)}</b>`:''}<small>${esc(r.querySelector('small')?.textContent||'')}</small></div></div>`).join('');
}
(function(){const o=loadStockComments;loadStockComments=async function(p){await o(p);showTopComments()}})();
(function(){const o=submitStockComment;submitStockComment=async function(e){await o(e);loadCmt()}})();
/* স্টক ম্যানেজমেন্টে: কয়টি প্রোডাক্টে কমেন্ট হয়েছে */
async function renderCommentSummary(){
 const host=document.getElementById('adminContent');if(!host)return;
 let j;try{j=await jf('/api/excel-stock/comment-summary')}catch(e){return}
 if(!j||j.error)return;
 let box=document.getElementById('cmtSummary');
 if(!box){box=document.createElement('section');box.id='cmtSummary';host.prepend(box)}
 const fmt=t=>new Date(t*1000).toLocaleString('en-US',{day:'numeric',month:'short',hour:'numeric',minute:'2-digit'});
 const chips=(j.products||[]).map(p=>`<button type="button" class="cs-chip" onclick="openStockProductFromRow('${esc(p.product_code)}')">${esc(p.product_code)}<i>${bn(p.n)}</i></button>`).join('')||'<small>এখনো কোনো প্রোডাক্টে কমেন্ট হয়নি</small>';
 const rows=(j.comments||[]).map(x=>`<div class="stock-comment-row"><b>${esc(x.product_code)} · ${esc(x.size)} · ${bn(x.sold_qty)} পিস</b><small>${esc(x.sheet)} · ${esc(x.username)} · ${esc(fmt(x.created_at))}</small>${x.message?`<p>${esc(x.message)}</p>`:''}${j.can_delete?`<button class="pill" onclick="deleteStockComment(${Number(x.id)||0})">মুছুন</button>`:''}</div>`).join('');
 box.innerHTML=`<details class="comments-overview"><summary><span><b>🧾 বিক্রির কমেন্ট — এক নজরে</b><small>সবাই কমেন্ট দেখতে পারবে; মুছতে পারবেন শুধু অ্যাডমিন।</small></span><i>＋</i></summary><div class="comments-overview-body"><div class="cs-stats"><div><b>${bn(j.total_products||0)}</b><span>প্রোডাক্টে কমেন্ট</span></div><div><b>${bn(j.total_comments||0)}</b><span>মোট কমেন্ট</span></div><div><b>${bn(j.total_sold||0)}</b><span>মোট বিক্রি (পিস)</span></div></div><div class="cs-chips">${chips}</div>${rows?`<details><summary style="cursor:pointer;font-weight:600;margin-bottom:10px">সাম্প্রতিক কমেন্ট দেখুন</summary><div class="cs-list">${rows}</div></details>`:''}</div></details>`;
}
</script>
<script>
(async function showSiteNotice(){try{const r=await fetch('/api/notice',{cache:'no-store'}),n=await r.json();if(!n.enabled||!String(n.text||'').trim())return;const shade=document.createElement('div');shade.id='siteNoticeShade';const card=document.createElement('div');card.id='siteNoticeCard';const title=document.createElement('h3');title.textContent='📢 নোটিশ';const close=document.createElement('button');close.id='siteNoticeClose';close.type='button';close.setAttribute('aria-label','নোটিশ বন্ধ করুন');close.title='বন্ধ করুন';close.textContent='×';const body=document.createElement('div');body.textContent=n.text;const dismiss=()=>shade.remove();close.addEventListener('click',dismiss);shade.addEventListener('click',e=>{if(e.target===shade)dismiss()});card.append(close,title,body);shade.append(card);document.body.append(shade);setTimeout(dismiss,Math.max(1,Math.min(300,Number(n.duration_seconds||6)))*1000)}catch(e){}})();

/* ===== POS (বিক্রয় কাউন্টার) ===== */
const POS={cart:[],res:[],hi:0,pend:-1,seq:0,busy:false,built:false,paidTouched:false,t:null,last:null,tot:0,geoData:null,geoDistricts:[],geoLoading:null,historyRows:[],customerMatches:[],canEditPrice:false,phoneSeq:0,phoneTimer:null,deliveryRates:{delivery_inside_dhaka:'0',delivery_dhaka_suburbs:'0',delivery_outside_dhaka:'0'}};
const BD_GEO_URL='https://iqbalhasandev.github.io/bangladesh-geo-json/bangladesh-geo.json';
function posMethodChange(){const meth=$('posMeth')?.value||'Cash';if(meth==='COD'){POS.paidTouched=true;$('posPaid').value='0'}else{POS.paidTouched=false}$('posPaid').value=meth==='COD'?'0':(POS.tot||'');posCalc()}
const posM=n=>'৳ '+bp(Math.round((Number(n)||0)*100)/100);
async function posLoadCompanyRates(companyKey){try{const key=companyKey||$('posCompany')?.value||'alien_tees';const r=await fetch('/api/company?company='+encodeURIComponent(key),{cache:'no-store'}),j=await r.json(),c=j.company||{};for(const k of Object.keys(POS.deliveryRates))POS.deliveryRates[k]=String(c[k]??'0');posDeliveryZoneChange(false)}catch(e){console.warn('Delivery rates could not be loaded',e)}}
function posCompanyChange(){const key=$('posCompany')?.value||'alien_tees';posLoadCompanyRates(key)}
function posSetOptions(sel,items,placeholder){if(!sel)return;sel.innerHTML='<option value="">'+esc(placeholder)+'</option>'+(items||[]).map((x,i)=>'<option value="'+i+'">'+esc(typeof x==='string'?x:(x.name||x.bn_name||''))+'</option>').join('');sel.disabled=!(items||[]).length}
function posSetDatalist(id,items){const list=$(id);if(!list)return;list.innerHTML=(items||[]).map(x=>'<option value="'+esc(typeof x==='string'?x:(x.name||x.bn_name||''))+'"></option>').join('')}
function posGeoLabel(x){return String((x&&x.name)|| (x&&x.bn_name)||'')}
function posGeoMatches(kind){const input=$(kind==='district'?'posDistrictSearch':'posUpazilaSearch');const q=String(input?.value||'').trim().toLocaleLowerCase();let items=[];if(kind==='district')items=POS.geoDistricts||[];else{const xi=$('posDistrict')?.value;const d=xi!==''&&xi!=null?(POS.geoDistricts||[])[Number(xi)]:null;items=d?(d.upazilas||[]):[];}return items.map((item,index)=>({item,index})).filter(x=>!q||posGeoLabel(x.item).toLocaleLowerCase().includes(q)||String(x.item.bn_name||'').toLocaleLowerCase().includes(q)).slice(0,40)}
function posHideGeoSuggestions(kind){const box=$(kind==='district'?'posDistrictSuggestions':'posUpazilaSuggestions');if(box)box.hidden=true}
function posShowGeoSuggestions(kind){const input=$(kind==='district'?'posDistrictSearch':'posUpazilaSearch');const box=$(kind==='district'?'posDistrictSuggestions':'posUpazilaSuggestions');if(!input||!box||input.disabled){if(box)box.hidden=true;return}const matches=posGeoMatches(kind);box.innerHTML=matches.map(x=>`<button type="button" class="pos-geo-option" onmousedown="event.preventDefault()" onclick="posSelectGeo('${kind}',${x.index})">${esc(posGeoLabel(x.item))}${x.item.bn_name&&x.item.bn_name!==x.item.name?`<small>${esc(x.item.bn_name)}</small>`:''}</button>`).join('');box.hidden=matches.length===0}
function posSelectGeo(kind,index){if(kind==='district'){const item=(POS.geoDistricts||[])[Number(index)];if(!item)return;const input=$('posDistrictSearch'),hidden=$('posDistrict');input.value=posGeoLabel(item);hidden.value=String(index);posDistrictChange();posHideGeoSuggestions('district');setTimeout(()=>{const up=$('posUpazilaSearch');if(up&&!up.disabled)up.focus()},0);return}const xi=$('posDistrict')?.value;const d=xi!==''&&xi!=null?(POS.geoDistricts||[])[Number(xi)]:null;const item=d?(d.upazilas||[])[Number(index)]:null;if(!item)return;$('posUpazilaSearch').value=posGeoLabel(item);$('posUpazila').value=String(index);posHideGeoSuggestions('upazila')}
function posGeoSuggestionKeydown(e,kind){if(e.key==='Escape'){posHideGeoSuggestions(kind);return}if(e.key==='Enter'){const box=$(kind==='district'?'posDistrictSuggestions':'posUpazilaSuggestions');if(box&&!box.hidden){const first=box.querySelector('button');if(first){e.preventDefault();first.click()}}}}
function posDistrictSearchChange(){const input=$('posDistrictSearch'),hidden=$('posDistrict');if(!input||!hidden)return;const q=input.value.trim().toLocaleLowerCase();const i=q?(POS.geoDistricts||[]).findIndex(x=>posGeoLabel(x).toLocaleLowerCase()===q||String(x.bn_name||'').trim().toLocaleLowerCase()===q):-1;if(i>=0){if(hidden.value!==String(i)){hidden.value=String(i);posDistrictChange()}posHideGeoSuggestions('district');return}if(hidden.value!==''){hidden.value='';posDistrictChange()}posShowGeoSuggestions('district')}
function posUpazilaSearchChange(){const input=$('posUpazilaSearch'),hidden=$('posUpazila');if(!input||!hidden)return;const xi=$('posDistrict')?.value;const d=xi!==''&&xi!=null?(POS.geoDistricts||[])[Number(xi)]:null;const items=d?(d.upazilas||[]):[];const q=input.value.trim().toLocaleLowerCase();const i=q?items.findIndex(x=>posGeoLabel(x).toLocaleLowerCase()===q||String(x.bn_name||'').trim().toLocaleLowerCase()===q):-1;if(i>=0){hidden.value=String(i);posHideGeoSuggestions('upazila');return}hidden.value='';posShowGeoSuggestions('upazila')}
async function posLoadGeo(){if(POS.geoDistricts.length)return true;if(POS.geoLoading)return POS.geoLoading;POS.geoLoading=(async()=>{try{const r=await fetch(BD_GEO_URL,{cache:'force-cache'});if(!r.ok)throw new Error('Geography data unavailable');const data=await r.json();if(!Array.isArray(data)||!data.some(x=>Array.isArray(x.districts)))throw new Error('Invalid geography data');POS.geoData=data;POS.geoDistricts=data.flatMap(x=>Array.isArray(x.districts)?x.districts:[]);if(!POS.geoDistricts.length)throw new Error('No districts found');posSetOptions($('posDistrict'),POS.geoDistricts,'Select district');posSetOptions($('posUpazila'),[],'Select district first');$('posGeoFallback').hidden=true;$('posGeoSelects').hidden=false;return true}catch(e){console.warn('Bangladesh address list unavailable; manual district/thana entry enabled.',e);$('posGeoSelects').hidden=true;$('posGeoFallback').hidden=false;return false}})();return POS.geoLoading}
function posDistrictChange(){const xi=$('posDistrict')?.value;const z=POS.geoDistricts&&xi!==''?POS.geoDistricts[Number(xi)]:null;const items=z?(z.upazilas||[]):[];posSetOptions($('posUpazila'),items,z?'Select thana / upazila':'Select district first');const up=$('posUpazilaSearch');if(up){up.value='';up.disabled=!items.length;up.placeholder=items.length?'Type or select a thana / upazila…':'Select district first'}const upHidden=$('posUpazila');if(upHidden)upHidden.value='';posHideGeoSuggestions('upazila')}
// Address selection is optional and must never silently apply a delivery fee.
function posSuggestDeliveryZone(){return}
function posDeliveryZoneChange(touched=true){const sel=$('posDeliveryZone'),inp=$('posDeliveryCharge');if(!sel||!inp)return;if(touched){sel.dataset.touched='1';inp.dataset.manual='0';const k={inside:'delivery_inside_dhaka',suburbs:'delivery_dhaka_suburbs',outside:'delivery_outside_dhaka'}[sel.value];if(sel.value==='none')inp.value='';else inp.value=String(Math.max(0,Number(POS.deliveryRates[k])||0));}else if(sel.value!=='none'&&inp.dataset.manual!=='1'){const k={inside:'delivery_inside_dhaka',suburbs:'delivery_dhaka_suburbs',outside:'delivery_outside_dhaka'}[sel.value];inp.value=String(Math.max(0,Number(POS.deliveryRates[k])||0));}const n=inp.value.trim()===''?0:Math.max(0,Number(inp.value)||0);const txt=$('posDeliveryPreview');if(txt)txt.textContent=posM(n);posCalc()}
function posDeliveryManualChange(){const inp=$('posDeliveryCharge');if(!inp)return;inp.dataset.manual='1';const n=inp.value.trim()===''?0:Math.max(0,Number(inp.value)||0);const txt=$('posDeliveryPreview');if(txt)txt.textContent=posM(n);posCalc()}
function posComposeAddress(){const v=id=>String($(id)?.value||'').trim();let district='',upazila='';if(!$('posGeoSelects')?.hidden&&POS.geoDistricts.length){const xi=v('posDistrict'),ui=v('posUpazila');const x=xi!==''?POS.geoDistricts[Number(xi)]:null;const u=x&&ui!==''?(x.upazilas||[])[Number(ui)]:null;district=x?posGeoLabel(x):v('posDistrictSearch');upazila=u?posGeoLabel(u):v('posUpazilaSearch')}else{district=v('posDistrictText');upazila=v('posUpazilaText')}return [['District',district],['Thana/Upazila',upazila],['Address',v('posVillage')]].filter(x=>x[1]).map(x=>x[0]+': '+x[1]).join(', ')}
function posBuild(){
 $('pos').innerHTML=`<div class="pos-top"><b>🛒 POS · বিক্রয় কাউন্টার</b><span class="pos-chip" id="posToday">আজকের বিক্রি: —</span><span class="sp"></span><span class="pos-chip">👤 ${esc(me||'')}</span><button class="pos-back" onclick="closePOS()">← মূল পেজ</button></div>
<div class="pos-main"><section class="pos-left">
<div class="pos-search"><span>🔍</span><input id="posQ" type="text" autocomplete="off" placeholder="কোড বা নাম লিখে খুঁজুন… (F2)" oninput="posType()" onkeydown="posKey(event)"></div>
<div class="pos-hints"><span><kbd>↑↓</kbd>বাছুন</span><span><kbd>Enter</kbd>যোগ</span><span><kbd>1-9</kbd>সাইজ</span><span><kbd>F2</kbd>সার্চ</span><span><kbd>F4</kbd>ফোন</span><span><kbd>F1</kbd>পেমেন্ট</span><span><kbd>F8</kbd>বিক্রি</span></div>
<div class="pos-results" id="posRes"></div>
<div class="pos-cart"><div class="pos-cart-h">কার্ট <span class="pos-badge" id="posCnt">0</span></div><div id="posCart"></div></div>
</section><aside class="pos-right">
<h4>কোম্পানি / মেমো</h4>
<label>বিক্রির কোম্পানি<select id="posCompany" onchange="posCompanyChange()"><option value="alien_tees" selected>Alien Tees</option><option value="bodyfitters">Bodyfitters</option></select></label>
<h4>কাস্টমার</h4>
<label>Phone / Search<input id="posPhone" inputmode="tel" autocomplete="tel" placeholder="ফোনের প্রথম ৩ সংখ্যা লিখুন… (F4)" oninput="posPhoneInput()"></label>
<div id="posHistory" class="pos-customer-history" hidden></div>
<label>কাস্টমারের নাম<input id="posName" placeholder="কাস্টমারের নাম"></label>
<h4>ডেলিভারির ঠিকানা — ধাপে ধাপে নির্বাচন করুন</h4>
<div id="posGeoSelects">
<label>District (Optional)<div class="pos-geo-field"><input id="posDistrictSearch" autocomplete="off" placeholder="Type to search district…" oninput="posDistrictSearchChange()" onfocus="posShowGeoSuggestions('district')" onblur="posHideGeoSuggestions('district')" onkeydown="posGeoSuggestionKeydown(event,'district')"><div id="posDistrictSuggestions" class="pos-geo-suggestions" hidden></div></div><select id="posDistrict" onchange="posDistrictChange()" hidden aria-hidden="true"><option value="">Select district</option></select></label>
<label>Thana / Upazila (Optional)<div class="pos-geo-field"><input id="posUpazilaSearch" autocomplete="off" placeholder="Select district first" oninput="posUpazilaSearchChange()" onfocus="posShowGeoSuggestions('upazila')" onblur="posHideGeoSuggestions('upazila')" onkeydown="posGeoSuggestionKeydown(event,'upazila')" disabled><div id="posUpazilaSuggestions" class="pos-geo-suggestions" hidden></div></div><select id="posUpazila" disabled hidden aria-hidden="true"><option value="">Select district first</option></select></label>
</div>
<div id="posGeoFallback" hidden>
<label>District (Optional)<input id="posDistrictText" placeholder="District"></label>
<label>Thana / Upazila (Optional)<input id="posUpazilaText" placeholder="Thana / Upazila"></label>
</div>
<label>Address<input id="posVillage" placeholder="Village, area, road or full address"></label>
<h4>ডেলিভারি চার্জ</h4>
<label>ডেলিভারি এলাকা (ঐচ্ছিক)<select id="posDeliveryZone" onchange="posDeliveryZoneChange()"><option value="none">চার্জ নেই / নিজে সংগ্রহ</option><option value="inside">ঢাকার মধ্যে</option><option value="suburbs">ঢাকার আশেপাশে (সাভার / গাজীপুর / নারায়ণগঞ্জ)</option><option value="outside">ঢাকার বাইরে</option></select></label>
<label>ডেলিভারি চার্জ (ঐচ্ছিক, টাকা)<input id="posDeliveryCharge" type="number" min="0" max="100000" step="1" value="" placeholder="চার্জ না থাকলে ফাঁকা রাখুন" oninput="posDeliveryManualChange()"></label>
<h4>ডিসকাউন্ট</h4>
<div class="pos-two"><select id="posDT" onchange="posCalc()"><option value="none">ডিসকাউন্ট নেই</option><option value="percent">শতকরা (%)</option><option value="amount">টাকা (৳)</option></select><input id="posDV" type="number" min="0" step="any" value="0" oninput="posCalc()" style="display:none"></div>
<h4>সারাংশ</h4>
<div class="pos-sum"><div><span>সাবটোটাল</span><b id="posSub">৳ 0</b></div><div><span>ডিসকাউন্ট</span><b id="posDis">৳ 0</b></div><div><span>ডেলিভারি চার্জ</span><b id="posDeliveryPreview">৳ 0</b></div><div class="tot"><span>সর্বমোট</span><b id="posTot">৳ 0</b></div></div>
<h4>পেমেন্ট</h4>
<label>মাধ্যম (F1)<select id="posMeth" onchange="posMethodChange()"><option value="Cash">নগদ (Cash)</option><option value="bKash">bKash</option><option value="Nagad">Nagad</option><option value="Card">কার্ড</option><option value="COD" selected>Cash on Delivery (COD)</option></select></label>
<div class="pos-two"><label>পরিশোধ<input id="posPaid" type="number" min="0" step="any" oninput="POS.paidTouched=true;posCalc()"></label><label><span id="posChL">ফেরত</span><input id="posCh" readonly></label></div>
<button class="pos-pay" id="posPay" onclick="posPay()">✔ বিক্রি সম্পন্ন (F8)</button>
<button class="pos-clear" onclick="posClear(true)">কার্ট খালি করুন</button>
<div id="posMsg" class="pos-msg" hidden></div>
</aside></div><div id="posRcpt" hidden></div>`;
 posLoadGeo();posLoadCompanyRates();posMethodChange();
}
window.openPOS = async function openPOS(){
 try{
  if(!me){needLogin();return}
  if(!allowed('pos_use')&&!isAdminLevelJS()){toast('POS permission is disabled for this account.',true);return}
  const layer=$('pos');if(!layer)throw new Error('POS panel element is missing');
  if(!POS.built){posBuild();POS.built=true}
  layer.classList.add('on');layer.style.removeProperty('display');layer.style.visibility='visible';layer.style.pointerEvents='auto';
  document.body.style.overflow='hidden';
  posRenderCart();posCalc();loadPosToday();const results=$('posRes');if(results)results.innerHTML='<div class="pos-empty">Type a product code or name to search.</div>';POS.res=[];POS.hi=0;POS.pend=-1;
  setTimeout(()=>{const q=$('posQ');if(q){q.focus({preventScroll:true})}},80);
 }catch(err){console.error('POS open failed:',err);toast('POS could not be opened. Check the browser console.',true)}
};
// POS button: use the normal inline handler as the primary action; keep the button clickable.
(function bindHeaderPOS(){
 const bind=()=>{const b=document.getElementById('posBtn');if(!b)return;b.style.pointerEvents='auto';b.style.position='relative';b.style.zIndex='2147482003';b.type='button';};
 bind();document.addEventListener('DOMContentLoaded',bind);setTimeout(bind,0);
})();
function closePOS(){$('pos').classList.remove('on');$('pos').style.display='';$('pos').style.visibility='';$('pos').style.pointerEvents='';document.body.style.overflow=''}
async function loadPosToday(){const j=await jf('/api/pos/summary');if(j._ok&&$('posToday'))$('posToday').textContent='আজকের বিক্রি: '+posM(j.total)+' · '+bp(j.count)+'টি'}
function posType(){clearTimeout(POS.t);POS.pend=-1;POS.t=setTimeout(()=>posSearch($('posQ').value),130)}
async function posSearch(q){
 const query=String(q||'').trim(), my=++POS.seq;
 POS.res=[];POS.hi=0;POS.pend=-1;
 if(!query){const box=$('posRes');if(box)box.innerHTML='<div class="pos-empty">Type a product code or name to search.</div>';return}
 const j=await jf('/api/pos/products?q='+encodeURIComponent(query));
 if(my!==POS.seq)return;
 if(!j._ok){$('posRes').innerHTML='<div class="pos-empty">'+esc(j.error||'Products could not be loaded')+'</div>';POS.res=[];return}
 POS.res=j.items||[];POS.canEditPrice=!!j.can_edit_price;POS.hi=0;POS.pend=-1;posRenderRes();
}
function posRenderRes(){
 const box=$('posRes');
 if(!POS.res.length){box.innerHTML='<div class="pos-empty">কোনো প্রোডাক্ট পাওয়া যায়নি</div>';return}
 box.innerHTML=POS.res.map((p,i)=>{
  const sz=Object.entries(p.sizes||{}).filter(([,q])=>q>0);p._sz=sz;
  const chips=sz.length?sz.map(([s,q],j)=>`<button type="button" onclick="event.stopPropagation();posAddIdx(${i},${j})" title="স্টক ${bp(q)}টি">${POS.pend===i&&j<9?`<em>${j+1}</em>`:''}${esc(s)}<i>${bp(q)}</i></button>`).join(''):'<span class="pos-out">স্টক নেই</span>';
  return `<div class="pos-row ${i===POS.hi?'hl':''}" data-i="${i}" onclick="posSetHi(${i})"><img src="${p.img?im(p.img):''}" alt="" loading="lazy"><div class="pi"><b>${esc(p.id)}<span class="pos-store">${esc(p.store)}</span></b><small>${esc(p.title)}</small><div class="pos-sz">${chips}</div></div><div class="pp">${posM(p.price)}</div></div>`}).join('');
 posMarkHi();
}
function posMarkHi(){const rows=$('posRes').querySelectorAll('.pos-row');rows.forEach((r,i)=>r.classList.toggle('hl',i===POS.hi));const h=rows[POS.hi];if(h)h.scrollIntoView({block:'nearest'})}
function posSetHi(i){if(POS.pend>=0&&POS.pend!==i){POS.pend=-1;POS.hi=i;posRenderRes();return}POS.hi=i;posMarkHi()}
function posMove(d){if(!POS.res.length)return;const was=POS.pend;POS.pend=-1;POS.hi=(POS.hi+d+POS.res.length)%POS.res.length;if(was>=0)posRenderRes();else posMarkHi()}
function posEnter(){
 const p=POS.res[POS.hi];if(!p)return;const sz=p._sz||[];
 if(!sz.length)return toast('এই প্রোডাক্টের স্টক নেই',true);
 if(sz.length===1)return posAddIdx(POS.hi,0);
 POS.pend=POS.hi;posRenderRes();toast('সাইজ বাছুন: 1-'+Math.min(9,sz.length)+' চাপুন বা সাইজে ক্লিক করুন');
}
function posKey(e){
 if(e.key==='ArrowDown'){e.preventDefault();posMove(1)}
 else if(e.key==='ArrowUp'){e.preventDefault();posMove(-1)}
 else if(e.key==='Enter'){e.preventDefault();posEnter()}
 else if(POS.pend>=0&&/^[1-9]$/.test(e.key)){e.preventDefault();posAddIdx(POS.pend,Number(e.key)-1)}
 else if(e.key==='Escape'&&$('posQ').value){$('posQ').value='';posSearch('')}
}
function posAddIdx(i,j){const p=POS.res[i];const x=p&&(p._sz||[])[j];if(!x)return;posAdd(p,x[0],x[1])}
function posAdd(p,s,max){
 const l=POS.cart.find(x=>x.key===p.key&&x.size===s);
 if(l){if(l.qty>=max)return toast('স্টকে আর নেই (মোট '+bp(max)+'টি)',true);l.qty++;l.max=max}
 else POS.cart.unshift({key:p.key,id:p.id,title:p.title,img:p.img,size:s,price:Number(p.price)||0,basePrice:Number(p.price)||0,qty:1,max});
 POS.pend=-1;posRenderCart();posCalc();
 const q=$('posQ');q.value='';q.focus();posSearch('');
}
function posRenderCart(){
 const box=$('posCart');$('posCnt').textContent=bp(POS.cart.reduce((a,l)=>a+l.qty,0));
 if(!POS.cart.length){box.innerHTML='<div class="pos-empty">কার্ট খালি — প্রোডাক্ট খুঁজে সাইজ বেছে যোগ করুন</div>';return}
 box.innerHTML=POS.cart.map((l,i)=>`<div class="pos-item"><img src="${l.img?im(l.img):''}" alt=""><div class="pt"><b>${esc(l.id)} · ${esc(l.size)}</b><small>${esc(l.title)}</small>${POS.canEditPrice?`<label class="pos-price-edit">প্রতি পিস দাম (৳)<input type="number" inputmode="decimal" min="0" max="1000000" step="0.01" value="${Number(l.price)||0}" onchange="posSetPrice(${i},this.value)" aria-label="${esc(l.title)}-এর unit price"></label>`:`<small>${posM(l.price)} প্রতি পিস</small>`}<small>স্টক ${bp(l.max)}টি</small><div class="pos-qty"><button type="button" onclick="posQty(${i},-1)">−</button><input type="number" min="1" max="${l.max}" value="${l.qty}" onchange="posQtySet(${i},this.value)"><button type="button" onclick="posQty(${i},1)">+</button></div></div><div class="pl"><button type="button" class="pos-del" onclick="posDel(${i})" title="মুছুন">🗑</button><b>${posM(l.price*l.qty)}</b><small>${bp(l.qty)} × ${posM(l.price)}</small></div></div>`).join('');
}
function posQty(i,d){posQtySet(i,(POS.cart[i]?.qty||1)+d)}
function posSetPrice(i,value){const l=POS.cart[i];if(!l)return;if(!POS.canEditPrice)return toast('দাম পরিবর্তনের অনুমতি নেই',true);const raw=String(value??'').trim();let n=raw===''?Number(l.basePrice||0):Number(raw);if(!Number.isFinite(n)||n<0||n>1000000){toast('দাম 0 থেকে 1000000 টাকার মধ্যে হতে হবে',true);posRenderCart();return}l.price=Math.round(n*100)/100;posRenderCart();posCalc()}
function posQtySet(i,v){const l=POS.cart[i];if(!l)return;let n=Math.floor(Number(v)||1);if(n<1)n=1;if(n>l.max){n=l.max;toast('স্টকে মাত্র '+bp(l.max)+'টি আছে',true)}l.qty=n;posRenderCart();posCalc()}
function posDel(i){POS.cart.splice(i,1);posRenderCart();posCalc()}
function posCalc(){
 const sub=POS.cart.reduce((a,l)=>a+l.price*l.qty,0),t=$('posDT').value,v=Math.max(0,Number($('posDV').value)||0),delivery=Math.max(0,Number($('posDeliveryCharge')?.value)||0);
 let dis=t==='percent'?sub*Math.min(v,100)/100:t==='amount'?Math.min(v,sub):0;
 dis=Math.round(dis*100)/100;const tot=Math.max(0,Math.round((sub-dis+delivery)*100)/100);POS.tot=tot;
 $('posDV').style.display=t==='none'?'none':'';
 $('posSub').textContent=posM(sub);$('posDis').textContent=dis?'− '+posM(dis):posM(0);if($('posDeliveryPreview'))$('posDeliveryPreview').textContent=posM(delivery);$('posTot').textContent=posM(tot);
 if(!POS.paidTouched)$('posPaid').value=tot||'';
 const d=(Number($('posPaid').value)||0)-tot;
 $('posChL').textContent=d<0?'বাকি':'ফেরত';$('posCh').value=String(Math.round(Math.abs(d)*100)/100);$('posCh').classList.toggle('due',d<0);
 $('posPay').disabled=!POS.cart.length||POS.busy;
}
function posPhoneInput(){
 clearTimeout(POS.phoneTimer);const phone=String($('posPhone')?.value||''),digits=phone.replace(/\D/g,'');const box=$('posHistory');if(!box)return;
 if(digits.length<3){POS.phoneSeq++;POS.historyRows=[];POS.customerMatches=[];box.hidden=true;box.innerHTML='';return}
 box.hidden=false;box.innerHTML='<div class="pos-history-muted">আগের অর্ডার খোঁজা হচ্ছে…</div>';const seq=++POS.phoneSeq;
 POS.phoneTimer=setTimeout(async()=>{const j=await jf('/api/pos/customer-history?phone='+encodeURIComponent(phone));if(seq!==POS.phoneSeq)return;POS.historyRows=j.orders||[];
 if(!j._ok){box.innerHTML='<div class="pos-history-muted">আগের অর্ডার দেখা যাচ্ছে না · '+esc(j.error||'Search failed')+'</div>';return}
 POS.customerMatches=j.customers||[];
 const suggestions=POS.customerMatches.length?'<div class="pos-customer-suggestions"><b>মিলে যাওয়া কাস্টমার</b>'+POS.customerMatches.map((c,i)=>`<button type="button" onclick="posUseHistoryCustomer(${i},true)"><span>${esc(c.name||'Customer')}</span><small>${esc(c.phone||'')}${c.order_no?' · '+esc(c.order_no):''}</small></button>`).join('')+'</div>':'';
 if(!POS.historyRows.length){box.innerHTML=suggestions+'<div class="pos-history-muted">এই নম্বরে আগের অর্ডার পাওয়া যায়নি</div>';return}
 box.innerHTML=suggestions+'<div class="pos-history-title">আগের অর্ডার · '+bp(POS.historyRows.length)+'টি</div>'+POS.historyRows.map((o,i)=>`<div class="pos-history-row"><div class="pos-history-order"><b>${esc(o.no||'Order')}</b><small>${esc(o.date||'')}</small></div><div class="pos-history-products">${esc(o.items_text||'পণ্যের বিবরণ নেই')}</div><div class="pos-history-money"><span>মোট <b>${posM(o.total)}</b></span><span>বাকি <b class="${Number(o.due)>0?'is-due':''}">${posM(o.due)}</b></span></div><button type="button" onclick="posUseHistoryCustomer(${i},false)">কাস্টমার তথ্য নিন</button></div>`).join('');
 },140);
}
function posUseHistoryCustomer(index,isSuggestion=false){const o=isSuggestion?POS.customerMatches[index]:POS.historyRows[index];if(!o)return;if(o.phone)$('posPhone').value=o.phone;if(o.name)$('posName').value=o.name;if(o.address&&!String($('posVillage').value||'').trim())$('posVillage').value=o.address;toast('আগের অর্ডার থেকে কাস্টমার তথ্য নেওয়া হয়েছে')}
function posClear(ask){
 if(ask&&POS.cart.length&&!confirm('কার্ট খালি করবেন?'))return;
 POS.cart=[];POS.paidTouched=($('posMeth')?.value==='COD');
 ['posPhone','posName','posVillage','posDistrictText','posUpazilaText'].forEach(k=>{if($(k))$(k).value=''});POS.phoneSeq++;POS.historyRows=[];if($('posHistory')){$('posHistory').hidden=true;$('posHistory').innerHTML=''}
 if($('posDistrict'))$('posDistrict').value='';posDistrictChange();
 if($('posDeliveryZone')){$('posDeliveryZone').dataset.touched='';$('posDeliveryZone').value='none'}
 if($('posDeliveryCharge')){$('posDeliveryCharge').value='';$('posDeliveryCharge').dataset.manual='0'}posDeliveryZoneChange(false);
 $('posDT').value='none';$('posDV').value='0';$('posPaid').value=$('posMeth')?.value==='COD'?'0':'';$('posMsg').hidden=true;
 posRenderCart();posCalc();
}
async function posPay(){
 if(POS.busy||!POS.cart.length)return;
 POS.busy=true;$('posPay').disabled=true;$('posPay').textContent='সেভ হচ্ছে…';
 const deliveryRaw=$('posDeliveryCharge')?.value??'';
 const body={company:$('posCompany')?.value||'alien_tees',items:POS.cart.map(l=>({key:l.key,size:l.size,qty:l.qty,price:l.price})),customer:{phone:$('posPhone').value,name:$('posName').value,address:posComposeAddress()},delivery_zone:$('posDeliveryZone')?.value||'none',delivery_charge:deliveryRaw.trim()===''?null:Number(deliveryRaw),discount:{type:$('posDT').value,value:Number($('posDV').value)||0},method:$('posMeth').value,paid:Number($('posPaid').value)||0};
 const j=await post('/api/pos/sale',body);
 POS.busy=false;$('posPay').textContent='✔ বিক্রি সম্পন্ন (F8)';
 if(!j._ok||!j.ok){const m=$('posMsg');m.hidden=false;m.textContent=j.error||'বিক্রি সেভ হয়নি';toast(j.error||'বিক্রি সেভ হয়নি',true);posCalc();posSearch($('posQ').value);return}
 POS.last=j;posClear(false);posRcpt(j);loadPosToday();loadProducts();
}
function posRcptHtml(j){
 const c=j.customer||{};
 return `<div class="r-h"><b>${esc(j.company_name||'Alien Tees')}</b><small>${esc(j.invoice)} · ${esc(new Date(j.created_at*1000).toLocaleString('bn-BD'))}</small></div>`
 +((c.name||c.phone)?`<div class="r-s"><span>${esc(c.name||'')}</span><span>${esc(c.phone||'')}</span></div>`:'')
 +`<table>${(j.lines||[]).map(l=>`<tr><td>${esc(l.id)} · ${esc(l.size)}<br><small>${esc(l.title)}</small></td><td>${bp(l.qty)} × ${bp(l.price)}</td><td style="text-align:right">${bp(l.price*l.qty)}</td></tr>`).join('')}</table>`
 +`<div class="r-s"><span>সাবটোটাল</span><span>${posM(j.subtotal)}</span></div>`
 +(j.discount?`<div class="r-s"><span>ডিসকাউন্ট</span><span>− ${posM(j.discount)}</span></div>`:'')
 +(j.delivery_charge?`<div class="r-s"><span>ডেলিভারি চার্জ</span><span>${posM(j.delivery_charge)}</span></div>`:'')
 +`<div class="r-s tot"><span>মোট</span><span>${posM(j.total)}</span></div>`
 +`<div class="r-s"><span>পরিশোধ (${esc(j.method)})</span><span>${posM(j.paid)}</span></div>`
 +(j.due>0?`<div class="r-s"><span>বাকি</span><span>${posM(j.due)}</span></div>`:j.change>0?`<div class="r-s"><span>ফেরত</span><span>${posM(j.change)}</span></div>`:'')
 +(j.qr_data?`<div class="r-qr"><img src="${Array.isArray(j.qr_data)?j.qr_data[0]:j.qr_data}" alt="Memo QR"><b>Memo QR</b><small>Scan to view memo · ${esc(j.invoice||'')}</small></div>`:'')
 +`<p class="r-f">ধন্যবাদ!</p>`;
}
function posRcpt(j){
 const r=$('posRcpt');r.hidden=false;
 r.innerHTML=`<div class="pos-rc-box"><div class="pos-rc-ok">✔ বিক্রি সম্পন্ন · স্টক আপডেট হয়েছে</div><div id="posRcBody">${posRcptHtml(j)}</div><div class="pos-rc-act"><button type="button" onclick="openMemos('pos',[],['${esc(j.invoice)}'])">🖨 মেমো PDF / প্রিন্ট</button><button type="button" class="pri" onclick="posRcClose()">নতুন বিক্রি</button></div></div>`;
 setTimeout(()=>r.querySelector('.pri')?.focus(),30);
}
function posRcClose(){$('posRcpt').hidden=true;$('posQ').focus();posSearch($('posQ').value)}
function posPrint(){
 if(!POS.last)return;
 const w=window.open('','_blank','width=380,height=640');
 if(!w)return toast('পপআপ ব্লক করা আছে — অনুমতি দিন',true);
 w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(POS.last.invoice)}</title><style>
body{font:13px system-ui,sans-serif;margin:12px;width:300px;color:#111;background:#fff}
table{width:100%;border-collapse:collapse}td{padding:4px 0;border-bottom:1px dashed #999;vertical-align:top}
small{color:#555}.r-s{display:flex;justify-content:space-between;margin:4px 0}.tot{font-weight:800;font-size:15px}
.r-h{text-align:center;margin-bottom:8px}.r-h b{display:block;font-size:15px}.r-qr{display:flex;flex-direction:column;justify-content:center;align-items:center;gap:5px;margin:12px auto;text-align:center;font-size:9px;max-width:190px;padding:9px;border:1px solid #d5e2d8;border-radius:10px;background:#f6faf7}.r-qr>b,.r-qr>small{display:block}.r-qr img{display:block;width:118px;height:118px;max-width:100%;object-fit:contain;image-rendering:pixelated;background:#fff}.r-qr small{font-size:8px}.r-f{text-align:center}
</style></head><body>${posRcptHtml(POS.last)}</body></html>`);
 w.document.close();w.focus();setTimeout(()=>w.print(),300);
}
document.addEventListener('keydown',e=>{
 const pos=$('pos');if(!pos)return;
 if(!pos.classList.contains('on')){if(e.key==='F9'&&me&&allowed('pos_use')){e.preventDefault();openPOS()}return}
 if(e.key==='F2'){e.preventDefault();$('posQ').focus();$('posQ').select()}
 else if(e.key==='F4'){e.preventDefault();$('posPhone').focus()}
 else if(e.key==='F1'){e.preventDefault();$('posMeth').focus()}
 else if(e.key==='F8'){e.preventDefault();if($('posRcpt').hidden)posPay()}
 else if(e.key==='Escape'&&!$('posRcpt').hidden){posRcClose()}
});

</script>
</body></html>"""

init()
# Optional CLIP model is loaded only after image search starts; app startup remains fast.
threading.Thread(target=loader, daemon=True).start()
# ===================== POS (বিক্রয় কাউন্টার) =====================
# বিক্রি হলে Excel stock workbook থেকে সাইজ-ভিত্তিক পরিমাণ কাটা হয়; আগের সব stock রেকর্ড/ইতিহাস অপরিবর্তিত থাকে।
_POS_SNAP = {"t": 0.0}
_POS_METHODS = ("Cash", "bKash", "Nagad", "Card", "COD")

def _pos_stock_keys(p):
    ks = [str(p["id"]).strip().lower()]
    for part in re.split(r"[-_ /]", str(p["id"]).lower()) + re.findall(r"[a-z0-9]+", str(p["title"]).lower()):
        if len(part) > 2 and re.search(r"\d", part) and re.search(r"[a-z]", part): ks.append(part)
    return ks

def _pos_code_for(p):
    stock = C.get("stock") or {}
    for k in _pos_stock_keys(p):
        if k in stock: return k
    return None

@app.get("/api/pos/products")
@perm_auth("pos_use")
def pos_products():
    q = request.args.get("q", "").strip().lower()
    toks = [t for t in re.split(r"\s+", q) if t]
    pr = prices(); out = []
    for p in C["products"]:
        key = p["store"] + "|" + p["id"]
        did = display_product_id(p) or str(p["id"])
        hay = (did + " " + str(p["title"]) + " " + str(p["id"]) + " " + str(p["store"])).lower()
        if toks and not all(t in hay for t in toks): continue
        sizes = dict((C["stk"].get(key) or {}).get("sizes") or {})
        out.append(dict(key=key, id=did, title=p["title"], store=p["store"],
                        price=pr.get(key, p["price"]), img=p.get("img", ""),
                        sizes=sizes, total=sum(sizes.values())))
    out.sort(key=lambda x: (0 if x["id"].lower() == q else 1, 0 if x["total"] > 0 else 1))
    r = jsonify(items=out[:30], can_edit_price=bool(is_admin_level() or has_perm("price_edit"))); r.headers["Cache-Control"] = "no-store"; return r

@app.get("/api/pos/summary")
@perm_auth("pos_use")
def pos_summary():
    start = datetime.now(DHAKA_TZ).replace(hour=0, minute=0, second=0, microsecond=0).timestamp()
    with db() as c:
        if is_admin_level():
            row = c.execute("SELECT COUNT(*) n, COALESCE(SUM(total),0) t FROM pos_sales WHERE created_at>=?", (start,)).fetchone()
        else:
            row = c.execute("SELECT COUNT(*) n, COALESCE(SUM(total),0) t FROM pos_sales WHERE created_at>=? AND username=?", (start, session.get("user"))).fetchone()
    return jsonify(ok=True, count=row["n"], total=round(row["t"], 2))

@app.post("/api/pos/sale")
@perm_auth("pos_use")
def pos_sale():
    j = request.get_json(silent=True) or {}
    company_key = _company_key(j.get("company") or "alien_tees")
    company_info = company_profile(company_key)
    items = j.get("items")
    if not isinstance(items, list) or not items or len(items) > 100:
        return jsonify(error="কার্ট খালি বা অনেক বড়"), 400
    pmap = {p["store"] + "|" + p["id"]: p for p in C["products"]}
    pr = prices()
    merged = {}
    for it in items:
        if not isinstance(it, dict): return jsonify(error="কার্টের তথ্য ঠিক নেই"), 400
        key = str(it.get("key", "")); size = normalize_stock_size(it.get("size", ""))
        try: qty = int(it.get("qty"))
        except (TypeError, ValueError): return jsonify(error="পরিমাণ পূর্ণসংখ্যা হতে হবে"), 400
        if qty < 1 or qty > 1000 or not size: return jsonify(error="পরিমাণ বা সাইজ ঠিক নেই"), 400
        p = pmap.get(key)
        if not p: return jsonify(error="প্রোডাক্ট ক্যাটালগে পাওয়া যায়নি; রিফ্রেশ করে আবার চেষ্টা করুন"), 404
        did = display_product_id(p) or str(p["id"])
        code = _pos_code_for(p)
        if not code: return jsonify(error=did + ": স্টক শিটে এই কোড নেই"), 409
        unit_price = float(pr.get(key, p["price"]))
        if "price" in it and (is_admin_level() or has_perm("price_edit")):
            try: unit_price = float(it.get("price"))
            except (TypeError, ValueError): return jsonify(error="পণ্যের দাম সঠিক সংখ্যা হতে হবে"), 400
            if not (0 <= unit_price <= 1000000): return jsonify(error="পণ্যের দাম 0 থেকে 1000000 টাকার মধ্যে হতে হবে"), 400
        # Include price in the merge key so differently priced lines remain individually auditable.
        mk = (key, size.lower(), round(unit_price, 2))
        if mk in merged: merged[mk]["qty"] += qty
        else: merged[mk] = dict(key=key, code=code, size=size, qty=qty, price=round(unit_price, 2), title=p["title"], id=did)
    lines = list(merged.values())
    subtotal = round(sum(l["price"] * l["qty"] for l in lines), 2)
    d = j.get("discount") if isinstance(j.get("discount"), dict) else {}
    dtype = str(d.get("type", "none")); 
    try: dval = max(0.0, float(d.get("value") or 0))
    except (TypeError, ValueError): dval = 0.0
    discount = round(subtotal * min(dval, 100.0) / 100.0, 2) if dtype == "percent" else round(min(dval, subtotal), 2) if dtype == "amount" else 0.0
    # Delivery is opt-in: use the exact amount selected/entered in POS when supplied.
    # For older clients that do not send delivery_charge, keep the My Company rate behavior.
    delivery_zone = str(j.get("delivery_zone", "none") or "none").strip().lower()
    delivery_keys = {"inside": "delivery_inside_dhaka", "suburbs": "delivery_dhaka_suburbs", "outside": "delivery_outside_dhaka"}
    if delivery_zone not in delivery_keys: delivery_zone = "none"
    delivery_charge = 0.0
    if "delivery_charge" in j:
        raw_delivery = j.get("delivery_charge")
        if raw_delivery is not None and str(raw_delivery).strip() != "":
            try: delivery_charge = float(raw_delivery)
            except (TypeError, ValueError): return jsonify(error="ডেলিভারি চার্জ সঠিক সংখ্যা হতে হবে"), 400
            if delivery_charge < 0 or delivery_charge > 100000: return jsonify(error="ডেলিভারি চার্জ 0 থেকে 100000 টাকার মধ্যে হতে হবে"), 400
        # blank/null means no fee, even when a destination address was selected
    elif delivery_zone in delivery_keys:
        try: delivery_charge = max(0.0, min(100000.0, float(company_info.get(delivery_keys[delivery_zone], "0") or 0)))
        except (TypeError, ValueError): delivery_charge = 0.0
    total = round(max(0.0, subtotal - discount) + delivery_charge, 2)
    try: paid = max(0.0, float(j.get("paid", total)))
    except (TypeError, ValueError): paid = total
    method = str(j.get("method", "Cash")); method = method if method in _POS_METHODS else "Cash"
    cu = j.get("customer") if isinstance(j.get("customer"), dict) else {}
    cname, cphone, caddr = (str(cu.get(k, "")).strip()[:1200 if k == "address" else 200] for k in ("name", "phone", "address"))
    if not os.path.isfile(EXCEL_STOCK_FILE): return jsonify(error="স্টক Excel ফাইল পাওয়া যায়নি; আগে Stock Excel আপলোড করুন"), 400
    user = session.get("user", ""); now = datetime.now(DHAKA_TZ); tmp = None
    try:
        with EXCEL_LOCK:
            wb = load_workbook(EXCEL_STOCK_FILE)
            taken, problems, edits_plan = {}, [], []
            for l in lines:
                cells, avail = [], 0
                for ws in wb.worksheets:
                    headers = excel_headers(ws)
                    if not headers or not headers[0]: continue
                    col = next((i + 1 for i, h in enumerate(headers) if i > 0 and h and h.lower() not in SKIPH and normalize_stock_size(h).lower() == l["size"].lower()), None)
                    if not col: continue
                    for rn in range(2, ws.max_row + 1):
                        if excel_cell_text(ws.cell(rn, 1)).strip().lower() != l["code"]: continue
                        try: have = int(float(ws.cell(rn, col).value or 0))
                        except (TypeError, ValueError): have = 0
                        have -= taken.get((ws.title, rn, col), 0)
                        if have > 0: cells.append((ws, rn, col, have, headers[col - 1])); avail += have
                if avail < l["qty"]:
                    problems.append(f'{l["id"]} ({l["size"]}): স্টকে আছে {avail}টি, কার্টে {l["qty"]}টি'); continue
                need = l["qty"]
                for ws, rn, col, have, hname in cells:
                    if need <= 0: break
                    use = min(have, need); need -= use
                    old = int(float(ws.cell(rn, col).value or 0))
                    taken[(ws.title, rn, col)] = taken.get((ws.title, rn, col), 0) + use
                    edits_plan.append((ws, rn, col, l["code"], hname, old, use))
            if problems:
                wb.close(); return jsonify(error="স্টক যথেষ্ট নেই — " + "; ".join(problems)), 409
            prefix = f"POS-{now:%y%m%d}-"
            with db() as c:
                prior_invoices = [str(r[0] or "") for r in c.execute(
                    "SELECT invoice FROM pos_sales WHERE invoice LIKE ? UNION ALL SELECT invoice FROM pos_deleted_sales WHERE invoice LIKE ?",
                    (prefix + "%", prefix + "%")).fetchall()]
            used_numbers=[]
            for prior in prior_invoices:
                try: used_numbers.append(int(prior.rsplit("-", 1)[1]))
                except (ValueError, IndexError): pass
            n = max(used_numbers or [0]) + 1
            invoice = f"{prefix}{n:03d}"
            hist, running = [], {}
            for ws, rn, col, code, hname, old, use in edits_plan:
                k = (ws.title, rn, col); cur = running.get(k, old); new = cur - use; running[k] = new
                ws.cell(rn, col).value = new
                hist.append((ws.title, rn, excel_cell_text(ws.cell(rn, 1)).strip(), hname, cur, new, user, "POS বিক্রি " + invoice, time.time()))
            fd, tmp = tempfile.mkstemp(prefix="stock-pos-", suffix=".xlsx", dir=INSTANCE); os.close(fd)
            wb.save(tmp); wb.close()
            if time.time() - _POS_SNAP["t"] > 900:   # ব্যাকআপ ১৫ মিনিটে একবার, যাতে প্রতিটি বিক্রিতে ফাইল না জমে
                snapshot_excel_stock("before-pos", user); _POS_SNAP["t"] = time.time()
            os.replace(tmp, EXCEL_STOCK_FILE); tmp = None
            created = time.time()
            with db() as c:
                c.executemany("INSERT INTO excel_stock_history(sheet,row_number,product_code,size,old_qty,new_qty,username,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)", hist)
                c.execute("INSERT INTO excel_stock_batches(username,change_count,plus_units,minus_units,created_at) VALUES(?,?,?,?,?)", (user, len(hist), 0, sum(h[4] - h[5] for h in hist), created))
                c.execute("INSERT INTO pos_sales(invoice,username,customer_name,phone,address,subtotal,discount,total,paid,method,items,created_at,delivery_charge,delivery_zone,company_key,phone_digits) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                          (invoice, user, cname, cphone, caddr, subtotal, discount, total, paid, method,
                           json.dumps([dict(key=l["key"], id=l["id"], title=l["title"], size=l["size"], qty=l["qty"], price=l["price"]) for l in lines], ensure_ascii=False), created, delivery_charge, delivery_zone, company_key, re.sub(r"\D", "", cphone)))
                sale_row = c.execute("SELECT id FROM pos_sales WHERE invoice=?", (invoice,)).fetchone()
                sale_id = int(sale_row["id"])
                token_row = c.execute("SELECT token FROM pos_memo_tokens WHERE sale_id=?", (sale_id,)).fetchone()
                if not token_row:
                    c.execute("INSERT OR IGNORE INTO pos_memo_tokens(sale_id,token,created_at) VALUES(?,?,?)", (sale_id, secrets.token_urlsafe(32), time.time()))
                    token_row = c.execute("SELECT token FROM pos_memo_tokens WHERE sale_id=?", (sale_id,)).fetchone()
                memo_token = str(token_row["token"] if token_row else "")
        refresh_stock_from_excel()
        log_activity("pos_sale", invoice, f"৳{total} · {sum(l['qty'] for l in lines)} পিস")
        qr_url = request.host_url.rstrip("/") + "/memo/pos/" + str(sale_id) + "?t=" + memo_token
        qr_memo = {"qr_url":qr_url,"order_no":invoice,"date":now.strftime("%d %b %Y, %I:%M %p"),"kind":method,"delivery_zone":delivery_zone,
                   "name":cname,"phone":cphone,"address":caddr,
                   "items":[dict(title=l["title"],code=l["id"],variation=l["size"],qty=l["qty"],price=l["price"]) for l in lines],
                   "subtotal":subtotal,"discount":discount,"delivery":delivery_charge,"advance":paid,"total":total,"due":max(0.0,round(total-paid,2))}
        qr_data = _memo_qr_data_uri(qr_memo, company_info)
        return jsonify(ok=True, invoice=invoice, company_key=company_key, company_name=company_info.get("name") or COMPANY_NAMES.get(company_key, "Alien Tees"),
                       subtotal=subtotal, discount=discount, delivery_charge=delivery_charge, delivery_zone=delivery_zone, total=total, paid=paid, method=method,
                       change=round(max(0.0, paid - total), 2), due=round(max(0.0, total - paid), 2), created_at=created,
                       qr_data=qr_data,
                       customer=dict(name=cname, phone=cphone, address=caddr),
                       lines=[dict(id=l["id"], title=l["title"], size=l["size"], qty=l["qty"], price=l["price"]) for l in lines])
    except Exception as e:
        if tmp and os.path.exists(tmp): os.remove(tmp)
        return jsonify(error="বিক্রি সেভ করা যায়নি: " + str(e)), 400

@app.get("/api/pos/customer-history")
@perm_auth("pos_use")
def pos_customer_history():
    """Find prior POS and imported Order Sheet purchases by normalized Bangladesh phone digits."""
    phone = str(request.args.get("phone") or "").strip()[:80]
    digits = re.sub(r"\D", "", phone)
    if len(digits) < 3:
        return jsonify(ok=True, orders=[], customers=[])
    # Match the digits anywhere while also translating Bangladesh local and +880 prefixes.
    variants = [digits]
    if digits.startswith("0") and len(digits) >= 2: variants.append("880" + digits[1:])
    if digits.startswith("880") and len(digits) >= 4: variants.append("0" + digits[3:])
    if len(digits) >= 10: variants.append(digits[-10:])
    variants = list(dict.fromkeys(x for x in variants if len(x) >= 3))
    phone_expr = "REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(COALESCE(phone,''),' ',''),'-',''),'+',''),'.',''),'(',''),')',''),'/',''),char(9),''),char(10),'')"
    prefix_patterns = [x + "%" for x in variants]
    fallback_patterns = ["%" + x + "%" for x in variants]
    indexed_where = "(" + " OR ".join(["phone_digits LIKE ?"] * len(prefix_patterns)) + ")"
    legacy_where = "(" + " OR ".join([phone_expr + " LIKE ?"] * len(fallback_patterns)) + ")"
    out = []
    with db() as c:
        # First run a prefix-only query that can use phone_digits indexes.
        pos_rows = list(c.execute(f"SELECT * FROM pos_sales WHERE {indexed_where} ORDER BY created_at DESC LIMIT 100", prefix_patterns).fetchall())
        pos_seen = {int(r["id"]) for r in pos_rows}
        remaining = max(0, 100-len(pos_rows))
        if remaining:
            # Only un-migrated legacy rows need the slower punctuation-tolerant fallback.
            legacy_pos = c.execute(f"SELECT * FROM pos_sales WHERE COALESCE(phone_digits,'')='' AND {legacy_where} ORDER BY created_at DESC LIMIT ?", fallback_patterns+[remaining]).fetchall()
            pos_rows.extend(r for r in legacy_pos if int(r["id"]) not in pos_seen)
        pos_rows.sort(key=lambda r: float(r["created_at"] or 0), reverse=True)
        pos_rows = pos_rows[:100]
        for row in pos_rows:
            r = dict(row)
            items = _pos_items(r)
            out.append({"source":"pos", "no":str(r.get("invoice") or ("POS-"+str(r.get("id")))),
                        "date":datetime.fromtimestamp(float(r.get("created_at") or 0), DHAKA_TZ).strftime("%Y-%m-%d %H:%M") if r.get("created_at") else "",
                        "name":str(r.get("customer_name") or ""), "phone":str(r.get("phone") or ""), "address":str(r.get("address") or ""),
                        "items_text":", ".join(f"{x.get('title') or x.get('id') or 'Product'} {x.get('size') or ''} × {int(_num(x.get('qty')))}".strip() for x in items),
                        "total":round(_num(r.get("total")),2), "paid":round(_num(r.get("paid")),2),
                        "due":round(max(0,_num(r.get("total"))-_num(r.get("paid"))),2), "method":str(r.get("method") or ""),
                        "status":"Due" if _num(r.get("total"))-_num(r.get("paid"))>0.009 else "Paid", "_ts":float(r.get("created_at") or 0)})
        online_rows = list(c.execute(f"SELECT * FROM order_history WHERE {indexed_where} ORDER BY id DESC LIMIT 500", prefix_patterns).fetchall())
        online_seen = {int(r["id"]) for r in online_rows}
        remaining_online = max(0, 500-len(online_rows))
        if remaining_online:
            legacy_online = c.execute(f"SELECT * FROM order_history WHERE COALESCE(phone_digits,'')='' AND {legacy_where} ORDER BY id DESC LIMIT ?", fallback_patterns+[remaining_online]).fetchall()
            online_rows.extend(r for r in legacy_online if int(r["id"]) not in online_seen)
        online_rows.sort(key=lambda r: int(r["id"] or 0), reverse=True)
        online_rows = online_rows[:500]
        grouped = {}
        for row in online_rows:
            d = dict(row); no = _order_no(d, d.get("id")); cons = str(d.get("cons_id") or "").strip()
            key = "cons:"+cons if cons else "ord:"+no
            try:
                raw_order_date = str(d.get("order_date") or "").strip()
                _order_dt = None
                if raw_order_date:
                    try:
                        _order_dt = datetime.fromisoformat(raw_order_date[:10])
                    except ValueError:
                        for _fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                            try:
                                _order_dt = datetime.strptime(raw_order_date[:10], _fmt)
                                break
                            except ValueError:
                                continue
                order_ts = _order_dt.replace(tzinfo=DHAKA_TZ).timestamp() if _order_dt else 0.0
            except Exception: order_ts = 0.0
            price = _num(d.get("product_price")); delivery = _num(d.get("delivery_charge")); paid = _num(d.get("advance_paid"))
            g = grouped.get(key)
            if g is None:
                g = {"source":"online", "no":no, "date":str(d.get("order_date") or ""), "name":str(d.get("customer_name") or ""),
                     "phone":str(d.get("phone") or ""), "address":str(d.get("address") or ""), "_items":[], "_subtotal":0.0,
                     "delivery":delivery, "paid":paid, "method":str(d.get("payment_status") or ""),
                     "courier_status":str(d.get("delivery_status") or "Pending"), "_ts":order_ts,
                     "_collectible":_num(d.get("collectible_amount")) if str(d.get("collectible_amount") or "").strip() else None}
                grouped[key] = g
            else:
                g["_ts"] = max(float(g.get("_ts") or 0), order_ts)
            code = str(d.get("product_code") or ""); size = str(d.get("size") or "")
            g["_items"].append((code, size, price))
            g["_subtotal"] += price
        for g in grouped.values():
            total = round(g["_subtotal"] + g["delivery"],2)
            due = max(0.0, float(g["_collectible"]) if g["_collectible"] is not None else total-g["paid"])
            g["items_text"] = ", ".join((code+" "+size).strip() for code,size,_ in g["_items"] if code or size)
            g["total"] = round(total,2); g["due"] = round(due,2)
            g["status"] = g["courier_status"] or ("Due" if due>0.009 else "Paid")
            for k in ("_items","_subtotal","_collectible","courier_status"): g.pop(k,None)
            out.append(g)
    out.sort(key=lambda x: (float(x.get("_ts") or 0), str(x.get("date") or "")), reverse=True)
    # Build one autocomplete suggestion per stored phone. Clicking it fills the customer form.
    customers, seen_phones = [], set()
    for item in out:
        stored_phone = str(item.get("phone") or "").strip()
        phone_key = re.sub(r"\D", "", stored_phone)
        if not phone_key or phone_key in seen_phones: continue
        seen_phones.add(phone_key)
        customers.append({"name": item.get("name") or "Customer", "phone": stored_phone,
                          "address": item.get("address") or "", "order_no": item.get("no") or "",
                          "total": item.get("total", 0), "status": item.get("status") or ""})
        if len(customers) >= 8: break
    for item in out: item.pop("_ts", None)
    return jsonify(ok=True, query=digits, orders=out[:12], customers=customers)

# ===================== My Company / Memo / POS অর্ডার তালিকা ও রিপোর্ট =====================
# শুধু নতুন route ও settings key যোগ হয়েছে; আগের কোনো টেবিল/ডাটা বদলানো হয়নি।
COMPANY_DEFAULTS = {
    "name": "Alien Tees", "phone": "", "email": "", "website": "", "tin": "", "reg_no": "",
    "address": "", "company_note": "", "invoice_note": "This is computer generated invoice",
    "delivery_inside_dhaka": "0", "delivery_dhaka_suburbs": "0", "delivery_outside_dhaka": "0", "qr_mode": "text",
}
COMPANY_NAMES = {"alien_tees": "Alien Tees", "bodyfitters": "Bodyfitters"}

def _company_key(value=None):
    raw = str(value or "alien_tees").strip().casefold().replace(" ", "").replace("-", "").replace("_", "")
    return "bodyfitters" if raw in ("bodyfitters", "bodyfittars", "bodyfit", "body") else "alien_tees"

def _saved_company_profiles():
    try: saved = json.loads(get_setting("company_profiles") or "{}")
    except Exception: saved = {}
    if not isinstance(saved, dict): saved = {}
    try: legacy = json.loads(get_setting("company_profile") or "{}")
    except Exception: legacy = {}
    if not isinstance(legacy, dict): legacy = {}
    alien = dict(COMPANY_DEFAULTS)
    alien.update({k: str(v) for k, v in legacy.items() if k in COMPANY_DEFAULTS})
    body = dict(COMPANY_DEFAULTS); body["name"] = "Bodyfitters"
    for key, target in (("alien_tees", alien), ("bodyfitters", body)):
        val = saved.get(key)
        if isinstance(val, dict): target.update({k: str(v) for k, v in val.items() if k in COMPANY_DEFAULTS})
    return {"alien_tees": alien, "bodyfitters": body}

def company_profile(company_key="alien_tees"):
    key = _company_key(company_key)
    return dict(_saved_company_profiles()[key])

def _saved_company_logos():
    try: saved = json.loads(get_setting("company_logos") or "{}")
    except Exception: saved = {}
    if not isinstance(saved, dict): saved = {}
    # Keep the original logo as Alien Tees' logo for backwards compatibility.
    if "alien_tees" not in saved:
        try:
            legacy = json.loads(get_setting("company_logo") or "{}")
            if isinstance(legacy, dict) and legacy.get("data"): saved["alien_tees"] = legacy
        except Exception: pass
    return saved

def _company_logo(company_key="alien_tees"):
    try: d = _saved_company_logos().get(_company_key(company_key), {})
    except Exception: d = {}
    if isinstance(d, dict) and d.get("data"):
        try: return d.get("mime", "image/png"), base64.b64decode(d["data"])
        except Exception: return None
    return None

def _company_logo_version(company_key="alien_tees"):
    try:
        versions = json.loads(get_setting("company_logo_versions") or "{}")
        return int(float(versions.get(_company_key(company_key), 0) or 0))
    except Exception:
        return int(float(get_setting("company_logo_v") or 0))

def _erp_purchase_memo_template(company_key="alien_tees"):
    defaults = {"title": "PURCHASE MEMO", "subtitle": "Supplier invoice, received item quantities and payment summary", "footer": "Alien Tees · Purchase Management", "color": "#0D7552", "show_logo": True}
    try: legacy = json.loads(get_setting("erp_purchase_memo_template") or "{}")
    except Exception: legacy = {}
    try: saved = json.loads(get_setting("erp_purchase_memo_templates") or "{}")
    except Exception: saved = {}
    if not isinstance(legacy, dict): legacy = {}
    if not isinstance(saved, dict): saved = {}
    key = _company_key(company_key)
    out = dict(defaults)
    if key == "alien_tees": out.update({k:v for k,v in legacy.items() if k in defaults})
    cfg = saved.get(key)
    if isinstance(cfg, dict): out.update({k:v for k,v in cfg.items() if k in defaults})
    if key == "bodyfitters" and not isinstance(cfg, dict): out.update(title="BODYFITTERS PURCHASE MEMO", footer="Bodyfitters · Purchase Management")
    out["show_logo"] = bool(out["show_logo"])
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", str(out.get("color") or "")): out["color"] = defaults["color"]
    return out

def _sales_memo_template(company_key="alien_tees"):
    """Independent editable template for POS/online sale memos, migrated from older settings once."""
    key = _company_key(company_key)
    try: saved = json.loads(get_setting("erp_sales_memo_templates") or "{}")
    except Exception: saved = {}
    if not isinstance(saved, dict): saved = {}
    cfg = dict(_erp_purchase_memo_template(key))
    # Convert untouched purchase defaults to sensible sales-memo defaults without discarding old custom text.
    if key == "alien_tees" and str(cfg.get("title") or "").strip().upper() == "PURCHASE MEMO":
        cfg["title"] = "SALES MEMO"
    elif key == "bodyfitters" and str(cfg.get("title") or "").strip().upper() == "BODYFITTERS PURCHASE MEMO":
        cfg["title"] = "BODYFITTERS SALES MEMO"
    if str(cfg.get("subtitle") or "").strip() == "Supplier invoice, received item quantities and payment summary":
        cfg["subtitle"] = "Customer invoice, itemized sale and payment summary"
    if str(cfg.get("footer") or "").strip() in ("Alien Tees · Purchase Management", "Bodyfitters · Purchase Management"):
        co = company_profile(key)
        cfg["footer"] = str(co.get("invoice_note") or (co.get("name") or COMPANY_NAMES[key]))
    custom = saved.get(key)
    if isinstance(custom, dict): cfg.update({k:v for k,v in custom.items() if k in ("title", "subtitle", "footer", "color", "show_logo")})
    cfg["show_logo"] = bool(cfg.get("show_logo", True))
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", str(cfg.get("color") or "")): cfg["color"] = "#0D7552"
    return cfg

@app.get("/api/company")
def company_get():
    key = _company_key(request.args.get("company"))
    lg = _company_logo(key)
    return jsonify(company=company_profile(key), company_key=key, company_name=COMPANY_NAMES[key],
                   companies=[{"key":k,"name":v} for k,v in COMPANY_NAMES.items()],
                   has_logo=bool(lg), logo_v=_company_logo_version(key),
                   purchase_memo_template=_erp_purchase_memo_template(key),
                   sales_memo_template=_sales_memo_template(key))

@app.get("/api/company/logo")
def company_logo():
    key = _company_key(request.args.get("company"))
    lg = _company_logo(key)
    if not lg: return jsonify(error="no_logo"), 404
    r = Response(lg[1], mimetype=lg[0]); r.headers["Cache-Control"] = "private, max-age=300"; return r

@app.post("/api/admin/company")
@admin_only
def company_save():
    f = request.form if request.form else (request.get_json(silent=True) or {})
    key = _company_key(f.get("company") or request.args.get("company"))
    cur = company_profile(key)
    for k, lim in (("name", 120), ("phone", 60), ("email", 120), ("website", 160), ("tin", 80), ("reg_no", 80), ("address", 400), ("company_note", 1000), ("invoice_note", 200), ("qr_mode", 10)):
        if k in f: cur[k] = str(f.get(k) or "").strip()[:lim]
    for k in ("delivery_inside_dhaka", "delivery_dhaka_suburbs", "delivery_outside_dhaka"):
        if k in f:
            raw = str(f.get(k) or "0").replace(",", "").strip()
            try: amount = float(raw or 0)
            except (TypeError, ValueError): return jsonify(error="Delivery charges must be valid numbers."), 400
            if amount < 0 or amount > 100000: return jsonify(error="Delivery charges must be between 0 and 100000 BDT."), 400
            cur[k] = (f"{amount:.2f}".rstrip("0").rstrip(".") if amount else "0")
    cur["qr_mode"] = "pdf" if str(cur.get("qr_mode", "text")).lower() == "pdf" else "text"
    if not cur["name"]: return jsonify(error="Company name দিতে হবে"), 400
    # Company editor controls sales/POS memo design separately; existing Purchase PDF settings remain untouched.
    memo_cfg = _sales_memo_template(key)
    for field, lim in (("title", 100), ("subtitle", 220), ("footer", 500)):
        form_key = "purchase_memo_" + field  # keep compatibility with the existing form field names
        if form_key in f: memo_cfg[field] = str(f.get(form_key) or "").strip()[:lim]
    memo_color = str(f.get("purchase_memo_color", memo_cfg["color"]) or memo_cfg["color"]).strip()
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", memo_color): return jsonify(error="মেমোর রং সঠিকভাবে নির্বাচন করুন"), 400
    memo_cfg["color"] = memo_color.upper()
    if "purchase_memo_show_logo" in f: memo_cfg["show_logo"] = str(f.get("purchase_memo_show_logo", "")).lower() in ("1", "true", "yes", "on")
    profiles = _saved_company_profiles(); profiles[key] = cur
    set_setting("company_profiles", json.dumps(profiles, ensure_ascii=False))
    if key == "alien_tees": set_setting("company_profile", json.dumps(cur, ensure_ascii=False))
    try: sales_templates = json.loads(get_setting("erp_sales_memo_templates") or "{}")
    except Exception: sales_templates = {}
    if not isinstance(sales_templates, dict): sales_templates = {}
    sales_templates[key] = memo_cfg
    set_setting("erp_sales_memo_templates", json.dumps(sales_templates, ensure_ascii=False))

    logos = _saved_company_logos(); versions = {}
    try: versions = json.loads(get_setting("company_logo_versions") or "{}")
    except Exception: versions = {}
    if not isinstance(versions, dict): versions = {}
    if str(f.get("remove_logo", "")).lower() in ("1", "true", "yes", "on"):
        logos.pop(key, None); versions[key] = int(time.time())
        if key == "alien_tees": set_setting("company_logo", "")
    up = request.files.get("logo")
    if up and up.filename:
        raw = up.read()
        if len(raw) > 2 * 1024 * 1024: return jsonify(error="Logo সর্বোচ্চ 2MB হতে পারবে"), 400
        try:
            im = Image.open(io.BytesIO(raw)); im.verify(); im = Image.open(io.BytesIO(raw))
            if (im.format or "").upper() not in ("PNG", "JPEG", "WEBP"): return jsonify(error="Logo শুধু PNG, JPG বা WebP হতে পারবে"), 400
            im = im.convert("RGBA") if im.mode in ("P", "LA", "RGBA") else im.convert("RGB")
            im.thumbnail((700, 700)); buf = io.BytesIO(); im.save(buf, "PNG")
        except Exception: return jsonify(error="Logo ছবিটি পড়া যায়নি"), 400
        logos[key] = {"mime":"image/png", "data":base64.b64encode(buf.getvalue()).decode()}
        versions[key] = int(time.time())
        if key == "alien_tees": set_setting("company_logo", json.dumps(logos[key]))
    set_setting("company_logos", json.dumps(logos, ensure_ascii=False))
    set_setting("company_logo_versions", json.dumps(versions, ensure_ascii=False))
    set_setting("company_logo_v", str(time.time()))
    log_activity("company_profile", key, "updated")
    return jsonify(ok=True, company=cur, company_key=key)

# ---------- সাধারণ helper ----------
# ---------- সাধারণ helper ----------
def _num(v):
    try:
        raw = re.sub(r"[^0-9.\-]", "", str(v if v is not None else "").replace(",", ""))
        return float(raw) if raw not in ("", "-", ".", "-.") else 0.0
    except Exception: return 0.0

def _norm_key(k): return re.sub(r"[^a-z0-9]", "", str(k).lower())

def _prod_maps():
    by_key, by_code = {}, {}
    for p in (C.get("products") or []):
        by_key[p["store"] + "|" + p["id"]] = p
        did = (display_product_id(p) or "").lower()
        if did and did not in by_code: by_code[did] = p
    return by_key, by_code

def _day_bounds(frm, to):
    s = e = None
    try:
        if frm: s = datetime.strptime(frm, "%Y-%m-%d").replace(tzinfo=DHAKA_TZ).timestamp()
        if to: e = (datetime.strptime(to, "%Y-%m-%d").replace(tzinfo=DHAKA_TZ) + timedelta(days=1)).timestamp()
    except ValueError: pass
    return s, e

def _valid_day(v):
    v = str(v or "").strip()
    return v if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v) else ""

def _order_no(rowd, rid):
    try: rj = json.loads(rowd.get("row_json") or "{}")
    except Exception: rj = {}
    if isinstance(rj, dict):
        for k, v in rj.items():
            if _norm_key(k) in ("orderno", "orderid", "ordernumber", "invoice", "invoiceno") and str(v or "").strip():
                return str(v).strip()
    return f"ORD-{rid}"

def _row_json_get(rowd, *names):
    try: rj = json.loads(rowd.get("row_json") or "{}")
    except Exception: rj = {}
    want = {_norm_key(n) for n in names}
    if isinstance(rj, dict):
        for k, v in rj.items():
            if _norm_key(k) in want and str(v or "").strip(): return str(v).strip()
    return ""

def _pos_items(row):
    try: it = json.loads(row["items"] or "[]")
    except Exception: it = []
    return [x for x in it if isinstance(x, dict)]

def _norm_pos(row, by_key):
    items = _pos_items(row)
    pieces = sum(int(_num(x.get("qty"))) for x in items)
    first = by_key.get(str(items[0].get("key", ""))) if items else None
    total, paid = float(row["total"] or 0), float(row["paid"] or 0)
    return dict(id=row["id"], source="pos", no=row["invoice"], ts=row["created_at"],
                date=datetime.fromtimestamp(row["created_at"], DHAKA_TZ).strftime("%Y-%m-%d %H:%M"),
                customer=row["customer_name"] or "", phone=row["phone"] or "", address=row["address"] or "",
                items_text=", ".join(f'{x.get("id","")} {x.get("size","")}×{int(_num(x.get("qty")))}' for x in items),
                pieces=pieces, subtotal=float(row["subtotal"] or 0), discount=float(row["discount"] or 0), delivery=float(row["delivery_charge"] or 0),
                delivery_zone=row["delivery_zone"] or "",
                total=total, paid=paid, due=round(max(0.0, total - paid), 2), method=row["method"] or "",
                status="Due" if total - paid > 0.009 else "Paid", user=row["username"] or "",
                img=(first or {}).get("img", ""), cons_id="",
                _lines=[(str(x.get("id", "")), str(x.get("size", "")), int(_num(x.get("qty"))), float(_num(x.get("price")))) for x in items])

def _norm_online(row, by_code):
    d = dict(row)
    code = str(d.get("product_code") or "").strip(); price = _num(d.get("product_price"))
    dc, adv = _num(d.get("delivery_charge")), _num(d.get("advance_paid"))
    bill = price + dc
    due = _num(d.get("collectible_amount")) if str(d.get("collectible_amount") or "").strip() else max(0.0, bill - adv)
    p = by_code.get(code.lower()) if code else None
    return dict(id=d["id"], source="online", no=_order_no(d, d["id"]), ts=0, date=str(d.get("order_date") or ""),
                customer=d.get("customer_name") or "", phone=d.get("phone") or "", address=d.get("address") or "",
                items_text=(code + " " + str(d.get("size") or "")).strip(), pieces=1, subtotal=price, discount=0.0,
                delivery=dc, total=round(bill, 2), paid=adv, due=round(due, 2),
                method=d.get("payment_status") or "", status=d.get("delivery_status") or "Pending",
                user=d.get("assigned_to") or "", img=(p or {}).get("img", d.get("image_url") or "") if p else (d.get("image_url") or ""),
                cons_id=d.get("cons_id") or "",
                _lines=[(code, str(d.get("size") or ""), 1, price)], _raw=d)

def _query_orders(a, ids=None, nos=None, source=None):
    source = source or ("online" if str(a.get("source", "pos")) == "online" else "pos")
    frm, to = _valid_day(a.get("from")), _valid_day(a.get("to"))
    q = str(a.get("q", "")).strip().lower()[:80]
    user = str(a.get("user", "")).strip()[:60]; method = str(a.get("method", "")).strip()[:40]; status = str(a.get("status", "")).strip()[:40]
    by_key, by_code = _prod_maps(); rows = []
    with db() as c:
        if source == "pos":
            where, args = [], []
            s, e = _day_bounds(frm, to)
            if s is not None: where.append("created_at>=?"); args.append(s)
            if e is not None: where.append("created_at<?"); args.append(e)
            if user: where.append("username=?"); args.append(user)
            if method: where.append("method=?"); args.append(method)
            if ids: where.append("id IN (%s)" % ",".join("?" * len(ids))); args += list(ids)
            if nos: where.append("invoice IN (%s)" % ",".join("?" * len(nos))); args += list(nos)
            sql = "SELECT * FROM pos_sales" + (" WHERE " + " AND ".join(where) if where else "") + " ORDER BY created_at DESC, id DESC"
            rows = [_norm_pos(r, by_key) for r in c.execute(sql, args).fetchall()]
            if status: rows = [r for r in rows if r["status"].lower() == status.lower()]
        else:
            where, args = [], []
            if frm: where.append("order_date>=?"); args.append(frm)
            if to: where.append("order_date<=?"); args.append(to)
            if user: where.append("assigned_to=?"); args.append(user)
            if method: where.append("payment_status=?"); args.append(method)
            if status: where.append("LOWER(COALESCE(delivery_status,''))=?"); args.append(status.lower())
            if ids: where.append("id IN (%s)" % ",".join("?" * len(ids))); args += list(ids)
            sql = "SELECT * FROM order_history" + (" WHERE " + " AND ".join(where) if where else "") + " ORDER BY order_date DESC, id DESC"
            rows = [_norm_online(r, by_code) for r in c.execute(sql, args).fetchall()]
    if q:
        rows = [r for r in rows if q in " ".join(str(r.get(k) or "") for k in ("no", "customer", "phone", "address", "items_text", "cons_id", "user", "method")).lower()
                or (len(re.sub(r"\D", "", q)) >= 6 and re.sub(r"\D", "", q) in re.sub(r"\D", "", str(r.get("phone"))))]
    return source, rows

def _summarize(rows):
    def grp(keyf):
        g = {}
        for r in rows:
            k = keyf(r) or "—"; x = g.setdefault(k, dict(name=k, orders=0, pieces=0, total=0.0)); x["orders"] += 1; x["pieces"] += r["pieces"]; x["total"] += r["total"]
        return sorted(g.values(), key=lambda x: -x["total"])
    prod = {}
    for r in rows:
        for code, size, qty, price in r["_lines"]:
            x = prod.setdefault(code or "—", dict(code=code or "—", pieces=0, total=0.0)); x["pieces"] += qty; x["total"] += qty * price
    days = {}
    for r in rows:
        k = str(r["date"])[:10]; x = days.setdefault(k, dict(day=k, orders=0, total=0.0)); x["orders"] += 1; x["total"] += r["total"]
    n = len(rows); tot = sum(r["total"] for r in rows)
    return dict(orders=n, pieces=sum(r["pieces"] for r in rows), subtotal=round(sum(r["subtotal"] for r in rows), 2),
                discount=round(sum(r["discount"] for r in rows), 2), delivery=round(sum(r["delivery"] for r in rows), 2),
                total=round(tot, 2), paid=round(sum(r["paid"] for r in rows), 2), due=round(sum(r["due"] for r in rows), 2),
                avg=round(tot / n, 2) if n else 0.0,
                by_method=[dict(x, total=round(x["total"], 2)) for x in grp(lambda r: r["method"])],
                by_status=[dict(x, total=round(x["total"], 2)) for x in grp(lambda r: r["status"])],
                by_user=[dict(x, total=round(x["total"], 2)) for x in grp(lambda r: r["user"])],
                top_products=[dict(x, total=round(x["total"], 2)) for x in sorted(prod.values(), key=lambda x: -x["pieces"])[:10]],
                by_day=[dict(x, total=round(x["total"], 2)) for x in sorted(days.values(), key=lambda x: x["day"])][-62:])

def _public_row(r): return {k: v for k, v in r.items() if not k.startswith("_")}

def _pos_orders_access_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if session.get("user") and (is_admin_level() or has_perm("pos_orders") or has_perm("pos_delete")):
            return fn(*args, **kwargs)
        return jsonify(error="permission_denied"), 403
    return wrapped

@app.get("/api/admin/pos-orders")
@_pos_orders_access_required
def pos_orders_list():
    if str(request.args.get("source", "pos")) == "online" and not (is_admin_level() or has_perm("pos_orders")):
        return jsonify(error="permission_denied"), 403
    source, rows = _query_orders(request.args)
    try: off = max(0, int(request.args.get("offset", 0) or 0))
    except ValueError: off = 0
    page = rows[off:off + 500]
    with db() as c:
        users = sorted({r[0] for r in c.execute("SELECT DISTINCT username FROM pos_sales WHERE username<>''")} if source == "pos" else {r[0] for r in c.execute("SELECT DISTINCT assigned_to FROM order_history WHERE COALESCE(assigned_to,'')<>''")}, key=str.lower)
        methods = sorted({r[0] for r in c.execute("SELECT DISTINCT method FROM pos_sales WHERE method<>''")} if source == "pos" else {r[0] for r in c.execute("SELECT DISTINCT payment_status FROM order_history WHERE COALESCE(payment_status,'')<>''")}, key=str.lower)
        statuses = ["Paid", "Due"] if source == "pos" else sorted({r[0] for r in c.execute("SELECT DISTINCT delivery_status FROM order_history WHERE COALESCE(delivery_status,'')<>''")}, key=str.lower)
    r = jsonify(source=source, rows=[_public_row(x) for x in page], total=len(rows), offset=off, summary=_summarize(rows),
                users=users, methods=methods, statuses=statuses)
    r.headers["Cache-Control"] = "no-store"; return r



@app.post("/api/admin/pos-orders/delete")
@perm_auth("pos_delete")
def pos_orders_delete():
    """Delete POS sale(s) only after restoring the exact recorded stock movement.

    Sales are also copied to pos_deleted_sales for audit/recovery. If the Excel stock
    workbook no longer matches the recorded history, the operation fails safely.
    """
    j = request.get_json(silent=True) or {}
    mode = str(j.get("mode") or "selected").strip().lower()
    reason = str(j.get("reason") or "").strip()[:500]
    if len(reason) < 3:
        return jsonify(ok=False, error="মুছে ফেলার কারণ অন্তত ৩ অক্ষরে লিখুন।"), 400
    if not os.path.isfile(EXCEL_STOCK_FILE):
        return jsonify(ok=False, error="Stock Excel পাওয়া যায়নি; স্টক নিরাপদে ফেরানো সম্ভব নয়।"), 409

    filters = j.get("filters") if isinstance(j.get("filters"), dict) else {}
    filters = dict(filters)
    filters["source"] = "pos"
    if mode == "all_filtered":
        _, matches = _query_orders(filters, source="pos")
        target_ids = [int(r["id"]) for r in matches]
    elif mode == "selected":
        raw_ids = j.get("ids") if isinstance(j.get("ids"), list) else []
        target_ids=[]
        for value in raw_ids[:2000]:
            try:
                rid=int(value)
                if rid>0 and rid not in target_ids: target_ids.append(rid)
            except (TypeError, ValueError): pass
    else:
        return jsonify(ok=False, error="Delete mode সঠিক নয়।"), 400
    if not target_ids:
        return jsonify(ok=False, error="মুছতে POS বিক্রি সিলেক্ট করুন।"), 400
    if len(target_ids) > 2000:
        return jsonify(ok=False, error="একবারে সর্বোচ্চ ২০০০টি POS বিক্রি মুছতে পারবেন। তারিখ/ফিল্টার দিয়ে ভাগ করুন।"), 413

    with db() as c:
        placeholders=",".join("?" for _ in target_ids)
        sale_rows=[dict(r) for r in c.execute(f"SELECT * FROM pos_sales WHERE id IN ({placeholders}) ORDER BY id", target_ids).fetchall()]
    if not sale_rows:
        return jsonify(ok=False, error="সিলেক্ট করা POS বিক্রি পাওয়া যায়নি।"), 404

    username = str(session.get("user") or "unknown")
    created_at = time.time()
    try:
        # Prepare the complete inverse stock plan in memory before modifying the workbook.
        per_invoice_cell = {}  # (invoice, sheet, row, col, product_code, size) -> units to restore
        invoice_units = {}
        with EXCEL_LOCK:
            wb = load_workbook(EXCEL_STOCK_FILE)
            try:
                with db() as c:
                    for sale in sale_rows:
                        invoice=str(sale.get("invoice") or "")
                        hrows=[dict(r) for r in c.execute(
                            "SELECT * FROM excel_stock_history WHERE note=? AND old_qty>new_qty ORDER BY id",
                            ("POS বিক্রি " + invoice,)).fetchall()]
                        try: items=json.loads(sale.get("items") or "[]")
                        except Exception: items=[]
                        sold_units=sum(max(0,int(_num(x.get("qty")))) for x in items if isinstance(x,dict))
                        history_units=sum(max(0,int(_num(h.get("old_qty"))-_num(h.get("new_qty")))) for h in hrows)
                        if not hrows or history_units != sold_units:
                            raise ValueError(f"{invoice}: পুরোনো Stock History সম্পূর্ণ নয় (বিক্রি {sold_units}, history {history_units})। আগে Stock Reconciliation করে নিন; কোনো রেকর্ড মুছা হয়নি।")
                        invoice_units[invoice]=history_units
                        for h in hrows:
                            sheet=str(h.get("sheet") or "")
                            if sheet not in wb.sheetnames:
                                raise ValueError(f"{invoice}: Stock Sheet '{sheet}' পাওয়া যায়নি; delete বাতিল করা হয়েছে।")
                            ws=wb[sheet]
                            headers=excel_headers(ws)
                            size=normalize_stock_size(h.get("size") or "")
                            col=next((i+1 for i,label in enumerate(headers[1:],start=1)
                                      if normalize_stock_size(label).casefold()==size.casefold()),None)
                            if not col:
                                raise ValueError(f"{invoice}: Size '{size}' sheet '{sheet}'-এ নেই; delete বাতিল করা হয়েছে।")
                            code=str(h.get("product_code") or "").strip().casefold()
                            rowno=int(h.get("row_number") or 0)
                            target_row=None
                            if 2 <= rowno <= ws.max_row and excel_cell_text(ws.cell(rowno,1)).strip().casefold()==code:
                                target_row=rowno
                            else:
                                candidates=[rn for rn in range(2,ws.max_row+1)
                                            if excel_cell_text(ws.cell(rn,1)).strip().casefold()==code]
                                if len(candidates)==1: target_row=candidates[0]
                                elif not candidates:
                                    raise ValueError(f"{invoice}: product code '{code}' sheet '{sheet}'-এ নেই; delete বাতিল করা হয়েছে।")
                                else:
                                    raise ValueError(f"{invoice}: product code '{code}' একাধিক row-তে আছে এবং row বদলেছে; delete নিরাপদ নয়।")
                            try: qty_to_restore=max(0,int(_num(h.get("old_qty"))-_num(h.get("new_qty"))))
                            except Exception: qty_to_restore=0
                            key=(invoice,sheet,target_row,col,excel_cell_text(ws.cell(target_row,1)).strip(),size)
                            per_invoice_cell[key]=per_invoice_cell.get(key,0)+qty_to_restore

                # Back up the workbook before editing and write the stock additions atomically.
                backup_name=snapshot_excel_stock("before-pos-delete", username)
                if not backup_name:
                    raise ValueError("Stock Excel backup তৈরি করা যায়নি; delete বাতিল করা হয়েছে।")
                grouped = {}
                for (invoice,sheet,rowno,col,code,size),qty in per_invoice_cell.items():
                    ck=(sheet,rowno,col)
                    grouped[ck]=grouped.get(ck,0)+qty
                cell_original={}
                for (sheet,rowno,col),inc in grouped.items():
                    ws=wb[sheet]
                    try: oldqty=int(float(ws.cell(rowno,col).value or 0))
                    except (TypeError,ValueError): oldqty=0
                    cell_original[(sheet,rowno,col)]=oldqty
                    ws.cell(rowno,col).value=oldqty+inc
                os.makedirs(INSTANCE,exist_ok=True)
                fd,tmp=tempfile.mkstemp(prefix="stock-pos-delete-",suffix=".xlsx",dir=INSTANCE);os.close(fd)
                wb.save(tmp)
                os.replace(tmp,EXCEL_STOCK_FILE);tmp=None
            finally:
                try: wb.close()
                except Exception: pass

        try:
            with db() as c:
                # Preserve deleted transaction details for accountability and audit.
                for sale in sale_rows:
                    c.execute("INSERT INTO pos_deleted_sales(original_sale_id,invoice,deleted_by,deleted_at,reason,sale_json,restored_units) VALUES(?,?,?,?,?,?,?)",
                              (int(sale["id"]),str(sale.get("invoice") or ""),username,created_at,reason,
                               json.dumps(sale,ensure_ascii=False,default=str),int(invoice_units.get(str(sale.get("invoice") or ""),0))))
                # Log the reverse stock movement with accurate old/new values.
                running=dict(cell_original)
                reversal_rows=[]
                for (invoice,sheet,rowno,col,code,size),qty in sorted(per_invoice_cell.items()):
                    old=running[(sheet,rowno,col)];new=old+qty;running[(sheet,rowno,col)]=new
                    reversal_rows.append((sheet,rowno,code,size,old,new,username,"POS DELETE reversal "+invoice+" · "+reason,created_at))
                if reversal_rows:
                    c.executemany("INSERT INTO excel_stock_history(sheet,row_number,product_code,size,old_qty,new_qty,username,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)",reversal_rows)
                total_restored=sum(invoice_units.values())
                c.execute("INSERT INTO excel_stock_batches(username,change_count,plus_units,minus_units,created_at) VALUES(?,?,?,?,?)",
                          (username,len(reversal_rows),total_restored,0,created_at))
                sale_ids=[int(s["id"]) for s in sale_rows]
                ph=",".join("?" for _ in sale_ids)
                c.execute(f"DELETE FROM pos_memo_tokens WHERE sale_id IN ({ph})",sale_ids)
                c.execute(f"DELETE FROM pos_sales WHERE id IN ({ph})",sale_ids)
        except Exception:
            # Database transaction failed after workbook update: restore the backed-up workbook.
            try:
                backup_path=os.path.join(EXCEL_BACKUP_DIR,backup_name)
                if os.path.isfile(backup_path): shutil.copy2(backup_path,EXCEL_STOCK_FILE)
            except Exception as rollback_exc: print("POS delete stock rollback failed:",rollback_exc)
            refresh_stock_from_excel()
            raise
        refresh_stock_from_excel()
        log_activity("pos_delete", ",".join(str(s.get("invoice") or "") for s in sale_rows[:20]),
                     f"deleted={len(sale_rows)}; restored_units={sum(invoice_units.values())}; by={username}; reason={reason}")
        return jsonify(ok=True, deleted=len(sale_rows), restored_units=sum(invoice_units.values()),
                       message=f"{len(sale_rows)}টি POS বিক্রি মুছে দেওয়া হয়েছে এবং {sum(invoice_units.values())} পিস stock ফেরত যোগ হয়েছে। Audit record রাখা হয়েছে।")
    except Exception as exc:
        print("POS sales delete failed:",repr(exc))
        return jsonify(ok=False,error=str(exc)[:600] or "POS বিক্রি মুছা যায়নি; কোনো রেকর্ড পরিবর্তন হয়নি।"),409

@app.get("/api/admin/pos-orders/export.xlsx")
@_pos_orders_access_required
def pos_orders_export():
    if str(request.args.get("source", "pos")) == "online" and not (is_admin_level() or has_perm("pos_orders")):
        return jsonify(error="permission_denied"), 403
    ids = [int(x) for x in re.findall(r"\d+", request.args.get("ids", ""))][:2000]
    source, rows = _query_orders(request.args, ids=ids or None)
    sm = _summarize(rows)
    wb = Workbook(); ws = wb.active; ws.title = "Summary"
    frm, to = _valid_day(request.args.get("from")), _valid_day(request.args.get("to"))
    ws.append(["Report", "POS বিক্রি" if source == "pos" else "Online / Courier অর্ডার"]); ws.append(["From", frm or "সব"]); ws.append(["To", to or "সব"])
    ws.append(["Selected only", "Yes" if ids else "No"]); ws.append([])
    for k, lab in (("orders", "Orders"), ("pieces", "Pieces"), ("subtotal", "Product total"), ("discount", "Discount"), ("delivery", "Delivery charge"), ("total", "Total"), ("paid", "Paid / Advance"), ("due", "Due / Collectible"), ("avg", "Average order")):
        ws.append([lab, sm[k]])
    ol = wb.create_sheet("Orders")
    ol.append(["Order No", "Date", "Customer", "Phone", "Address", "Products", "Pieces", "Product total", "Discount", "Delivery", "Total", "Paid / Advance", "Due / Collectible", "Payment", "Status", "User / Assignee", "Consignment"])
    for r in rows: ol.append([r["no"], r["date"], r["customer"], r["phone"], r["address"], r["items_text"], r["pieces"], r["subtotal"], r["discount"], r["delivery"], r["total"], r["paid"], r["due"], r["method"], r["status"], r["user"], r["cons_id"]])
    it = wb.create_sheet("Items"); it.append(["Order No", "Date", "Customer", "Phone", "Product code", "Size", "Qty", "Price", "Line total"])
    for r in rows:
        for code, size, qty, price in r["_lines"]: it.append([r["no"], r["date"], r["customer"], r["phone"], code, size, qty, price, round(qty * price, 2)])
    def sheet(title, head, data):
        w = wb.create_sheet(title); w.append(head)
        for d in data: w.append(d)
    sheet("By Product", ["Product code", "Pieces", "Amount"], [[x["code"], x["pieces"], x["total"]] for x in _summarize_all_products(rows)])
    sheet("By Day", ["Day", "Orders", "Amount"], [[x["day"], x["orders"], x["total"]] for x in _summarize_days(rows)])
    sheet("By Payment", ["Payment", "Orders", "Pieces", "Amount"], [[x["name"], x["orders"], x["pieces"], x["total"]] for x in sm["by_method"]])
    sheet("By Status", ["Status", "Orders", "Pieces", "Amount"], [[x["name"], x["orders"], x["pieces"], x["total"]] for x in sm["by_status"]])
    sheet("By User", ["User / Assignee", "Orders", "Pieces", "Amount"], [[x["name"], x["orders"], x["pieces"], x["total"]] for x in sm["by_user"]])
    from openpyxl.styles import Font, PatternFill
    for w in wb.worksheets:
        for cell in w[1 if w.title != "Summary" else 1]: cell.font = Font(bold=True)
        if w.title != "Summary":
            for cell in w[1]: cell.fill = PatternFill("solid", fgColor="DDEFE5")
        for col in w.columns:
            width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            w.column_dimensions[col[0].column_letter].width = min(48, max(10, width + 2))
        w.freeze_panes = "A2" if w.title != "Summary" else None
    out = io.BytesIO(); wb.save(out); out.seek(0)
    log_activity("pos_orders_export", source, f"{len(rows)} orders {frm}..{to}")
    nm = "POS" if source == "pos" else "Online-Orders"
    return send_file(out, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True,
                     download_name=f"Alien-Tees-{nm}-Report{('-' + frm) if frm else ''}{('-to-' + to) if to else ''}.xlsx")

def _summarize_all_products(rows):
    prod = {}
    for r in rows:
        for code, size, qty, price in r["_lines"]:
            x = prod.setdefault(code or "—", dict(code=code or "—", pieces=0, total=0.0)); x["pieces"] += qty; x["total"] = round(x["total"] + qty * price, 2)
    return sorted(prod.values(), key=lambda x: -x["pieces"])

def _summarize_days(rows):
    days = {}
    for r in rows:
        k = str(r["date"])[:10]; x = days.setdefault(k, dict(day=k, orders=0, total=0.0)); x["orders"] += 1; x["total"] = round(x["total"] + r["total"], 2)
    return sorted(days.values(), key=lambda x: x["day"])

def _memo_qr_payloads(memo, company):
    """Return exactly one compact, readable plain-text memo QR payload.

    A single QR has a finite capacity (about 2.9 KB in byte mode at L
    correction). Keep labels compact while preserving customer, item and
    payment details. Never embed a URL or split one memo across many QR codes.
    """
    def clean(value, limit=220):
        value = re.sub(r"[\r\n\t]+", " ", str(value or "")).strip()
        value = re.sub(r"(?i)\b(?:https?://|www\.)\S+", "[link omitted]", value)
        value = re.sub(r"\s{2,}", " ", value)
        return value[:limit]

    def cash(value):
        try: return f"{float(value or 0):.2f}"
        except (TypeError, ValueError): return "0.00"

    template = memo.get("memo_template") if isinstance(memo.get("memo_template"), dict) else {}
    co_name = clean(company.get("name") or "Company", 80)
    title = clean(memo.get("title") or template.get("title") or "SALES MEMO", 80)
    order_no = clean(memo.get("number") or memo.get("order_no") or "—", 60)
    customer = clean(memo.get("customer") or memo.get("name") or "Walk-in Customer", 80)
    phone = clean(memo.get("phone") or "—", 40)
    address = clean(memo.get("address") or "—", 120)
    date = clean(memo.get("date") or "—", 40)
    method = clean(memo.get("method") or memo.get("kind") or "—", 40)
    out = [f"{co_name} | {title}", f"Invoice: {order_no} | Date: {date}",
           f"Customer: {customer} | Phone: {phone}", f"Address: {address}", f"Payment: {method}"]
    seller=clean(memo.get("seller") or memo.get("created_by") or "", 80)
    if seller: out.append(f"Created by: {seller}")
    for label, key in (("Co phone", "phone"), ("Co address", "address"), ("Email", "email"), ("TIN", "tin"), ("Reg", "reg_no")):
        val = clean(company.get(key), 100)
        if val: out.append(f"{label}: {val}")
    items = memo.get("items") or []
    out.append("ITEMS:")
    for i, it in enumerate(items, 1):
        if not isinstance(it, dict): continue
        name = clean(it.get("title") or it.get("name") or it.get("code") or "Product", 100)
        code = clean(it.get("code") or it.get("id") or "", 35)
        size = clean(it.get("variation") or it.get("size") or "—", 45)
        try: qty = max(0, int(_num(it.get("qty")) or 0))
        except Exception: qty = 0
        price = _num(it.get("price")); total = qty * price
        row = f"{i}. {name}"
        if code and code != name: row += f" [{code}]"
        row += f" | Size:{size} | Qty:{qty} | Rate:{cash(price)} | Line:{cash(total)}"
        out.append(row)
    subtotal = _num(memo.get("subtotal")); discount = _num(memo.get("discount"))
    delivery = _num(memo.get("delivery")); total = _num(memo.get("total"))
    paid = _num(memo.get("advance", memo.get("paid", 0)))
    due = _num(memo.get("due", max(0, total-paid)))
    out.extend([f"Subtotal:{cash(subtotal)} | Discount:{cash(discount)} | Delivery:{cash(delivery)}",
                f"TOTAL:{cash(total)} | Paid:{cash(paid)} | DUE:{cash(due)}"])
    zone = clean(memo.get("delivery_zone_label") or memo.get("delivery_zone") or "", 50)
    if zone: out.append(f"Delivery area: {zone}")
    note = clean(memo.get("notes") or memo.get("customer_note") or memo.get("office_note") or company.get("company_note") or "", 180)
    footer = clean(memo.get("footer") or template.get("footer") or company.get("invoice_note") or "Thank you for your business", 120)
    if note: out.append(f"Note: {note}")
    if footer: out.append(footer)
    payload = "\n".join(x for x in out if x)
    # Prefer one QR with all memo details. Fail clearly instead of silently
    # creating multiple QR codes or dropping item/customer/payment information.
    if len(payload.encode("utf-8")) > 2850:
        # Remove optional company metadata and long footer/note first, retaining
        # every item row and all required invoice/payment fields.
        optional = {"Co phone:", "Co address:", "Email:", "TIN:", "Reg:"}
        compact = [x for x in out if not any(x.startswith(k) for k in optional)]
        payload = "\n".join(x for x in compact if x)
    return [payload]

def _memo_qr_data_uri(memo, company):
    # Per-company setting: either plain memo text in the QR itself, or a direct memo PDF URL.
    qr_mode = "pdf" if str((company or {}).get("qr_mode", "text")).lower() == "pdf" else "text"
    direct_pdf_url = str((memo or {}).get("qr_url") or "").strip()
    payloads = [direct_pdf_url] if qr_mode == "pdf" and direct_pdf_url.startswith(("https://", "http://")) else _memo_qr_payloads(memo, company)
    def svg_data_uri(text_payload):
        # The qrcode package's SVG encoder does not need PNG/renderPM native dependencies.
        try:
            import qrcode
            from qrcode.image.svg import SvgPathImage
            qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
            qr.add_data(text_payload); qr.make(fit=True)
            raw = qr.make_image(image_factory=SvgPathImage).to_string()
            if isinstance(raw, str): raw = raw.encode("utf-8")
            raw = re.sub(rb"<\?xml[^>]*\?>\s*", b"", raw, count=1)
            raw = re.sub(rb"<!DOCTYPE\s+svg[\s\S]*?>\s*", b"", raw, count=1)
            return "data:image/svg+xml;base64," + base64.b64encode(raw).decode("ascii")
        except Exception:
            from reportlab.graphics.barcode import qr as qr_barcode
            from reportlab.graphics.shapes import Drawing
            from reportlab.graphics import renderSVG
            widget = qr_barcode.QrCodeWidget(text_payload, barLevel="L")
            x1, y1, x2, y2 = widget.getBounds()
            drawing = Drawing(x2 - x1, y2 - y1); drawing.add(widget)
            raw = renderSVG.drawToString(drawing)
            if isinstance(raw, str): raw = raw.encode("utf-8")
            raw = re.sub(rb"<\?xml[^>]*\?>\s*", b"", raw, count=1)
            raw = re.sub(rb"<!DOCTYPE\s+svg[\s\S]*?>\s*", b"", raw, count=1)
            return "data:image/svg+xml;base64," + base64.b64encode(raw).decode("ascii")

    try:
        rendered = [svg_data_uri(text_payload) for text_payload in payloads]
        return rendered[0] if len(rendered) == 1 else rendered
    except Exception as exc:
        print("plain memo text QR generation failed:", str(exc)[:180])
        return ""

def memo_auth(f):
    @wraps(f)
    def w(*a, **k):
        if session.get("user") and (is_admin_level() or has_perm("pos_orders") or has_perm("pos_use") or has_perm("pos_delete")): return f(*a, **k)
        return jsonify(error="permission_denied"), 403
    return w

def _fmt_date(s):
    s = str(s or "")[:10]
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", s)
    return f"{m.group(3)} / {m.group(2)} / {m.group(1)}" if m else s

@app.post("/api/admin/memos")
@memo_auth
def memo_data():
    j = request.get_json(silent=True) or {}
    source = "online" if j.get("source") == "online" else "pos"
    ids = [int(x) for x in (j.get("ids") or []) if str(x).isdigit()][:300]
    nos = [str(x)[:40] for x in (j.get("nos") or [])][:300]
    if not ids and not nos: return jsonify(error="কোনো অর্ডার সিলেক্ট করা হয়নি"), 400
    if source == "online" and not (is_admin_level() or has_perm("pos_orders")): return jsonify(error="permission_denied"), 403
    if source == "pos" and not (is_admin_level() or has_perm("pos_orders") or has_perm("pos_delete")):
        nos = [n for n in nos]; ids = []   # শুধু pos_use থাকলে invoice নম্বর দিয়ে নিজের বিক্রি
    _, rows = _query_orders({}, ids=ids or None, nos=nos or None, source=source)
    if source == "pos" and not (is_admin_level() or has_perm("pos_orders") or has_perm("pos_delete")):
        rows = [r for r in rows if r["user"] == session.get("user")]
    by_key, by_code = _prod_maps(); memos = []; co = company_profile("alien_tees")
    for r in rows:
        items = []
        if source == "pos":
            posrow = _pos_row(r["id"]) or {}
            company_key = _company_key(posrow.get("company_key") or "alien_tees")
            memo_company = company_profile(company_key); memo_logo = _company_logo(company_key)
            try: raw = json.loads(posrow.get("items") or "[]")
            except Exception: raw = []
            for x in raw:
                p = by_key.get(str(x.get("key", ""))) or {}
                items.append(dict(title=str(x.get("title") or ""), code=str(x.get("id") or ""), variation=str(x.get("size") or ""), qty=int(_num(x.get("qty")) or 1), price=_num(x.get("price")), img=p.get("img", "")))
            zone_names = {"inside": "ঢাকার মধ্যে", "suburbs": "ঢাকার আশেপাশে", "outside": "ঢাকার বাইরে"}
            memos.append(dict(source="pos", id=r["id"], order_no=r["no"], kind=r["method"] or "Cash", date=_fmt_date(r["date"]),
                              company_key=company_key, company=memo_company, memo_template=_sales_memo_template(company_key), has_logo=bool(memo_logo) and _sales_memo_template(company_key).get("show_logo", True), logo_v=_company_logo_version(company_key),
                              name=r["customer"], phone=r["phone"], address=r["address"], seller=(posrow.get("username") or r.get("user") or ""), items=items,
                              customer_note="", office_note="", subtotal=r["subtotal"], discount=r["discount"], delivery=r["delivery"],
                              delivery_zone_label=zone_names.get(posrow.get("delivery_zone") or "", ""),
                              advance=r["paid"], total=r["total"], due=r["due"]))
            sale_id = int(r["id"])
            with db() as c:
                token_row = c.execute("SELECT token FROM pos_memo_tokens WHERE sale_id=?", (sale_id,)).fetchone()
                if not token_row:
                    c.execute("INSERT OR IGNORE INTO pos_memo_tokens(sale_id,token,created_at) VALUES(?,?,?)", (sale_id, secrets.token_urlsafe(24), time.time()))
                    token_row = c.execute("SELECT token FROM pos_memo_tokens WHERE sale_id=?", (sale_id,)).fetchone()
            if token_row:
                memos[-1]["qr_url"] = request.host_url.rstrip("/") + "/memo/pos/" + str(sale_id) + "?t=" + str(token_row["token"])
        else:
            d = r["_raw"]; p = by_code.get(str(d.get("product_code") or "").lower()) or {}
            items.append(dict(title=(p.get("title") or ""), code=str(d.get("product_code") or ""), variation=str(d.get("size") or ""), qty=1, price=_num(d.get("product_price")), img=p.get("img") or d.get("image_url") or ""))
            memo_company = company_profile("alien_tees"); memo_logo = _company_logo("alien_tees")
            memos.append(dict(source="online", id=r["id"], order_no=r["no"], kind=r["method"] or "COD", date=_fmt_date(r["date"]),
                              company_key="alien_tees", company=memo_company, memo_template=_sales_memo_template("alien_tees"), has_logo=bool(memo_logo) and _sales_memo_template("alien_tees").get("show_logo", True), logo_v=_company_logo_version("alien_tees"),
                              name=r["customer"], phone=r["phone"], address=r["address"], seller=(d.get("assigned_to") or d.get("uploaded_by") or r.get("user") or ""), items=items,
                              customer_note=_row_json_get(d, "customer note", "customernote"), office_note=str(d.get("note") or ""),
                              subtotal=r["subtotal"], discount=0.0, delivery=r["delivery"], advance=r["paid"], total=r["total"], due=r["due"]))
            online_id = int(r["id"])
            with db() as c:
                token_row = c.execute("SELECT token FROM order_history_memo_tokens WHERE order_id=?", (online_id,)).fetchone()
                if not token_row:
                    c.execute("INSERT OR IGNORE INTO order_history_memo_tokens(order_id,token,created_at) VALUES(?,?,?)", (online_id, secrets.token_urlsafe(32), time.time()))
                    token_row = c.execute("SELECT token FROM order_history_memo_tokens WHERE order_id=?", (online_id,)).fetchone()
            if token_row:
                memos[-1]["qr_url"] = request.host_url.rstrip("/") + "/memo/online/" + str(online_id) + "?t=" + str(token_row["token"])
    for memo in memos:
        memo["qr_data"] = _memo_qr_data_uri(memo, memo.get("company") or co)
    lg = _company_logo("alien_tees")
    return jsonify(company=co, has_logo=bool(lg), logo_v=_company_logo_version("alien_tees"), memos=memos)

def _pos_row(rid):
    with db() as c:
        r = c.execute("SELECT * FROM pos_sales WHERE id=?", (rid,)).fetchone()
    return dict(r) if r else None


def _make_scannable_memo_pdf(memo, company, logo_data=None):
    """Build a clean, self-contained PDF for a QR scan; no HTML/web chrome is returned."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.lib.enums import TA_RIGHT
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, KeepTogether, Flowable
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.graphics.barcode import qr as qr_barcode
        from reportlab.graphics.shapes import Drawing
        from reportlab.graphics import renderPDF
    except ImportError:
        return Response("PDF memo support is not installed on this server. Install reportlab.", status=500, mimetype="text/plain")

    # Use a Latin-capable base font and switch Bengali text runs to Noto Bengali.
    # NotoSansBengali contains Bengali glyphs but has little/no Latin coverage; using it
    # for the entire paragraph makes English labels disappear on some PDF viewers.
    font_regular, font_bold = "Helvetica", "Helvetica-Bold"
    latin_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    latin_bold_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    bangla_path = "/usr/share/fonts/truetype/noto/NotoSansBengali-Regular.ttf"
    bangla_bold_path = "/usr/share/fonts/truetype/noto/NotoSansBengali-Bold.ttf"
    try:
        if os.path.isfile(latin_path):
            if "QRMemoLatin" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("QRMemoLatin", latin_path))
            font_regular = "QRMemoLatin"
        if os.path.isfile(latin_bold_path):
            if "QRMemoLatinBold" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("QRMemoLatinBold", latin_bold_path))
            font_bold = "QRMemoLatinBold"
        if os.path.isfile(bangla_path) and "QRMemoBangla" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("QRMemoBangla", bangla_path))
        if os.path.isfile(bangla_bold_path) and "QRMemoBanglaBold" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("QRMemoBanglaBold", bangla_bold_path))
    except Exception as font_exc:
        print("memo font registration warning:", str(font_exc)[:120])

    dark = colors.HexColor("#18382b")
    memo_color = str(memo.get("color") or "#145b43")
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", memo_color): memo_color = "#145b43"
    green = colors.HexColor(memo_color)
    pale = colors.HexColor("#eaf5ee"); line = colors.HexColor("#d7e4da")
    muted = colors.HexColor("#65776c"); white = colors.white
    out = io.BytesIO()
    company_name = str(company.get("name") or "Alien Tees")
    doc = SimpleDocTemplate(out, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm,
                            topMargin=31*mm, bottomMargin=17*mm,
                            title=f"{memo.get('title') or 'Sales Memo'} {memo.get('number') or ''}", author=company_name)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="QRTitle", parent=styles["Title"], fontName=font_bold, fontSize=21, leading=25, alignment=0, textColor=dark, spaceAfter=5))
    styles.add(ParagraphStyle(name="QRSub", parent=styles["Normal"], fontName=font_regular, fontSize=8.5, leading=12, textColor=muted))
    styles.add(ParagraphStyle(name="QRCell", parent=styles["Normal"], fontName=font_regular, fontSize=9, leading=12, textColor=dark, wordWrap="CJK"))
    styles.add(ParagraphStyle(name="QRCellBold", parent=styles["QRCell"], fontName=font_bold))
    styles.add(ParagraphStyle(name="QRHead", parent=styles["QRCellBold"], fontSize=7.5, leading=9, textColor=white))
    styles.add(ParagraphStyle(name="QRRight", parent=styles["QRCell"], alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name="QRRightBold", parent=styles["QRCellBold"], alignment=TA_RIGHT))
    def h(v, bold=False):
        from html import escape
        text = escape(str(v if v is not None else "—"), quote=True)
        bangla_font = "QRMemoBanglaBold" if bold else "QRMemoBangla"
        if bangla_font in pdfmetrics.getRegisteredFontNames():
            text = re.sub(r"([\u0980-\u09FF]+)", lambda m: '<font name="' + bangla_font + '">' + m.group(1) + '</font>', text)
        return text
    def para(v, style="QRCell"):
        return Paragraph(h(v, style in ("QRCellBold", "QRHead", "QRRightBold")), styles[style])
    def cash(v): return "BDT " + format(float(v or 0), ",.2f")

    class MemoQR(Flowable):
        def __init__(self, payload, size=26*mm, caption="SCAN TO READ MEMO TEXT"):
            super().__init__(); self.size=size; self.caption=caption; self.widget=qr_barcode.QrCodeWidget(payload,barLevel="L")
            self.x1,self.y1,self.x2,self.y2=self.widget.getBounds()
            self.drawing=Drawing(self.x2-self.x1,self.y2-self.y1);self.drawing.add(self.widget)
        def wrap(self, aw, ah): return self.size, self.size+2*mm
        def draw(self):
            c=self.canv;c.saveState();c.setFillColor(white);c.roundRect(-1*mm,-1*mm,self.size+2*mm,self.size+2*mm,2*mm,fill=1,stroke=0)
            c.translate(0,0);c.scale(self.size/(self.x2-self.x1),self.size/(self.y2-self.y1));renderPDF.draw(self.drawing,c,-self.x1,-self.y1);c.restoreState()

    def on_page(canvas, document):
        canvas.saveState(); w,hh=A4;canvas.setFillColor(green);canvas.rect(0,hh-25*mm,w,25*mm,fill=1,stroke=0)
        left=15*mm
        if logo_data:
            try:
                reader=ImageReader(io.BytesIO(logo_data[1]));canvas.drawImage(reader,15*mm,hh-22*mm,width=18*mm,height=18*mm,preserveAspectRatio=True,anchor='c',mask='auto');left=37*mm
            except Exception: pass
        canvas.setFillColor(white);canvas.setFont("Helvetica-Bold",15);canvas.drawString(left,hh-10*mm,company_name[:55])
        contact=" · ".join(str(company.get(k) or "") for k in ("phone","email") if company.get(k))
        canvas.setFont("Helvetica",8);canvas.drawString(left,hh-18*mm,contact[:100])
        if company.get("address"):
            canvas.drawRightString(w-15*mm,hh-18*mm,str(company.get("address"))[:75])
        canvas.setStrokeColor(line);canvas.line(15*mm,12*mm,w-15*mm,12*mm)
        canvas.setFillColor(muted);canvas.setFont("Helvetica",7);canvas.drawString(15*mm,7*mm,str(memo.get("footer") or company.get("invoice_note") or "Computer generated memo")[:110]);canvas.drawRightString(w-15*mm,7*mm,f"Page {document.page}");canvas.restoreState()

    story=[Spacer(1,2*mm),Paragraph(h(memo.get("title") or "SALES MEMO", True),styles["QRTitle"]),
           Paragraph(h(memo.get("subtitle") or "Complete itemized memo · Keep this document for your records"),styles["QRSub"]),Spacer(1,5*mm)]
    meta=[
        [para("INVOICE / ORDER NO.","QRCellBold"),para(memo.get("number") or "—","QRCellBold"),para("DATE", "QRCellBold"),para(memo.get("date") or "—")],
        [para("CUSTOMER", "QRCellBold"),para(memo.get("customer") or "Walk-in Customer"),para("PHONE", "QRCellBold"),para(memo.get("phone") or "—")],
        [para("ADDRESS", "QRCellBold"),para(memo.get("address") or "—"),para("PAYMENT METHOD", "QRCellBold"),para(memo.get("method") or "—")],
        [para("CREATED BY", "QRCellBold"),para(memo.get("seller") or "—"),para("COMPANY", "QRCellBold"),para(company_name)],
    ]
    mt=Table(meta,colWidths=[31*mm,57*mm,30*mm,57*mm],hAlign="LEFT")
    mt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),pale),("GRID",(0,0),(-1,-1),0.45,line),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)]))
    story.extend([mt,Spacer(1,5*mm),Paragraph("ITEM DETAILS",styles["QRCellBold"]),Spacer(1,2*mm)])
    item_rows=[[para("#","QRHead"),para("PRODUCT / CODE","QRHead"),para("SIZE / VARIANT","QRHead"),para("QTY","QRHead"),para("UNIT PRICE","QRHead"),para("LINE TOTAL","QRHead")]]
    for idx,it in enumerate(memo.get("items") or [],1):
        qty=int(_num(it.get("qty")) or 0);price=_num(it.get("price"));
        product=(str(it.get("title") or it.get("name") or "Product") + ("\nCode: " + str(it.get("code")) if it.get("code") else ""))
        item_rows.append([para(str(idx)),para(product),para(it.get("size") or it.get("variation") or "—"),para(str(qty),"QRRight"),para(cash(price),"QRRight"),para(cash(price*qty),"QRRight")])
    itbl=Table(item_rows,colWidths=[8*mm,62*mm,32*mm,14*mm,30*mm,29*mm],repeatRows=1,hAlign="LEFT")
    itbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),green),("GRID",(0,0),(-1,-1),0.4,line),("ROWBACKGROUNDS",(0,1),(-1,-1),[white,colors.HexColor("#f8fbf8")]),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
    story.extend([itbl,Spacer(1,5*mm)])
    due=max(0.0,round(_num(memo.get("due")),2));totals_data=[
        [para("Product Subtotal"),para(cash(memo.get("subtotal")),"QRRight")],
        [para("Discount"),para(cash(memo.get("discount")),"QRRight")],
        [para("Delivery Charge"),para(cash(memo.get("delivery")),"QRRight")],
        [para("TOTAL PAYABLE","QRCellBold"),para(cash(memo.get("total")),"QRRightBold")],
        [para("PAID / ADVANCE"),para(cash(memo.get("paid")),"QRRight")],
        [para("BALANCE DUE","QRCellBold"),para(cash(due),"QRRightBold")],
    ]
    totals=Table(totals_data,colWidths=[45*mm,43*mm],hAlign="RIGHT")
    totals.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.4,line),("BACKGROUND",(0,3),(-1,3),pale),("BACKGROUND",(0,5),(-1,5),colors.HexColor("#fff4e5")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
    qr_payload = (_memo_qr_payloads(memo, company) or ["SALES MEMO"])[0]
    qr=Table([[MemoQR(qr_payload, size=24*mm, caption="")]],colWidths=[45*mm],hAlign="CENTER")
    qr.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#f5faf6")),("BOX",(0,0),(-1,-1),0.8,line),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),2*mm),("RIGHTPADDING",(0,0),(-1,-1),2*mm),("TOPPADDING",(0,0),(-1,-1),2*mm),("BOTTOMPADDING",(0,0),(-1,-1),2*mm)]))
    panel=Table([[qr,totals]],colWidths=[65*mm,100*mm],hAlign="LEFT")
    # Place the single complete-memo QR in the former summary position and PAYMENT SUMMARY at the former QR position.
    panel.setStyle(TableStyle([("ALIGN",(0,0),(0,0),"LEFT"),("ALIGN",(1,0),(1,0),"RIGHT")]))
    panel.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),1*mm),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    story.extend([panel,Spacer(1,4*mm)])
    if memo.get("notes") or company.get("company_note"):
        story.extend([Paragraph("NOTES",styles["QRCellBold"]),Paragraph(h(memo.get("notes") or company.get("company_note")),styles["QRSub"]),Spacer(1,3*mm)])
    story.append(Paragraph("Thank you for your business.",styles["QRSub"]))
    doc.build(story,onFirstPage=on_page,onLaterPages=on_page);out.seek(0)
    safe=re.sub(r"[^A-Za-z0-9_-]+","-",str(memo.get("number") or "memo")).strip("-")[:70] or "memo"
    return send_file(out,mimetype="application/pdf",as_attachment=False,download_name=f"Memo-{safe}.pdf",conditional=False,
                     headers={"Cache-Control":"private, no-store","X-Robots-Tag":"noindex, nofollow","X-Content-Type-Options":"nosniff"})

@app.get("/memo/pos/<int:sale_id>")
def public_pos_memo(sale_id):
    """Token-protected direct PDF endpoint for QR scans; returns no web-page UI."""
    supplied = str(request.args.get("t") or "")
    if not supplied or len(supplied) > 180: return Response("Memo not found", status=404, mimetype="text/plain")
    with db() as c:
        sale_row=c.execute("SELECT * FROM pos_sales WHERE id=?",(sale_id,)).fetchone()
        token_row=c.execute("SELECT token FROM pos_memo_tokens WHERE sale_id=?",(sale_id,)).fetchone()
    if not sale_row or not token_row or not secrets.compare_digest(supplied,str(token_row["token"])):
        return Response("Memo not found",status=404,mimetype="text/plain")
    sale=dict(sale_row); key=_company_key(sale.get("company_key") or "alien_tees"); co=company_profile(key); mt=_sales_memo_template(key); logo=_company_logo(key) if mt.get("show_logo", True) else None
    try: raw_items=json.loads(sale.get("items") or "[]")
    except Exception: raw_items=[]
    items=[{"title":str(x.get("title") or x.get("id") or "Product"),"code":str(x.get("id") or ""),"size":str(x.get("size") or "—"),"qty":int(_num(x.get("qty")) or 0),"price":_num(x.get("price"))} for x in raw_items if isinstance(x,dict)]
    total=_num(sale.get("total"));paid=_num(sale.get("paid"));due=max(0.0,round(total-paid,2))
    created=datetime.fromtimestamp(float(sale.get("created_at") or time.time()),DHAKA_TZ).strftime("%d %b %Y, %I:%M %p")
    url=request.host_url.rstrip("/")+"/memo/pos/"+str(sale_id)+"?t="+supplied
    memo={"title":mt.get("title") or "SALES MEMO","subtitle":mt.get("subtitle"),"footer":mt.get("footer"),"color":mt.get("color"),"number":sale.get("invoice") or sale_id,"date":created,"customer":sale.get("customer_name") or "Walk-in Customer","phone":sale.get("phone") or "—","address":sale.get("address") or "—","method":sale.get("method") or "Cash","seller":sale.get("username") or "—","items":items,"subtotal":sale.get("subtotal"),"discount":sale.get("discount"),"delivery":sale.get("delivery_charge"),"total":total,"paid":paid,"due":due,"notes":co.get("company_note"),"qr_url":url}
    return _make_scannable_memo_pdf(memo,co,logo)

@app.get("/memo/online/<int:order_id>")
def public_online_order_memo(order_id):
    """Token-protected direct PDF for a legacy Order Sheet row's memo QR."""
    supplied=str(request.args.get("t") or "")
    if not supplied or len(supplied)>180:return Response("Memo not found",status=404,mimetype="text/plain")
    with db() as c:
        row=c.execute("SELECT * FROM order_history WHERE id=?",(order_id,)).fetchone()
        token=c.execute("SELECT token FROM order_history_memo_tokens WHERE order_id=?",(order_id,)).fetchone()
    if not row or not token or not secrets.compare_digest(supplied,str(token["token"])):
        return Response("Memo not found",status=404,mimetype="text/plain")
    d=dict(row);_,by_code=_prod_maps();code=str(d.get("product_code") or "");pdt=by_code.get(code.lower()) or {}
    price=_num(d.get("product_price"));delivery=_num(d.get("delivery_charge"));paid=_num(d.get("advance_paid"));total=price+delivery
    due=_num(d.get("collectible_amount")) if str(d.get("collectible_amount") or "").strip() else max(0,total-paid)
    key="alien_tees";co=company_profile(key);mt=_sales_memo_template(key);logo=_company_logo(key) if mt.get("show_logo", True) else None
    items=[{"title":pdt.get("title") or code or "Product","code":code,"size":d.get("size") or "—","qty":1,"price":price}]
    no=_order_no(d,order_id);url=request.host_url.rstrip("/")+"/memo/online/"+str(order_id)+"?t="+supplied
    memo={"title":mt.get("title") or "ONLINE ORDER MEMO","subtitle":mt.get("subtitle"),"footer":mt.get("footer"),"color":mt.get("color"),"number":no,"date":str(d.get("order_date") or "—"),"customer":d.get("customer_name") or "—","phone":d.get("phone") or "—","address":d.get("address") or "—","method":d.get("payment_status") or "COD","seller":d.get("assigned_to") or d.get("uploaded_by") or "—","items":items,"subtotal":price,"discount":0,"delivery":delivery,"total":total,"paid":paid,"due":due,"notes":d.get("note") or co.get("company_note"),"qr_url":url}
    return _make_scannable_memo_pdf(memo,co,logo)


# ===================== Purchase / Supplier / Expense Management =====================

# ===================== Purchase / Supplier / Expense Management =====================
# Additive, permission-gated modules. Uses new erp_* tables and does not reset existing data.
def erp_module_auth(*permissions):
    def decorate(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            if not session.get("user"):
                return jsonify(error="login_required"), 401
            if is_admin_level() or any(user_perms().get(p, False) for p in permissions):
                return fn(*args, **kwargs)
            return jsonify(error="permission_denied"), 403
        return wrapped
    return decorate

def _erp_num(value, default=None):
    try:
        raw = str(value if value is not None else "").replace(",", "").strip()
        number = float(raw)
        if not (number == number and abs(number) != float("inf")): return default
        return number
    except (TypeError, ValueError):
        return default

def _erp_clean_day(value):
    value = str(value or "").strip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value): return ""
    try: datetime.strptime(value, "%Y-%m-%d")
    except ValueError: return ""
    return value

def _erp_due_supplier(supplier_id, c=None):
    own = c is None
    conn = c or db()
    try:
        sup = conn.execute("SELECT opening_balance FROM erp_suppliers WHERE id=?", (supplier_id,)).fetchone()
        if not sup: return 0.0
        purchases = conn.execute("SELECT COALESCE(SUM(total),0) FROM erp_purchases WHERE supplier_id=?", (supplier_id,)).fetchone()[0]
        payments = conn.execute("SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments WHERE supplier_id=?", (supplier_id,)).fetchone()[0]
        return round(float(sup["opening_balance"] or 0) + float(purchases or 0) - float(payments or 0), 2)
    finally:
        if own: conn.close()

def _erp_payment_totals(c, supplier_id=None, purchase_id=None):
    sql = "SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments WHERE 1=1"; args=[]
    if supplier_id is not None: sql += " AND supplier_id=?"; args.append(supplier_id)
    if purchase_id is not None: sql += " AND purchase_id=?"; args.append(purchase_id)
    return float(c.execute(sql, args).fetchone()[0] or 0)

def _erp_date_where(alias_date, frm="", to=""):
    sql=""; args=[]
    frm=_erp_clean_day(frm); to=_erp_clean_day(to)
    if frm: sql += f" AND {alias_date}>=?"; args.append(frm)
    if to: sql += f" AND {alias_date}<=?"; args.append(to)
    return sql,args

def _erp_purchase_records(q="", frm="", to="", limit=300):
    q=str(q or "").strip().lower(); limit=max(1,min(1000,int(limit or 300)))
    sql="""SELECT p.*,
       COALESCE((SELECT COUNT(*) FROM erp_purchase_items i WHERE i.purchase_id=p.id),0) item_count,
       COALESCE((SELECT SUM(i.quantity) FROM erp_purchase_items i WHERE i.purchase_id=p.id),0) quantity,
       COALESCE((SELECT SUM(pm.amount) FROM erp_supplier_payments pm WHERE pm.purchase_id=p.id),0) paid,
       COALESCE((SELECT COUNT(*) FROM erp_purchase_items sx WHERE sx.purchase_id=p.id AND (trim(sx.stock_sheet)='' OR trim(sx.item_code)='' OR trim(sx.size)='')),0) missing_stock_fields
       FROM erp_purchases p WHERE 1=1"""
    args=[]
    if q:
        sql += " AND (lower(p.purchase_no) LIKE ? OR lower(p.supplier_name) LIKE ? OR lower(p.invoice_no) LIKE ? OR CAST(p.id AS TEXT)=?)"
        pat="%"+q+"%"; args.extend([pat,pat,pat,q])
    extra, date_args = _erp_date_where("p.purchase_date",frm,to); sql += extra; args.extend(date_args)
    sql += " ORDER BY p.id DESC LIMIT ?"; args.append(limit)
    with db() as c: rows=[dict(r) for r in c.execute(sql,args).fetchall()]
    for r in rows:
        r["paid"]=round(float(r.get("paid") or 0),2);r["due"]=round(float(r.get("total") or 0)-r["paid"],2);r["stock_receivable"]=int(r.get("missing_stock_fields") or 0)==0
    with db() as c:
        totals=c.execute("SELECT COALESCE(SUM(total),0) total, COUNT(*) count FROM erp_purchases").fetchone()
        paid=c.execute("SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments").fetchone()[0]
        due=c.execute("SELECT COALESCE(SUM(opening_balance),0) FROM erp_suppliers").fetchone()[0] + float(totals["total"] or 0) - float(paid or 0)
    return rows,{"total":round(float(totals["total"] or 0),2),"count":int(totals["count"] or 0),"paid":round(float(paid or 0),2),"due":round(max(0,float(due or 0)),2)}

def _erp_expense_records(q="", frm="", to="", limit=300):
    q=str(q or "").strip().lower(); limit=max(1,min(1000,int(limit or 300)))
    sql="SELECT * FROM erp_expenses WHERE 1=1"; args=[]
    if q:
        pat="%"+q+"%";sql += " AND (lower(title) LIKE ? OR lower(category) LIKE ? OR lower(paid_to) LIKE ? OR lower(expense_no) LIKE ? OR lower(reference) LIKE ?)";args.extend([pat,pat,pat,pat,pat])
    extra,date_args=_erp_date_where("expense_date",frm,to);sql+=extra;args.extend(date_args)
    sql += " ORDER BY id DESC LIMIT ?";args.append(limit)
    with db() as c: rows=[dict(r) for r in c.execute(sql,args).fetchall()]
    today=datetime.now(DHAKA_TZ).date();month_prefix=today.strftime("%Y-%m")
    with db() as c:
        counts=c.execute("SELECT COUNT(*) count, COALESCE(SUM(CASE WHEN status='Approved' THEN amount ELSE 0 END),0) approved, COALESCE(SUM(CASE WHEN status='Pending' THEN amount ELSE 0 END),0) pending, COALESCE(SUM(CASE WHEN status='Approved' AND substr(expense_date,1,7)=? THEN amount ELSE 0 END),0) month_total FROM erp_expenses",(month_prefix,)).fetchone()
    return rows,{k:(round(float(counts[k] or 0),2) if k!='count' else int(counts[k] or 0)) for k in ('approved','pending','month_total','count')}

def _erp_supplier_rows(include_inactive=False):
    with db() as c:
        sql="SELECT * FROM erp_suppliers"+("" if include_inactive else " WHERE active=1")+" ORDER BY active DESC,name COLLATE NOCASE"
        raw=[dict(r) for r in c.execute(sql).fetchall()]
        sums={r["supplier_id"]:dict(purchases=float(r["purchases"] or 0),paid=float(r["paid"] or 0)) for r in c.execute("SELECT s.id supplier_id, COALESCE((SELECT SUM(total) FROM erp_purchases p WHERE p.supplier_id=s.id),0) purchases, COALESCE((SELECT SUM(amount) FROM erp_supplier_payments x WHERE x.supplier_id=s.id),0) paid FROM erp_suppliers s").fetchall()}
        ptotal=float(c.execute("SELECT COALESCE(SUM(total),0) FROM erp_purchases").fetchone()[0] or 0)
        paidtotal=float(c.execute("SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments").fetchone()[0] or 0)
    out=[]
    for r in raw:
        z=sums.get(r["id"],{"purchases":0.0,"paid":0.0});r.update(z);r["due"]=round(float(r.get("opening_balance") or 0)+z["purchases"]-z["paid"],2);r["purchases"]=round(z["purchases"],2);r["paid"]=round(z["paid"],2);out.append(r)
    total_due=round(sum(r["due"] for r in out),2)
    return out,{"active_count":sum(1 for r in out if r["active"]),"purchase_total":round(ptotal,2),"paid_total":round(paidtotal,2),"due_total":total_due}

def _erp_product_title_from_code(code):
    """Resolve a product code against the current catalogue/workbook for invoice display."""
    target = str(code or "").strip().casefold()
    if not target:
        return ""
    try:
        products = view()
    except Exception:
        products = []
    def matches(product):
        title = str(product.get("title") or "").strip()
        ids = [product.get("id"), product.get("display_id")]
        if any(str(v or "").strip().casefold() == target for v in ids):
            return title if title and title.casefold() != target else ""
        if title and re.match(r"^" + re.escape(str(code).strip()) + r"(?=$|[\s/_-])", title, re.I):
            return title
        return ""
    for product in products:
        resolved = matches(product)
        if resolved:
            return resolved[:180]
    try:
        for product in excel_product_catalog_from_stock():
            resolved = matches(product)
            if resolved:
                return resolved[:180]
    except Exception:
        pass
    return str(code or "").strip()[:180]


def _erp_validate_purchase_items(items):
    if not isinstance(items,list) or not items or len(items)>100: raise ValueError("Add between 1 and 100 purchase items.")
    clean=[]; subtotal=0.0
    for idx,x in enumerate(items,1):
        if not isinstance(x,dict): raise ValueError(f"Purchase line {idx} is invalid.")
        code=str(x.get("code") or "").strip()[:100];name=str(x.get("name") or "").strip()[:180];size=str(x.get("size") or "").strip()[:50];sheet=str(x.get("sheet") or "").strip()[:100]
        try: qty=int(x.get("qty"))
        except (TypeError,ValueError): raise ValueError(f"Purchase line {idx}: quantity must be a whole number.")
        cost=_erp_num(x.get("unit_cost"),None)
        if not code: raise ValueError(f"Purchase line {idx}: product code is required.")
        # Product name is optional in the entry form. Resolve it from the catalog for reports/memos.
        if not name: name=_erp_product_title_from_code(code) or code
        if qty<1 or qty>100000: raise ValueError(f"Purchase line {idx}: quantity must be between 1 and 100000.")
        if cost is None or cost<0 or cost>100000000: raise ValueError(f"Purchase line {idx}: unit cost is invalid.")
        line=round(qty*cost,2);subtotal=round(subtotal+line,2)
        clean.append({"code":code,"name":name,"size":size,"sheet":sheet,"qty":qty,"unit_cost":round(cost,2),"line_total":line})
    return clean,subtotal

def _erp_validate_stock_items(items):
    if not os.path.isfile(EXCEL_STOCK_FILE): raise ValueError("The Excel stock workbook is not available. Upload it in Stock Management first.")
    with EXCEL_LOCK:
        wb=load_workbook(EXCEL_STOCK_FILE,read_only=True,data_only=True)
        try:
            for idx,x in enumerate(items,1):
                sheet,code,size=x.get("sheet"),x.get("code"),normalize_stock_size(x.get("size"))
                if not sheet or not code or not size: raise ValueError(f"Purchase line {idx}: choose stock sheet, product code and size to receive stock.")
                if sheet not in wb.sheetnames: raise ValueError(f"Purchase line {idx}: stock sheet '{sheet}' was not found.")
                ws=wb[sheet]; headers=[normalize_stock_size(h) for h in excel_headers(ws)]
                if size not in headers[1:]: raise ValueError(f"Purchase line {idx}: size '{size}' was not found in sheet '{sheet}'.")
                found=False
                for row in ws.iter_rows(min_row=2,max_col=1,values_only=False):
                    if row and excel_cell_text(row[0]).strip().casefold()==code.casefold(): found=True
                if not found: raise ValueError(f"Purchase line {idx}: product code '{code}' was not found in sheet '{sheet}'.")
        finally: wb.close()


def _erp_apply_purchase_stock(purchase_id):
    with db() as c:
        purchase=c.execute("SELECT * FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
        if not purchase: raise ValueError("Purchase record was not found.")
        if purchase["stock_received"]: return {"received":True,"unchanged":True}
        items=[dict(r) for r in c.execute("SELECT * FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(purchase_id,)).fetchall()]
    if not items: raise ValueError("No purchase items are available to receive.")
    _erp_validate_stock_items([{"sheet":x["stock_sheet"],"code":x["item_code"],"size":x["size"]} for x in items])
    tmp=None;edits=[]
    try:
        with EXCEL_LOCK:
            # Re-check while holding the stock workbook lock so rapid double-clicks
            # cannot apply the same purchase receipt to stock twice.
            with db() as c:
                fresh=c.execute("SELECT stock_received FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
            if not fresh: raise ValueError("Purchase record was not found.")
            if fresh["stock_received"]: return {"received":True,"unchanged":True}
            wb=load_workbook(EXCEL_STOCK_FILE)
            try:
                targets={}
                for x in items:
                    ws=wb[x["stock_sheet"]]; headers=[normalize_stock_size(h) for h in excel_headers(ws)]; size=normalize_stock_size(x["size"])
                    if size not in headers[1:]: raise ValueError(f"Size '{size}' was not found on sheet '{ws.title}'.")
                    col=headers.index(size)+1; found_rows=[]
                    for rn in range(2,ws.max_row+1):
                        if excel_cell_text(ws.cell(rn,1)).strip().casefold()==str(x["item_code"]).strip().casefold(): found_rows.append(rn)
                    if not found_rows: raise ValueError(f"Product code '{x['item_code']}' was not found on sheet '{ws.title}'.")
                    rn=found_rows[-1]; key=(ws.title,rn,col)
                    if key not in targets:
                        cell=ws.cell(rn,col)
                        try: old=int(float(cell.value or 0))
                        except (TypeError,ValueError): old=0
                        targets[key]={"sheet":ws.title,"row":rn,"code":x["item_code"],"size":size,"col":col,"old":old,"qty":0}
                    targets[key]["qty"]+=int(x["quantity"])
                for item in targets.values():
                    ws=wb[item["sheet"]];new=item["old"]+item["qty"]
                    if new<0 or new>10000000: raise ValueError("The resulting stock quantity is outside the allowed range.")
                    ws.cell(item["row"],item["col"]).value=new
                    edits.append((item["sheet"],item["row"],item["code"],item["size"],item["old"],new,session.get("user"),f"Purchase receipt {purchase['purchase_no']}",time.time()))
                os.makedirs(INSTANCE,exist_ok=True);fd,tmp=tempfile.mkstemp(prefix="purchase-stock-",suffix=".xlsx",dir=INSTANCE);os.close(fd)
                snapshot_excel_stock("before-purchase-receipt",session.get("user",""))
                wb.save(tmp)
            finally: wb.close()
            os.replace(tmp,EXCEL_STOCK_FILE);tmp=None
            # Mark the receipt inside the same lock before another request may enter.
            with db() as c:
                c.executemany("INSERT INTO excel_stock_history(sheet,row_number,product_code,size,old_qty,new_qty,username,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)",edits)
                c.execute("INSERT INTO excel_stock_batches(username,change_count,plus_units,minus_units,created_at) VALUES(?,?,?,?,?)",(session.get("user"),len(edits),sum(max(0,x[5]-x[4]) for x in edits),sum(max(0,x[4]-x[5]) for x in edits),time.time()))
                c.execute("UPDATE erp_purchases SET stock_received=1 WHERE id=? AND stock_received=0",(purchase_id,))
        refresh_stock_from_excel();log_activity("purchase_stock_received",str(purchase_id),f"Received stock for {purchase['purchase_no']}")
        return {"received":True,"changes":len(edits)}
    except Exception:
        if tmp and os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass
        raise

@app.get("/api/erp/meta")
@erp_module_auth("purchase_view","purchase_manage","supplier_manage","expense_view","expense_manage")
def erp_meta():
    sheets=[]
    if os.path.isfile(EXCEL_STOCK_FILE):
        try:
            with EXCEL_LOCK:
                wb=load_workbook(EXCEL_STOCK_FILE,read_only=True,data_only=True);sheets=list(wb.sheetnames);wb.close()
        except Exception: sheets=[]
    return jsonify(sheets=sheets,has_stock=bool(sheets))

@app.get("/api/erp/catalog-options")
@erp_module_auth("purchase_view","purchase_manage")
def erp_catalog_options():
    """Small safe catalogue payload used by the searchable Product Name input."""
    out=[]; seen=set()
    try: products=list(view() or [])
    except Exception: products=[]
    try: products.extend(excel_product_catalog_from_stock())
    except Exception: pass
    for p in products:
        title=str(p.get("title") or p.get("name") or "").strip()
        code=str(p.get("id") or p.get("display_id") or "").strip()
        if not title: continue
        key=(title.casefold(),code.casefold())
        if key in seen: continue
        seen.add(key)
        out.append({"title":title[:180],"code":code[:100],"display_id":str(p.get("display_id") or "")[:100],"img":str(p.get("img") or "")[:1000],"store":str(p.get("store") or "")[:100]})
        if len(out)>=6000: break
    return jsonify(ok=True,products=out)

@app.get("/api/erp/stock-item-lookup")
@login_required
def erp_stock_item_lookup():
    """Resolve a code or catalogue name to its real worksheet row and current sizes."""
    perms=user_perms()
    if not is_admin_level() and not any(perms.get(k,False) for k in ("stock_dashboard","stock_manage","stock_add","stock_set","purchase_manage","purchase_view")):
        return jsonify(ok=False,error="permission_denied",message="Stock lookup is not enabled for this account."),403
    q=str(request.args.get("q") or request.args.get("code") or "").strip()[:240]
    if not q: return jsonify(ok=False,error="Enter a product name or code.",matches=[]),400
    if not os.path.isfile(EXCEL_STOCK_FILE): return jsonify(ok=False,error="The stock workbook is not available.",matches=[]),404
    qf=q.casefold(); products=[]
    try: products.extend(list(view() or []))
    except Exception: pass
    try: products.extend(excel_product_catalog_from_stock())
    except Exception: pass
    chosen=None
    def _product_ids(p):
        vals=[p.get("id"),p.get("display_id"),display_product_id(p)]
        title=str(p.get("title") or p.get("name") or "").strip()
        m=re.match(r"^([A-Za-z0-9]+)(?=\s|[-_/]|$)",title)
        if m: vals.append(m.group(1))
        return {str(v).strip().casefold() for v in vals if str(v or "").strip()}
    # Prefer exact title/code, then a unique name prefix match; this avoids choosing a vague partial match.
    for p in products:
        title=str(p.get("title") or p.get("name") or "").strip()
        if qf in _product_ids(p) or title.casefold()==qf:
            chosen=p;break
    if chosen is None:
        name_matches=[p for p in products if str(p.get("title") or p.get("name") or "").strip().casefold().startswith(qf)] if len(qf)>=3 else []
        if len(name_matches)==1: chosen=name_matches[0]
    candidates={qf}
    if chosen:
        candidates |= _product_ids(chosen)
    matches=[]
    try:
        with EXCEL_LOCK:
            wb=load_workbook(EXCEL_STOCK_FILE,read_only=True,data_only=True)
            for ws in wb.worksheets:
                headers=excel_headers(ws)
                if not headers or not headers[0]: continue
                size_headers=[]
                for col,header in enumerate(headers[1:],start=1):
                    h=normalize_stock_size(header);hf=h.casefold()
                    if not h or hf in SKIPH or any(t in hf for t in ("price","rate","amount","cost","total","sum","name","title","image","photo","picture","link","url","দাম","মূল্য","টাকা")):
                        continue
                    size_headers.append((col,h))
                # Use formatted cell values for codes (including zero-padded codes) and cell.value for quantities.
                for rn,cells in enumerate(ws.iter_rows(min_row=2,values_only=False),start=2):
                    if not cells: continue
                    code=excel_cell_text(cells[0]).strip()
                    if not code or code.casefold() not in candidates: continue
                    sizes=[]
                    for col,h in size_headers:
                        raw=cells[col].value if col<len(cells) else None
                        try: qty=max(0,int(float(str(raw or 0).replace(",",""))))
                        except (TypeError,ValueError): qty=0
                        sizes.append({"size":h,"qty":qty})
                    if sizes:
                        matches.append({"sheet":ws.title,"code":code,"row":rn,"sizes":sizes,"product_name":str((chosen or {}).get("title") or (chosen or {}).get("name") or _erp_product_title_from_code(code) or code)[:180],"product_img":str((chosen or {}).get("img") or "")[:1000]})
            wb.close()
    except Exception as exc:
        print("stock item lookup failed:",str(exc)[:200])
        return jsonify(ok=False,error="Could not inspect stock workbook size columns.",matches=[]),500
    product_payload=None
    if chosen:
        product_payload={"title":str(chosen.get("title") or chosen.get("name") or q)[:180],"code":str(chosen.get("id") or chosen.get("display_id") or q)[:100],"img":str(chosen.get("img") or "")[:1000],"store":str(chosen.get("store") or "")[:100]}
    return jsonify(ok=True,query=q,product=product_payload,matches=matches,found=bool(matches))

@app.get("/api/erp/suppliers")
@erp_module_auth("purchase_view","purchase_manage","supplier_manage")
def erp_supplier_list():
    include_inactive=request.args.get("include_inactive")=="1"
    rows,sm=_erp_supplier_rows(include_inactive)
    q=request.args.get("q"," ").strip().casefold()
    if q and q.strip(): rows=[r for r in rows if q in (str(r.get("name",""))+" "+str(r.get("phone",""))+" "+str(r.get("email",""))).casefold()]
    return jsonify(suppliers=rows,summary=sm)

@app.post("/api/erp/suppliers")
@erp_module_auth("supplier_manage","purchase_manage")
def erp_supplier_create():
    j=request.get_json(silent=True) or {};name=str(j.get("name") or "").strip()[:160]
    if not name:return jsonify(error="Supplier name is required."),400
    opening=_erp_num(j.get("opening_balance",0),None)
    if opening is None or opening<0 or opening>1000000000:return jsonify(error="Opening balance must be a valid non-negative amount."),400
    now=time.time()
    with db() as c:
        cur=c.execute("INSERT INTO erp_suppliers(name,phone,email,address,opening_balance,notes,active,created_by,created_at) VALUES(?,?,?,?,?,?,1,?,?)",(name,str(j.get("phone") or "").strip()[:60],str(j.get("email") or "").strip()[:160],str(j.get("address") or "").strip()[:500],round(opening,2),str(j.get("notes") or "").strip()[:1000],session.get("user"),now))
        sid=cur.lastrowid
    log_activity("supplier_add",name,f"Supplier #{sid} added")
    return jsonify(ok=True,id=sid)

@app.post("/api/erp/suppliers/<int:supplier_id>")
@erp_module_auth("supplier_manage","purchase_manage")
def erp_supplier_update(supplier_id):
    j=request.get_json(silent=True) or {};active=1 if j.get("active") else 0
    with db() as c:
        r=c.execute("SELECT id,name FROM erp_suppliers WHERE id=?",(supplier_id,)).fetchone()
        if not r:return jsonify(error="Supplier not found."),404
        c.execute("UPDATE erp_suppliers SET active=? WHERE id=?",(active,supplier_id))
    log_activity("supplier_status",r["name"],"activated" if active else "deactivated")
    return jsonify(ok=True)

@app.post("/api/erp/suppliers/<int:supplier_id>/payments")
@erp_module_auth("supplier_manage","purchase_manage")
def erp_supplier_payment(supplier_id):
    j=request.get_json(silent=True) or {};amount=_erp_num(j.get("amount"),None)
    if amount is None or amount<=0 or amount>1000000000:return jsonify(error="Payment amount is invalid."),400
    today=datetime.now(DHAKA_TZ).date().isoformat();pay_date=_erp_clean_day(j.get("payment_date")) or today;method=str(j.get("payment_method") or "Cash").strip()[:50];ref=str(j.get("reference") or "").strip()[:120];notes=str(j.get("notes") or "").strip()[:800]
    with db() as c:
        supplier=c.execute("SELECT * FROM erp_suppliers WHERE id=?",(supplier_id,)).fetchone()
        if not supplier:return jsonify(error="Supplier not found."),404
        due=_erp_due_supplier(supplier_id,c)
        if amount>due+0.01:return jsonify(error=f"Payment exceeds current supplier due ({due:.2f} BDT)."),400
        remaining=round(amount,2);allocations=[]
        # Allocate payment FIFO to invoices so their balances update consistently.
        purchases=c.execute("SELECT p.id,p.total,COALESCE((SELECT SUM(amount) FROM erp_supplier_payments x WHERE x.purchase_id=p.id),0) paid FROM erp_purchases p WHERE p.supplier_id=? ORDER BY p.purchase_date,p.id",(supplier_id,)).fetchall()
        for p in purchases:
            invoice_due=max(0.0,round(float(p["total"] or 0)-float(p["paid"] or 0),2))
            if invoice_due<=0 or remaining<=0: continue
            a=min(invoice_due,remaining);allocations.append((p["id"],a));remaining=round(remaining-a,2)
        # Any remainder reduces the supplier opening balance / unallocated balance.
        if remaining>0: allocations.append((None,remaining))
        for purchase_id,amt in allocations:
            c.execute("INSERT INTO erp_supplier_payments(supplier_id,purchase_id,amount,payment_date,payment_method,reference,notes,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(supplier_id,purchase_id,amt,pay_date,method,ref,notes,session.get("user"),time.time()))
    log_activity("supplier_payment",supplier["name"],f"Paid {amount:.2f} via {method}; ref {ref}")
    return jsonify(ok=True,amount=round(amount,2))

@app.get("/api/erp/purchases")
@erp_module_auth("purchase_view","purchase_manage")
def erp_purchase_list():
    rows,sm=_erp_purchase_records(request.args.get("q",""),request.args.get("from",""),request.args.get("to",""),request.args.get("limit",300))
    return jsonify(rows=rows,summary=sm)

@app.get("/api/erp/purchases/<int:purchase_id>")
@erp_module_auth("purchase_view","purchase_manage")
def erp_purchase_detail(purchase_id):
    with db() as c:
        r=c.execute("SELECT * FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
        if not r:return jsonify(error="Purchase not found."),404
        p=dict(r);p["paid"]=_erp_payment_totals(c,purchase_id=purchase_id);p["due"]=round(p["total"]-p["paid"],2)
        items=[dict(x) for x in c.execute("SELECT * FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(purchase_id,)).fetchall()]
    return jsonify(purchase=p,items=items)

@app.post("/api/erp/purchases")
@erp_module_auth("purchase_manage")
def erp_purchase_create():
    j=request.get_json(silent=True) or {}
    date=_erp_clean_day(j.get("purchase_date"))
    if not date:return jsonify(error="Choose a valid purchase date."),400
    try:items,subtotal=_erp_validate_purchase_items(j.get("items"))
    except ValueError as e:return jsonify(error=str(e)),400
    discount=_erp_num(j.get("discount",0),None);paid=_erp_num(j.get("paid_amount",0),None)
    if discount is None or discount<0 or discount>subtotal:return jsonify(error="Discount must be between zero and the subtotal."),400
    total=round(subtotal-discount,2)
    if paid is None or paid<0 or paid>total+0.01:return jsonify(error="Initial payment cannot exceed the purchase total."),400
    # Supplier purchase entries always update the linked Excel stock automatically.
    receive=True
    try:_erp_validate_stock_items(items)
    except ValueError as e:return jsonify(error=str(e)),400
    with db() as c:
        supplier=c.execute("SELECT * FROM erp_suppliers WHERE id=? AND active=1",(int(j.get("supplier_id") or 0),)).fetchone()
        if not supplier:return jsonify(error="Choose an active supplier."),400
        now=time.time();purchase_no="PUR-"+datetime.now(DHAKA_TZ).strftime("%Y%m%d")+"-"+str(time.time_ns())[-6:]
        cur=c.execute("INSERT INTO erp_purchases(purchase_no,supplier_id,supplier_name,invoice_no,purchase_date,subtotal,discount,total,notes,stock_received,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,0,?,?)",(purchase_no,supplier["id"],supplier["name"],str(j.get("invoice_no") or "").strip()[:100],date,subtotal,round(discount,2),total,str(j.get("notes") or "").strip()[:2000],session.get("user"),now))
        pid=cur.lastrowid
        c.executemany("INSERT INTO erp_purchase_items(purchase_id,stock_sheet,item_code,item_name,size,quantity,unit_cost,line_total) VALUES(?,?,?,?,?,?,?,?)",[(pid,x["sheet"],x["code"],x["name"],x["size"],x["qty"],x["unit_cost"],x["line_total"]) for x in items])
        if paid>0:c.execute("INSERT INTO erp_supplier_payments(supplier_id,purchase_id,amount,payment_date,payment_method,reference,notes,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(supplier["id"],pid,round(paid,2),date,str(j.get("payment_method") or "Cash")[:50],str(j.get("invoice_no") or "")[:120],"Initial purchase payment",session.get("user"),now))
    stock_received=False;stock_error=""
    if receive:
        try: _erp_apply_purchase_stock(pid);stock_received=True
        except Exception as e: stock_error=str(e)[:400]
    log_activity("purchase_add",purchase_no,f"Supplier: {supplier['name']}; total {total:.2f}; stock_received={stock_received}")
    return jsonify(ok=True,id=pid,purchase_no=purchase_no,total=total,stock_received=stock_received,stock_error=stock_error)

@app.post("/api/erp/purchases/<int:purchase_id>/payments")
@erp_module_auth("purchase_manage","supplier_manage")
def erp_purchase_payment(purchase_id):
    j=request.get_json(silent=True) or {};amount=_erp_num(j.get("amount"),None)
    if amount is None or amount<=0:return jsonify(error="Payment amount must be greater than zero."),400
    with db() as c:
        p=c.execute("SELECT * FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
        if not p:return jsonify(error="Purchase not found."),404
        paid=_erp_payment_totals(c,purchase_id=purchase_id);due=round(float(p["total"] or 0)-paid,2)
        if amount>due+0.01:return jsonify(error=f"Payment exceeds invoice due ({due:.2f} BDT)."),400
        c.execute("INSERT INTO erp_supplier_payments(supplier_id,purchase_id,amount,payment_date,payment_method,reference,notes,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(p["supplier_id"],purchase_id,round(amount,2),_erp_clean_day(j.get("payment_date")) or datetime.now(DHAKA_TZ).date().isoformat(),str(j.get("payment_method") or "Cash")[:50],str(j.get("reference") or "")[:120],str(j.get("notes") or "")[:800],session.get("user"),time.time()))
    log_activity("purchase_payment",p["purchase_no"],f"Payment {amount:.2f}; {j.get('payment_method') or 'Cash'}")
    return jsonify(ok=True,due=round(due-amount,2))

def _erp_reverse_purchase_stock_for_delete(purchase_id):
    """Subtract a received purchase from stock before deletion; fail safely if it would go negative."""
    if not os.path.isfile(EXCEL_STOCK_FILE):
        raise ValueError("The purchase was received into stock, but the stock workbook is unavailable. Restore the workbook before deleting this purchase.")
    with db() as c:
        purchase=c.execute("SELECT * FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
        if not purchase: raise ValueError("Purchase record was not found.")
        items=[dict(r) for r in c.execute("SELECT * FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(purchase_id,)).fetchall()]
    if not purchase["stock_received"]: return {"reversed":False,"changes":0}
    if not items: raise ValueError("Purchase items are missing; stock reversal cannot be verified.")
    _erp_validate_stock_items([{"sheet":x["stock_sheet"],"code":x["item_code"],"size":x["size"]} for x in items])
    edits=[];tmp=None
    with EXCEL_LOCK:
        wb=load_workbook(EXCEL_STOCK_FILE)
        try:
            targets={}
            for x in items:
                if x["stock_sheet"] not in wb.sheetnames: raise ValueError(f"Stock sheet '{x['stock_sheet']}' is missing; purchase was not deleted.")
                ws=wb[x["stock_sheet"]];headers=[normalize_stock_size(h) for h in excel_headers(ws)];size=normalize_stock_size(x["size"])
                if size not in headers[1:]: raise ValueError(f"Size '{size}' is missing from stock sheet '{ws.title}'.")
                col=headers.index(size)+1;found=[]
                for rn in range(2,ws.max_row+1):
                    if excel_cell_text(ws.cell(rn,1)).strip().casefold()==str(x["item_code"]).strip().casefold():found.append(rn)
                if not found: raise ValueError(f"Product code '{x['item_code']}' was not found in stock sheet '{ws.title}'.")
                rn=found[-1];key=(ws.title,rn,col)
                if key not in targets:
                    cell=ws.cell(rn,col)
                    try:old=int(float(cell.value or 0))
                    except (TypeError,ValueError):old=0
                    targets[key]={"sheet":ws.title,"row":rn,"code":x["item_code"],"size":size,"col":col,"old":old,"qty":0}
                targets[key]["qty"]+=int(x["quantity"])
            for item in targets.values():
                new_qty=item["old"]-item["qty"]
                if new_qty<0:
                    raise ValueError(f"Cannot delete safely: stock for {item['code']} / {item['size']} is {item['old']}, but this purchase added {item['qty']}. Adjust inventory first.")
                wb[item["sheet"]].cell(item["row"],item["col"]).value=new_qty
                edits.append((item["sheet"],item["row"],item["code"],item["size"],item["old"],new_qty,session.get("user"),f"Reversed deleted purchase {purchase['purchase_no']}",time.time()))
            snapshot_excel_stock("before-purchase-delete",session.get("user",""))
            os.makedirs(INSTANCE,exist_ok=True);fd,tmp=tempfile.mkstemp(prefix="purchase-delete-reversal-",suffix=".xlsx",dir=INSTANCE);os.close(fd)
            wb.save(tmp)
        finally: wb.close()
        try:
            os.replace(tmp,EXCEL_STOCK_FILE);tmp=None
            with db() as c:
                c.executemany("INSERT INTO excel_stock_history(sheet,row_number,product_code,size,old_qty,new_qty,username,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)",edits)
                c.execute("INSERT INTO excel_stock_batches(username,change_count,plus_units,minus_units,created_at) VALUES(?,?,?,?,?)",(session.get("user"),len(edits),0,sum(max(0,x[4]-x[5]) for x in edits),time.time()))
            refresh_stock_from_excel()
        finally:
            if tmp and os.path.exists(tmp):
                try: os.remove(tmp)
                except OSError: pass
    return {"reversed":True,"changes":len(edits)}

@app.route("/api/erp/purchases/<int:purchase_id>", methods=["DELETE"])
@admin_only
def erp_purchase_delete(purchase_id):
    with db() as c:
        p=c.execute("SELECT * FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
        if not p:return jsonify(error="Purchase record not found."),404
        purchase_no=p["purchase_no"]
        stock_received=bool(p["stock_received"])
    reversed_info={"reversed":False,"changes":0}
    if stock_received:
        try: reversed_info=_erp_reverse_purchase_stock_for_delete(purchase_id)
        except ValueError as e:return jsonify(error=str(e)),409
        except Exception as e:return jsonify(error="Stock could not be safely reversed; purchase was not deleted: "+str(e)[:220]),500
    with db() as c:
        fresh=c.execute("SELECT id FROM erp_purchases WHERE id=?",(purchase_id,)).fetchone()
        if not fresh:return jsonify(error="Purchase record was already deleted."),404
        c.execute("DELETE FROM erp_supplier_payments WHERE purchase_id=?",(purchase_id,))
        c.execute("DELETE FROM erp_purchase_items WHERE purchase_id=?",(purchase_id,))
        c.execute("DELETE FROM erp_purchases WHERE id=?",(purchase_id,))
    log_activity("purchase_delete",purchase_no,f"Deleted purchase; stock_reversed={bool(reversed_info.get('reversed'))}; lines={reversed_info.get('changes',0)}")
    return jsonify(ok=True,deleted=True,stock_reversed=bool(reversed_info.get("reversed")),stock_changes=int(reversed_info.get("changes",0)))

@app.post("/api/erp/purchases/<int:purchase_id>/receive-stock")
@erp_module_auth("purchase_manage")
def erp_purchase_receive_stock(purchase_id):
    try:r=_erp_apply_purchase_stock(purchase_id)
    except ValueError as e:return jsonify(error=str(e)),400
    except Exception as e:return jsonify(error="Could not update stock workbook: "+str(e)[:250]),500
    return jsonify(ok=True,**r)

@app.get("/api/erp/expenses")
@erp_module_auth("expense_view","expense_manage")
def erp_expense_list():
    rows,sm=_erp_expense_records(request.args.get("q",""),request.args.get("from",""),request.args.get("to",""),request.args.get("limit",300))
    return jsonify(rows=rows,summary=sm)

@app.post("/api/erp/expenses")
@erp_module_auth("expense_manage")
def erp_expense_create():
    j=request.get_json(silent=True) or {};title=str(j.get("title") or "").strip()[:180];category=str(j.get("category") or "Other").strip()[:100]
    amount=_erp_num(j.get("amount"),None);date=_erp_clean_day(j.get("expense_date"));status=str(j.get("status") or "Pending")
    if not title or not category or not date or amount is None or amount<=0 or amount>1000000000:return jsonify(error="Title, category, valid date and amount are required."),400
    if status not in ("Pending","Approved"):status="Pending"
    now=time.time();no="EXP-"+datetime.now(DHAKA_TZ).strftime("%Y%m%d")+"-"+str(time.time_ns())[-6:]
    with db() as c:cur=c.execute("INSERT INTO erp_expenses(expense_no,title,category,amount,expense_date,payment_method,paid_to,reference,notes,status,created_by,created_at,updated_by,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(no,title,category,round(amount,2),date,str(j.get("payment_method") or "Cash")[:50],str(j.get("paid_to") or "").strip()[:160],str(j.get("reference") or "").strip()[:120],str(j.get("notes") or "").strip()[:2000],status,session.get("user"),now,session.get("user"),now));eid=cur.lastrowid
    log_activity("expense_add",no,f"{title}; {amount:.2f} BDT; {status}")
    return jsonify(ok=True,id=eid,expense_no=no)

@app.post("/api/erp/expenses/<int:expense_id>/status")
@erp_module_auth("expense_manage")
def erp_expense_status(expense_id):
    j=request.get_json(silent=True) or {};status=str(j.get("status") or "")
    if status not in ("Pending","Approved","Rejected"):return jsonify(error="Invalid expense status."),400
    with db() as c:
        r=c.execute("SELECT expense_no,title FROM erp_expenses WHERE id=?",(expense_id,)).fetchone()
        if not r:return jsonify(error="Expense not found."),404
        c.execute("UPDATE erp_expenses SET status=?,updated_by=?,updated_at=? WHERE id=?",(status,session.get("user"),time.time(),expense_id))
    log_activity("expense_status",r["expense_no"],f"{r['title']} → {status}")
    return jsonify(ok=True)

@app.route("/api/erp/expenses/<int:expense_id>", methods=["DELETE"])
@admin_only
def erp_expense_delete(expense_id):
    with db() as c:
        r=c.execute("SELECT expense_no,title,amount,status FROM erp_expenses WHERE id=?",(expense_id,)).fetchone()
        if not r:return jsonify(error="Expense record not found."),404
        c.execute("DELETE FROM erp_expenses WHERE id=?",(expense_id,))
    log_activity("expense_delete",r["expense_no"],f"Deleted {r['title']}; amount {float(r['amount'] or 0):.2f} BDT; previous status {r['status']}")
    return jsonify(ok=True,deleted=True)

@app.get("/api/erp/summary")
@erp_module_auth("purchase_view","purchase_manage","expense_view","expense_manage")
def erp_summary():
    frm=_erp_clean_day(request.args.get("from"));to=_erp_clean_day(request.args.get("to"));pw=" WHERE 1=1";pa=[]
    extra,args=_erp_date_where("purchase_date",frm,to);pw+=extra;pa.extend(args)
    ew=" WHERE 1=1";ea=[];extra,args=_erp_date_where("expense_date",frm,to);ew+=extra;ea.extend(args)
    with db() as c:
        pr=c.execute("SELECT COALESCE(SUM(total),0) total, COUNT(*) count FROM erp_purchases"+pw,pa).fetchone()
        paid_all=c.execute("SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments WHERE 1=1"+_erp_date_where("payment_date",frm,to)[0],_erp_date_where("payment_date",frm,to)[1]).fetchone()[0]
        er=c.execute("SELECT COALESCE(SUM(CASE WHEN status='Approved' THEN amount ELSE 0 END),0) approved, COALESCE(SUM(CASE WHEN status='Pending' THEN amount ELSE 0 END),0) pending FROM erp_expenses"+ew,ea).fetchone()
        all_purchase=float(c.execute("SELECT COALESCE(SUM(total),0) FROM erp_purchases").fetchone()[0] or 0)
        all_paid=float(c.execute("SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments").fetchone()[0] or 0)
        opening=float(c.execute("SELECT COALESCE(SUM(opening_balance),0) FROM erp_suppliers").fetchone()[0] or 0)
        cats=c.execute("SELECT category,COUNT(*) count,COALESCE(SUM(CASE WHEN status='Approved' THEN amount ELSE 0 END),0) approved,COALESCE(SUM(CASE WHEN status='Pending' THEN amount ELSE 0 END),0) pending FROM erp_expenses"+ew+" GROUP BY category ORDER BY approved DESC,pending DESC,category",ea).fetchall()
        sup_rows=c.execute("SELECT s.id,s.name,s.opening_balance,COALESCE((SELECT SUM(total) FROM erp_purchases p WHERE p.supplier_id=s.id),0) purchases,COALESCE((SELECT SUM(amount) FROM erp_supplier_payments x WHERE x.supplier_id=s.id),0) paid FROM erp_suppliers s WHERE s.active=1 ORDER BY s.name COLLATE NOCASE").fetchall()
    # Overall supplier due is global; date range only filters flow totals, not lifetime liabilities.
    srows=[dict(id=int(x["id"]),name=x["name"],purchases=round(float(x["purchases"] or 0),2),paid=round(float(x["paid"] or 0),2),due=round(float(x["opening_balance"] or 0)+float(x["purchases"] or 0)-float(x["paid"] or 0),2)) for x in sup_rows]
    sm={"purchase_total":round(float(pr["total"] or 0),2),"purchase_count":int(pr["count"] or 0),"supplier_paid":round(float(paid_all or 0),2),"supplier_due":round(opening+all_purchase-all_paid,2),"expenses_approved":round(float(er["approved"] or 0),2),"expenses_pending":round(float(er["pending"] or 0),2)}
    return jsonify(summary=sm,expense_categories=[dict(x) for x in cats],supplier_dues=srows)

@app.get("/api/erp/purchases/<int:purchase_id>/memo.pdf")
def erp_purchase_memo_pdf(purchase_id):
    # Normal staff access still requires an appropriate permission. A customer QR token may view only this memo PDF.
    authorized = bool(session.get("user") and (is_admin_level() or any(user_perms().get(x, False) for x in ("purchase_view", "purchase_manage", "supplier_manage"))))
    if not authorized:
        supplied = str(request.args.get("t") or "")
        with db() as c: token_row = c.execute("SELECT token FROM erp_purchase_memo_tokens WHERE purchase_id=?", (purchase_id,)).fetchone()
        if not supplied or not token_row or not secrets.compare_digest(supplied, str(token_row["token"])):
            return jsonify(error="permission_denied"), 403
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Flowable
        from reportlab.lib.utils import ImageReader
        from reportlab.graphics.barcode import qr as qr_barcode
        from reportlab.graphics.shapes import Drawing
        from reportlab.graphics import renderPDF
    except ImportError:
        return jsonify(error="PDF memo dependency is missing. Install it with: pip install reportlab"), 500

    with db() as c:
        raw=c.execute("SELECT p.*,s.name AS supplier_name,s.phone AS supplier_phone,s.email AS supplier_email,s.address AS supplier_address FROM erp_purchases p JOIN erp_suppliers s ON s.id=p.supplier_id WHERE p.id=?",(purchase_id,)).fetchone()
        if not raw: return jsonify(error="Purchase record not found."),404
        p=dict(raw)
        p["paid"]=_erp_payment_totals(c,purchase_id=purchase_id)
        p["due"]=round(float(p.get("total") or 0)-float(p.get("paid") or 0),2)
        items=[dict(x) for x in c.execute("SELECT item_name,item_code,size,quantity,unit_cost,line_total FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(purchase_id,)).fetchall()]
        payments=[dict(x) for x in c.execute("SELECT payment_date,amount,payment_method,reference,created_by FROM erp_supplier_payments WHERE purchase_id=? ORDER BY payment_date,id",(purchase_id,)).fetchall()]
    company = company_profile()
    memo_cfg = _erp_purchase_memo_template()
    logo_data = _company_logo() if memo_cfg.get("show_logo", True) else None
    memo_title = str(memo_cfg.get("title") or "PURCHASE MEMO")[:100]
    memo_subtitle = str(memo_cfg.get("subtitle") or "")[:220]
    memo_footer = str(memo_cfg.get("footer") or "")[:500]
    # Stable, unguessable scan link: the printed QR opens a responsive full purchase memo.
    with db() as c:
        c.execute("CREATE TABLE IF NOT EXISTS erp_purchase_memo_tokens(purchase_id INTEGER PRIMARY KEY, token TEXT UNIQUE NOT NULL, created_at REAL NOT NULL)")
        token_row = c.execute("SELECT token FROM erp_purchase_memo_tokens WHERE purchase_id=?", (purchase_id,)).fetchone()
        if not token_row:
            c.execute("INSERT OR IGNORE INTO erp_purchase_memo_tokens(purchase_id,token,created_at) VALUES(?,?,?)", (purchase_id, secrets.token_urlsafe(32), time.time()))
            token_row = c.execute("SELECT token FROM erp_purchase_memo_tokens WHERE purchase_id=?", (purchase_id,)).fetchone()
    purchase_qr_url = request.host_url.rstrip("/") + "/api/erp/purchases/" + str(purchase_id) + "/memo.pdf?t=" + str(token_row["token"] if token_row else "")
    purchase_notes = str(p.get("notes") or "").strip()
    payment_text = "; ".join(
        f"{x.get('payment_date') or ''} {x.get('payment_method') or ''} BDT {_num(x.get('amount')):,.2f}"
        + (f" Ref {x.get('reference')}" if x.get('reference') else "")
        for x in payments
    )
    if payment_text: purchase_notes = (purchase_notes + " | " if purchase_notes else "") + "Payments: " + payment_text
    purchase_qr_memo = {
        "title": memo_title, "number": p.get("purchase_no"), "date": p.get("purchase_date"),
        "customer": p.get("supplier_name"), "phone": p.get("supplier_phone"),
        "address": p.get("supplier_address"), "method": "Supplier Purchase",
        "items": [{"title": x.get("item_name") or x.get("item_code"), "code": x.get("item_code"),
                   "variation": x.get("size"), "qty": x.get("quantity"), "price": x.get("unit_cost")}
                  for x in items],
        "subtotal": p.get("subtotal"), "discount": p.get("discount"), "delivery": 0,
        "total": p.get("total"), "advance": p.get("paid"), "due": p.get("due"), "notes": purchase_notes,
    }
    purchase_qr_payloads = ([purchase_qr_url] if str(company.get("qr_mode", "text")).lower() == "pdf" and purchase_qr_url.startswith(("https://", "http://")) else (_memo_qr_payloads(purchase_qr_memo, company) or ["PURCHASE MEMO"]))

    class PurchaseMemoQR(Flowable):
        """Vector QR widget so the printed code stays crisp and does not need PNG rendering."""
        def __init__(self, payload, size=26*mm, caption="SCAN TO READ MEMO TEXT"):
            super().__init__()
            self.size = size
            self.caption = caption
            widget = qr_barcode.QrCodeWidget(payload, barLevel="L")
            self.x1, self.y1, self.x2, self.y2 = widget.getBounds()
            self.drawing = Drawing(self.x2-self.x1, self.y2-self.y1)
            self.drawing.add(widget)
        def wrap(self, availWidth, availHeight):
            return (self.size, self.size + 2*mm)
        def draw(self):
            canvas = self.canv
            canvas.saveState()
            canvas.setFillColor(colors.white)
            canvas.roundRect(-1*mm, -1*mm, self.size+2*mm, self.size+2*mm, 2*mm, fill=1, stroke=0)
            canvas.scale(self.size/(self.x2-self.x1), self.size/(self.y2-self.y1))
            renderPDF.draw(self.drawing, canvas, -self.x1, -self.y1)
            canvas.restoreState()

    def esc_pdf(value):
        value=str(value if value is not None else "—")
        return value.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    def money_pdf(value):
        return "BDT " + format(float(value or 0), ",.2f")

    out=io.BytesIO()
    doc=SimpleDocTemplate(out,pagesize=A4,rightMargin=16*mm,leftMargin=16*mm,topMargin=33*mm,bottomMargin=19*mm,title=f"{memo_title} {p['purchase_no']}",author=str(company.get("name") or "Alien Tees"))
    green=colors.HexColor(memo_cfg.get("color") or "#0D7552");dark=colors.HexColor("#172B24");muted=colors.HexColor("#5C7067");pale=colors.HexColor("#EAF5EE");line=colors.HexColor("#D8E5DD");white=colors.white
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name="MemoTitle",parent=styles["Title"],fontName="Helvetica-Bold",fontSize=17,leading=21,textColor=dark,alignment=TA_LEFT,spaceAfter=4))
    styles.add(ParagraphStyle(name="MemoSub",parent=styles["Normal"],fontName="Helvetica",fontSize=8.5,leading=12,textColor=muted))
    styles.add(ParagraphStyle(name="MemoCell",parent=styles["Normal"],fontName="Helvetica",fontSize=8.5,leading=11,textColor=dark,wordWrap="CJK"))
    styles.add(ParagraphStyle(name="MemoCellBold",parent=styles["MemoCell"],fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="MemoHeader",parent=styles["MemoCellBold"],textColor=colors.white,fontSize=7.5,leading=9))
    styles.add(ParagraphStyle(name="MemoRight",parent=styles["MemoCell"],alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name="MemoRightBold",parent=styles["MemoCellBold"],alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name="MemoSection",parent=styles["Heading2"],fontName="Helvetica-Bold",fontSize=10.5,leading=13,textColor=green,spaceBefore=8,spaceAfter=6))
    def para(t,style="MemoCell"):
        return Paragraph(esc_pdf(t),styles[style])
    def on_page(canvas,document):
        canvas.saveState();w,h=A4
        canvas.setFillColor(green);canvas.rect(0,h-25*mm,w,25*mm,fill=1,stroke=0)
        left_x=16*mm
        if logo_data:
            try:
                _mime, logo_bytes = logo_data
                logo_reader=ImageReader(io.BytesIO(logo_bytes))
                canvas.drawImage(logo_reader,15*mm,h-22*mm,width=19*mm,height=19*mm,preserveAspectRatio=True,anchor='c',mask='auto')
                left_x=38*mm
            except Exception: left_x=16*mm
        company_name=str(company.get("name") or "Alien Tees")
        canvas.setFillColor(white);canvas.setFont("Helvetica-Bold",16);canvas.drawString(left_x,h-10*mm,company_name[:55])
        canvas.setFont("Helvetica",9);canvas.drawRightString(w-16*mm,h-10*mm,memo_title[:40])
        line2=str(company.get("phone") or company.get("website") or "Supplier Purchase Record")
        canvas.setFont("Helvetica",8);canvas.drawString(left_x,h-18*mm,line2[:70])
        canvas.drawRightString(w-16*mm,h-18*mm,"Generated: "+datetime.now(DHAKA_TZ).strftime("%d %b %Y, %I:%M %p"))
        canvas.setStrokeColor(line);canvas.line(16*mm,13*mm,w-16*mm,13*mm)
        canvas.setFillColor(muted);canvas.setFont("Helvetica",7.5);canvas.drawString(16*mm,8*mm,(memo_footer or (company_name+" · Purchase Management"))[:100])
        canvas.drawRightString(w-16*mm,8*mm,f"Page {document.page}")
        canvas.restoreState()

    story=[Spacer(1,2*mm),Paragraph(esc_pdf(memo_title),styles["MemoTitle"]),Paragraph(esc_pdf(memo_subtitle),styles["MemoSub"]),Spacer(1,5*mm)]
    meta_data=[
        [para("PURCHASE NO.","MemoCellBold"),para(p.get("purchase_no"),"MemoCellBold"),para("PURCHASE DATE","MemoCellBold"),para(p.get("purchase_date"))],
        [para("SUPPLIER","MemoCellBold"),para(p.get("supplier_name")),para("SUPPLIER INVOICE","MemoCellBold"),para(p.get("invoice_no") or "Not provided")],
        [para("PHONE","MemoCellBold"),para(p.get("supplier_phone") or "Not provided"),para("STOCK STATUS","MemoCellBold"),para("Received" if p.get("stock_received") else "Not received")],
        [para("ADDRESS","MemoCellBold"),para(p.get("supplier_address") or "Not provided"),para("RECORDED BY","MemoCellBold"),para(p.get("created_by") or "System")],
    ]
    meta=Table(meta_data,colWidths=[25*mm,64*mm,31*mm,58*mm],hAlign="LEFT")
    meta.setStyle(TableStyle([("BACKGROUND",(0,0),(0,-1),pale),("BACKGROUND",(2,0),(2,-1),pale),("GRID",(0,0),(-1,-1),0.45,line),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)]))
    story.extend([meta,Spacer(1,5*mm),Paragraph("PURCHASED ITEMS",styles["MemoSection"])])
    headers=[para("#","MemoHeader"),para("PRODUCT CODE","MemoHeader"),para("PRODUCT NAME","MemoHeader"),para("SIZE / VARIANT","MemoHeader"),para("QTY","MemoHeader"),para("UNIT COST","MemoHeader"),para("LINE TOTAL","MemoHeader")]
    rows=[headers]
    for idx,x in enumerate(items,1):
        name=str(x.get("item_name") or "").strip();code=str(x.get("item_code") or "").strip();name=_erp_product_title_from_code(code) if not name or name.casefold()==code.casefold() else name;rows.append([para(str(idx)),para(code or "—","MemoCellBold"),para(name or code or "—"),para(x.get("size") or "—"),para(str(x.get("quantity") or 0),"MemoRight"),para(money_pdf(x.get("unit_cost")),"MemoRight"),para(money_pdf(x.get("line_total")),"MemoRight")])
    if len(rows)==1: rows.append([para("—"),para("—"),para("No item rows recorded"),para("—"),para("0"),para(money_pdf(0)),para(money_pdf(0))])
    item_table=Table(rows,colWidths=[7*mm,30*mm,47*mm,25*mm,11*mm,27*mm,31*mm],repeatRows=1,hAlign="LEFT")
    item_table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),green),("TEXTCOLOR",(0,0),(-1,0),white),("GRID",(0,0),(-1,-1),0.45,line),("ROWBACKGROUNDS",(0,1),(-1,-1),[white,colors.HexColor("#F8FBF9")]),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)]))
    story.extend([item_table,Spacer(1,5*mm)])
    totals_data=[
        [para("Subtotal"),para(money_pdf(p.get("subtotal")),"MemoRight")],
        [para("Discount"),para(money_pdf(p.get("discount")),"MemoRight")],
        [para("INVOICE TOTAL","MemoCellBold"),para(money_pdf(p.get("total")),"MemoRightBold")],
        [para("Paid to Date"),para(money_pdf(p.get("paid")),"MemoRight")],
        [para("BALANCE DUE","MemoCellBold"),para(money_pdf(p.get("due")),"MemoRightBold")],
    ]
    totals=Table(totals_data,colWidths=[47*mm,45*mm],hAlign="RIGHT")
    totals.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.4,line),("BACKGROUND",(0,2),(-1,2),pale),("BACKGROUND",(0,4),(-1,4),colors.HexColor("#FFF4E5")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
    purchase_qr_payload = (purchase_qr_payloads or ["PURCHASE MEMO"])[0]
    qr_card = Table([[PurchaseMemoQR(purchase_qr_payload, size=24*mm, caption="")]], colWidths=[50*mm], hAlign="CENTER")
    qr_card.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F5FAF6")),
        ("BOX",(0,0),(-1,-1),0.8,colors.HexColor("#CFE2D4")),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),1*mm),
        ("RIGHTPADDING",(0,0),(-1,-1),1*mm),
        ("TOPPADDING",(0,0),(-1,-1),2*mm),
        ("BOTTOMPADDING",(0,0),(-1,-1),3*mm),
    ]))
    summary_panel = Table([[qr_card, totals]], colWidths=[60*mm, 105*mm], hAlign="LEFT")
    # Place the purchase QR in the former summary position and PAYMENT SUMMARY at the former QR position.
    summary_panel.setStyle(TableStyle([("ALIGN",(0,0),(0,0),"LEFT"),("ALIGN",(1,0),(1,0),"RIGHT")]))
    summary_panel.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("ALIGN",(0,0),(0,0),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),2*mm),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    story.extend([summary_panel, Spacer(1, 3*mm)])
    if p.get("notes"): 
        story.extend([Spacer(1,4*mm),Paragraph("NOTES",styles["MemoSection"]),Paragraph(esc_pdf(p.get("notes")).replace("\n","<br/>"),styles["MemoCell"])])
    if payments:
        story.extend([Spacer(1,4*mm),Paragraph("PAYMENT RECORD",styles["MemoSection"])])
        pay_rows=[[para("DATE","MemoCellBold"),para("METHOD","MemoCellBold"),para("REFERENCE","MemoCellBold"),para("AMOUNT","MemoCellBold")]]
        for pay in payments:
            pay_rows.append([para(pay.get("payment_date")),para(pay.get("payment_method")),para(pay.get("reference") or "—"),para(money_pdf(pay.get("amount")),"MemoRight")])
        pay_table=Table(pay_rows,colWidths=[30*mm,35*mm,68*mm,45*mm],repeatRows=1)
        pay_table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),pale),("GRID",(0,0),(-1,-1),0.4,line),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
        story.append(pay_table)
    story.extend([Spacer(1,12*mm)])
    sig_data=[[para("____________________________"),para("____________________________"),para("____________________________")],[para("PREPARED BY","MemoSub"),para("AUTHORIZED BY","MemoSub"),para("RECEIVED BY","MemoSub")]]
    sig=Table(sig_data,colWidths=[59*mm,59*mm,59*mm],hAlign="LEFT")
    sig.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    story.append(KeepTogether(sig))
    doc.build(story,onFirstPage=on_page,onLaterPages=on_page)
    out.seek(0)
    safe_no=re.sub(r"[^A-Za-z0-9_-]+","-",str(p.get("purchase_no") or purchase_id)).strip("-")[:70] or str(purchase_id)
    return send_file(out,mimetype="application/pdf",as_attachment=(request.args.get("download") == "1"),download_name=f"Purchase-Memo-{safe_no}.pdf",conditional=False)

@app.get("/memo/purchase/<int:purchase_id>")
def public_erp_purchase_memo(purchase_id):
    """Backward-compatible token URL: validate token then return the PDF document itself."""
    supplied=str(request.args.get("t") or "")
    with db() as c: token_row=c.execute("SELECT token FROM erp_purchase_memo_tokens WHERE purchase_id=?",(purchase_id,)).fetchone()
    if not supplied or not token_row or not secrets.compare_digest(supplied,str(token_row["token"])):
        return Response("Memo not found",status=404,mimetype="text/plain")
    return erp_purchase_memo_pdf(purchase_id)

@app.get("/api/erp/purchases/export.xlsx")
@erp_module_auth("purchase_view","purchase_manage")
def erp_purchase_export():
    frm=_erp_clean_day(request.args.get("from"));to=_erp_clean_day(request.args.get("to"));rows,_=_erp_purchase_records("",frm,to,1000)
    wb=Workbook();ws=wb.active;ws.title="Purchases";ws.append(["Purchase No.","Date","Supplier","Supplier Invoice","Item Count","Quantity","Subtotal","Discount","Invoice Total","Paid on Invoice","Invoice Due","Stock Received","Created By","Notes"])
    with db() as c:
        for r in rows:
            notes=c.execute("SELECT item_name,item_code,size,quantity,unit_cost,line_total FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(r["id"],)).fetchall()
            item_names="; ".join(f"{x['item_name']} [{x['item_code'] or ''} / {x['size'] or ''}] ×{x['quantity']} @ {x['unit_cost']}" for x in notes)
            ws.append([r["purchase_no"],r["purchase_date"],r["supplier_name"],r["invoice_no"],r["item_count"],r["quantity"],r["subtotal"],r["discount"],r["total"],r["paid"],r["due"],"Yes" if r["stock_received"] else "No",r["created_by"],r["notes"]+((" | "+item_names) if item_names else "")])
    _erp_style_workbook(wb);out=io.BytesIO();wb.save(out);out.seek(0);return send_file(out,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name="Alien-Tees-Purchase-Report.xlsx")

@app.get("/api/erp/expenses/export.xlsx")
@erp_module_auth("expense_view","expense_manage")
def erp_expense_export():
    frm=_erp_clean_day(request.args.get("from"));to=_erp_clean_day(request.args.get("to"));rows,_=_erp_expense_records("",frm,to,1000)
    wb=Workbook();ws=wb.active;ws.title="Expenses";ws.append(["Expense No.","Date","Title","Category","Amount (BDT)","Payment Method","Paid To","Reference","Status","Created By","Notes"])
    for r in rows:ws.append([r["expense_no"],r["expense_date"],r["title"],r["category"],r["amount"],r["payment_method"],r["paid_to"],r["reference"],r["status"],r["created_by"],r["notes"]])
    _erp_style_workbook(wb);out=io.BytesIO();wb.save(out);out.seek(0);return send_file(out,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name="Alien-Tees-Expense-Report.xlsx")

@app.get("/api/erp/suppliers/<int:supplier_id>/purchases")
@erp_module_auth("purchase_view","purchase_manage","supplier_manage")
def erp_supplier_purchase_history(supplier_id):
    try: limit=max(1,min(1000,int(request.args.get("limit",1000) or 1000)))
    except (TypeError,ValueError): limit=1000
    with db() as c:
        supplier=c.execute("SELECT id,name,phone,email,address,opening_balance FROM erp_suppliers WHERE id=?",(supplier_id,)).fetchone()
        if not supplier:return jsonify(error="Supplier not found."),404
        rows=[dict(r) for r in c.execute("""SELECT p.*,
          COALESCE((SELECT COUNT(*) FROM erp_purchase_items i WHERE i.purchase_id=p.id),0) item_count,
          COALESCE((SELECT SUM(i.quantity) FROM erp_purchase_items i WHERE i.purchase_id=p.id),0) quantity,
          COALESCE((SELECT SUM(pm.amount) FROM erp_supplier_payments pm WHERE pm.purchase_id=p.id),0) paid
          FROM erp_purchases p WHERE p.supplier_id=? ORDER BY p.purchase_date DESC,p.id DESC LIMIT ?""",(supplier_id,limit)).fetchall()]
        for row in rows:
            row["paid"]=round(float(row.get("paid") or 0),2);row["due"]=round(float(row.get("total") or 0)-row["paid"],2)
            row["items"]=[dict(x) for x in c.execute("SELECT item_name,item_code,size,quantity,unit_cost,line_total,stock_sheet FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(row["id"],)).fetchall()]
    return jsonify(supplier=dict(supplier),purchases=rows,count=len(rows))

@app.get("/api/erp/suppliers/<int:supplier_id>/purchases/export.xlsx")
@erp_module_auth("purchase_view","purchase_manage","supplier_manage")
def erp_supplier_purchase_export(supplier_id):
    with db() as c:
        supplier=c.execute("SELECT * FROM erp_suppliers WHERE id=?",(supplier_id,)).fetchone()
        if not supplier:return jsonify(error="Supplier not found."),404
        purchases=[dict(r) for r in c.execute("SELECT * FROM erp_purchases WHERE supplier_id=? ORDER BY purchase_date DESC,id DESC",(supplier_id,)).fetchall()]
        items=[]
        for purchase in purchases:
            paid=float(c.execute("SELECT COALESCE(SUM(amount),0) FROM erp_supplier_payments WHERE purchase_id=?",(purchase["id"],)).fetchone()[0] or 0)
            purchase["paid"]=round(paid,2);purchase["due"]=round(float(purchase.get("total") or 0)-paid,2)
            for item in c.execute("SELECT item_name,item_code,size,quantity,unit_cost,line_total,stock_sheet FROM erp_purchase_items WHERE purchase_id=? ORDER BY id",(purchase["id"],)).fetchall():
                items.append([purchase["purchase_no"],purchase["purchase_date"],purchase["invoice_no"],item["stock_sheet"],item["item_code"],item["item_name"],item["size"],item["quantity"],item["unit_cost"],item["line_total"]])
        payments=[dict(r) for r in c.execute("SELECT p.purchase_no,x.amount,x.payment_date,x.payment_method,x.reference,x.notes,x.created_by FROM erp_supplier_payments x LEFT JOIN erp_purchases p ON p.id=x.purchase_id WHERE x.supplier_id=? ORDER BY x.payment_date DESC,x.id DESC",(supplier_id,)).fetchall()]
    wb=Workbook();ws=wb.active;ws.title="Purchase Invoices";ws.append(["Purchase No.","Date","Supplier Invoice","Item Count","Quantity","Subtotal","Discount","Invoice Total","Paid","Invoice Due","Stock Received","Created By","Notes"])
    with db() as c:
        for p in purchases:
            counts=c.execute("SELECT COUNT(*) count,COALESCE(SUM(quantity),0) qty FROM erp_purchase_items WHERE purchase_id=?",(p["id"],)).fetchone()
            ws.append([p["purchase_no"],p["purchase_date"],p["invoice_no"],counts["count"],counts["qty"],p["subtotal"],p["discount"],p["total"],p["paid"],p["due"],"Yes" if p["stock_received"] else "No",p["created_by"],p["notes"]])
    wi=wb.create_sheet("Purchased Items");wi.append(["Purchase No.","Date","Supplier Invoice","Stock Sheet","Product Code","Product Name","Size / Variant","Quantity","Unit Cost (BDT)","Line Total (BDT)"])
    for row in items:wi.append(row)
    wp=wb.create_sheet("Payments");wp.append(["Purchase No.","Payment Date","Amount (BDT)","Method","Reference","Notes","Recorded By"])
    for x in payments:wp.append([x["purchase_no"] or "Supplier-level / opening balance",x["payment_date"],x["amount"],x["payment_method"],x["reference"],x["notes"],x["created_by"]])
    _erp_style_workbook(wb);out=io.BytesIO();wb.save(out);out.seek(0)
    safe_name=re.sub(r"[^A-Za-z0-9_-]+","-",str(supplier["name"] or "Supplier")).strip("-")[:60] or f"Supplier-{supplier_id}"
    return send_file(out,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name=f"{safe_name}-Purchase-History.xlsx")

@app.get("/api/erp/suppliers/export.xlsx")
@erp_module_auth("purchase_view","purchase_manage","supplier_manage")
def erp_supplier_export():
    rows,_=_erp_supplier_rows(True);wb=Workbook();ws=wb.active;ws.title="Suppliers";ws.append(["Supplier","Phone","Email","Address","Opening Balance","Purchases","Payments","Current Due","Status","Notes"])
    for r in rows:ws.append([r["name"],r["phone"],r["email"],r["address"],r["opening_balance"],r["purchases"],r["paid"],r["due"],"Active" if r["active"] else "Inactive",r["notes"]])
    _erp_style_workbook(wb);out=io.BytesIO();wb.save(out);out.seek(0);return send_file(out,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name="Alien-Tees-Suppliers.xlsx")

def _erp_style_workbook(wb):
    from openpyxl.styles import Font, PatternFill, Alignment
    for ws in wb.worksheets:
        ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:cell.font=Font(bold=True,color="FFFFFF");cell.fill=PatternFill("solid",fgColor="0D7552");cell.alignment=Alignment(vertical="center")
        for col in ws.columns:
            width=max((len(str(cell.value)) for cell in col if cell.value is not None),default=8)
            ws.column_dimensions[col[0].column_letter].width=min(48,max(12,width+2))


# ===================== HR MANAGEMENT MODULE =====================
def hr_module_auth(permission="hr_view"):
    def deco(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if not session.get("user"):
                return jsonify(ok=False, error="Please sign in to access HR Management."), 401
            if session.get("role") in ("admin", "super_admin") or has_perm(permission) or (permission == "hr_view" and (has_perm("hr_manage") or has_perm("hr_payroll") or has_perm("hr_leave_approve"))):
                return f(*args, **kwargs)
            return jsonify(ok=False, error="You do not have permission to access this HR feature."), 403
        return wrapped
    return deco

_HR_UPLOAD_DIR = os.path.join(INSTANCE, "hr_uploads")
os.makedirs(_HR_UPLOAD_DIR, exist_ok=True)

def _hr_save_photo(data_url, prefix):
    if not data_url:
        return ""
    if not isinstance(data_url, str) or "," not in data_url or not data_url.startswith("data:image/"):
        raise ValueError("Invalid image data.")
    header, payload = data_url.split(",", 1)
    if len(payload) > 7 * 1024 * 1024:
        raise ValueError("Image is too large. Please use an image under 5 MB.")
    try:
        raw = base64.b64decode(payload, validate=True)
        im = Image.open(io.BytesIO(raw)).convert("RGB")
        im.thumbnail((1000, 1000))
        out = io.BytesIO(); im.save(out, "JPEG", quality=86, optimize=True)
        filename = f"hr_{prefix}_{int(time.time()*1000)}_{os.urandom(3).hex()}.jpg"
        path = os.path.join(_HR_UPLOAD_DIR, filename)
        with open(path, "wb") as fh: fh.write(out.getvalue())
        return filename
    except Exception as exc:
        raise ValueError("The uploaded photo could not be processed.") from exc

def _hr_json_employee(row):
    d = dict(row)
    d["salary"] = round(float(d.get("salary") or 0), 2)
    d["overtime_rate"] = round(float(d.get("overtime_rate") or 0), 2)
    d["shift_minutes"] = int(d.get("shift_minutes") or 0)
    d["shift_hours"] = round(d["shift_minutes"] / 60, 2) if d["shift_minutes"] else 0
    return d

def _hr_shift_minutes():
    """Configured normal daily shift duration; default is 8 hours."""
    try: minutes = int(float(get_setting("hr_shift_minutes") or 480))
    except (TypeError, ValueError): minutes = 480
    return max(60, min(1440, minutes))

def _hr_late_rules():
    """Global monthly late-arrival grace and deduction-per-extra-late-day policy."""
    try: allowed_days = int(float(get_setting("hr_late_allowed_days") or 0))
    except (TypeError, ValueError): allowed_days = 3
    try: deduction_per_day = float(get_setting("hr_late_deduction_per_day") or 0)
    except (TypeError, ValueError): deduction_per_day = 100.0
    return {"allowed_days":max(0,min(31,allowed_days)),"deduction_per_day":round(max(0.0,min(1e9,deduction_per_day)),2)}

def _hr_attendance_minutes(check_in, check_out):
    try:
        a = datetime.strptime(check_in, "%H:%M"); b = datetime.strptime(check_out, "%H:%M")
        delta = int(round((b-a).total_seconds()/60))
        if delta < 0: delta += 24 * 60
        return max(0, min(24 * 60, delta))
    except Exception:
        return 0

def _hr_attendance_hours(check_in, check_out):
    return round(_hr_attendance_minutes(check_in, check_out) / 60, 2)

def _hr_duration_label(minutes):
    try: minutes = max(0, int(round(float(minutes or 0))))
    except (TypeError, ValueError): minutes = 0
    return f"{minutes // 60}h {minutes % 60:02d}m"

@app.post("/api/hr/settings/shift-hours")
@hr_module_auth("hr_manage")
def hr_shift_hours_save():
    j = request.get_json(silent=True) or {}
    try: hours = float(j.get("hours"))
    except (TypeError, ValueError): return jsonify(ok=False, error="Enter standard shift hours between 1 and 24."), 400
    if not (1 <= hours <= 24): return jsonify(ok=False, error="Standard shift hours must be between 1 and 24."), 400
    minutes = int(round(hours * 60))
    with db() as c:
        c.execute("INSERT OR REPLACE INTO settings(k,v) VALUES('hr_shift_minutes',?)", (str(minutes),))
    log_activity("hr_shift_settings", "HR Management", f"standard_shift_minutes={minutes}")
    return jsonify(ok=True, shift_minutes=minutes, shift_hours=round(minutes/60, 2))

@app.post("/api/hr/settings/late-rules")
@hr_module_auth("hr_manage")
def hr_late_rules_save():
    j=request.get_json(silent=True) or {}
    try:
        allowed_days=int(j.get("allowed_days"))
        deduction=float(j.get("deduction_per_day"))
        if allowed_days<0 or allowed_days>31 or not (0<=deduction<=1e9): raise ValueError
    except (TypeError,ValueError):
        return jsonify(ok=False,error="Allowed late days must be 0–31 and deduction must be a valid non-negative amount."),400
    with db() as c:
        c.execute("INSERT OR REPLACE INTO settings(k,v) VALUES('hr_late_allowed_days',?)",(str(allowed_days),))
        c.execute("INSERT OR REPLACE INTO settings(k,v) VALUES('hr_late_deduction_per_day',?)",(str(round(deduction,2)),))
    log_activity("hr_late_rules", "HR Management", f"allowed_days={allowed_days};deduction_per_day={round(deduction,2)}")
    return jsonify(ok=True,allowed_days=allowed_days,deduction_per_day=round(deduction,2))

@app.get("/api/hr/overview")
@hr_module_auth("hr_view")
def hr_overview():
    today = datetime.now(DHAKA_TZ).strftime("%Y-%m-%d")
    month = datetime.now(DHAKA_TZ).strftime("%Y-%m")
    with db() as c:
        total = c.execute("SELECT COUNT(*) FROM hr_employees").fetchone()[0]
        active = c.execute("SELECT COUNT(*) FROM hr_employees WHERE status='Active'").fetchone()[0]
        present = c.execute("SELECT COUNT(*) FROM hr_attendance WHERE attendance_date=? AND status IN ('Present','Late','Half Day')", (today,)).fetchone()[0]
        pending = c.execute("SELECT COUNT(*) FROM hr_leaves WHERE status='Pending'").fetchone()[0]
        payroll_total = c.execute("SELECT COALESCE(SUM(net_pay),0) FROM hr_payroll WHERE pay_month=?", (month,)).fetchone()[0]
        employees = [dict(r) for r in c.execute("SELECT * FROM hr_employees ORDER BY CASE WHEN status='Active' THEN 0 ELSE 1 END,full_name COLLATE NOCASE LIMIT 500").fetchall()]
        attendance = [dict(r) for r in c.execute("SELECT a.*,e.employee_code,e.full_name,e.department FROM hr_attendance a JOIN hr_employees e ON e.id=a.employee_id ORDER BY a.attendance_date DESC,a.id DESC LIMIT 150").fetchall()]
        leaves = [dict(r) for r in c.execute("SELECT l.*,e.employee_code,e.full_name,e.department FROM hr_leaves l JOIN hr_employees e ON e.id=l.employee_id ORDER BY CASE l.status WHEN 'Pending' THEN 0 ELSE 1 END,l.created_at DESC LIMIT 150").fetchall()]
        payroll = [dict(r) for r in c.execute("SELECT p.*,e.employee_code,e.full_name,e.department FROM hr_payroll p JOIN hr_employees e ON e.id=p.employee_id ORDER BY p.pay_month DESC,p.id DESC LIMIT 150").fetchall()]
    for e in employees: e["salary"] = round(float(e.get("salary") or 0), 2)
    late_rules=_hr_late_rules()
    return jsonify(ok=True, summary={"employees": total, "active_employees": active, "present_today": present, "pending_leaves": pending, "payroll_month_total": round(float(payroll_total or 0), 2), "shift_hours": round(_hr_shift_minutes()/60, 2), "late_allowed_days":late_rules["allowed_days"], "late_deduction_per_day":late_rules["deduction_per_day"]}, employees=employees, attendance=attendance, leaves=leaves, payroll=payroll)

@app.get("/api/hr/employees")
@hr_module_auth("hr_view")
def hr_employees_list():
    q = str(request.args.get("q", "") or "").strip().lower()
    with db() as c:
        if q:
            rows = c.execute("SELECT * FROM hr_employees WHERE lower(employee_code) LIKE ? OR lower(full_name) LIKE ? OR lower(department) LIKE ? OR lower(phone) LIKE ? ORDER BY full_name LIMIT 500", tuple([f"%{q}%"]*4)).fetchall()
        else:
            rows = c.execute("SELECT * FROM hr_employees ORDER BY full_name LIMIT 500").fetchall()
    return jsonify(ok=True, employees=[_hr_json_employee(r) for r in rows])

@app.get("/api/hr/employees/<int:employee_id>")
@hr_module_auth("hr_view")
def hr_employee_detail(employee_id):
    with db() as c: row = c.execute("SELECT * FROM hr_employees WHERE id=?", (employee_id,)).fetchone()
    if not row: return jsonify(ok=False, error="Employee not found."), 404
    return jsonify(ok=True, employee=_hr_json_employee(row))

@app.post("/api/hr/employees")
@hr_module_auth("hr_manage")
def hr_employee_save():
    j = request.get_json(silent=True) or {}
    name = str(j.get("full_name") or "").strip()[:120]
    if not name: return jsonify(ok=False, error="Employee full name is required."), 400
    try:
        salary = float(j.get("salary") or 0)
        if not (salary >= 0 and salary < 1e12): raise ValueError
    except (TypeError, ValueError): return jsonify(ok=False, error="Enter a valid monthly salary."), 400
    try:
        overtime_rate = float(j.get("overtime_rate") or 0)
        if not (0 <= overtime_rate < 1e9): raise ValueError
    except (TypeError, ValueError): return jsonify(ok=False, error="Overtime hourly rate must be a valid non-negative amount."), 400
    try:
        shift_hours = float(j.get("shift_hours") or 0)
        if not (shift_hours == 0 or 1 <= shift_hours <= 24): raise ValueError
    except (TypeError, ValueError): return jsonify(ok=False, error="Daily shift hours must be 0 (use global setting) or between 1 and 24."), 400
    shift_minutes = int(round(shift_hours * 60)) if shift_hours else 0
    status = str(j.get("status") or "Active")
    if status not in ("Active", "Inactive"): status = "Active"
    code = str(j.get("employee_code") or "").strip()[:40]
    emp_id = int(j.get("id") or 0)
    now = time.time(); photo_file = ""
    try: photo_file = _hr_save_photo(j.get("photo_data"), "employee") if j.get("photo_data") else ""
    except ValueError as e: return jsonify(ok=False, error=str(e)), 400
    data = (name, str(j.get("phone") or "").strip()[:40], str(j.get("email") or "").strip()[:180], str(j.get("department") or "").strip()[:100], str(j.get("designation") or "").strip()[:100], str(j.get("employment_type") or "Full-time")[:40], str(j.get("join_date") or "")[:10], salary, overtime_rate, shift_minutes, status, str(j.get("address") or "")[:500], str(j.get("emergency_contact") or "")[:100], str(j.get("notes") or "")[:800])
    try:
        with db() as c:
            if emp_id:
                old = c.execute("SELECT * FROM hr_employees WHERE id=?", (emp_id,)).fetchone()
                if not old: return jsonify(ok=False, error="Employee record not found."), 404
                code = code or old["employee_code"]
                photo_file = photo_file or old["photo_file"]
                c.execute("UPDATE hr_employees SET employee_code=?,full_name=?,phone=?,email=?,department=?,designation=?,employment_type=?,join_date=?,salary=?,overtime_rate=?,shift_minutes=?,status=?,address=?,emergency_contact=?,notes=?,photo_file=?,updated_at=? WHERE id=?", (code, *data, photo_file, now, emp_id))
                saved_id = emp_id
            else:
                code = code or ("AT-" + datetime.now(DHAKA_TZ).strftime("%y%m%d") + "-" + os.urandom(3).hex().upper())
                c.execute("INSERT INTO hr_employees(employee_code,full_name,phone,email,department,designation,employment_type,join_date,salary,overtime_rate,shift_minutes,status,address,emergency_contact,notes,photo_file,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (code, *data, photo_file, session.get("user", ""), now, now))
                saved_id = c.execute("SELECT last_insert_rowid()").fetchone()[0]
    except sqlite3.IntegrityError:
        return jsonify(ok=False, error="That Employee Code is already in use."), 409
    log_activity("hr_employee_save", code, f"employee_id={saved_id}")
    return jsonify(ok=True, id=saved_id, employee_code=code)

@app.post("/api/hr/employees/<int:employee_id>/status")
@hr_module_auth("hr_manage")
def hr_employee_status(employee_id):
    j = request.get_json(silent=True) or {}; status = str(j.get("status") or "")
    if status not in ("Active", "Inactive"): return jsonify(ok=False, error="Status must be Active or Inactive."), 400
    with db() as c:
        cur = c.execute("UPDATE hr_employees SET status=?,updated_at=? WHERE id=?", (status, time.time(), employee_id))
        if not cur.rowcount: return jsonify(ok=False, error="Employee not found."), 404
    log_activity("hr_employee_status", str(employee_id), status)
    return jsonify(ok=True)

@app.get("/api/hr/photo/<path:filename>")
@hr_module_auth("hr_view")
def hr_photo(filename):
    if not re.fullmatch(r"hr_(?:employee|attendance)_[A-Za-z0-9_-]+\.jpg", filename): return jsonify(ok=False, error="Invalid photo name."), 400
    path = os.path.join(_HR_UPLOAD_DIR, filename)
    if not os.path.isfile(path): return jsonify(ok=False, error="Photo not found."), 404
    return send_file(path, mimetype="image/jpeg", conditional=True, max_age=300)

@app.get("/api/hr/attendance-report")
@hr_module_auth("hr_view")
def hr_attendance_report():
    month = str(request.args.get("month") or datetime.now(DHAKA_TZ).strftime("%Y-%m")).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", month): return jsonify(ok=False, error="Select a valid month."), 400
    try:
        month_start = datetime.strptime(month, "%Y-%m").replace(day=1).strftime("%Y-%m-%d")
        next_month_dt = (datetime.strptime(month, "%Y-%m").replace(day=28) + timedelta(days=4)).replace(day=1)
        month_end = next_month_dt.strftime("%Y-%m-%d")
        employee_id = int(request.args.get("employee_id") or 0)
        if employee_id < 0: raise ValueError
    except (TypeError, ValueError): return jsonify(ok=False, error="Invalid employee or month filter."), 400
    sql = "SELECT a.*,e.employee_code,e.full_name,e.department FROM hr_attendance a JOIN hr_employees e ON e.id=a.employee_id WHERE a.attendance_date>=? AND a.attendance_date<?"
    args = [month_start, month_end]
    if employee_id:
        sql += " AND a.employee_id=?"; args.append(employee_id)
    sql += " ORDER BY a.attendance_date DESC,e.full_name COLLATE NOCASE,a.id DESC"
    with db() as c:
        rows = [dict(r) for r in c.execute(sql, args).fetchall()]
    for row in rows:
        if not row.get("worked_minutes") and row.get("work_hours"):
            row["worked_minutes"] = int(round(float(row["work_hours"] or 0)*60))
        row["shift_minutes"] = int(row.get("shift_minutes") or _hr_shift_minutes())
        row["overtime_minutes"] = int(row.get("overtime_minutes") or max(0, int(row.get("worked_minutes") or 0)-row["shift_minutes"]))
    return jsonify(ok=True, month=month, employee_id=employee_id or None, rows=rows,
                   total_worked_minutes=sum(int(r.get("worked_minutes") or 0) for r in rows),
                   total_overtime_minutes=sum(int(r.get("overtime_minutes") or 0) for r in rows))

@app.get("/api/hr/attendance-export.xlsx")
@hr_module_auth("hr_view")
def hr_attendance_export_xlsx():
    month = str(request.args.get("month") or datetime.now(DHAKA_TZ).strftime("%Y-%m")).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", month): return jsonify(ok=False, error="Select a valid month."), 400
    try:
        month_start = datetime.strptime(month, "%Y-%m").replace(day=1).strftime("%Y-%m-%d")
        next_month_dt = (datetime.strptime(month, "%Y-%m").replace(day=28) + timedelta(days=4)).replace(day=1)
        month_end = next_month_dt.strftime("%Y-%m-%d")
        employee_id = int(request.args.get("employee_id") or 0)
        if employee_id < 0: raise ValueError
    except (TypeError, ValueError): return jsonify(ok=False, error="Invalid employee or month filter."), 400
    sql = "SELECT a.attendance_date,e.employee_code,e.full_name,e.department,e.designation,a.check_in,a.check_out,a.worked_minutes,a.shift_minutes,a.overtime_minutes,a.status,a.notes,a.photo_file FROM hr_attendance a JOIN hr_employees e ON e.id=a.employee_id WHERE a.attendance_date>=? AND a.attendance_date<?"
    args = [month_start, month_end]
    employee_name = "All Employees"
    if employee_id:
        sql += " AND a.employee_id=?"; args.append(employee_id)
    sql += " ORDER BY a.attendance_date,e.full_name COLLATE NOCASE,a.id"
    wb = Workbook(); ws = wb.active; ws.title = "Monthly Attendance"
    ws.append(["ALIEN TEES - Monthly Attendance Report"])
    ws.append(["Month", month, "Employee ID", employee_id or "All Employees"])
    ws.append(["Date", "Employee Code", "Employee Name", "Department", "Designation", "Check In", "Check Out", "Worked Duration", "Worked Minutes", "Shift / OT Limit", "Shift Minutes", "Overtime Duration", "Overtime Minutes", "Status", "Notes", "Photo File"])
    with db() as c:
        if employee_id:
            er = c.execute("SELECT full_name FROM hr_employees WHERE id=?", (employee_id,)).fetchone()
            if not er: return jsonify(ok=False, error="Employee not found."), 404
            employee_name = er["full_name"]
        for r in c.execute(sql, args):
            d = dict(r); worked = int(d.get("worked_minutes") or round(float(d.get("work_hours") or 0)*60)); shift = int(d.get("shift_minutes") or _hr_shift_minutes()); overtime = int(d.get("overtime_minutes") or max(0, worked-shift))
            ws.append([d["attendance_date"],d["employee_code"],d["full_name"],d["department"],d["designation"],d["check_in"],d["check_out"],_hr_duration_label(worked),worked,_hr_duration_label(shift),shift,_hr_duration_label(overtime),overtime,d["status"],d["notes"],d["photo_file"]])
    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:P{max(3,ws.max_row)}"
    _erp_style_workbook(wb); out=io.BytesIO(); wb.save(out); out.seek(0)
    safe_emp = re.sub(r"[^A-Za-z0-9_-]+", "_", employee_name).strip("_") or "Employee"
    log_activity("hr_attendance_export", month, f"employee={employee_id or 'all'};rows={max(0,ws.max_row-3)}")
    filename=f"Alien-Tees-Attendance-{month}-{safe_emp}.xlsx"
    return send_file(out,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name=filename)

@app.post("/api/hr/attendance")
@hr_module_auth("hr_manage")
def hr_attendance_save():
    j = request.get_json(silent=True) or {}; emp = int(j.get("employee_id") or 0)
    date = str(j.get("attendance_date") or datetime.now(DHAKA_TZ).strftime("%Y-%m-%d"))[:10]
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date): return jsonify(ok=False, error="Select a valid attendance date."), 400
    action = str(j.get("action") or "manual")
    status = str(j.get("status") or "Present")
    if status not in ("Present", "Late", "Half Day", "Absent", "Leave", "Holiday"): status = "Present"
    now_text = datetime.now(DHAKA_TZ).strftime("%H:%M")
    with db() as c:
        employee = c.execute("SELECT id,shift_minutes FROM hr_employees WHERE id=? AND status='Active'", (emp,)).fetchone()
        if not employee: return jsonify(ok=False, error="Choose an active employee."), 400
        old = c.execute("SELECT * FROM hr_attendance WHERE employee_id=? AND attendance_date=?", (emp, date)).fetchone()
        cin = str(j.get("check_in") or (old["check_in"] if old else ""))[:5]
        cout = str(j.get("check_out") or (old["check_out"] if old else ""))[:5]
        if action == "clock_in": cin = old["check_in"] if old and old["check_in"] else now_text
        elif action == "clock_out":
            if not (old and old["check_in"]): return jsonify(ok=False, error="Clock in before clocking out."), 400
            cout = now_text
        if cin and not re.fullmatch(r"\d{2}:\d{2}", cin): return jsonify(ok=False, error="Invalid clock-in time."), 400
        if cout and not re.fullmatch(r"\d{2}:\d{2}", cout): return jsonify(ok=False, error="Invalid clock-out time."), 400
        try: photo = _hr_save_photo(j.get("photo_data"), "attendance") if j.get("photo_data") else (old["photo_file"] if old else "")
        except ValueError as e: return jsonify(ok=False, error=str(e)), 400
        worked_minutes = _hr_attendance_minutes(cin, cout)
        hours = round(worked_minutes / 60, 2)
        shift_minutes = int(employee["shift_minutes"] or 0) or _hr_shift_minutes()
        overtime_minutes = max(0, worked_minutes - shift_minutes)
        now = time.time(); notes = str(j.get("notes") or (old["notes"] if old else ""))[:500]
        if old:
            c.execute("UPDATE hr_attendance SET check_in=?,check_out=?,status=?,work_hours=?,worked_minutes=?,shift_minutes=?,overtime_minutes=?,notes=?,photo_file=?,updated_at=? WHERE id=?", (cin, cout, status, hours, worked_minutes, shift_minutes, overtime_minutes, notes, photo, now, old["id"]))
            att_id = old["id"]
        else:
            c.execute("INSERT INTO hr_attendance(employee_id,attendance_date,check_in,check_out,status,work_hours,worked_minutes,shift_minutes,overtime_minutes,notes,photo_file,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (emp,date,cin,cout,status,hours,worked_minutes,shift_minutes,overtime_minutes,notes,photo,session.get("user", ""),now,now))
            att_id = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        row = dict(c.execute("SELECT * FROM hr_attendance WHERE id=?", (att_id,)).fetchone())
    log_activity("hr_attendance_save", str(emp), f"date={date};action={action};status={status}")
    return jsonify(ok=True, attendance=row)

@app.post("/api/hr/leaves")
@hr_module_auth("hr_manage")
def hr_leave_create():
    j = request.get_json(silent=True) or {}; emp = int(j.get("employee_id") or 0)
    start = str(j.get("start_date") or "")[:10]; end = str(j.get("end_date") or "")[:10]
    try:
        a = datetime.strptime(start, "%Y-%m-%d"); b = datetime.strptime(end, "%Y-%m-%d")
        days = (b-a).days + 1
        if days < 1 or days > 366: raise ValueError
    except Exception: return jsonify(ok=False, error="Enter valid leave dates; end date must not be before start date."), 400
    with db() as c:
        if not c.execute("SELECT 1 FROM hr_employees WHERE id=?", (emp,)).fetchone(): return jsonify(ok=False, error="Employee not found."), 404
        c.execute("INSERT INTO hr_leaves(employee_id,leave_type,start_date,end_date,days,reason,status,created_by,created_at) VALUES(?,?,?,?,?,?, 'Pending',?,?)", (emp, str(j.get("leave_type") or "Casual")[:60], start, end, days, str(j.get("reason") or "").strip()[:1000], session.get("user", ""), time.time()))
        leave_id = c.execute("SELECT last_insert_rowid()").fetchone()[0]
    log_activity("hr_leave_create", str(leave_id), f"employee={emp};dates={start}:{end}")
    return jsonify(ok=True, id=leave_id, days=days)

@app.post("/api/hr/leaves/<int:leave_id>/review")
@hr_module_auth("hr_leave_approve")
def hr_leave_review(leave_id):
    j = request.get_json(silent=True) or {}; status = str(j.get("status") or "")
    if status not in ("Approved", "Rejected"): return jsonify(ok=False, error="Choose Approved or Rejected."), 400
    with db() as c:
        cur = c.execute("UPDATE hr_leaves SET status=?,reviewed_by=?,reviewed_at=? WHERE id=? AND status='Pending'", (status, session.get("user", ""), time.time(), leave_id))
        if not cur.rowcount:
            exists = c.execute("SELECT 1 FROM hr_leaves WHERE id=?", (leave_id,)).fetchone()
            return jsonify(ok=False, error="Leave request not found or already reviewed." if exists else "Leave request not found."), 409 if exists else 404
    log_activity("hr_leave_review", str(leave_id), status)
    return jsonify(ok=True)

@app.get("/api/hr/payroll/preview")
@hr_module_auth("hr_payroll")
def hr_payroll_preview():
    try:
        emp = int(request.args.get("employee_id") or 0)
        month = str(request.args.get("pay_month") or "")[:7]
        if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month): raise ValueError
        month_start = datetime.strptime(month + "-01", "%Y-%m-%d")
        next_month = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return jsonify(ok=False, error="Select a valid employee and payroll month."), 400
    with db() as c:
        employee = c.execute("SELECT id,salary,overtime_rate FROM hr_employees WHERE id=?", (emp,)).fetchone()
        if not employee: return jsonify(ok=False, error="Employee not found."), 404
        old = c.execute("SELECT * FROM hr_payroll WHERE employee_id=? AND pay_month=?", (emp, month)).fetchone()
        ot = c.execute("SELECT COALESCE(SUM(overtime_minutes),0) FROM hr_attendance WHERE employee_id=? AND attendance_date>=? AND attendance_date<?", (emp, month_start.strftime("%Y-%m-%d"), next_month)).fetchone()[0]
        late_count = c.execute("SELECT COUNT(*) FROM hr_attendance WHERE employee_id=? AND attendance_date>=? AND attendance_date<? AND status='Late'", (emp, month_start.strftime("%Y-%m-%d"), next_month)).fetchone()[0]
    employee_rate = round(float(employee["overtime_rate"] or 0), 2)
    measured_hours = round(float(ot or 0) / 60.0, 2)
    late_rules=_hr_late_rules(); live_late_days=int(late_count or 0)
    live_excess=max(0,live_late_days-late_rules["allowed_days"])
    live_late_deduction=round(live_excess*late_rules["deduction_per_day"],2)
    if old:
        r = dict(old)
        # Older payroll rows predate overtime/late snapshots, so preview those values from current attendance.
        if float(r.get("overtime_rate") or 0) or float(r.get("overtime_hours") or 0) or float(r.get("overtime_pay") or 0):
            ot_hours = float(r.get("overtime_hours") or 0); ot_rate = float(r.get("overtime_rate") or 0); ot_pay = float(r.get("overtime_pay") or 0)
        else:
            ot_hours = measured_hours; ot_rate = employee_rate; ot_pay = round(float(ot or 0) * ot_rate / 60.0, 2)
        saved_basic = float(r.get("basic_salary") or 0)
        if saved_basic <= 0 and float(employee["salary"] or 0) > 0:
            saved_basic = float(employee["salary"] or 0)
        snapshot_days=r.get("late_days",-1); snapshot_allowed=r.get("late_allowed_days",-1); snapshot_deduction=r.get("late_deduction",-1)
        try: snapshot_days=int(snapshot_days if snapshot_days is not None else -1)
        except (TypeError,ValueError): snapshot_days=-1
        try: snapshot_allowed=int(snapshot_allowed if snapshot_allowed is not None else -1)
        except (TypeError,ValueError): snapshot_allowed=-1
        try: snapshot_deduction=float(snapshot_deduction if snapshot_deduction is not None else -1)
        except (TypeError,ValueError): snapshot_deduction=-1
        late_days,late_deduction=(snapshot_days,round(snapshot_deduction,2)) if snapshot_days>=0 and snapshot_deduction>=0 else (live_late_days,live_late_deduction)
        late_allowed=snapshot_allowed if snapshot_allowed>=0 else late_rules["allowed_days"]
        late_excess=max(0,late_days-late_allowed)
        preview = {"basic_salary":saved_basic,"allowances":float(r.get("allowances") or 0),"bonus":float(r.get("bonus") or 0),"deductions":float(r.get("deductions") or 0),"status":r.get("status") or "Draft","paid_date":r.get("paid_date") or "","payment_method":r.get("payment_method") or "Bank Transfer","reference":r.get("reference") or "","notes":r.get("notes") or "","overtime_hours":ot_hours,"overtime_rate":ot_rate,"overtime_pay":ot_pay,"late_days":late_days,"late_allowed_days":late_allowed,"late_excess_days":late_excess,"late_deduction":late_deduction,"existing":True}
    else:
        preview = {"basic_salary":float(employee["salary"] or 0),"allowances":0,"bonus":0,"deductions":0,"status":"Draft","paid_date":"","payment_method":"Bank Transfer","reference":"","notes":"","overtime_hours":measured_hours,"overtime_rate":employee_rate,"overtime_pay":round(float(ot or 0)*employee_rate/60.0,2),"late_days":live_late_days,"late_allowed_days":late_rules["allowed_days"],"late_excess_days":live_excess,"late_deduction":live_late_deduction,"existing":False}
    return jsonify(ok=True, preview=preview)

@app.post("/api/hr/payroll")
@hr_module_auth("hr_payroll")
def hr_payroll_save():
    j = request.get_json(silent=True) or {}; emp = int(j.get("employee_id") or 0); month = str(j.get("pay_month") or "")[:7]
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month): return jsonify(ok=False, error="Enter a valid payroll month."), 400
    try:
        allowance = float(j.get("allowances") or 0); bonus = float(j.get("bonus") or 0); deductions = float(j.get("deductions") or 0)
        if any(not (0 <= n < 1e12) for n in (allowance, bonus, deductions)): raise ValueError
    except (TypeError, ValueError): return jsonify(ok=False, error="Salary values must be valid non-negative amounts."), 400
    status = str(j.get("status") or "Draft")
    if status not in ("Draft", "Paid"): status = "Draft"
    paid_date = str(j.get("paid_date") or "")[:10] if status == "Paid" else ""
    if status == "Paid" and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", paid_date): paid_date = datetime.now(DHAKA_TZ).strftime("%Y-%m-%d")
    try:
        month_start = datetime.strptime(month + "-01", "%Y-%m-%d")
        next_month = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1).strftime("%Y-%m-%d")
        late_rules=_hr_late_rules()
        with db() as c:
            er = c.execute("SELECT id,salary,overtime_rate FROM hr_employees WHERE id=?", (emp,)).fetchone()
            if not er: return jsonify(ok=False, error="Employee not found."), 404
            basic_raw = j.get("basic_salary", None)
            basic = float(basic_raw) if basic_raw is not None and str(basic_raw).strip() else float(er["salary"] or 0)
            # If an uninitialised/old form submits zero, retain the employee's configured salary.
            if basic == 0 and float(er["salary"] or 0) > 0: basic = float(er["salary"] or 0)
            if not (0 <= basic < 1e12): raise ValueError
            ot_minutes = float(c.execute("SELECT COALESCE(SUM(overtime_minutes),0) FROM hr_attendance WHERE employee_id=? AND attendance_date>=? AND attendance_date<?", (emp, month_start.strftime("%Y-%m-%d"), next_month)).fetchone()[0] or 0)
            late_days = int(c.execute("SELECT COUNT(*) FROM hr_attendance WHERE employee_id=? AND attendance_date>=? AND attendance_date<? AND status='Late'", (emp, month_start.strftime("%Y-%m-%d"), next_month)).fetchone()[0] or 0)
            late_allowed_days=late_rules["allowed_days"]
            late_excess_days=max(0,late_days-late_allowed_days)
            late_deduction=round(late_excess_days*late_rules["deduction_per_day"],2)
            overtime_hours = round(ot_minutes / 60.0, 2)
            overtime_rate = round(float(er["overtime_rate"] or 0), 2)
            overtime_pay = round(ot_minutes * overtime_rate / 60.0, 2)
            net = round(basic + allowance + bonus + overtime_pay - deductions - late_deduction, 2)
            if net < 0: return jsonify(ok=False, error="Manual deductions plus late-entry deductions cannot exceed gross pay."), 400
            old = c.execute("SELECT id FROM hr_payroll WHERE employee_id=? AND pay_month=?", (emp, month)).fetchone()
            now=time.time(); args=(basic,allowance,bonus,deductions,overtime_hours,overtime_rate,overtime_pay,late_days,late_allowed_days,late_deduction,net,status,paid_date,str(j.get("payment_method") or "Bank Transfer")[:40],str(j.get("reference") or "")[:120],str(j.get("notes") or "")[:500],session.get("user", ""),now,now)
            if old:
                c.execute("UPDATE hr_payroll SET basic_salary=?,allowances=?,bonus=?,deductions=?,overtime_hours=?,overtime_rate=?,overtime_pay=?,late_days=?,late_allowed_days=?,late_deduction=?,net_pay=?,status=?,paid_date=?,payment_method=?,reference=?,notes=?,created_by=?,created_at=?,updated_at=? WHERE id=?", (*args,old["id"]))
                payroll_id=old["id"]
            else:
                c.execute("INSERT INTO hr_payroll(employee_id,pay_month,basic_salary,allowances,bonus,deductions,overtime_hours,overtime_rate,overtime_pay,late_days,late_allowed_days,late_deduction,net_pay,status,paid_date,payment_method,reference,notes,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (emp,month,*args))
                payroll_id=c.execute("SELECT last_insert_rowid()").fetchone()[0]
    except (TypeError, ValueError):
        return jsonify(ok=False, error="Salary values must be valid non-negative amounts."), 400
    log_activity("hr_payroll_save", str(payroll_id), f"employee={emp};month={month};net={net};overtime_hours={overtime_hours};overtime_pay={overtime_pay};late_days={late_days};late_deduction={late_deduction};status={status}")
    return jsonify(ok=True, id=payroll_id, net_pay=net, overtime_hours=overtime_hours, overtime_rate=overtime_rate, overtime_pay=overtime_pay, late_days=late_days, late_excess_days=late_excess_days, late_deduction=late_deduction)

@app.get("/api/hr/payslip/<int:payroll_id>")
@hr_module_auth("hr_payroll")
def hr_payslip(payroll_id):
    with db() as c:
        row=c.execute("SELECT p.*,e.employee_code,e.full_name,e.phone,e.email,e.department,e.designation,e.join_date,e.employment_type,e.salary AS employee_salary FROM hr_payroll p JOIN hr_employees e ON e.id=p.employee_id WHERE p.id=?", (payroll_id,)).fetchone()
    if not row: return Response("Payslip not found", status=404, mimetype="text/plain")
    r=dict(row)
    esc_html=lambda x: __import__('html').escape(str(x if x is not None else ''))
    money=lambda x: "BDT " + format(float(x or 0), ",.2f")
    basic=float(r.get("basic_salary") or 0)
    if basic<=0 and float(r.get("employee_salary") or 0)>0: basic=float(r.get("employee_salary") or 0)
    allowances=float(r.get("allowances") or 0)
    bonus=float(r.get("bonus") or 0)
    ot_hours=float(r.get("overtime_hours") or 0)
    ot_rate=float(r.get("overtime_rate") or 0)
    ot_pay=float(r.get("overtime_pay") or 0)
    deductions=float(r.get("deductions") or 0)
    try: late_days=int(r.get("late_days") if r.get("late_days") is not None else -1)
    except (TypeError,ValueError): late_days=-1
    try: late_deduction=float(r.get("late_deduction") if r.get("late_deduction") is not None else -1)
    except (TypeError,ValueError): late_deduction=-1
    late_rules=_hr_late_rules()
    try: late_allowed=int(r.get("late_allowed_days") if r.get("late_allowed_days") is not None else -1)
    except (TypeError,ValueError): late_allowed=-1
    if late_allowed<0: late_allowed=late_rules["allowed_days"]
    legacy_late_snapshot=late_days<0 or late_deduction<0 or r.get("late_allowed_days",-1)<0
    if legacy_late_snapshot:
        try:
            mstart=datetime.strptime(str(r.get("pay_month") or ""),"%Y-%m")
            mnext=(mstart.replace(day=28)+timedelta(days=4)).replace(day=1).strftime("%Y-%m-%d")
            with db() as c:
                late_days=int(c.execute("SELECT COUNT(*) FROM hr_attendance WHERE employee_id=? AND attendance_date>=? AND attendance_date<? AND status='Late'",(r["employee_id"],mstart.strftime("%Y-%m-%d"),mnext)).fetchone()[0] or 0)
        except Exception: late_days=max(0,late_days)
        late_deduction=round(max(0,late_days-late_allowed)*late_rules["deduction_per_day"],2)
    late_days=max(0,late_days); late_excess=max(0,late_days-late_allowed); late_deduction=max(0.0,late_deduction)
    gross=basic+allowances+bonus+ot_pay
    net=round(gross-deductions-late_deduction,2) if legacy_late_snapshot else float(r.get("net_pay") or 0)
    logo_data = _company_logo()
    try: logo_v = int(float(get_setting("company_logo_v") or 0))
    except (TypeError, ValueError): logo_v = 0
    logo_html = f'<img class="company-logo" src="/api/company/logo?v={logo_v}" alt="ALIEN TEES logo">' if logo_data else '<div class="company-logo-fallback" aria-label="ALIEN TEES logo">AT</div>'
    html_doc=f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Payslip {esc_html(r['employee_code'])} · {esc_html(r['pay_month'])}</title><style>
:root{{--green:#0d7552;--dark:#173b2d;--muted:#687a70;--line:#dce8df;--pale:#f2f8f3}}
*{{box-sizing:border-box}}body{{font:14px/1.5 Arial,system-ui,sans-serif;color:var(--dark);margin:0;padding:24px;background:#edf3ee}}
.actions{{display:flex;justify-content:flex-end;gap:8px;margin:0 auto 12px;max-width:850px}}button{{padding:11px 17px;border:0;border-radius:9px;background:var(--green);color:white;font-weight:700;cursor:pointer}}
.slip{{max-width:850px;margin:auto;background:#fff;padding:32px;border:1px solid var(--line);border-top:7px solid var(--green);box-shadow:0 12px 32px #163d2b12}}
header{{display:flex;justify-content:space-between;align-items:flex-start;gap:18px;border-bottom:2px solid var(--line);padding-bottom:20px}}
.brandblock{{display:flex;align-items:center;gap:14px;min-width:0}}.company-logo,.company-logo-fallback{{width:68px;height:68px;flex:0 0 68px;object-fit:contain;border:1px solid var(--line);border-radius:12px;background:#fff;padding:5px}}
.company-logo-fallback{{display:grid;place-items:center;background:linear-gradient(145deg,var(--green),#123d30);color:white;font-size:26px;font-weight:900;padding:0}}
.company-name{{font-size:16px;letter-spacing:.12em;font-weight:900;color:var(--green);margin-bottom:2px}}h1{{margin:2px 0;font-size:28px;line-height:1.15;letter-spacing:.04em}}h2{{font-size:15px;margin:0 0 8px}}h3{{font-size:17px;margin:6px 0}}small,.muted{{color:var(--muted)}}.topmeta{{text-align:right;overflow-wrap:anywhere}}
.pill-label{{display:inline-block;padding:5px 9px;border-radius:999px;background:var(--pale);border:1px solid var(--line);font-size:11px;font-weight:700}}.grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin:18px 0}}
.info{{min-width:0;padding:13px 14px;border:1px solid var(--line);border-radius:10px;background:#fff;overflow-wrap:anywhere}}.info small{{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px}}.info b{{font-size:14px}}
table{{width:100%;border-collapse:collapse;margin:14px 0 18px}}th,td{{padding:11px 12px;border-bottom:1px solid #e4ece6;text-align:left;vertical-align:top}}thead th{{background:#eff6f0;color:#355546;font-size:11px;text-transform:uppercase;letter-spacing:.05em}}td:last-child,th:last-child{{text-align:right;white-space:nowrap}}tbody tr.subtotal td{{font-weight:800;background:#f7faf7;border-top:1px solid #cfded3}}tbody tr.deduction td:last-child{{color:#9b3434}}
.net{{margin-top:18px;padding:17px 18px;background:#e7f5eb;border:1px solid #cce6d3;border-radius:12px;display:flex;align-items:center;justify-content:space-between;gap:12px;font-size:20px;font-weight:800;color:#123e2c}}.net strong{{font-size:24px;text-align:right;overflow-wrap:anywhere}}
.payment{{padding:14px 15px;background:#f8fbf8;border:1px solid var(--line);border-radius:10px;margin-top:16px;overflow-wrap:anywhere}}.payment p{{margin:5px 0}}.sign{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:22px;margin-top:54px;page-break-inside:avoid}}.sign div{{border-top:1px solid #84948a;padding-top:8px;text-align:center;font-size:11px;color:#516458}}
.footer{{border-top:1px solid var(--line);margin-top:24px;padding-top:10px;color:#7a887f;font-size:10px;display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}}
@media(max-width:600px){{body{{padding:8px}}.slip{{padding:16px 13px;border-top-width:5px}}header{{gap:8px}}.brandblock{{gap:9px}}.company-logo,.company-logo-fallback{{width:48px;height:48px;flex-basis:48px}}.company-name{{font-size:12px}}h1{{font-size:22px}}.topmeta{{font-size:12px}}.grid{{grid-template-columns:1fr;gap:8px}}td,th{{padding:9px 7px;font-size:12px}}.net{{font-size:16px;padding:14px}}.net strong{{font-size:19px}}.sign{{gap:10px;margin-top:38px}}}}
@page{{size:A4;margin:12mm}}@media print{{body{{background:#fff;padding:0;font-size:11px;-webkit-print-color-adjust:exact;print-color-adjust:exact}}.actions{{display:none!important}}.slip{{max-width:none;width:100%;padding:0;border:0;border-top:4px solid var(--green);box-shadow:none}}header{{padding-bottom:12px}}.grid{{margin:12px 0}}.info{{padding:9px 10px}}th,td{{padding:8px 9px}}.net{{break-inside:avoid}}.sign{{margin-top:42px}}.payment,.grid,table{{break-inside:avoid}}}}
</style></head><body><div class="actions"><button type="button" onclick="window.print()">🖨 Print / Save as PDF</button></div><main class="slip">
<header><div class="brandblock">{logo_html}<div><div class="company-name">ALIEN TEES</div><small>HUMAN RESOURCES · SALARY STATEMENT</small><h1>PAYSLIP</h1><p class="muted" style="margin:4px 0 0">Payroll month: <b>{esc_html(r['pay_month'])}</b></p></div></div><div class="topmeta"><span class="pill-label">{esc_html(r['status'])}</span><p><b>Employee ID</b><br>{esc_html(r['employee_code'])}</p><small>Payslip No. #{int(r['id'])}<br>Paid date: {esc_html(r['paid_date'] or '—')}</small></div></header>
<section class="grid"><div class="info"><small>Employee Name</small><b>{esc_html(r['full_name'])}</b><p class="muted" style="margin:4px 0 0">{esc_html(r['designation'] or '—')} · {esc_html(r['department'] or '—')}</p></div><div class="info"><small>Contact Information</small><b>{esc_html(r['phone'] or '—')}</b><p class="muted" style="margin:4px 0 0">{esc_html(r['email'] or '—')}</p></div><div class="info"><small>Employment Details</small><b>{esc_html(r['employment_type'] or '—')}</b><p class="muted" style="margin:4px 0 0">Joining date: {esc_html(r['join_date'] or '—')}</p></div><div class="info"><small>Overtime Summary</small><b>{ot_hours:,.2f} hours</b><p class="muted" style="margin:4px 0 0">Rate: {money(ot_rate)} per hour</p></div></section>
<h2>Earnings & Deductions</h2><table><thead><tr><th>Description</th><th>Details</th><th>Amount (BDT)</th></tr></thead><tbody><tr><td>Basic Salary</td><td>Monthly contracted salary</td><td>{money(basic)}</td></tr><tr><td>Allowances</td><td>Additional allowances</td><td>{money(allowances)}</td></tr><tr><td>Bonus</td><td>Performance / other bonus</td><td>{money(bonus)}</td></tr><tr><td>Overtime Pay</td><td>{ot_hours:,.2f} hours × {money(ot_rate)} / hour</td><td>{money(ot_pay)}</td></tr><tr class="subtotal"><td colspan="2">Gross Earnings</td><td>{money(gross)}</td></tr><tr class="deduction"><td>Manual / Other Deductions</td><td>Adjustments entered in payroll</td><td>− {money(deductions)}</td></tr><tr class="deduction"><td>Late Attendance Deduction</td><td>{late_days} late day(s), {late_allowed} monthly grace day(s), {late_excess} extra late day(s)</td><td>− {money(late_deduction)}</td></tr></tbody></table>
<div class="net"><span>NET PAYABLE</span><strong>{money(net)}</strong></div><div class="payment"><h2>Payment Information</h2><p><b>Payment method:</b> {esc_html(r['payment_method'] or '—')}</p><p><b>Reference / Transaction ID:</b> {esc_html(r['reference'] or '—')}</p><p><b>Payment status:</b> {esc_html(r['status'])} &nbsp; <b>Paid date:</b> {esc_html(r['paid_date'] or '—')}</p><p><b>Notes:</b> {esc_html(r['notes'] or '—')}</p></div>
<div class="sign"><div>Prepared By<br><br>{esc_html(r.get('created_by') or 'HR / Payroll')}</div><div>Authorized Signature</div><div>Employee Signature</div></div><div class="footer"><span>ALIEN TEES · Human Resources</span><span>Computer-generated salary statement · BDT</span><span>Slip #{int(r['id'])}</span></div></main></body></html>'''
    return Response(html_doc, mimetype="text/html")

@app.get("/api/hr/export.xlsx")
@hr_module_auth("hr_view")
def hr_export_xlsx():
    wb=Workbook(); ws=wb.active; ws.title="Employees"; ws.append(["Employee Code","Full Name","Phone","Email","Department","Designation","Employment Type","Join Date","Monthly Salary","OT Rate / Hour","Status","Address","Emergency Contact"])
    with db() as c:
        employees=c.execute("SELECT employee_code,full_name,phone,email,department,designation,employment_type,join_date,salary,overtime_rate,status,address,emergency_contact FROM hr_employees ORDER BY full_name").fetchall()
        for r in employees: ws.append(list(r))
        wa=wb.create_sheet("Attendance"); wa.append(["Date","Employee Code","Name","Department","Check In","Check Out","Worked Duration","Worked Minutes","Standard Shift","Shift Minutes","Overtime Duration","Overtime Minutes","Status","Notes","Photo File"])
        for r in c.execute("SELECT a.attendance_date,e.employee_code,e.full_name,e.department,a.check_in,a.check_out,a.worked_minutes,a.shift_minutes,a.overtime_minutes,a.status,a.notes,a.photo_file FROM hr_attendance a JOIN hr_employees e ON e.id=a.employee_id ORDER BY a.attendance_date DESC,a.id DESC"):
            worked = int(r[6] or round(float(r[6] or 0)))
            shift = int(r[7] or 480)
            overtime = int(r[8] or 0)
            wa.append([r[0],r[1],r[2],r[3],r[4],r[5],_hr_duration_label(worked),worked,_hr_duration_label(shift),shift,_hr_duration_label(overtime),overtime,r[9],r[10],r[11]])
        wl=wb.create_sheet("Leave Requests"); wl.append(["Employee Code","Name","Leave Type","Start Date","End Date","Days","Reason","Status","Reviewed By","Created By"])
        for r in c.execute("SELECT e.employee_code,e.full_name,l.leave_type,l.start_date,l.end_date,l.days,l.reason,l.status,l.reviewed_by,l.created_by FROM hr_leaves l JOIN hr_employees e ON e.id=l.employee_id ORDER BY l.created_at DESC"): wl.append(list(r))
        wp=wb.create_sheet("Payroll"); wp.append(["Employee Code","Name","Month","Basic Salary","Allowances","Bonus","OT Hours","OT Rate / Hour","OT Pay","Late Days","Allowed Late Days","Late Deduction","Other Deductions","Net Pay","Status","Paid Date","Payment Method","Reference","Notes"])
        for r in c.execute("SELECT e.employee_code,e.full_name,p.pay_month,p.basic_salary,p.allowances,p.bonus,p.overtime_hours,p.overtime_rate,p.overtime_pay,p.late_days,p.late_allowed_days,p.late_deduction,p.deductions,p.net_pay,p.status,p.paid_date,p.payment_method,p.reference,p.notes FROM hr_payroll p JOIN hr_employees e ON e.id=p.employee_id ORDER BY p.pay_month DESC"): wp.append(list(r))
    _erp_style_workbook(wb); out=io.BytesIO(); wb.save(out); out.seek(0)
    log_activity("hr_export", "HR Reports", "Excel workbook with Employees, Attendance, Leaves and Payroll")
    return send_file(out,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",as_attachment=True,download_name="Alien-Tees-HR-Reports.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), threaded=True)






























