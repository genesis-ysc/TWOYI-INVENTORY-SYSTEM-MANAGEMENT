#!/usr/bin/env python3
"""TWOYI INVENTORY MANAGEMENT - web app, SQLite only, no extra installs.
Run:  python3 twoyi_inventory.py   then open http://localhost:8000
Data is stored in twoyi_inventory.db (next to this file)."""
import json, sqlite3, os, re, csv, io, tempfile, base64, zipfile, secrets, hashlib
from xml.etree import ElementTree as ET
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from http.cookies import SimpleCookie

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "twoyi_inventory.db")
FIELDS = ["category", "brand", "model", "serial", "department", "accountable",
          "status", "problem", "returned_date", "deployed_date"]
LABELS = ["Category", "Brand", "Model", "Serial Number", "Department", "Accountable",
          "Status", "Problem", "Returned Date", "Deployed Date"]
HEADER_MAP = {  # normalized header text -> field name
    "category": "category", "brand": "brand", "model": "model",
    "serial number": "serial", "serial": "serial", "serialnumber": "serial",
    "department": "department", "accountable": "accountable", "user": "accountable",
    "status": "status", "problem": "problem", "issue": "problem",
    "returned": "returned_date", "returned date": "returned_date", "returneddate": "returned_date",
    "deployed date": "deployed_date", "deployed": "deployed_date", "deployeddate": "deployed_date",
}
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

def _col_letters(ref):
    return re.match(r"[A-Z]+", ref).group()

def _colnum(letters):
    n = 0
    for ch in letters: n = n * 26 + (ord(ch) - 64)
    return n

def xlsx_excel_date(serial):
    try:
        from datetime import timedelta
        d = datetime(1899, 12, 30) + timedelta(days=float(serial))
        return d.strftime("%Y-%m-%d")
    except Exception:
        return str(serial)

def parse_xlsx(data):
    """Return list of rows (list of cell strings) from the first sheet of an .xlsx file."""
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            texts = si.findall(".//" + NS + "t")
            shared.append("".join(t.text or "" for t in texts))
    sheet_name = "xl/worksheets/sheet1.xml"
    if sheet_name not in z.namelist():
        candidates = [n for n in z.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")]
        if not candidates: raise ValueError("No sheet found in file")
        sheet_name = sorted(candidates)[0]
    root = ET.fromstring(z.read(sheet_name))
    rows_out = []
    for row in root.findall(".//" + NS + "row"):
        cells = {}
        maxcol = 0
        for c in row.findall(NS + "c"):
            ref = c.get("r") or ""
            col = _colnum(_col_letters(ref)) if ref else len(cells) + 1
            maxcol = max(maxcol, col)
            t = c.get("t")
            v = c.find(NS + "v")
            if t == "inlineStr":
                isnode = c.find(NS + "is")
                texts = isnode.findall(".//" + NS + "t") if isnode is not None else []
                val = "".join(x.text or "" for x in texts)
            elif t == "s" and v is not None:
                idx = int(v.text)
                val = shared[idx] if 0 <= idx < len(shared) else ""
            elif v is not None:
                val = v.text or ""
            else:
                val = ""
            cells[col] = val
        rows_out.append([cells.get(i, "") for i in range(1, maxcol + 1)])
    return rows_out

def build_template_xlsx():
    ct = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    wb = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Items" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
    wbrels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
    def esc(s): return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    def row_xml(i, vals):
        cells = "".join(
            '<c r="%s%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (chr(65 + j), i, esc(v))
            for j, v in enumerate(vals))
        return '<row r="%d">%s</row>' % (i, cells)
    example = ["Laptop", "Dell", "Latitude 5420", "SN-0001", "IT", "Juan Dela Cruz", "Deployed", "", "", "2026-01-15"]
    sheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>%s%s</sheetData></worksheet>""" % (
        row_xml(1, LABELS), row_xml(2, example))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wbrels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return buf.getvalue()

def import_rows(rows):
    if not rows: raise ValueError("File is empty")
    header = [re.sub(r"\s+", " ", (h or "").strip().lower()) for h in rows[0]]
    colmap = {}
    for i, h in enumerate(header):
        key = h.replace(" ", "")
        f = HEADER_MAP.get(h) or HEADER_MAP.get(key)
        if f: colmap[i] = f
    if not colmap:
        raise ValueError("No recognizable columns found. Use the template headers: " + ", ".join(LABELS))
    inserted = updated = skipped = 0
    errs = []
    with db() as c:
        for rnum, row in enumerate(rows[1:], start=2):
            if not any((row[i] if i < len(row) else "").strip() for i in colmap): continue
            d = {f: "" for f in FIELDS}
            for i, f in colmap.items():
                if i < len(row): d[f] = str(row[i]).strip()
            for dk in ("returned_date", "deployed_date"):
                v = d.get(dk, "")
                if re.fullmatch(r"\d+(\.\d+)?", v or ""):
                    d[dk] = xlsx_excel_date(v)
            if d["category"] not in ("Laptop", "Mobile Phone"):
                skipped += 1; errs.append("Row %d: category must be Laptop or Mobile Phone" % rnum); continue
            if d["status"] not in ("Returned", "Deployed"):
                d["status"] = d["status"] or "Deployed"
                if d["status"] not in ("Returned", "Deployed"):
                    skipped += 1; errs.append("Row %d: status must be Returned or Deployed" % rnum); continue
            try:
                existing = None
                if d["serial"]:
                    existing = c.execute("SELECT * FROM items WHERE serial=?", (d["serial"],)).fetchone()
                if existing:
                    merged = {f: (d[f] or existing[f] or "") for f in FIELDS}
                    save(existing["id"], merged); updated += 1
                else:
                    save(None, d); inserted += 1
            except ValueError as e:
                skipped += 1; errs.append("Row %d: %s" % (rnum, e))
    return {"inserted": inserted, "updated": updated, "skipped": skipped, "errors": errs[:25]}


def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS items(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          category TEXT NOT NULL CHECK(category IN ('Laptop','Mobile Phone')),
          brand TEXT, model TEXT, serial TEXT UNIQUE, department TEXT,
          accountable TEXT, status TEXT NOT NULL CHECK(status IN ('Returned','Deployed')),
          problem TEXT, returned_date TEXT, deployed_date TEXT,
          created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS history(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          item_id INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
          field TEXT NOT NULL, old_value TEXT, new_value TEXT, changed_at TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS h_item ON history(item_id);""")
        # Auto-heal: if a 'users' table already exists from an older/incompatible
        # version of this app (e.g. missing full_name), drop and recreate it along
        # with 'sessions'. This never touches 'items' or 'history', so inventory
        # data is never affected.
        cols = {r["name"] for r in c.execute("PRAGMA table_info(users)")}
        if cols and "full_name" not in cols:
            c.executescript("DROP TABLE IF EXISTS sessions; DROP TABLE IF EXISTS users;")
            cols = set()
        if not cols:
            c.executescript("""
            CREATE TABLE users(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              full_name TEXT NOT NULL,
              email TEXT NOT NULL UNIQUE,
              pass_hash TEXT NOT NULL,
              pass_salt TEXT NOT NULL,
              role TEXT NOT NULL CHECK(role IN ('admin','user')),
              created_at TEXT NOT NULL);
            CREATE TABLE sessions(
              sid TEXT PRIMARY KEY,
              user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              created_at TEXT NOT NULL, expires_at TEXT NOT NULL);""")
        else:
            c.executescript("""
            CREATE TABLE IF NOT EXISTS sessions(
              sid TEXT PRIMARY KEY,
              user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              created_at TEXT NOT NULL, expires_at TEXT NOT NULL);""")

def now(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

SESSION_DAYS = 14

def hash_pw(password, salt=None):
    salt = salt or secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 200000).hex()
    return h, salt

def verify_pw(password, salt, hashed):
    h, _ = hash_pw(password, salt)
    return secrets.compare_digest(h, hashed)

def register_user(full_name, email, password):
    full_name = (full_name or "").strip()
    email = (email or "").strip().lower()
    password = password or ""
    if not full_name: raise ValueError("Full name is required")
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email): raise ValueError("Enter a valid email address")
    if len(password) < 6: raise ValueError("Password must be at least 6 characters")
    h, salt = hash_pw(password)
    t = now()
    with db() as c:
        is_first = c.execute("SELECT COUNT(*) n FROM users").fetchone()["n"] == 0
        role = "admin" if is_first else "user"
        try:
            cur = c.execute(
                "INSERT INTO users(full_name,email,pass_hash,pass_salt,role,created_at) VALUES(?,?,?,?,?,?)",
                (full_name, email, h, salt, role, t))
        except sqlite3.IntegrityError:
            raise ValueError("That email is already registered")
        return cur.lastrowid, role

def login_user(email, password):
    email = (email or "").strip().lower()
    with db() as c:
        u = c.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
    if not u or not verify_pw(password or "", u["pass_salt"], u["pass_hash"]):
        raise ValueError("Incorrect email or password")
    return dict(u)

def create_session(user_id):
    sid = secrets.token_hex(24)
    t = datetime.now()
    with db() as c:
        c.execute("INSERT INTO sessions(sid,user_id,created_at,expires_at) VALUES(?,?,?,?)",
                  (sid, user_id, t.strftime("%Y-%m-%d %H:%M:%S"),
                   (t + timedelta(days=SESSION_DAYS)).strftime("%Y-%m-%d %H:%M:%S")))
    return sid

def session_user(handler):
    sc = SimpleCookie()
    sc.load(handler.headers.get("Cookie", "") or "")
    if "sid" not in sc: return None
    sid = sc["sid"].value
    with db() as c:
        row = c.execute(
            "SELECT u.id,u.full_name,u.email,u.role FROM sessions s JOIN users u ON u.id=s.user_id "
            "WHERE s.sid=? AND s.expires_at>?", (sid, now())).fetchone()
    return dict(row) if row else None

def clean(d):
    return {f: (str(d.get(f) or "").strip()) for f in FIELDS}

def save(iid, d):
    d = clean(d)
    if d["category"] not in ("Laptop", "Mobile Phone"): raise ValueError("Category must be Laptop or Mobile Phone")
    if d["status"] not in ("Returned", "Deployed"): raise ValueError("Status must be Returned or Deployed")
    t = now()
    with db() as c:
        try:
            if iid is None:
                cur = c.execute("INSERT INTO items(%s,created_at,updated_at) VALUES(%s)" %
                    (",".join(FIELDS), ",".join("?" * (len(FIELDS) + 2))),
                    [d[f] or None for f in FIELDS] + [t, t])
                iid = cur.lastrowid
                for f in FIELDS:
                    if d[f]: c.execute("INSERT INTO history(item_id,field,old_value,new_value,changed_at) VALUES(?,?,?,?,?)",
                                       (iid, f, None, d[f], t))
            else:
                old = c.execute("SELECT * FROM items WHERE id=?", (iid,)).fetchone()
                if not old: raise ValueError("Item not found")
                for f in FIELDS:
                    if (old[f] or "") != d[f]:
                        c.execute("INSERT INTO history(item_id,field,old_value,new_value,changed_at) VALUES(?,?,?,?,?)",
                                  (iid, f, old[f], d[f], t))
                c.execute("UPDATE items SET %s,updated_at=? WHERE id=?" % ",".join(f + "=?" for f in FIELDS),
                          [d[f] or None for f in FIELDS] + [t, iid])
        except sqlite3.IntegrityError:
            raise ValueError("Serial number already exists")
    return iid

class H(BaseHTTPRequestHandler):
    def send(self, code, body, ctype="application/json", cookie=None):
        b = body if isinstance(body, bytes) else (body if isinstance(body, str) else json.dumps(body)).encode()
        self.send_response(code); self.send_header("Content-Type", ctype + "; charset=utf-8")
        if cookie: self.send_header("Set-Cookie", cookie)
        self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)
    def body(self):
        return json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0) or b"{}")
    def current_user(self):
        return session_user(self)
    def require_login(self):
        u = self.current_user()
        if not u: self.send(401, {"error": "Please log in"}); return None
        return u
    def require_admin(self):
        u = self.require_login()
        if u is None: return None
        if u["role"] != "admin":
            self.send(403, {"error": "View-only account: ask an admin to make changes"}); return None
        return u
    def download(self, data, name, ctype):
        self.send_response(200); self.send_header("Content-Type", ctype)
        self.send_header("Content-Disposition", 'attachment; filename="%s"' % name)
        self.send_header("Content-Length", str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        if self.path == "/": return self.send(200, PAGE, "text/html")
        if self.path == "/api/me":
            u = self.current_user()
            return self.send(200, u) if u else self.send(401, {"error": "not logged in"})
        if self.path in ("/export/items.csv", "/export/history.csv", "/export/template.xlsx",
                          "/export/database.db", "/api/items", "/api/activity", "/api/users") \
           or re.fullmatch(r"/api/items/(\d+)/history", self.path):
            if self.require_login() is None: return
        if self.path in ("/export/items.csv", "/export/history.csv"):
            out = io.StringIO(); w = csv.writer(out)
            with db() as c:
                if "items" in self.path:
                    w.writerow(["ID", "Category", "Brand", "Model", "Serial Number", "Department", "Accountable",
                                "Status", "Problem", "Returned Date", "Deployed Date", "Last Updated"])
                    for r in c.execute("SELECT * FROM items ORDER BY id"):
                        w.writerow([r["id"]] + [r[f] or "" for f in FIELDS] + [r["updated_at"]])
                else:
                    w.writerow(["Changed At", "Item ID", "Serial Number", "Brand", "Model", "Field", "Old Value", "New Value"])
                    for r in c.execute("SELECT h.*,i.serial,i.brand,i.model FROM history h JOIN items i ON i.id=h.item_id ORDER BY h.id"):
                        w.writerow([r["changed_at"], r["item_id"], r["serial"] or "", r["brand"] or "", r["model"] or "",
                                    r["field"], r["old_value"] or "", r["new_value"] or ""])
            return self.download(("\ufeff" + out.getvalue()).encode("utf-8"),
                                 "twoyi_%s_%s.csv" % ("items" if "items" in self.path else "history", stamp), "text/csv; charset=utf-8")
        if self.path == "/export/template.xlsx":
            return self.download(build_template_xlsx(), "twoyi_import_template.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if self.path == "/export/database.db":
            tmp = os.path.join(tempfile.gettempdir(), "twoyi_copy_%s.db" % os.getpid())
            src = db(); dst = sqlite3.connect(tmp); src.backup(dst); dst.close(); src.close()
            data = open(tmp, "rb").read(); os.remove(tmp)
            return self.download(data, "twoyi_inventory_%s.db" % stamp, "application/x-sqlite3")
        m = re.fullmatch(r"/api/items/(\d+)/history", self.path)
        with db() as c:
            if self.path == "/api/items":
                return self.send(200, [dict(r) for r in c.execute("SELECT * FROM items ORDER BY id DESC")])
            if self.path == "/api/activity":
                return self.send(200, [dict(r) for r in c.execute(
                    "SELECT h.*,i.serial,i.brand,i.model FROM history h JOIN items i ON i.id=h.item_id ORDER BY h.id DESC LIMIT 12")])
            if self.path == "/api/users":
                me = self.current_user()
                if me["role"] != "admin": return self.send(403, {"error": "Admins only"})
                return self.send(200, [dict(r) for r in c.execute(
                    "SELECT id,full_name,email,role,created_at FROM users ORDER BY id")])
            if m:
                return self.send(200, [dict(r) for r in c.execute(
                    "SELECT * FROM history WHERE item_id=? ORDER BY id DESC", (m[1],))])
        self.send(404, {"error": "not found"})
    def do_POST(self):
        if self.path == "/api/register":
            try:
                b = self.body()
                uid, role = register_user(b.get("full_name"), b.get("email"), b.get("password"))
                sid = create_session(uid)
                cookie = "sid=%s; Path=/; Max-Age=%d; HttpOnly; SameSite=Lax" % (sid, SESSION_DAYS * 86400)
                return self.send(200, {"id": uid, "role": role}, cookie=cookie)
            except ValueError as e:
                return self.send(400, {"error": str(e)})
        if self.path == "/api/login":
            try:
                b = self.body()
                u = login_user(b.get("email"), b.get("password"))
                sid = create_session(u["id"])
                cookie = "sid=%s; Path=/; Max-Age=%d; HttpOnly; SameSite=Lax" % (sid, SESSION_DAYS * 86400)
                return self.send(200, {"id": u["id"], "full_name": u["full_name"], "email": u["email"], "role": u["role"]}, cookie=cookie)
            except ValueError as e:
                return self.send(400, {"error": str(e)})
        if self.path == "/api/logout":
            sc = SimpleCookie(); sc.load(self.headers.get("Cookie", "") or "")
            if "sid" in sc:
                with db() as c: c.execute("DELETE FROM sessions WHERE sid=?", (sc["sid"].value,))
            return self.send(200, {"ok": True}, cookie="sid=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")
        if self.path == "/api/import":
            if self.require_admin() is None: return
            try:
                b = self.body()
                raw = base64.b64decode(b.get("content_base64", ""))
                name = (b.get("filename") or "").lower()
                if name.endswith(".csv"):
                    text = raw.decode("utf-8-sig", errors="replace")
                    rows = list(csv.reader(io.StringIO(text)))
                else:
                    rows = parse_xlsx(raw)
                return self.send(200, import_rows(rows))
            except Exception as e:
                return self.send(400, {"error": str(e)})
        if self.require_admin() is None: return
        try: self.send(200, {"id": save(None, self.body())})
        except ValueError as e: self.send(400, {"error": str(e)})
    def do_PUT(self):
        mu = re.fullmatch(r"/api/users/(\d+)", self.path)
        if mu:
            if self.require_admin() is None: return
            role = self.body().get("role")
            if role not in ("admin", "user"): return self.send(400, {"error": "role must be admin or user"})
            with db() as c: c.execute("UPDATE users SET role=? WHERE id=?", (role, mu[1]))
            return self.send(200, {"ok": True})
        if self.require_admin() is None: return
        m = re.fullmatch(r"/api/items/(\d+)", self.path)
        try: self.send(200, {"id": save(int(m[1]), self.body())})
        except ValueError as e: self.send(400, {"error": str(e)})
    def do_DELETE(self):
        mu = re.fullmatch(r"/api/users/(\d+)", self.path)
        if mu:
            me = self.require_admin()
            if me is None: return
            if int(mu[1]) == me["id"]: return self.send(400, {"error": "You can't delete your own account"})
            with db() as c: c.execute("DELETE FROM users WHERE id=?", (mu[1],))
            return self.send(200, {"ok": True})
        if self.require_admin() is None: return
        if self.path == "/api/items":
            ids = self.body().get("ids", [])
            with db() as c:
                c.executemany("DELETE FROM items WHERE id=?", [(i,) for i in ids])
            return self.send(200, {"ok": True, "deleted": len(ids)})
        m = re.fullmatch(r"/api/items/(\d+)", self.path)
        with db() as c: c.execute("DELETE FROM items WHERE id=?", (m[1],))
        self.send(200, {"ok": True})
    def log_message(self, *a): pass

PAGE = r"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>TWOYI Inventory Management</title><style>
:root{--o:#f97316;--od:#c2410c;--ol:#fff7ed;--ob:#fed7aa;--t:#1f2937}
*{box-sizing:border-box}body{margin:0;font-family:system-ui,Segoe UI,Arial,sans-serif;background:var(--ol);color:var(--t)}
header{background:linear-gradient(90deg,var(--od),var(--o));color:#fff;padding:12px 24px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.logo{height:64px;width:auto;background:#fff;border-radius:12px;padding:4px;box-shadow:0 2px 6px #0003}header h1{font-size:20px;margin:0;letter-spacing:1px}header small{opacity:.85;display:block;font-weight:400;letter-spacing:0}
main{padding:20px 24px}.bar{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap}
input,select,textarea{padding:8px 10px;border:1px solid var(--ob);border-radius:8px;font:inherit;width:100%;background:#fff}
.bar input,.bar select{width:auto}.bar input{flex:1;min-width:200px}
button{background:var(--o);color:#fff;border:0;padding:8px 14px;border-radius:8px;font:inherit;cursor:pointer;font-weight:600}
button:hover{background:var(--od)}button.g{background:#fff;color:var(--od);border:1px solid var(--o)}button.g:hover{background:var(--ob)}
.w{overflow-x:auto;background:#fff;border-radius:12px;box-shadow:0 2px 10px #f9731633}
table{border-collapse:collapse;width:100%;font-size:14px}th{background:var(--o);color:#fff;text-align:left;padding:10px;white-space:nowrap}
td{padding:9px 10px;border-bottom:1px solid var(--ob);white-space:nowrap}tr:hover td{background:var(--ol)}
.b{padding:3px 10px;border-radius:99px;font-size:12px;font-weight:700}.Deployed{background:#dcfce7;color:#166534}.Returned{background:#ffedd5;color:var(--od)}
.m{position:fixed;inset:0;background:#0007;display:none;align-items:center;justify-content:center;padding:14px;z-index:9}.m.on{display:flex}
.card{background:#fff;border-radius:14px;padding:20px;width:640px;max-width:100%;max-height:92vh;overflow:auto;border-top:6px solid var(--o)}
.card h2{margin:0 0 12px;color:var(--od)}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.grid label{font-size:12px;font-weight:700;color:var(--od)}
.full{grid-column:1/-1}.act{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}
.err{color:#b91c1c;font-size:13px;margin-top:8px}.h{border-left:3px solid var(--o);padding:6px 10px;margin:8px 0;background:var(--ol);border-radius:0 8px 8px 0;font-size:13px}
.h s{color:#9a3412}.h b{color:var(--od)}.h time{float:right;color:#6b7280;font-size:12px}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}.chips button{padding:4px 10px;font-size:12px}
.tabs{display:flex;gap:6px;margin-bottom:14px}.tabs button{background:#fff;color:var(--od);border:1px solid var(--o)}.tabs button.on{background:var(--o);color:#fff}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}
.st{background:#fff;border-radius:12px;padding:14px;border-left:6px solid var(--o);box-shadow:0 2px 8px #f9731622}.st b{font-size:28px;color:var(--od);display:block}.st span{font-size:12px;color:#6b7280;font-weight:600;text-transform:uppercase}
.pan{background:#fff;border-radius:12px;padding:16px;box-shadow:0 2px 8px #f9731622}.pan h3{margin:0 0 10px;color:var(--od);font-size:15px}
.two{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}.br{display:flex;align-items:center;gap:8px;margin:6px 0;font-size:13px}.br i{display:block;height:14px;background:var(--o);border-radius:4px;min-width:3px}.br em{width:120px;font-style:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ex{margin-left:auto;display:flex;gap:6px;flex-wrap:wrap}.hide{display:none}
.who{margin-left:auto;display:flex;align-items:center;gap:10px;color:#fff;font-size:13px}
.who .role{background:#fff3;padding:2px 10px;border-radius:99px;font-weight:700;text-transform:uppercase;font-size:11px}
.who button{background:#fff2;border:1px solid #fff5;padding:5px 12px}
.auth{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,var(--ol),var(--ob))}
.auth .card{width:380px;text-align:center}.auth img{height:70px;margin-bottom:10px}
.auth h2{margin-top:0}.auth .err{min-height:18px}
.swap{font-size:13px;margin-top:12px;color:#6b7280}.swap a{color:var(--od);font-weight:700;cursor:pointer;text-decoration:underline}
.vo{background:#fef3c7;color:#92400e;border:1px solid #fcd34d;padding:8px 14px;border-radius:8px;font-size:13px;margin-bottom:14px}
@media(max-width:560px){.grid{grid-template-columns:1fr}.two{grid-template-columns:1fr}}</style></head><body>
<div class="auth" id="authScreen"><div class="card">
<img src="data:image/png;base64,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" alt="Yardstick Coffee logo">
<h2 id="authTitle">Sign In</h2>
<div id="loginForm">
<div class="grid"><div class="full"><label>Email</label><input id="l_email" type="email" placeholder="you@company.com"></div>
<div class="full"><label>Password</label><input id="l_pass" type="password" placeholder="••••••••"></div></div>
<div class="err" id="l_err"></div>
<div class="act" style="justify-content:center;margin-top:14px"><button onclick="doLogin()" style="width:100%">Sign In</button></div>
<div class="swap">No account yet? <a onclick="showAuth('register')">Register</a></div>
</div>
<div id="registerForm" class="hide">
<div class="grid"><div class="full"><label>Full Name</label><input id="r_name" placeholder="Juan Dela Cruz"></div>
<div class="full"><label>Email (used as username)</label><input id="r_email" type="email" placeholder="you@company.com"></div>
<div class="full"><label>Password</label><input id="r_pass" type="password" placeholder="At least 6 characters"></div></div>
<div class="err" id="r_err"></div>
<div class="act" style="justify-content:center;margin-top:14px"><button onclick="doRegister()" style="width:100%">Create Account</button></div>
<div class="swap">Already registered? <a onclick="showAuth('login')">Sign In</a></div>
</div>
</div></div>
<div id="app" class="hide">
<header><img class="logo" src="data:image/png;base64,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" alt="Yardstick Coffee logo">
<h1>TWOYI INVENTORY MANAGEMENT<small>One Yard &middot; Laptops &amp; Mobile Phones</small></h1><div class="who"><span id="whoName"></span><span class="role" id="whoRole"></span><button onclick="doLogout()">Log Out</button></div></header>
<main><div class="tabs"><button id="t1" class="on" onclick="tab(1)">Dashboard</button><button id="t2" onclick="tab(2)">All Items</button><button id="t3" class="hide" onclick="tab(3)">Users</button>
<div class="ex"><button class="g admin-only" onclick="location='/export/template.xlsx'">⬇ Import Template</button><button class="admin-only" onclick="$('fi').click()">⬆ Import Excel</button><input type="file" id="fi" accept=".xlsx,.csv" class="hide" onchange="doImport(this)">
<button class="g" onclick="location='/export/items.csv'">⬇ Items (CSV)</button><button class="g" onclick="location='/export/history.csv'">⬇ History (CSV)</button><button class="g" onclick="location='/export/database.db'">⬇ SQLite DB</button></div></div>
<section id="dash"><div class="cards" id="cards"></div>
<div class="two"><div class="pan"><h3>Items per Department</h3><div id="bd"></div></div><div class="pan"><h3>Category &amp; Status Breakdown</h3><div id="bc"></div></div></div>
<div class="two"><div class="pan"><h3>Items with Problems</h3><div id="bp"></div></div><div class="pan"><h3>Recent Activity</h3><div id="ra"></div></div></div></section>
<div class="vo hide" id="voBanner">You are signed in with a <b>view-only</b> account. Ask an admin to make changes.</div>
<section id="all" class="hide"><div class="bar"><input id="q" placeholder="Search anything (serial, user, brand, department...)" oninput="draw()">
<select id="fc" onchange="draw()"><option value="">All categories</option><option>Laptop</option><option>Mobile Phone</option></select>
<select id="fs" onchange="draw()"><option value="">All status</option><option>Deployed</option><option>Returned</option></select>
<button class="admin-only" onclick="edit()">+ Add Item</button><button class="g admin-only" id="delsel" onclick="delSel()" style="display:none">Delete Selected</button><span id="cnt" style="align-self:center;font-weight:600;color:var(--od)"></span></div>
<div class="w"><table><thead><tr><th class="admin-only"><input type="checkbox" id="selall" onchange="selAll(this)"></th><th>Category</th><th>Brand</th><th>Model</th><th>Serial Number</th><th>Department</th><th>Accountable</th><th>Status</th><th>Problem</th><th>Returned Date</th><th>Deployed Date</th><th></th></tr></thead><tbody id="tb"></tbody></table></div></section>
<section id="users" class="hide"><div class="w"><table><thead><tr><th>Full Name</th><th>Email</th><th>Role</th><th>Joined</th><th></th></tr></thead><tbody id="ub"></tbody></table></div></section>
</main>
<div class="m" id="m1"><div class="card"><h2 id="mt"></h2><div class="grid">
<div><label>Category</label><select id="category"><option>Laptop</option><option>Mobile Phone</option></select></div>
<div><label>Status</label><select id="status"><option>Deployed</option><option>Returned</option></select></div>
<div><label>Brand</label><input id="brand"></div><div><label>Model</label><input id="model"></div>
<div><label>Serial Number</label><input id="serial"></div><div><label>Department</label><input id="department"></div>
<div class="full"><label>Accountable (current user)</label><input id="accountable"></div>
<div class="full"><label>Problem</label><textarea id="problem" rows="2"></textarea></div>
<div><label>Returned Date</label><input type="date" id="returned_date"></div><div><label>Deployed Date</label><input type="date" id="deployed_date"></div></div>
<div class="err" id="er"></div><div class="act"><button class="g" onclick="cl()">Cancel</button><button onclick="sv()">Save</button></div></div></div>
<div class="m" id="m3"><div class="card"><h2>Import Result</h2><div id="imr"></div><div class="act"><button onclick="cl()">OK</button></div></div></div>
<div class="m" id="m2"><div class="card"><h2 id="ht"></h2><div class="chips" id="hc"></div><div id="hl"></div><div class="act"><button class="g" onclick="cl()">Close</button></div></div></div>
</div>
<script>
const F=["category","brand","model","serial","department","accountable","status","problem","returned_date","deployed_date"];
const L={category:"Category",brand:"Brand",model:"Model",serial:"Serial Number",department:"Department",accountable:"Accountable",status:"Status",problem:"Problem",returned_date:"Returned Date",deployed_date:"Deployed Date"};
let D=[],cur=null,HH=[],ME=null;const $=i=>document.getElementById(i);
const e=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
async function load(){D=await (await fetch("/api/items")).json();draw();dash()}
function tab(n){$("dash").classList.toggle("hide",n!=1);$("all").classList.toggle("hide",n!=2);$("users").classList.toggle("hide",n!=3);
$("t1").classList.toggle("on",n==1);$("t2").classList.toggle("on",n==2);$("t3").classList.toggle("on",n==3);
if(n==3)loadUsers()}
function showAuth(which){$("authTitle").textContent=which=="register"?"Create Account":"Sign In";
$("loginForm").classList.toggle("hide",which=="register");$("registerForm").classList.toggle("hide",which!="register");
$("l_err").textContent="";$("r_err").textContent=""}
async function doLogin(){const r=await fetch("/api/login",{method:"POST",body:JSON.stringify({email:$("l_email").value,password:$("l_pass").value})});
const j=await r.json();if(!r.ok){$("l_err").textContent=j.error;return}await boot()}
async function doRegister(){const r=await fetch("/api/register",{method:"POST",body:JSON.stringify({full_name:$("r_name").value,email:$("r_email").value,password:$("r_pass").value})});
const j=await r.json();if(!r.ok){$("r_err").textContent=j.error;return}await boot()}
async function doLogout(){await fetch("/api/logout",{method:"POST"});ME=null;$("app").classList.add("hide");$("authScreen").style.display="flex";showAuth("login")}
function applyRole(){const admin=ME.role=="admin";
$("whoName").textContent=ME.full_name;$("whoRole").textContent=ME.role;
$("t3").classList.toggle("hide",!admin);
$("voBanner").classList.toggle("hide",admin);
document.querySelectorAll(".admin-only").forEach(el=>el.style.display=admin?"":"none")}
async function boot(){const r=await fetch("/api/me");
if(!r.ok){$("authScreen").style.display="flex";$("app").classList.add("hide");return}
ME=await r.json();applyRole();$("authScreen").style.display="none";$("app").classList.remove("hide");load()}
async function loadUsers(){const r=await fetch("/api/users");if(!r.ok)return;const us=await r.json();
$("ub").innerHTML=us.map(u=>`<tr><td>${e(u.full_name)}</td><td>${e(u.email)}</td><td>
<select onchange="setRole(${u.id},this.value)" ${u.id==ME.id?"disabled":""}><option value="user" ${u.role=="user"?"selected":""}>user</option><option value="admin" ${u.role=="admin"?"selected":""}>admin</option></select>
</td><td>${e(u.created_at)}</td><td>${u.id==ME.id?"":`<button class="g" onclick="delUser(${u.id})">Remove</button>`}</td></tr>`).join("")||"No users yet."}
async function setRole(id,role){await fetch("/api/users/"+id,{method:"PUT",body:JSON.stringify({role})});loadUsers()}
async function delUser(id){if(confirm("Remove this user's account? They won't be able to sign in anymore."))
{await fetch("/api/users/"+id,{method:"DELETE"});loadUsers()}}
const bars=o=>{const m=Math.max(1,...Object.values(o));return Object.entries(o).sort((a,b)=>b[1]-a[1]).map(([k,v])=>`<div class="br"><em title="${e(k)}">${e(k)}</em><i style="width:${v/m*60}%"></i><b>${v}</b></div>`).join("")||"No data yet."};
const cnt=(a,f)=>a.reduce((o,r)=>{const k=r[f]||"(none)";o[k]=(o[k]||0)+1;return o},{});
async function dash(){const n=(f,v)=>D.filter(r=>r[f]==v).length,P=D.filter(r=>r.problem);
$("cards").innerHTML=[["Total Items",D.length],["Deployed",n("status","Deployed")],["Returned",n("status","Returned")],["Laptops",n("category","Laptop")],["Mobile Phones",n("category","Mobile Phone")],["With Problems",P.length]].map(([l,v])=>`<div class="st"><b>${v}</b><span>${l}</span></div>`).join("");
$("bd").innerHTML=bars(cnt(D,"department"));
$("bc").innerHTML=bars(D.reduce((o,r)=>{const k=r.category+" – "+r.status;o[k]=(o[k]||0)+1;return o},{}));
$("bp").innerHTML=P.slice(0,8).map(r=>`<div class="h"><b>${e(r.brand)} ${e(r.model)}</b> (${e(r.serial)}) – ${e(r.accountable)}<br>${e(r.problem)}</div>`).join("")||"No problems reported.";
const a=await (await fetch("/api/activity")).json();
$("ra").innerHTML=a.map(h=>`<div class="h"><time>${h.changed_at}</time><b>${e(h.brand)} ${e(h.model)}</b><br>${L[h.field]}: ${h.old_value?`<s>${e(h.old_value)}</s> → `:""}${e(h.new_value||"(cleared)")}</div>`).join("")||"No activity yet."}
let SEL=new Set();
function selAll(cb){SEL=cb.checked?new Set(D.map(r=>r.id)):new Set();draw()}
function selOne(cb,id){cb.checked?SEL.add(id):SEL.delete(id);$("delsel").style.display=SEL.size?"inline-block":"none"}
async function delSel(){if(!SEL.size)return;if(!confirm(`Delete ${SEL.size} selected item(s) AND their history?`))return;
await fetch("/api/items",{method:"DELETE",body:JSON.stringify({ids:[...SEL]})});SEL=new Set();$("delsel").style.display="none";load()}
function draw(){const q=$("q").value.toLowerCase(),c=$("fc").value,s=$("fs").value;
const V=D.filter(r=>(!c||r.category==c)&&(!s||r.status==s)&&(!q||F.some(f=>(r[f]||"").toLowerCase().includes(q))));$("cnt").textContent=V.length+" of "+D.length+" items";
const admin=ME&&ME.role=="admin";
$("tb").innerHTML=V.map(r=>
`<tr><td>${admin?`<input type="checkbox" ${SEL.has(r.id)?"checked":""} onchange="selOne(this,${r.id})">`:""}</td>${F.map(f=>f=="status"?`<td><span class="b ${r.status}">${r.status}</span></td>`:`<td>${e(r[f])}</td>`).join("")}
<td>${admin?`<button class="g" onclick="edit(${r.id})">Edit</button> `:""}<button class="g" onclick="hist(${r.id})">History</button> ${admin?`<button class="g" onclick="del(${r.id})">✕</button>`:""}</td></tr>`).join("")||`<tr><td colspan="12" style="text-align:center;padding:30px">No items yet.</td></tr>`;
$("selall").checked=D.length>0&&SEL.size===D.length;$("delsel").style.display=SEL.size?"inline-block":"none"}
function doImport(inp){const f=inp.files[0];if(!f)return;const r=new FileReader();
r.onload=async()=>{const b64=r.result.split(",")[1];
const res=await fetch("/api/import",{method:"POST",body:JSON.stringify({filename:f.name,content_base64:b64})});
const j=await res.json();inp.value="";
if(!res.ok){$("imr").innerHTML=`<div class="err">${e(j.error)}</div>`}
else{$("imr").innerHTML=`<div class="h"><b>${j.inserted}</b> added, <b>${j.updated}</b> updated (matched by Serial Number), <b>${j.skipped}</b> skipped.</div>`+
(j.errors.length?`<div class="err">${j.errors.map(e2=>e(e2)).join("<br>")}</div>`:"");load()}
$("m3").classList.add("on")};
r.readAsDataURL(f)}
function edit(id){cur=id||null;const r=D.find(x=>x.id==id)||{category:"Laptop",status:"Deployed"};
F.forEach(f=>$(f).value=r[f]||"");$("mt").textContent=id?"Edit Item":"Add Item";$("er").textContent="";$("m1").classList.add("on")}
async function sv(){const b={};F.forEach(f=>b[f]=$(f).value);
const r=await fetch(cur?"/api/items/"+cur:"/api/items",{method:cur?"PUT":"POST",body:JSON.stringify(b)});const j=await r.json();
if(!r.ok){$("er").textContent=j.error;return}cl();load()}
async function del(id){if(confirm("Delete this item AND its history?")){await fetch("/api/items/"+id,{method:"DELETE"});load()}}
function cl(){document.querySelectorAll(".m").forEach(m=>m.classList.remove("on"))}
async function hist(id){const r=D.find(x=>x.id==id);HH=await (await fetch(`/api/items/${id}/history`)).json();
$("ht").textContent=`History: ${r.brand||""} ${r.model||""} (${r.serial||"no serial"})`;
$("hc").innerHTML=`<button onclick="hf('')">All</button>`+F.map(f=>`<button class="g" onclick="hf('${f}')">${L[f]}</button>`).join("");hf("");$("m2").classList.add("on")}
function hf(f){const a=HH.filter(h=>!f||h.field==f);
$("hl").innerHTML=(f=="accountable"?`<div class="h"><b>Users in order:</b> ${e(a.slice().reverse().map(h=>h.new_value||"(blank)").join(" → "))}</div>`:"")+
(a.map(h=>`<div class="h"><time>${h.changed_at}</time><b>${L[h.field]}</b>: ${h.old_value?`<s>${e(h.old_value)}</s> → `:"(new) "}${e(h.new_value||"(cleared)")}</div>`).join("")||"No history.")}
boot()</script></body></html>"""

class Server(ThreadingHTTPServer):
    allow_reuse_address = True  # let the OS reuse the port right after a restart

if __name__ == "__main__":
    import sys
    port = int(os.environ.get("TWOYI_PORT", sys.argv[1] if len(sys.argv) > 1 else 8000))
    init()
    try:
        srv = Server(("0.0.0.0", port), H)
    except OSError as e:
        if e.errno == 98:
            print(f"\nPort {port} is already in use - likely an earlier run of this app still active.\n"
                  f"Either stop that process, or start this one on a different port, e.g.:\n"
                  f"  python3 {os.path.basename(__file__)} 8001\n")
            sys.exit(1)
        raise
    print(f"TWOYI Inventory running at http://localhost:{port}  (Ctrl+C to stop)")
    srv.serve_forever()
