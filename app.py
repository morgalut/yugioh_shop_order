# app.py
# Run:
#   pip install fastapi uvicorn openpyxl
#   uvicorn app:app --reload
# Open:
#   http://127.0.0.1:8000
#
# What’s new in this version:
# ✅ Customers saved in SQLite (unique names).
# ✅ Orders + order lines saved in SQLite (basket persists across server restarts).
# ✅ Basket shows each user separately (grouped).
# ✅ Add / Remove / Edit products (order lines) directly in the basket.
# ✅ Download Excel per user OR combined Excel (from DB).
#
# Notes:
# - “Basket” == all orders with status='OPEN'
# - Download does NOT clear basket (you can clear manually).

from __future__ import annotations

import io
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import List, Optional, Dict, Tuple

from fastapi import FastAPI, Form, Query
from fastapi.responses import HTMLResponse, StreamingResponse, RedirectResponse
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

app = FastAPI()
DB_PATH = "customers.db"

# ----------------------------
# SQLite helpers
# ----------------------------
def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def init_db() -> None:
    conn = db_connect()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at_utc TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER NOT NULL,
                pasted_text TEXT NOT NULL,
                global_note TEXT,
                status TEXT NOT NULL DEFAULT 'OPEN', -- OPEN / ARCHIVED
                created_at_utc TEXT NOT NULL,
                FOREIGN KEY(customer_id) REFERENCES customers(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_orders_customer_status
            ON orders(customer_id, status);

            CREATE TABLE IF NOT EXISTS order_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                card_name TEXT NOT NULL,
                qty INTEGER NOT NULL,
                display_item TEXT NOT NULL,
                rarity TEXT,
                notes TEXT,
                raw_line TEXT,
                created_at_utc TEXT NOT NULL,
                FOREIGN KEY(order_id) REFERENCES orders(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_lines_order
            ON order_lines(order_id);
            """
        )
        conn.commit()
    finally:
        conn.close()


@app.on_event("startup")
def _startup():
    init_db()


def _collapse_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


# --- customers ---
def get_customer_by_exact_name(name: str) -> Optional[sqlite3.Row]:
    name = _collapse_spaces(name)
    conn = db_connect()
    try:
        cur = conn.execute("SELECT id, name FROM customers WHERE name = ?", (name,))
        return cur.fetchone()
    finally:
        conn.close()


def create_customer(name: str) -> Tuple[bool, Optional[sqlite3.Row], str]:
    name = _collapse_spaces(name)
    if not name:
        return (False, None, "Empty name.")
    conn = db_connect()
    try:
        try:
            conn.execute(
                "INSERT INTO customers (name, created_at_utc) VALUES (?, ?)",
                (name, utc_now_iso()),
            )
            conn.commit()
            row = get_customer_by_exact_name(name)
            return (True, row, f'Created "{name}".')
        except sqlite3.IntegrityError:
            row = get_customer_by_exact_name(name)
            return (False, row, f'"{name}" already exists.')
    finally:
        conn.close()


def search_customers(q: str, limit: int = 20) -> List[sqlite3.Row]:
    q = _collapse_spaces(q)
    conn = db_connect()
    try:
        if not q:
            cur = conn.execute("SELECT id, name FROM customers ORDER BY name ASC LIMIT ?", (limit,))
            return cur.fetchall()
        like = f"%{q}%"
        cur = conn.execute(
            "SELECT id, name FROM customers WHERE name LIKE ? ORDER BY name ASC LIMIT ?",
            (like, limit),
        )
        return cur.fetchall()
    finally:
        conn.close()


# --- orders & lines ---
def create_order(customer_id: int, pasted_text: str, global_note: Optional[str]) -> int:
    conn = db_connect()
    try:
        cur = conn.execute(
            """
            INSERT INTO orders (customer_id, pasted_text, global_note, status, created_at_utc)
            VALUES (?, ?, ?, 'OPEN', ?)
            """,
            (customer_id, pasted_text, global_note, utc_now_iso()),
        )
        conn.commit()
        return int(cur.lastrowid)
    finally:
        conn.close()


def add_lines(order_id: int, lines: List["ParsedLine"]) -> None:
    conn = db_connect()
    try:
        now = utc_now_iso()
        conn.executemany(
            """
            INSERT INTO order_lines
            (order_id, card_name, qty, display_item, rarity, notes, raw_line, created_at_utc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    order_id,
                    ln.card_name_raw,
                    int(ln.quantity),
                    ln.display_item,
                    ln.rarity or "",
                    ln.notes or "",
                    ln.raw_line,
                    now,
                )
                for ln in lines
            ],
        )
        conn.commit()
    finally:
        conn.close()


def get_open_basket_grouped() -> List[Dict]:
    """
    Returns list of:
      {
        customer_id, customer_name,
        orders: [{order_id, created_at_utc, global_note, pasted_text}],
        lines:  [{line_id, order_id, card_name, qty, display_item, rarity, notes, raw_line}]
      }
    """
    conn = db_connect()
    try:
        # customers with any OPEN orders
        customers = conn.execute(
            """
            SELECT DISTINCT c.id AS customer_id, c.name AS customer_name
            FROM customers c
            JOIN orders o ON o.customer_id = c.id
            WHERE o.status = 'OPEN'
            ORDER BY c.name ASC
            """
        ).fetchall()

        result: List[Dict] = []
        for c in customers:
            cid = int(c["customer_id"])
            orders = conn.execute(
                """
                SELECT id AS order_id, created_at_utc, global_note, pasted_text
                FROM orders
                WHERE customer_id = ? AND status = 'OPEN'
                ORDER BY id DESC
                """,
                (cid,),
            ).fetchall()

            order_ids = [int(o["order_id"]) for o in orders]
            lines: List[sqlite3.Row] = []
            if order_ids:
                placeholders = ",".join(["?"] * len(order_ids))
                lines = conn.execute(
                    f"""
                    SELECT
                      l.id AS line_id,
                      l.order_id,
                      l.card_name,
                      l.qty,
                      l.display_item,
                      l.rarity,
                      l.notes,
                      l.raw_line
                    FROM order_lines l
                    WHERE l.order_id IN ({placeholders})
                    ORDER BY l.id ASC
                    """,
                    tuple(order_ids),
                ).fetchall()

            result.append(
                {
                    "customer_id": cid,
                    "customer_name": str(c["customer_name"]),
                    "orders": [dict(o) for o in orders],
                    "lines": [dict(l) for l in lines],
                }
            )
        return result
    finally:
        conn.close()


def delete_open_orders_for_customer(customer_id: int) -> None:
    conn = db_connect()
    try:
        conn.execute("DELETE FROM orders WHERE customer_id = ? AND status = 'OPEN'", (customer_id,))
        conn.commit()
    finally:
        conn.close()


def clear_basket() -> None:
    conn = db_connect()
    try:
        conn.execute("DELETE FROM orders WHERE status = 'OPEN'")
        conn.commit()
    finally:
        conn.close()


def update_line(line_id: int, card_name: str, qty: int, rarity: str, notes: str) -> None:
    card_name = _collapse_spaces(card_name)
    if qty < 1:
        qty = 1
    display_item = f"{card_name} {qty}"
    conn = db_connect()
    try:
        conn.execute(
            """
            UPDATE order_lines
            SET card_name = ?, qty = ?, display_item = ?, rarity = ?, notes = ?
            WHERE id = ?
            """,
            (card_name, int(qty), display_item, rarity or "", notes or "", int(line_id)),
        )
        conn.commit()
    finally:
        conn.close()


def delete_line(line_id: int) -> None:
    conn = db_connect()
    try:
        conn.execute("DELETE FROM order_lines WHERE id = ?", (int(line_id),))
        conn.commit()
    finally:
        conn.close()


def add_manual_line_to_customer_open(customer_id: int, card_name: str, qty: int, rarity: str, notes: str) -> None:
    """
    Adds a line under a new OPEN order for that customer (so it belongs to basket).
    """
    card_name = _collapse_spaces(card_name)
    if not card_name:
        return
    if qty < 1:
        qty = 1

    order_id = create_order(customer_id, pasted_text="(manual)", global_note=None)

    display_item = f"{card_name} {qty}"
    conn = db_connect()
    try:
        conn.execute(
            """
            INSERT INTO order_lines
            (order_id, card_name, qty, display_item, rarity, notes, raw_line, created_at_utc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                order_id,
                card_name,
                int(qty),
                display_item,
                rarity or "",
                notes or "",
                "(manual)",
                utc_now_iso(),
            ),
        )
        conn.commit()
    finally:
        conn.close()


def get_customer_open_orders_and_lines(customer_id: int) -> Tuple[str, List[Dict], List[Dict]]:
    conn = db_connect()
    try:
        c = conn.execute("SELECT id, name FROM customers WHERE id = ?", (int(customer_id),)).fetchone()
        if not c:
            raise ValueError("Customer not found.")

        orders = conn.execute(
            """
            SELECT id AS order_id, created_at_utc, global_note, pasted_text
            FROM orders
            WHERE customer_id = ? AND status = 'OPEN'
            ORDER BY id DESC
            """,
            (int(customer_id),),
        ).fetchall()
        order_ids = [int(o["order_id"]) for o in orders]

        lines: List[sqlite3.Row] = []
        if order_ids:
            placeholders = ",".join(["?"] * len(order_ids))
            lines = conn.execute(
                f"""
                SELECT
                  l.id AS line_id,
                  l.order_id,
                  l.card_name,
                  l.qty,
                  l.display_item,
                  l.rarity,
                  l.notes,
                  l.raw_line
                FROM order_lines l
                WHERE l.order_id IN ({placeholders})
                ORDER BY l.id ASC
                """,
                tuple(order_ids),
            ).fetchall()

        return (str(c["name"]), [dict(o) for o in orders], [dict(l) for l in lines])
    finally:
        conn.close()


# ----------------------------
# Parsing (Hebrew + mixed English)
# ----------------------------
RARITY_WORDS = re.compile(r"\b(Rare|Secret|Ultra|Super|Platinum|Common|SR|UR|SCR|CR)\b", re.IGNORECASE)
HEBREW_NOTE_KEYWORDS = ["קומון", "קומונים", "פלייסט", "פלייסטים", "גרסה", "לא אכפת", "לא משנה"]

DIVIDER_RE = re.compile(r"^[\-\–\—_=]{3,}\s*$")
PAREN_RE = re.compile(r"\(([^)]*)\)")

RE_TRAIL_X = re.compile(r"^(?P<name>.+?)\s*[x×]\s*(?P<qty>\d{1,3})\s*$", re.IGNORECASE)
RE_HE_QTYWORD = re.compile(r"^(?P<name>.+?)\s+(?:כמות|עותקים|יחידות)\s*(?P<qty>\d{1,3})\s*$")
RE_TRAIL_NUM = re.compile(r"^(?P<name>.+?)\s+(?P<qty>\d{1,3})\s*$")
RE_LEAD_NUM = re.compile(r"^(?P<qty>\d{1,3})\s+(?P<name>.+?)\s*$")
RE_LEAD_NUM_STUCK = re.compile(r"^(?P<qty>\d{1,3})(?P<name>[A-Za-z].+?)\s*$")

HEBREW_RANGE = re.compile(r"[\u0590-\u05FF]")
LATIN_RANGE = re.compile(r"[A-Za-z]")


@dataclass
class ParsedLine:
    raw_line: str
    card_name_raw: str
    quantity: int
    display_item: str
    rarity: Optional[str] = None
    notes: Optional[str] = None
    needs_review: bool = False


def split_candidates(text: str) -> List[str]:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\-_=\u2014\u2013]{3,}", "\n", text)
    raw_lines = [ln.strip() for ln in text.split("\n")]

    cleaned: List[str] = []
    for ln in raw_lines:
        if not ln:
            continue
        if DIVIDER_RE.match(ln):
            continue
        base = ln.replace(":", "").strip()
        if base in ("הזמנה", "הזמנות", "רשימה"):
            continue
        if ln.startswith("הזמנה") and not LATIN_RANGE.search(ln) and not re.search(r"\d", ln):
            continue
        cleaned.append(ln)
    return cleaned


def detect_global_note(lines: List[str]) -> Tuple[Optional[str], List[str]]:
    if not lines:
        return None, lines
    first = lines[0].strip()
    if not first:
        return None, lines
    if not HEBREW_RANGE.search(first):
        return None, lines
    if not any(k in first for k in HEBREW_NOTE_KEYWORDS):
        return None, lines
    if LATIN_RANGE.search(first) or re.search(r"\d", first):
        return None, lines
    return first, lines[1:]


def parse_line(line: str) -> ParsedLine:
    raw = line
    rarity: Optional[str] = None
    notes_parts: List[str] = []

    def _paren_repl(m: re.Match) -> str:
        nonlocal rarity
        inner = _collapse_spaces(m.group(1))
        if not inner:
            return ""
        if RARITY_WORDS.search(inner):
            rarity = inner
        else:
            notes_parts.append(inner)
        return ""

    line_wo_paren = PAREN_RE.sub(_paren_repl, line)
    line_wo_paren = _collapse_spaces(line_wo_paren)

    qty: Optional[int] = None
    name: Optional[str] = None
    for rx in (RE_TRAIL_X, RE_HE_QTYWORD, RE_TRAIL_NUM, RE_LEAD_NUM, RE_LEAD_NUM_STUCK):
        m = rx.match(line_wo_paren)
        if m:
            name = _collapse_spaces(m.group("name"))
            qty = int(m.group("qty"))
            break

    if qty is None or name is None:
        qty = 1
        name = line_wo_paren

    name = name.strip(" -•*:\t")
    name = _collapse_spaces(name)

    display = f"{name} {qty}"  # HARD RULE: qty after name
    needs_review = (not name) or (len(name) < 3) or (qty > 20)
    notes = _collapse_spaces(" | ".join(notes_parts)) if notes_parts else None

    return ParsedLine(raw, name, int(qty), display, rarity, notes, needs_review)


def parse_order(text: str) -> Tuple[Optional[str], List[ParsedLine]]:
    lines = split_candidates(text)
    global_note, remaining = detect_global_note(lines)
    parsed: List[ParsedLine] = []
    for ln in remaining:
        if "תודה" in ln and len(ln) <= 6:
            continue
        pl = parse_line(ln)
        if pl.card_name_raw:
            parsed.append(pl)
    return global_note, parsed


# ----------------------------
# Excel generation
# ----------------------------
def style_sheet(ws, col_widths: Dict[int, float]):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}1"

    header_font = Font(bold=True)
    header_align = Alignment(vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_align

    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wrap_cols = {3, 7, 8}  # CardName, Notes, RawLine
    top_wrap = Alignment(vertical="top", wrap_text=True)
    top_nowrap = Alignment(vertical="top", wrap_text=False)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = top_wrap if cell.column in wrap_cols else top_nowrap


def build_excel_from_lines(customer_name: str, customer_id: int, lines: List[Dict], orders: List[Dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "OrderLines"
    headers = ["CustomerName", "CustomerID", "CardName", "Qty", "DisplayItem", "Rarity", "Notes", "RawLine"]
    ws.append(headers)

    # Map global_note by order_id for better notes composition (if needed)
    order_note: Dict[int, str] = {}
    for o in orders:
        oid = int(o["order_id"])
        gn = (o.get("global_note") or "").strip()
        order_note[oid] = gn

    for l in lines:
        oid = int(l["order_id"])
        gn = order_note.get(oid, "")
        notes = (l.get("notes") or "").strip()
        if gn:
            notes = (notes + " | " if notes else "") + gn

        ws.append(
            [
                customer_name,
                customer_id,
                l["card_name"],
                int(l["qty"]),
                l["display_item"],
                l.get("rarity") or "",
                notes,
                l.get("raw_line") or "",
            ]
        )

    style_sheet(ws, {1: 22, 2: 12, 3: 38, 4: 6, 5: 42, 6: 20, 7: 40, 8: 50})

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_excel_combined_from_db(groups: List[Dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "CombinedOrderLines"
    headers = ["CustomerName", "CustomerID", "CardName", "Qty", "DisplayItem", "Rarity", "Notes", "RawLine"]
    ws.append(headers)

    totals: Dict[str, int] = {}

    for g in groups:
        cname = g["customer_name"]
        cid = g["customer_id"]

        order_note: Dict[int, str] = {}
        for o in g["orders"]:
            order_note[int(o["order_id"])] = (o.get("global_note") or "").strip()

        for l in g["lines"]:
            oid = int(l["order_id"])
            gn = order_note.get(oid, "")
            notes = (l.get("notes") or "").strip()
            if gn:
                notes = (notes + " | " if notes else "") + gn

            ws.append(
                [
                    cname,
                    cid,
                    l["card_name"],
                    int(l["qty"]),
                    l["display_item"],
                    l.get("rarity") or "",
                    notes,
                    l.get("raw_line") or "",
                ]
            )

            key = str(l["card_name"]).strip().lower()
            totals[key] = totals.get(key, 0) + int(l["qty"])

    style_sheet(ws, {1: 22, 2: 12, 3: 38, 4: 6, 5: 42, 6: 20, 7: 40, 8: 50})

    summary = wb.create_sheet("Summary")
    summary.append(["CardName", "TotalQty"])
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = "A1:B1"
    summary["A1"].font = Font(bold=True)
    summary["B1"].font = Font(bold=True)
    summary.column_dimensions["A"].width = 38
    summary.column_dimensions["B"].width = 10
    for key in sorted(totals.keys()):
        summary.append([key, totals[key]])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ----------------------------
# HTML
# ----------------------------
def page_layout(body: str) -> str:
    return f"""<!doctype html>
<html lang="he">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Yugioh Orders</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 18px; }}
    .card {{ border: 1px solid #ddd; padding: 12px; border-radius: 10px; margin-top: 12px; }}
    .row {{ display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
    input[type="text"] {{ padding: 8px; width: 360px; }}
    input[type="number"] {{ padding: 6px; width: 90px; }}
    textarea {{ width:100%; height: 220px; padding: 8px; }}
    button {{ padding: 8px 12px; cursor:pointer; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; }}
    th, td {{ border: 1px solid #ddd; padding: 8px; vertical-align: top; }}
    th {{ background: #f4f4f4; }}
    .muted {{ color:#666; }}
    .warn {{ background:#fff4d6; border:1px solid #f0d48a; padding:10px; border-radius:10px; }}
    .pill {{ display:inline-block; padding:2px 10px; border-radius:999px; background:#eee; margin:2px; }}
    .sectionTitle {{ display:flex; justify-content:space-between; align-items:center; gap:10px; flex-wrap:wrap; }}
    .small {{ font-size: 12px; }}
    .tight td {{ padding: 6px; }}

    /* Ping status badge */
    .pingBadge {{
      position: fixed;
      right: 14px;
      bottom: 14px;
      z-index: 9999;
      border: 1px solid #ddd;
      background: #fff;
      border-radius: 999px;
      padding: 8px 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.08);
      font-size: 12px;
      display: flex;
      gap: 8px;
      align-items: center;
    }}
    .dot {{
      width: 10px;
      height: 10px;
      border-radius: 50%;
      background: #bbb;
      display: inline-block;
    }}
    .dot.ok {{ background: #2ecc71; }}
    .dot.bad {{ background: #e74c3c; }}
    .pingError {{
      position: fixed;
      left: 18px;
      right: 18px;
      bottom: 56px;
      z-index: 9998;
      display: none;
    }}
  </style>
</head>
<body>

<div id="pingError" class="warn pingError">
  <b>Server connection problem.</b>
  <span class="muted">Refreshing may help. If this continues, the server may be down.</span>
</div>

<div class="pingBadge" title="Auto ping every 2 minutes">
  <span id="pingDot" class="dot"></span>
  <span id="pingText" class="muted">Checking server…</span>
</div>

{body}

<script>
(function() {{
  const dot = document.getElementById("pingDot");
  const text = document.getElementById("pingText");
  const err = document.getElementById("pingError");

  function setOk(msg) {{
    dot.classList.remove("bad");
    dot.classList.add("ok");
    text.textContent = msg;
    err.style.display = "none";
  }}

  function setBad(msg) {{
    dot.classList.remove("ok");
    dot.classList.add("bad");
    text.textContent = msg;
    err.style.display = "block";
  }}

  async function doPing() {{
    const controller = new AbortController();
    const t = setTimeout(() => controller.abort(), 4000); // 4s timeout

    try {{
      const res = await fetch("/ping", {{
        cache: "no-store",
        signal: controller.signal
      }});
      clearTimeout(t);

      if (!res.ok) {{
        setBad("Ping failed (HTTP " + res.status + ")");
        return;
      }}
      const data = await res.json();
      if (data && data.ok) {{
        setOk("Server OK");
      }} else {{
        setBad("Ping failed (bad response)");
      }}
    }} catch (e) {{
      clearTimeout(t);
      setBad("Server not reachable");
    }}
  }}

  // First ping immediately, then every 2 minutes
  doPing();
  setInterval(doPing, 120000);
}})();
</script>

</body></html>"""



@app.get("/", response_class=HTMLResponse)
def home(q: str = Query("", description="search customers")):
    results = search_customers(q, limit=15)
    result_pills = "".join([f'<span class="pill">{r["name"]} (ID {r["id"]})</span>' for r in results])

    groups = get_open_basket_grouped()
    basket_users = len(groups)
    basket_items = sum(len(g["lines"]) for g in groups)

    body = f"""
<h2>Paste Order → Basket</h2>

<div class="card">
  <div class="row">
    <form method="get" action="/" class="row">
      <label><b>Search users</b></label>
      <input type="text" name="q" value="{q.replace('"','&quot;')}" placeholder="Search by name (e.g. Mor)" />
      <button type="submit">Search</button>
    </form>
  </div>
  <div class="muted" style="margin-top:8px;">Results:</div>
  <div>{result_pills or "<span class='muted'>No results</span>"}</div>
</div>

<div class="card">
  <form method="post" action="/add_to_basket">
    <!-- Row 1: username -->
    <div class="row">
      <label><b>Recipient name (exact)</b></label>
      <input type="text" name="recipient_name" placeholder="e.g. Name Customer" required />
      <button type="submit" name="action" value="find">Find</button>
      <button type="submit" name="action" value="create">Create</button>
      <span class="muted">Use Create only if this exact name is new.</span>
    </div>

    <!-- Row 2: cards -->
    <div style="margin-top:10px;">
      <label><b>Cards (paste WhatsApp text)</b></label>
      <textarea name="order_text" placeholder="Paste the order here..." required></textarea>
    </div>

    <div class="row" style="margin-top:10px;">
      <button type="submit" name="action" value="add">Add to Basket</button>
      <a href="/basket">Open Basket</a>
      <span class="muted">Basket: {basket_users} users, {basket_items} items</span>
    </div>
  </form>
</div>

<p class="muted">
Flow: type recipient name → click Find (must exist) or Create (new) → paste cards → Add to Basket.
</p>
"""
    return page_layout(body)


@app.post("/add_to_basket", response_class=HTMLResponse)
def add_to_basket(
    recipient_name: str = Form(...),
    order_text: str = Form(...),
    action: str = Form(...),  # find | create | add
):
    recipient_name = _collapse_spaces(recipient_name)
    order_text = order_text or ""

    if action in ("find", "create"):
        row = get_customer_by_exact_name(recipient_name)
        if action == "find":
            if not row:
                body = f"""
<h2>User not found</h2>
<div class="warn">
  "{recipient_name}" is not in the database. Use <b>Create</b> to add a new unique name.
</div>
<div style="margin-top:10px;"><a href="/">← Back</a></div>
"""
                return page_layout(body)
            body = f"""
<h2>User found</h2>
<div class="card">
  <b>{row["name"]}</b> (ID {row["id"]}) is ready. Now paste cards and press <b>Add to Basket</b>.
</div>
<div style="margin-top:10px;"><a href="/">← Back</a></div>
"""
            return page_layout(body)

        created, new_row, msg = create_customer(recipient_name)
        if new_row is None:
            body = f"""
<h2>Cannot create</h2>
<div class="warn">{msg}</div>
<div style="margin-top:10px;"><a href="/">← Back</a></div>
"""
            return page_layout(body)

        body = f"""
<h2>Create result</h2>
<div class="card">
  {msg} (ID {new_row["id"]})
  <div class="muted" style="margin-top:8px;">Now paste cards and press <b>Add to Basket</b>.</div>
</div>
<div style="margin-top:10px;"><a href="/">← Back</a></div>
"""
        return page_layout(body)

    # action == add
    row = get_customer_by_exact_name(recipient_name)
    if not row:
        body = f"""
<h2>Cannot add to basket</h2>
<div class="warn">
  "{recipient_name}" is not in the database.
  Click <b>Create</b> first (for a new unique name) or <b>Find</b> if it already exists.
</div>
<div style="margin-top:10px;"><a href="/">← Back</a></div>
"""
        return page_layout(body)

    customer_id = int(row["id"])
    global_note, lines = parse_order(order_text)

    order_id = create_order(customer_id, pasted_text=order_text, global_note=global_note)
    add_lines(order_id, lines)

    return RedirectResponse("/basket", status_code=303)


@app.get("/basket", response_class=HTMLResponse)
def basket():
    groups = get_open_basket_grouped()
    if not groups:
        return page_layout("""
<h2>Basket</h2>
<div class="card">
  <div class="muted">Basket is empty.</div>
  <a href="/">← Back</a>
</div>
""")

    sections_html = ""
    for g in groups:
        cid = g["customer_id"]
        cname = g["customer_name"]

        # Build table rows with inline edit forms
        rows_html = ""
        for l in g["lines"]:
            line_id = int(l["line_id"])
            rows_html += f"""
<tr>
  <td class="small muted">#{line_id}</td>
  <td>
    <form method="post" action="/line_update" class="row" style="gap:6px;">
      <input type="hidden" name="line_id" value="{line_id}"/>
      <input type="hidden" name="customer_id" value="{cid}"/>
      <input type="text" name="card_name" value="{str(l["card_name"]).replace('"','&quot;')}" style="width:280px;" required/>
  </td>
  <td>
      <input type="number" name="qty" value="{int(l["qty"])}" min="1" />
  </td>
  <td>
      <input type="text" name="rarity" value="{str(l.get("rarity") or '').replace('"','&quot;')}" style="width:160px;" />
  </td>
  <td>
      <input type="text" name="notes" value="{str(l.get("notes") or '').replace('"','&quot;')}" style="width:260px;" />
  </td>
  <td>
      <button type="submit">Save</button>
    </form>
  </td>
  <td>
    <form method="post" action="/line_delete">
      <input type="hidden" name="line_id" value="{line_id}"/>
      <button type="submit">Delete</button>
    </form>
  </td>
</tr>
"""

        # Add manual line row
        add_row = f"""
<tr>
  <td class="small muted">+</td>
  <td colspan="6">
    <form method="post" action="/line_add" class="row" style="gap:6px;">
      <input type="hidden" name="customer_id" value="{cid}"/>
      <input type="text" name="card_name" placeholder="New card name" style="width:280px;" required/>
      <input type="number" name="qty" value="1" min="1" />
      <input type="text" name="rarity" placeholder="Rarity (optional)" style="width:160px;" />
      <input type="text" name="notes" placeholder="Notes (optional)" style="width:260px;" />
      <button type="submit">Add</button>
    </form>
  </td>
</tr>
"""

        sections_html += f"""
<div class="card">
  <div class="sectionTitle">
    <div>
      <b>{cname}</b> <span class="muted">(ID {cid})</span>
      <div class="muted small">Lines: {len(g["lines"])} | Orders (messages): {len(g["orders"])}</div>
    </div>
    <div class="row">
      <form method="post" action="/download_user">
        <input type="hidden" name="customer_id" value="{cid}"/>
        <button type="submit">Download Excel</button>
      </form>
      <form method="post" action="/remove_user">
        <input type="hidden" name="customer_id" value="{cid}"/>
        <button type="submit">Remove User From Basket</button>
      </form>
    </div>
  </div>

  <table class="tight">
    <thead>
      <tr>
        <th>#</th>
        <th>CardName</th>
        <th>Qty</th>
        <th>Rarity</th>
        <th>Notes</th>
        <th>Save</th>
        <th>Delete</th>
      </tr>
    </thead>
    <tbody>
      {rows_html or "<tr><td colspan='7' class='muted'>No items.</td></tr>"}
      {add_row}
    </tbody>
  </table>
</div>
"""

    basket_users = len(groups)
    basket_items = sum(len(g["lines"]) for g in groups)

    body = f"""
<h2>Basket</h2>

<div class="card">
  <div class="row">
    <a href="/">← Back</a>
    <span class="muted">Users: {basket_users} | Items: {basket_items}</span>
    <form method="post" action="/download_combined">
      <button type="submit">Download Combined Excel</button>
    </form>
    <form method="post" action="/clear_basket">
      <button type="submit">Clear Basket</button>
    </form>
  </div>
</div>

{sections_html}
"""
    return page_layout(body)


@app.post("/line_update")
def line_update(
    line_id: int = Form(...),
    customer_id: int = Form(...),
    card_name: str = Form(...),
    qty: int = Form(...),
    rarity: str = Form(""),
    notes: str = Form(""),
):
    update_line(line_id=int(line_id), card_name=card_name, qty=int(qty), rarity=rarity or "", notes=notes or "")
    return RedirectResponse("/basket", status_code=303)


@app.post("/line_delete")
def line_delete(line_id: int = Form(...)):
    delete_line(int(line_id))
    return RedirectResponse("/basket", status_code=303)


@app.post("/line_add")
def line_add(
    customer_id: int = Form(...),
    card_name: str = Form(...),
    qty: int = Form(...),
    rarity: str = Form(""),
    notes: str = Form(""),
):
    add_manual_line_to_customer_open(int(customer_id), card_name, int(qty), rarity or "", notes or "")
    return RedirectResponse("/basket", status_code=303)


@app.post("/remove_user")
def remove_user(customer_id: int = Form(...)):
    delete_open_orders_for_customer(int(customer_id))
    return RedirectResponse("/basket", status_code=303)


@app.post("/clear_basket")
def clear_basket_route():
    clear_basket()
    return RedirectResponse("/basket", status_code=303)


@app.post("/download_user")
def download_user(customer_id: int = Form(...)):
    cid = int(customer_id)
    cname, orders, lines = get_customer_open_orders_and_lines(cid)

    if not lines:
        return RedirectResponse("/basket", status_code=303)

    data = build_excel_from_lines(cname, cid, lines, orders)
    today = datetime.now().date().isoformat()
    safe_name = _collapse_spaces(cname).replace("/", "-").replace("\\", "-")
    filename = f"{today} - {safe_name}.xlsx"

    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/download_combined")
def download_combined():
    groups = get_open_basket_grouped()
    if not groups:
        return RedirectResponse("/basket", status_code=303)

    data = build_excel_combined_from_db(groups)
    today = datetime.now().date().isoformat()
    filename = f"{today} - Combined Orders.xlsx"

    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

@app.get("/ping")
def ping():
    return {"ok": True, "utc": utc_now_iso()}
