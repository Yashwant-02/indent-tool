"""
INDENT TOOL - single-file Streamlit app
=======================================

Phase 1  Tab "1 · Build Indent Workbook"
    Download the input template, fill it, upload it, choose the brand, download TR_INDENT_<BRAND>.xlsx.
    Required column names are in red in the template.
    Sheets   WH             HSN, Section, Barcode, Article, Color, Size, MRP, Qty
             Current Stock  Store Code, Brand, Section, Qty                      (+ Store Name, Article, Color, Size, Barcode)
             Sale           Store Code, Brand, Section, Qty        (optional)    (+ Store Name)
             Outward        Store Code, Section, Date, Qty         (optional)    (+ Store Name, Article, Color, Size, Barcode)
    - Only WH and Current Stock are needed. Sheet and column names ignore case and extra spaces.
    - WH and Outward have no Brand: WH must be of the chosen brand, and so must the Outward.
    - Only the chosen brand is used from Sale and Current Stock. Only WH sections are used.
    - Stores: the codes typed in the box. If the box is empty, every store of the brand in Current Stock.
    - Outward: if a date is chosen, only rows on or after that date are used (older Outward has already
      reached the store). Without a date, all Outward is used.
    - Stores are ranked by sale (by stock, highest first, when there is no Sale sheet).
      Summary rows show soh, and sale and outward when those sheets are used (sale/soh/out).
    - Each WH row also shows, in light red text in every store column, the SOH (and Outward) the store
      already has for the same article + color + size. It is text, so the Indent and SHORT formulas
      ignore it. Typing a qty over it makes a normal number. Without Article/Color/Size only the
      summary rows are shown.
    - The first line of the instructions block shows the brand, the Outward date and the store count.

Manual   Allocate quantities in the store columns of Sheet2 (column L onwards), save in Excel.

Phase 2  Tab "2 · Fill Platform Templates"
    Upload the allocated sheet, choose the indent date, download:
      FF Xpress Store Indent  FF_XPRESS_STORE_INDENT<DDMMYYYY>.xlsx
      PO Indent               PO_INDENT_<DD-Mon-YY>.xlsx
      Plain Data              INDENT_DATA_<DDMMYYYY>.xlsx  (all fields, to copy into any other format)

Files
    app.py                  this file (screen + all logic)
    requirements.txt        libraries, versions fixed so the app runs the same everywhere
    .streamlit/config.toml  light theme and upload limit
    templates/*.xlsx        blank platform templates (column mapping is in PLATFORMS below)

Run on a PC       pip install -r requirements.txt   then   streamlit run app.py
Streamlit Cloud   Put these files in a GitHub repo. On share.streamlit.io choose New app, pick the repo,
                  main file app.py, Advanced settings: Python 3.11 or newer, Deploy.
                  Uploaded files are processed in memory only; nothing is saved.

Adding a platform format
    Put the blank template in templates/ and add one entry to PLATFORMS. "columns" maps a template
    header (case and extra spaces ignored) to a field: store_code, store_name, ean, qty, mrp, hsn,
    section, article, color, size, indent_date, season_year - or {"const": "EA"} for a fixed value.
    "as": "int" is optional. Headers that are not listed stay blank.
"""
import datetime as dt
import io
import re
import zipfile
from copy import copy
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =================================================================== SETTINGS
TEMPLATE_DIR = Path(__file__).parent / "templates"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

WH_COLS = ["HSN", "Section", "Barcode", "Article", "Color", "Size", "MRP", "Qty"]
SHEETS = {   # all columns of each input sheet, in template order
    "WH": WH_COLS,
    "Sale": ["Store Code", "Store Name", "Brand", "Section", "Qty"],
    "Current Stock": ["Store Code", "Store Name", "Brand", "Section", "Article", "Color", "Size", "Barcode", "Qty"],
    "Outward": ["Store Code", "Store Name", "Section", "Date", "Article", "Color", "Size", "Barcode", "Qty"],
}
REQUIRED = {   # the other columns may be missing or empty
    "WH": WH_COLS,
    "Sale": ["Store Code", "Brand", "Section", "Qty"],
    "Current Stock": ["Store Code", "Brand", "Section", "Qty"],
    "Outward": ["Store Code", "Section", "Date", "Qty"],
}
SIZE_ORDER = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"]

FONT = "Aptos Narrow"
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YELLOW = PatternFill("solid", fgColor="FFFFFF00")
SOFT_RED = "FFE08A8A"   # soh/out text colour (faded on purpose)
CENTER = Alignment(horizontal="center", vertical="center")
FIRST_STORE_COL = 12  # L
MIN_LAST_COL = 62     # BJ (formula range end, kept as in original)

# Written in the empty block at the top left of the indent workbook.
INSTRUCTIONS = [
    "Type the qty to send in each store column (from column L). Leave it blank for none.",
    "Use whole numbers only.",
    "Indent (column J) is the total sent. If SHORT (column K) is negative, reduce the qty.",
    "Rows above show {top} for each store and section. Row 1 is the store rank by {rank}.",
    "Light red text = {art} of that article, color and size already at the store. "
    "View only. Type a qty over it to allocate.",
    "Do not delete or move the yellow header row, the store codes or columns A to K.",
    "Save the file in Excel, then upload it in tab 2.",
]

# Platform output formats (Phase 2). One blank template per platform in templates/.
PLATFORMS = [
    {
        "name": "FF Xpress Store Indent",
        "template_file": "FF_XPRESS_STORE_INDENT.xlsx",
        "sheet": "FF Pune Handewadi",
        "header_row": 2,
        "first_data_row": 3,
        "output_name": "FF_XPRESS_STORE_INDENT{date:%d%m%Y}.xlsx",
        "columns": {
            "INDENT DATE": {"field": "indent_date"},
            "STORE CODE": {"field": "store_code"},
            "STORE NAME": {"field": "store_name"},
            "EAN": {"field": "ean", "as": "int"},
            "QTY": {"field": "qty"},
            "UNIT OF MEASURE": {"const": "N/A"},
            "MRP": {"field": "mrp"},
            "HSN": {"field": "hsn", "as": "int"},
            "SEASON YEAR": {"field": "season_year"},
            "PRODUCT": {"field": "section"},
            "Style NO/ Name": {"field": "article"},
            "Colour": {"field": "color"},
            "Size": {"field": "size"},
        },
    },
    {
        "name": "PO Indent",
        "template_file": "PO_INDENT_TEMPLATE.xlsx",
        "sheet": "IM_RECORDS",
        "header_row": 1,
        "first_data_row": 2,
        "output_name": "PO_INDENT_{date:%d-%b-%y}.xlsx",
        "columns": {
            "Store/DC Code": {"field": "store_code"},
            "Substore": {"field": "store_code"},
            "EAN": {"field": "ean", "as": "int"},
            "Total Qty": {"field": "qty"},
            "Unit Of Measure": {"const": "EA"},
            "MRP": {"field": "mrp"},
            "HSN": {"field": "hsn", "as": "int"},
            "Season Year": {"field": "season_year"},
        },
    },
]

# Plain list of everything that was allocated (no platform template).
PLAIN_DATA = {
    "name": "Plain Data",
    "output_name": "INDENT_DATA_{date:%d%m%Y}.xlsx",
    "columns": {
        "Indent Date": {"field": "indent_date"},
        "Store Code": {"field": "store_code"},
        "Store Name": {"field": "store_name"},
        "EAN": {"field": "ean", "as": "int"},
        "HSN": {"field": "hsn", "as": "int"},
        "Section": {"field": "section"},
        "Article": {"field": "article"},
        "Color": {"field": "color"},
        "Size": {"field": "size"},
        "MRP": {"field": "mrp"},
        "Qty": {"field": "qty"},
        "Season Year": {"field": "season_year"},
    },
    "widths": [11, 11, 34, 16, 11, 16, 12, 16, 7, 9, 7, 11],
}
OUTPUTS = PLATFORMS + [PLAIN_DATA]


# =================================================================== PHASE 1: input template
def make_template() -> bytes:
    """Blank input file. Required column names are in red."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, cols in SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(cols)
        for i, col in enumerate(cols, 1):
            c = ws.cell(1, i)
            c.font = Font(name=FONT, bold=True, color="FFFF0000" if col in REQUIRED[name] else None)
            c.fill = YELLOW
            c.alignment = CENTER
            c.border = BORDER
            ws.column_dimensions[get_column_letter(i)].width = 18
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# =================================================================== PHASE 1: reading
def _clean_text(s):
    s = s.where(s.notna(), "").astype(str).str.strip()
    return s.replace({"nan": "", "None": "", "<NA>": ""})


def _code(v):
    """Store code -> text key (1234, 1234.0, '1234' and ' 1234.0 ' are the same; 'tbs1' -> 'TBS1')."""
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    v = str(v).strip()
    if re.fullmatch(r"\d+\.0+", v):
        v = v.split(".")[0]
    return v.upper()


def _num(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError):
        return v


def _norm(s):
    return re.sub(r"\s+", " ", str(s if s is not None else "")).strip().casefold()


def _item(article, color, size):
    """Article + color + size -> text key (case and spacing ignored). '' when there is no article."""
    if pd.isna(article) or str(article).strip() == "":
        return ""
    a = str(_num(str(article).strip()))
    c = "" if pd.isna(color) else re.sub(r"\s+", " ", str(color)).strip().upper()
    z = "" if pd.isna(size) else str(size).strip().upper()
    return f"{a}|{c}|{z}"


def _store_cell(code):
    """Digits-only code (no leading zero) is written as a number, else as text."""
    if code.isdigit() and not code.startswith("0"):
        return int(code)
    return code


def _to_date(v):
    """Excel date, Excel serial number or text (dd-mm-yyyy) -> date. None when it is not a date."""
    if v is None or (isinstance(v, float) and v != v) or str(v).strip() == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date() if 20000 < v < 80000 else None
    d = pd.to_datetime(str(v).strip(), dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.date()


def _read(xl, sheet, optional=False):
    """Read one input sheet (name and column names ignore case and extra spaces).
    An optional sheet that is missing or has no rows returns None."""
    actual = {_norm(n): n for n in xl.sheet_names}.get(_norm(sheet))
    if actual is None:
        if optional:
            return None
        raise ValueError(f"Sheet '{sheet}' was not found. Please use the input template.")
    df = xl.parse(actual, dtype=object)
    if optional and df.dropna(how="all").empty:
        return None
    seen = {_norm(c): c for c in df.columns}
    df = df.rename(columns={seen[_norm(c)]: c for c in SHEETS[sheet] if _norm(c) in seen})
    missing = [c for c in REQUIRED[sheet] if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing: {', '.join(missing)}.")
    for c in SHEETS[sheet]:
        if c not in df.columns:
            df[c] = ""
    return df[SHEETS[sheet]].dropna(how="all").copy()


def _clean_store(df):
    df["Store Code"] = df["Store Code"].map(_code)
    df["Store Name"] = _clean_text(df["Store Name"])
    df["Section"] = _clean_text(df["Section"])
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["_sk"] = df["Section"].str.upper()
    if "Brand" in df:
        df["Brand"] = _clean_text(df["Brand"])
    if "Article" in df:   # Current Stock and Outward
        df["Item"] = [_item(a, c, z) for a, c, z in zip(df["Article"], df["Color"], df["Size"])]
    return df[df["Store Code"] != ""]


def read_inputs(file):
    """Returns (wh, sale, stock, out, names). sale and out are None when that sheet is missing or empty."""
    xl = pd.ExcelFile(file)
    wh = _read(xl, "WH")
    if wh.empty:
        raise ValueError("The WH sheet is empty.")
    stock = _clean_store(_read(xl, "Current Stock"))
    if stock.empty:
        raise ValueError("The Current Stock sheet is empty.")
    sale, out = (_read(xl, name, optional=True) for name in ("Sale", "Outward"))
    sale, out = (None if df is None else _clean_store(df) for df in (sale, out))
    sale, out = (None if df is not None and df.empty else df for df in (sale, out))
    for c in ("HSN", "Section", "Barcode", "Article", "Color", "Size"):
        wh[c] = wh[c].where(wh[c].notna(), "")
    wh["Section"] = _clean_text(wh["Section"])
    wh["Color"] = _clean_text(wh["Color"])
    wh["Size"] = _clean_text(wh["Size"]).str.upper()
    wh["MRP"] = pd.to_numeric(wh["MRP"], errors="coerce").fillna(0)
    wh["Qty"] = pd.to_numeric(wh["Qty"], errors="coerce").fillna(0)
    for c in ("HSN", "Barcode", "Article"):
        wh[c] = wh[c].map(lambda v: _num(v) if str(v).strip() != "" else "")
    names = {}   # store name: first non-empty one from Current Stock, Sale, Outward
    for df in (stock, sale, out):
        if df is not None:
            for code, nm in zip(df["Store Code"], df["Store Name"]):
                if nm and code not in names:
                    names[code] = nm
    return wh, sale, stock, out, names


def brands_of(stock):
    """Brands found in Current Stock (first spelling of each, sorted)."""
    found = {}
    for b in stock["Brand"]:
        if b and _norm(b) not in found:
            found[_norm(b)] = b
    return [found[k] for k in sorted(found)]


def prepare(inputs, brand, codes_text="", from_date=None):
    """Keeps only what the indent needs: the chosen brand, the chosen stores, the WH sections
    and (for Outward) the dates from `from_date`."""
    wh, sale, stock, out, names = inputs
    key = _norm(brand)
    stock = stock[stock["Brand"].map(_norm) == key]
    if sale is not None:
        sale = sale[sale["Brand"].map(_norm) == key]
    typed = {_code(c) for c in re.split(r"[,;\s]+", codes_text) if c.strip()}
    codes = sorted(typed or set(stock["Store Code"]))
    if not codes:
        raise ValueError(f"No stores found in Current Stock for brand '{brand}'.")
    wh_secs = {s.upper() for s in wh["Section"] if s}

    def keep(df):
        return df[df["Store Code"].isin(codes) & df["_sk"].isin(wh_secs)]

    info = {"out_used": 0, "out_skipped": 0}
    if out is not None:
        if from_date:
            ok = out["Date"].map(_to_date).map(lambda d: d is not None and d >= from_date)
            info["out_skipped"] = int((~ok).sum())
            out = out[ok]
        out = keep(out)
        info["out_used"] = len(out)
    return {"wh": wh, "sale": None if sale is None else keep(sale), "stock": keep(stock), "out": out,
            "names": names, "codes": codes, "info": info}


# =================================================================== PHASE 1: calculation
def _fmt(v):
    return int(v) if float(v) == int(v) else round(float(v), 2)


def compute(p):
    """Returns (wh_rows, stores, sections, sale_pivot, stock_pivot, out_pivot, article_stock, article_out).
    Pivots are dicts: (store, SECTION) -> qty for the summary rows, (store, article|color|size) -> qty
    for the article cells. sale_pivot, out_pivot and article_out are None when that sheet is not used."""
    wh, sale, stock, out = p["wh"], p["sale"], p["stock"], p["out"]
    wh_g = (wh.groupby(["HSN", "Section", "Barcode", "Article", "Color", "Size", "MRP"],
                       dropna=False, sort=False, as_index=False)["Qty"].sum())

    def art_key(a):
        return (0, float(a), "") if isinstance(a, (int, float)) else (1, 0, str(a))

    def size_key(s):
        return SIZE_ORDER.index(s) if s in SIZE_ORDER else len(SIZE_ORDER)

    wh_g["_a"] = wh_g["Article"].map(art_key)
    wh_g["_c"] = wh_g["Color"].str.upper()
    wh_g["_s"] = wh_g["Size"].map(size_key)
    wh_g = wh_g.sort_values(["_a", "_c", "_s", "Size"], kind="stable").drop(columns=["_a", "_c", "_s"])

    secs = {}   # only the sections that are in WH
    for s in wh_g["Section"]:
        if s and s.upper() not in secs:
            secs[s.upper()] = s
    sections = [secs[k] for k in sorted(secs)]

    sp = None if sale is None else sale.groupby(["Store Code", "_sk"])["Qty"].sum().to_dict()
    kp = stock.groupby(["Store Code", "_sk"])["Qty"].sum().to_dict()
    kb = stock[stock["Item"] != ""].groupby(["Store Code", "Item"])["Qty"].sum().to_dict()
    op = ob = None
    if out is not None:
        op = out.groupby(["Store Code", "_sk"])["Qty"].sum().to_dict()
        ob = out[out["Item"] != ""].groupby(["Store Code", "Item"])["Qty"].sum().to_dict()

    stores = pd.DataFrame({"code": p["codes"]})
    stores["name"] = stores["code"].map(lambda c: p["names"].get(c, c))
    tot_stock = stock.groupby("Store Code")["Qty"].sum()
    stores["stock"] = stores["code"].map(lambda c: tot_stock.get(c, 0))
    if sale is not None:
        tot_sale = sale.groupby("Store Code")["Qty"].sum()
        stores["sale"] = stores["code"].map(lambda c: tot_sale.get(c, 0))
    if out is not None:
        tot_out = out.groupby("Store Code")["Qty"].sum()
        stores["out"] = stores["code"].map(lambda c: tot_out.get(c, 0))
    by = (["sale"] if sale is not None else []) + ["stock", "code"]   # highest first; code breaks a tie
    stores = stores.sort_values(by, ascending=[False] * (len(by) - 1) + [True], kind="stable").reset_index(drop=True)
    stores["rank"] = range(1, len(stores) + 1)
    return wh_g, stores, sections, sp, kp, op, kb, ob


# =================================================================== PHASE 1: workbook
def _write_instructions(ws, last_row, head, top, art, rank):
    """Instructions in the empty block A1:H<last_row>. The first line shows the brand, Outward date and store count."""
    lines = [s.format(top=top, art=art, rank=rank) for s in INSTRUCTIONS if art or "{art}" not in s]   # no light red line without article data
    lines = [head] + [f"{i}. {x}" for i, x in enumerate(lines, 1)]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    title = ws.cell(1, 1, "INSTRUCTIONS")
    title.font = Font(name=FONT, size=10, bold=True, color="FFFF0000")
    title.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=last_row, end_column=8)
    body = ws.cell(2, 1, "\n".join(lines))
    body.font = Font(name=FONT, size=9)
    body.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    rows = last_row - 1
    need = 12.5 * sum(-(-len(x) // 100) for x in lines) + 3   # points needed (long lines wrap)
    if 15 * rows < need:                  # few sections: make these rows a little taller
        h = round(need / rows + 0.5)
        for r in range(2, last_row + 1):
            ws.row_dimensions[r].height = h


def _hide_green_triangles(data, sqref):
    """openpyxl cannot write <ignoredErrors>, so it is added to the sheet xml (hides the green triangles)."""
    src, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            blob = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                flags = 'numberStoredAsText="1" formula="1" formulaRange="1" emptyCellReference="1" unlockedFormula="1"'
                tag = f'<ignoredErrors><ignoredError sqref="{sqref}" {flags}/></ignoredErrors>'
                blob = blob.decode().replace("</worksheet>", tag + "</worksheet>").encode()
            dst.writestr(item, blob)
    return out.getvalue()


def build_workbook(wh_g, stores, sections, sp, kp, op, kb, ob, head):
    n = len(sections)
    r_gt = 2
    r_sub = 3 + n
    r_name = 4 + n
    r_hdr = 5 + n
    r_data = 6 + n
    ns = len(stores)
    last_store_col = FIRST_STORE_COL + max(ns, 1) - 1
    last_col = max(last_store_col, 11)
    sum_end = get_column_letter(max(MIN_LAST_COL, last_store_col))

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet2"

    def put(r, c, v=None, size=11, bold=False, fill=None, border=True, color=None, wrap=False, fmt=None, center=True):
        cell = ws.cell(r, c, v)
        cell.font = Font(name=FONT, size=size, bold=bold, color=color)
        if center:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        if fill:
            cell.fill = fill
        if border:
            cell.border = BORDER
        if fmt:
            cell.number_format = fmt
        return cell

    # row 1
    for c in (9, 10, 11):
        put(1, c, ("Sale Wise Store Ranking" if sp is not None else "Stock Wise Store Ranking") if c == 9 else None,
            color="FFFF0000")
    ws.merge_cells(start_row=1, start_column=9, end_row=1, end_column=11)
    for i, row in stores.iterrows():
        put(1, FIRST_STORE_COL + i, int(row["rank"]), color="FFFF0000", center=False)

    # summary rows: sale/soh/out. Sale and out are left out when that sheet is not used (only Grand Total row is yellow)
    def pick(d, key):
        return 0 if d is None else d.get(key, 0)

    def join(s, k, o):
        parts = ([s] if sp is not None else []) + [k] + ([o] if op is not None else [])
        return "/".join(str(_fmt(v)) for v in parts)

    def cell_text(row, skey):
        if skey is None:
            return join(row.get("sale", 0), row["stock"], row.get("out", 0))
        key = (row["code"], skey)
        return join(pick(sp, key), pick(kp, key), pick(op, key))

    def total_text(skey):
        if skey is None:
            return join(stores["sale"].sum() if sp is not None else 0, stores["stock"].sum(),
                        stores["out"].sum() if op is not None else 0)
        total = lambda d: 0 if d is None else sum(v for (c, k), v in d.items() if k == skey)
        return join(total(sp), total(kp), total(op))

    def article_text(code, item):
        """SOH (and Outward) of this store for the article. It is text on purpose, so formulas ignore it."""
        s, o = kb.get((code, item), 0), (ob.get((code, item), 0) if ob is not None else 0)
        if not s and not o:
            return None
        return str(_fmt(s)) if ob is None else f"{_fmt(s)}/{_fmt(o)}"

    labels = ["Grand Total"] + sections
    keys = [None] + [s.upper() for s in sections]
    for k, (lab, key) in enumerate(zip(labels, keys)):
        r = r_gt + k
        fill = YELLOW if k == 0 else None
        put(r, 9, None, size=8)
        put(r, 10, total_text(key), size=8, fill=fill)
        put(r, 11, lab, size=8, fill=fill)
        for i, row in stores.iterrows():
            put(r, FIRST_STORE_COL + i, cell_text(row, key), size=8, fill=fill)

    # subtotal row
    for c in range(9, last_col + 1):
        L = get_column_letter(c)
        put(r_sub, c, f"=SUBTOTAL(9,{L}{r_data}:{L}111365)")

    # store names row
    put(r_name, 1, "WH STOCK", size=36, bold=True, color="FFFF0000")
    for c in range(2, 9):
        put(r_name, c)
    ws.merge_cells(start_row=r_name, start_column=1, end_row=r_name, end_column=8)
    put(r_name, 9, "Calculation", size=26, bold=True, color="FFFF0000")
    put(r_name, 10)
    put(r_name, 11)
    ws.merge_cells(start_row=r_name, start_column=9, end_row=r_name, end_column=11)
    for i, row in stores.iterrows():
        put(r_name, FIRST_STORE_COL + i, row["name"], size=8, bold=True, wrap=True)
    ws.row_dimensions[r_name].height = 56.25

    # header
    hdr = ["HSN", "Section", "Barcode", "Article", "SKU", "Color", "Size", "MRP", " Qty", "Indent", "SHORT"]
    for c, h in enumerate(hdr, 1):
        put(r_hdr, c, h, fill=YELLOW)
    for i, row in stores.iterrows():
        put(r_hdr, FIRST_STORE_COL + i, _store_cell(row["code"]), fill=YELLOW)

    # data
    for k, row in enumerate(wh_g.itertuples(index=False)):
        r = r_data + k
        qty = _fmt(row.Qty)
        vals = [row.HSN, row.Section, row.Barcode, row.Article, f"=D{r}&F{r}", row.Color,
                row.Size, _fmt(row.MRP), qty, f"=SUM(L{r}:{sum_end}{r})", f"=I{r}-J{r}"]
        for c, v in enumerate(vals, 1):
            put(r, c, v if v != "" else None, fmt="0" if c in (3, 5) else None)
        item = _item(row.Article, row.Color, row.Size)
        for i, code in enumerate(stores["code"]):
            txt = article_text(code, item) if item else None
            if txt:
                cell = ws.cell(r, FIRST_STORE_COL + i, txt)
                cell.quotePrefix = True
                cell.font = Font(name=FONT, size=9)
                cell.alignment = CENTER

    widths = {"A": 11.14, "B": 14.14, "C": 17.0, "E": 24.43, "F": 19.0, "G": 7.57, "K": 12.71}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    for c in [10] + list(range(FIRST_STORE_COL, FIRST_STORE_COL + ns)):   # widen only when the text is long
        longest = max(len(str(ws.cell(r, c).value or "")) for r in range(r_gt, r_gt + len(labels)))
        need = round(longest * 0.75 + 1.5, 2)
        if need > 8.43:
            ws.column_dimensions[get_column_letter(c)].width = need

    first, last = get_column_letter(FIRST_STORE_COL), get_column_letter(last_store_col)
    r_last = r_data + len(wh_g) - 1
    ws.conditional_formatting.add(f"{first}{r_data}:{last}{r_last}",
                                  FormulaRule(formula=[f"ISTEXT({first}{r_data})"], font=Font(color=SOFT_RED)))

    parts = ([("Sale", "sale")] if sp is not None else []) + [("SOH", "stock on hand")] \
        + ([("Out", "outward")] if op is not None else [])
    top = "/".join(a for a, _ in parts) + " (" + " / ".join(b for _, b in parts) + ")"
    art = "SOH/Out (stock on hand / outward)" if op is not None else "SOH (stock on hand)"
    _write_instructions(ws, r_sub, head, top, art if (kb or ob) else None, "sale" if sp is not None else "stock")

    bio = io.BytesIO()
    wb.save(bio)
    return _hide_green_triangles(bio.getvalue(), f"A1:{last}{r_last}")


def make_indent(inputs, brand, codes_text="", from_date=None):
    """Builds the indent workbook. Returns (file_name, bytes, stores_df, summary)."""
    p = prepare(inputs, brand, codes_text, from_date)
    wh_g, stores, sections, sp, kp, op, kb, ob = compute(p)
    if op is None:
        outward = "Outward: none"
    else:
        outward = f"Outward from: {from_date:%d-%m-%Y}" if from_date else "Outward from: All"
    head = f"Brand: {brand} | {outward} | Stores: {len(stores)}"
    data = build_workbook(wh_g, stores, sections, sp, kp, op, kb, ob, head)
    summary = {"brand": brand, "skus": len(wh_g), "stores": len(stores), "sections": sections,
               "wh_qty": _fmt(wh_g["Qty"].sum()), "sale": sp is not None, "outward": op is not None, **p["info"]}
    file_name = "TR_INDENT_" + re.sub(r"[^A-Za-z0-9]+", "_", brand).strip("_").upper() + ".xlsx"
    return file_name, data, stores, summary


# =================================================================== PHASE 2: read allocation
def _whole(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v


def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def read_allocation(file):
    """Returns (records, store_summary, totals, warnings)."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb["Sheet2"] if "Sheet2" in wb.sheetnames else wb.worksheets[0]

    hdr = None
    for r in range(1, min(ws.max_row, 60) + 1):
        if _norm(ws.cell(r, 1).value) == "hsn" and _norm(ws.cell(r, 2).value) == "section":
            hdr = r
            break
    if hdr is None:
        raise ValueError("Header row not found. Please upload the sheet made in step 1.")
    name_row = hdr - 1

    stores, seen, warnings = [], set(), []
    for c in range(FIRST_STORE_COL, ws.max_column + 1):
        code = ws.cell(hdr, c).value
        if _blank(code):
            continue
        code = _whole(code)
        if isinstance(code, str):
            code = code.strip()
        if str(code) in seen:
            warnings.append(f"Store code {code} is repeated (column {get_column_letter(c)}).")
        seen.add(str(code))
        nm = ws.cell(name_row, c).value
        stores.append((c, code, "" if nm is None else str(nm).strip()))
    if not stores:
        raise ValueError("No store columns found (from column L).")

    per_store = {c: [] for c, _, _ in stores}
    bad, over = [], []
    for r in range(hdr + 1, ws.max_row + 1):
        row_vals = {}
        total = 0
        for c, code, nm in stores:
            v = ws.cell(r, c).value
            if _blank(v) or isinstance(v, str):   # text = the red soh/out info, not an allocation
                continue
            if v == 0:
                continue
            if isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0 or v != int(v):
                bad.append(ws.cell(r, c).coordinate)
                continue
            row_vals[c] = int(v)
            total += int(v)
        if not row_vals:
            continue

        meta = {
            "hsn": _whole(ws.cell(r, 1).value),
            "section": ws.cell(r, 2).value,
            "ean": _whole(ws.cell(r, 3).value),
            "article": _whole(ws.cell(r, 4).value),
            "color": ws.cell(r, 6).value,
            "size": ws.cell(r, 7).value,
            "mrp": _whole(ws.cell(r, 8).value),
        }
        wh_qty = ws.cell(r, 9).value
        if isinstance(wh_qty, (int, float)) and total > wh_qty:
            over.append(f"Row {r} (Article {meta['article']}, {meta['color']}, {meta['size']}): sent {total}, WH Qty {_whole(wh_qty)}")
        for c, q in row_vals.items():
            per_store[c].append({**meta, "qty": q})

    records, store_rows = [], []
    for c, code, nm in stores:
        recs = per_store[c]
        for rec in recs:
            rec["store_code"] = code
            rec["store_name"] = nm
        records.extend(recs)
        store_rows.append({"Code": code, "Store": nm, "Rows": len(recs), "Qty": sum(x["qty"] for x in recs)})

    if bad:
        warnings.append(f"{len(bad)} cell(s) with a wrong qty were skipped: "
                        + ", ".join(bad[:10]) + (" ..." if len(bad) > 10 else ""))
    if over:
        warnings.append(f"{len(over)} row(s) have more qty than WH Qty:\n" + "\n".join(over[:10])
                        + ("\n..." if len(over) > 10 else ""))
    totals = {"stores": sum(1 for s in store_rows if s["Rows"]), "rows": len(records),
              "qty": sum(x["qty"] for x in records)}
    if not records:
        raise ValueError("No qty found in the store columns.")
    return records, store_rows, totals, warnings


# =================================================================== PHASE 2: outputs
def _resolve(spec, rec, indent_date):
    if "const" in spec:
        v = spec["const"]
    else:
        f = spec["field"]
        if f == "indent_date":
            v = dt.datetime.combine(indent_date, dt.time())
        elif f == "season_year":
            v = indent_date.year
        else:
            v = rec.get(f)
    if _blank(v):
        return None
    kind = spec.get("as")
    if kind == "int":
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return v
    return v


def fill_template(cfg, records, indent_date):
    """Fill a platform template. Returns (file_name, bytes, warnings)."""
    wb = openpyxl.load_workbook(TEMPLATE_DIR / cfg["template_file"])
    ws = wb[cfg["sheet"]] if cfg["sheet"] in wb.sheetnames else wb.worksheets[0]
    hdr_row, first = cfg["header_row"], cfg["first_data_row"]

    heads = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(hdr_row, c).value
        if v is not None:
            heads[_norm(v)] = c
    colmap, warnings = [], []
    for name, spec in cfg["columns"].items():
        c = heads.get(_norm(name))
        if c is None:
            warnings.append(f"{cfg['name']}: column '{name}' was not found in the template and was skipped.")
            continue
        colmap.append((c, spec))

    max_r, max_c = ws.max_row, ws.max_column
    proto = {c: copy(ws.cell(first, c)._style) for c in range(1, max_c + 1)}
    for r in range(first, max_r + 1):          # remove sample rows, keep formatting
        for c in range(1, max_c + 1):
            ws.cell(r, c).value = None

    for i, rec in enumerate(records):
        r = first + i
        if r > max_r:
            for c in range(1, max_c + 1):
                ws.cell(r, c)._style = copy(proto[c])
        for c, spec in colmap:
            v = _resolve(spec, rec, indent_date)
            if v is None:
                continue
            cell = ws.cell(r, c)
            cell.value = v
            if spec.get("field") == "indent_date" and cell.number_format == "General":
                cell.number_format = "d-mmm-yy"

    bio = io.BytesIO()
    wb.save(bio)
    return cfg["output_name"].format(date=indent_date), bio.getvalue(), warnings


def plain_data(cfg, records, indent_date):
    """All allocated rows with every field, no platform template. Returns (file_name, bytes, warnings)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Indent Data"
    heads = list(cfg["columns"])
    ws.append(heads)
    for i, w in enumerate(cfg["widths"], 1):
        c = ws.cell(1, i)
        c.font = Font(name=FONT, bold=True)
        c.fill = YELLOW
        c.alignment = CENTER
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, rec in enumerate(records, 2):
        for i, spec in enumerate(cfg["columns"].values(), 1):
            ws.cell(r, i, _resolve(spec, rec, indent_date))
    for r in range(2, len(records) + 2):
        ws.cell(r, heads.index("Indent Date") + 1).number_format = "d-mmm-yy"
        ws.cell(r, heads.index("EAN") + 1).number_format = "0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{len(records) + 1}"
    bio = io.BytesIO()
    wb.save(bio)
    return cfg["output_name"].format(date=indent_date), bio.getvalue(), []


def make_output(cfg, records, indent_date):
    if "template_file" in cfg:
        return fill_template(cfg, records, indent_date)
    return plain_data(cfg, records, indent_date)


# =================================================================== SCREEN
st.set_page_config(page_title="Indent Tool", layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
    /* ---- layout: use the full window ---- */
    .block-container {max-width: 100% !important; padding: 1.1rem 2.6rem 1.2rem 2.6rem;}
    #MainMenu, footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}

    /* ---- hero banner ---- */
    .hero {background: linear-gradient(115deg, #0B1B4D 0%, #1D4ED8 60%, #3B82F6 100%);
           border-radius: 16px; padding: 20px 28px; margin-bottom: 14px; color: #fff;
           display: flex; align-items: center; justify-content: space-between;
           box-shadow: 0 6px 18px rgba(29,78,216,.22);}
    .hero-title {font-size: 1.75rem; font-weight: 750; letter-spacing: .2px; line-height: 1.15;}
    .hero-sub {font-size: .98rem; opacity: .88; margin-top: 4px;}
    .hero-flow {display: flex; gap: 10px; align-items: center; font-size: .85rem;}
    .pill {background: rgba(255,255,255,.16); border: 1px solid rgba(255,255,255,.28);
           border-radius: 999px; padding: 6px 14px; font-weight: 600; white-space: nowrap;}
    .arrow {opacity: .7;}

    /* ---- tabs ---- */
    .stTabs [data-baseweb="tab-list"] {gap: 6px; border-bottom: 1px solid #DDE3EE;}
    .stTabs [data-baseweb="tab"] {height: 48px; padding: 0 22px; font-weight: 650; font-size: 1rem;
                                  border-radius: 10px 10px 0 0;}
    .stTabs [aria-selected="true"] {background: #fff; color: #1D4ED8;}

    /* ---- panels ---- */
    div[data-testid="stVerticalBlockBorderWrapper"] {background: #fff; border-radius: 14px;
                                                    border-color: #E3E8F1 !important;}
    .panel-title {font-size: 1.05rem; font-weight: 700; color: #0F172A; margin-bottom: 2px;}
    .panel-sub {font-size: .85rem; color: #64748B; margin-bottom: 10px;}
    .step {display: flex; align-items: center; gap: 10px; font-weight: 650; margin: 14px 0 6px 0;
           font-size: .98rem;}
    .step .num {background: #1D4ED8; color: #fff; width: 24px; height: 24px; border-radius: 50%;
                display: inline-flex; align-items: center; justify-content: center; font-size: .8rem;}
    .hint {font-size: .82rem; color: #64748B; margin: 2px 0 8px 34px;}
    .empty {border: 1.5px dashed #CBD5E1; border-radius: 12px; padding: 54px 20px; text-align: center;
            color: #64748B; background: #FAFBFD;}
    .empty b {color: #334155; font-size: 1.02rem;}

    /* ---- metrics ---- */
    div[data-testid="stMetric"] {background: #F6F8FC; border: 1px solid #E3E8F1; border-radius: 12px;
                                 padding: 12px 16px;}
    div[data-testid="stMetricLabel"] p {font-size: .82rem; color: #64748B; font-weight: 600;}
    div[data-testid="stMetricValue"] {font-size: 1.65rem; font-weight: 750; color: #0F172A;}

    /* ---- buttons: full width, rounded ---- */
    .stButton > button, .stDownloadButton > button {width: 100%; border-radius: 10px; font-weight: 650;
                                                    padding: .62rem 1rem;}
    div[data-testid="stFileUploader"] section {border-radius: 12px;}
    </style>
    <div class="hero">
      <div>
        <div class="hero-title">Indent Tool</div>
        <div class="hero-sub">Build the indent workbook, send stock to stores, then fill the platform upload templates.</div>
      </div>
      <div class="hero-flow">
        <span class="pill">WH + Sale + Stock + Outward</span><span class="arrow">&rarr;</span>
        <span class="pill">Indent workbook</span><span class="arrow">&rarr;</span>
        <span class="pill">Fill qty in Excel</span><span class="arrow">&rarr;</span>
        <span class="pill">Platform templates</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


def step(n, text, hint=None):
    st.markdown(f'<div class="step"><span class="num">{n}</span>{text}</div>', unsafe_allow_html=True)
    if hint:
        st.markdown(f'<div class="hint">{hint}</div>', unsafe_allow_html=True)


def empty(title, text):
    st.markdown(f'<div class="empty"><b>{title}</b><br>{text}</div>', unsafe_allow_html=True)


def table(df, height):
    """st.dataframe; a column mixing numbers and text (store code 1234 and TBS1) is shown as text."""
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object and df[c].dropna().map(type).nunique() > 1:
            df[c] = df[c].map(lambda v: v if v is None else str(v))
    st.dataframe(df, hide_index=True, height=height)


tab1, tab2 = st.tabs(["1 · Build Indent Workbook", "2 · Fill Platform Templates"])

# ------------------------------------------------------------------ TAB 1
@st.cache_data(show_spinner="Reading the file...")
def read_cached(data: bytes):
    return read_inputs(io.BytesIO(data))


with tab1:
    left, right = st.columns([5, 8], gap="large")

    with left:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Inputs</div>'
                        '<div class="panel-sub">One Excel file with the sheets WH, Sale, Current Stock '
                        'and Current Stock are required. Sale and Outward are optional. Column names in red are required.</div>',
                        unsafe_allow_html=True)
            step(1, "Download the input template",
                 "WH and Outward have no Brand column, so put only the chosen brand's stock in them. "
                 "Fill Sale and Outward only if you want to use them.")
            st.download_button("Download input template", make_template(),
                               file_name="Indent_Input_Template.xlsx", mime=XLSX, key="dl_tpl")
            step(2, "Upload the filled template")
            up1 = st.file_uploader("Filled input template", type=["xlsx"], key="up1",
                                   label_visibility="collapsed")

            inputs, brand, ready, p = None, None, False, None
            if up1 is not None:
                try:
                    inputs = read_cached(up1.getvalue())
                except Exception as e:
                    st.error(str(e))
            if inputs is not None:
                brands = brands_of(inputs[2])
                if not brands:
                    st.error("No Brand found in the Current Stock sheet.")
                else:
                    step(3, "Choose the brand")
                    brand = st.selectbox("Brand", brands, index=0 if len(brands) == 1 else None,
                                         placeholder="Choose a brand", key="brand", label_visibility="collapsed")
                    if brand:
                        st.caption(f"The WH and Outward sheets must have only {brand} stock.")
                    step(4, "Stores and Outward date (optional)")
                    codes_text = st.text_area("Store codes", key="codes", height=80,
                                              placeholder="Leave empty to use all stores of this brand in Current Stock. "
                                                          "Or type store codes, like 1001, 1002, TBS1")
                    from_date = st.date_input("Outward from date", value=None, format="DD/MM/YYYY", key="from_date",
                                              help="Only Outward on or after this date is used. "
                                                   "Older Outward has already reached the store. Leave empty to use all Outward.")
                    if brand:
                        try:
                            p = prepare(inputs, brand, codes_text, from_date)
                            n_sku = len(inputs[0].drop_duplicates(["HSN", "Section", "Barcode", "Article", "Color", "Size", "MRP"]))
                            n_st = len(p["codes"])
                            line = f"{brand} · {n_st} store{'s' if n_st != 1 else ''} · {n_sku} SKUs"
                            line += " · Sale: used" if p["sale"] is not None else " · Sale: not used"
                            if p["out"] is None:
                                line += " · Outward: not used"
                            else:
                                line += f" · Outward: {p['info']['out_used']} rows used, {p['info']['out_skipped']} skipped"
                            st.info(line)
                            ready = True
                        except ValueError as e:
                            st.error(str(e))
            step(5, "Generate")
            go1 = st.button("Build indent workbook", type="primary", key="go1", disabled=not ready)
            if go1:
                try:
                    with st.spinner("Building workbook..."):
                        st.session_state["r1"] = make_indent(inputs, brand, codes_text, from_date)
                except Exception as e:
                    st.session_state.pop("r1", None)
                    st.error(str(e))

    with right:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Result</div>', unsafe_allow_html=True)
            if "r1" not in st.session_state:
                empty("No workbook yet",
                      "Upload the filled input template, choose the brand and click <i>Build indent workbook</i>.<br>"
                      "The store ranking and the download will appear here.")
            else:
                fname, data, stores, summ = st.session_state["r1"]
                st.markdown(f'<div class="panel-sub">Brand: {summ["brand"]}</div>', unsafe_allow_html=True)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("WH SKUs", f"{summ['skus']:,}")
                m2.metric("WH Qty", f"{summ['wh_qty']:,}")
                m3.metric("Stores", summ["stores"])
                m4.metric("Sections", len(summ["sections"]))
                st.markdown(f'<div class="step" style="margin-top:12px">Store ranking (by {"sale" if summ["sale"] else "stock"}, highest first)</div>',
                            unsafe_allow_html=True)
                cols = ["Rank", "Code", "Store"] + (["Sale"] if summ["sale"] else []) + ["Stock"] + (["Outward"] if summ["outward"] else [])
                view = stores.rename(columns={"rank": "Rank", "code": "Code", "name": "Store",
                                              "sale": "Sale", "stock": "Stock", "out": "Outward"})[cols]
                table(view, min(38 * (len(view) + 1) + 4, 330))
                st.download_button(f"Download {fname}", data, file_name=fname,
                                   mime=XLSX, type="primary", key="dl1")

# ------------------------------------------------------------------ TAB 2
with tab2:
    names = [t["name"] for t in OUTPUTS]
    left, right = st.columns([5, 8], gap="large")

    with left:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Inputs</div>'
                        '<div class="panel-sub">Upload the indent sheet after allocating quantities to the stores.</div>',
                        unsafe_allow_html=True)
            step(1, "Upload the allocated indent sheet",
                 "Type the qty in the store columns of Sheet2 and save the file in Excel. The light red text is ignored.")
            up2 = st.file_uploader("Allocated indent sheet", type=["xlsx"], key="up2",
                                   label_visibility="collapsed")
            step(2, "Indent date and templates",
                 "Plain Data has every field in one simple sheet, to copy into any other format.")
            indent_date = st.date_input("Indent date", value=dt.date.today(), format="DD/MM/YYYY")
            chosen = st.multiselect("Templates to generate", names, default=names)
            step(3, "Generate")
            go2 = st.button("Fill templates", type="primary", key="go2",
                            disabled=(up2 is None or not chosen))
            if go2:
                try:
                    with st.spinner("Filling templates..."):
                        records, store_rows, totals, warns = read_allocation(up2)
                        files, fwarns = [], []
                        for t in OUTPUTS:
                            if t["name"] in chosen:
                                fn, blob, w = make_output(t, records, indent_date)
                                files.append((t["name"], fn, blob))
                                fwarns += w
                    st.session_state["r2"] = (up2.name, records, files, store_rows, totals, warns + fwarns)
                except Exception as e:
                    st.session_state.pop("r2", None)
                    st.error(str(e))

    with right:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Result</div>', unsafe_allow_html=True)
            if "r2" not in st.session_state:
                empty("No templates yet",
                      "Upload the allocated indent sheet and click <i>Fill templates</i>.<br>"
                      "The store-wise summary, a preview and the downloads will appear here.")
            else:
                fname, records, files, store_rows, totals, warns = st.session_state["r2"]
                st.markdown(f'<div class="panel-sub">Generated from {fname}</div>', unsafe_allow_html=True)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Stores allocated", totals["stores"])
                m2.metric("Rows", f"{totals['rows']:,}")
                m3.metric("Total Qty", f"{totals['qty']:,}")
                m4.metric("Templates", len(files))
                for w in warns:
                    st.warning(w)

                sc, pc = st.columns([4, 7], gap="medium")
                with sc:
                    st.markdown('<div class="step" style="margin-top:12px">Store-wise summary</div>',
                                unsafe_allow_html=True)
                    sdf = pd.DataFrame(store_rows)
                    table(sdf, min(38 * (len(sdf) + 1) + 4, 330))
                with pc:
                    st.markdown('<div class="step" style="margin-top:12px">Preview (first 100 rows)</div>',
                                unsafe_allow_html=True)
                    pv = pd.DataFrame(records[:100]).rename(columns={
                        "store_code": "Store Code", "store_name": "Store", "ean": "EAN", "qty": "Qty",
                        "mrp": "MRP", "hsn": "HSN", "section": "Section", "article": "Article",
                        "color": "Color", "size": "Size"})
                    cols = ["Store Code", "Store", "EAN", "Section", "Article", "Color", "Size", "MRP", "Qty"]
                    table(pv[cols], 330)

                st.markdown('<div class="step" style="margin-top:12px">Downloads</div>', unsafe_allow_html=True)
                dcols = st.columns(len(files))
                for i, (name, fn, blob) in enumerate(files):
                    dcols[i].download_button(f"Download {name}", blob, file_name=fn, mime=XLSX,
                                             type="primary", key=f"dl2_{i}")
                    dcols[i].caption(fn)
