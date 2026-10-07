"""
app.py - Proof-Carrying Data Analyst (local, no external AI, no internet).

1. Upload your CSV / Excel files (and optional .txt notes).
2. Check the auto-detected columns (fix them with the dropdowns if needed).
3. Pick one of 12 questions, then press "Get answer".

Every number comes with standalone pandas code. The app runs that code twice in
fresh Python processes and only marks the answer "verified" if both runs agree.
It answers ANSWERED, AMBIGUOUS or CANNOT_DETERMINE, and refuses when the data
cannot support a reliable number (no cost data, mixed currencies with no rate,
missing values, ambiguous dates, conflicting tables).

SETUP:  pip install pandas openpyxl
RUN:    python app.py
"""
import os
import re
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# ---------------------------------------------------------------------------
# Code that is placed at the top of every generated, re-runnable script
# ---------------------------------------------------------------------------
HEADER = r'''import warnings
warnings.filterwarnings("ignore")
import pandas as pd

def load(path, sheet=None, dedup=True):
    try:
        d = pd.read_csv(path, sep=None, engine="python") if sheet is None else pd.read_excel(path, sheet_name=sheet)
    except UnicodeDecodeError:
        d = pd.read_csv(path, sep=None, engine="python", encoding="latin-1")
    d.columns = [str(c).strip() for c in d.columns]
    d = d.loc[:, [c for c in d.columns if not (c.startswith("Unnamed") and d[c].isna().all())]]
    for c in d.columns:
        if not pd.api.types.is_numeric_dtype(d[c]):
            d[c] = d[c].map(lambda v: v.strip() if isinstance(v, str) else v)
    d = d.dropna(how="all")
    return d.drop_duplicates() if dedup else d

def clean(x):
    if pd.api.types.is_numeric_dtype(x):
        return pd.to_numeric(x, errors="coerce")
    return pd.to_numeric(x.astype(str).str.replace(r"[^0-9.\-]", "", regex=True), errors="coerce")

'''
exec(HEADER)  # defines load() and clean() for the app itself (identical to the generated code)

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]
NONE = "(none)"

ROLES = {  # role -> (table group, label)
    "o_id": ("o", "Order ID"), "o_cust": ("o", "Customer ID"), "o_date": ("o", "Order date"),
    "o_amt": ("o", "Amount"), "o_cur": ("o", "Currency"), "o_ship": ("o", "Ship country"),
    "c_id": ("c", "Customer ID"), "c_country": ("c", "Country"),
    "r_id": ("r", "Refund ID"), "r_order": ("r", "Order ID"), "r_amt": ("r", "Refund amount"),
}
GROUPS = {"o": "Orders / sales table", "c": "Customers table", "r": "Refunds table"}


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
class Data:
    def __init__(self):
        self.tables = {}   # name -> dict(df, raw, path, sheet)
        self.docs = {}     # name -> text

    def add(self, path):
        stem, ext = os.path.splitext(os.path.basename(path))
        ext = ext.lower()
        if ext in (".txt", ".md"):
            with open(path, encoding="utf-8", errors="replace") as f:
                self.docs[os.path.basename(path)] = f.read()
        elif ext == ".csv":
            self._put(stem, path, None)
        elif ext in (".xlsx", ".xlsm", ".xls"):
            for sh in pd.ExcelFile(path).sheet_names:
                self._put(f"{stem}[{sh}]", path, sh)
        else:
            raise ValueError("Unsupported file type: " + ext)

    def _put(self, name, path, sheet):
        raw = load(path, sheet, dedup=False)
        self.tables[name] = dict(name=name, df=raw.drop_duplicates(), raw=len(raw), path=path, sheet=sheet)

    def clear(self):
        self.tables.clear()
        self.docs.clear()


# ---------------------------------------------------------------------------
# Auto-detection of columns
# ---------------------------------------------------------------------------
def find(cols, *pats, exclude=None):
    for p in pats:
        for c in cols:
            if re.search(p, c, re.I) and not (exclude and re.search(exclude, c, re.I)):
                return c
    return ""


PATS = {
    "o_id": [r"^order[_ ]?id$", r"order.*(id|no|num)", r"^id$"],
    "o_cust": [r"(cust|client).*(id|no|num)", r"customer", r"client", r"buyer"],
    "o_date": [r"date", r"time"],
    "o_amt": [r"^amount$", r"amount", r"total", r"price", r"revenue", r"sales", r"value"],
    "o_cur": [r"curr", r"^cur"],
    "o_ship": [r"ship.*country", r"deliver", r"dest", r"country"],
    "c_id": [r"^customer[_ ]?id$", r"(cust|client).*(id|no|num)", r"^id$", r"customer", r"client"],
    "c_country": [r"^country$", r"country"],
    "r_id": [r"refund.*id", r"^id$"],
    "r_order": [r"order.*id", r"order"],
    "r_amt": [r"refund.*amount", r"amount", r"refund"],
}


def guess(data):
    M = {k: "" for k in ROLES}
    M.update(o_t="", c_t="", r_t="")
    names = list(data.tables)
    refund_names = [n for n in names if re.search("refund", n, re.I)
                    or any(re.search("refund", c, re.I) for c in data.tables[n]["df"].columns)]
    M["r_t"] = refund_names[0] if refund_names else ""

    def score(n):
        cols = list(data.tables[n]["df"].columns)
        return sum(bool(find(cols, *PATS[k], exclude="refund")) for k in ("o_id", "o_date", "o_amt", "o_cur", "o_ship"))

    cand = [n for n in names if n != M["r_t"]]
    if cand:
        M["o_t"] = max(cand, key=score)
    cust = [n for n in cand if n != M["o_t"]]
    if cust:
        M["c_t"] = max(cust, key=lambda n: bool(find(list(data.tables[n]["df"].columns), *PATS["c_id"]))
                       + bool(find(list(data.tables[n]["df"].columns), *PATS["c_country"])))
        cols = list(data.tables[M["c_t"]]["df"].columns)
        if not (find(cols, *PATS["c_id"]) and find(cols, *PATS["c_country"])):
            M["c_t"] = ""
    for k, (g, _) in ROLES.items():
        t = M[g + "_t"]
        if t:
            cols = list(data.tables[t]["df"].columns)
            ex = "refund" if k.startswith("o_") else None
            M[k] = find(cols, *PATS[k], exclude=ex)
    if M["r_t"] and not M["r_id"]:
        M["r_id"] = data.tables[M["r_t"]]["df"].columns[0]
    return M


def concepts(data):
    cost, rate = [], []
    for n, t in data.tables.items():
        for c in t["df"].columns:
            if re.search(r"cost|cogs|expense|margin|profit", c, re.I):
                cost.append(f"{n}.{c}")
            if re.search(r"(^|[^a-z])(fx|rate)([^a-z]|$)|exchange", c, re.I):
                rate.append(f"{n}.{c}")
    for n, text in data.docs.items():
        if re.search(r"cost of goods|cogs|unit cost|cost price|profit margin", text, re.I):
            cost.append(f"note {n}")
        if re.search(r"exchange rate|fx rate|conversion rate|1\s*(usd|eur|gbp|inr)\s*=", text, re.I):
            rate.append(f"note {n}")
    return cost, rate


# ---------------------------------------------------------------------------
# Running / verifying code
# ---------------------------------------------------------------------------
class Fail(Exception):
    pass


def run(code):
    p = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=90)
    return p.returncode, p.stdout.strip(), p.stderr.strip()


def prove(code):
    """Run the code twice in fresh processes. Returns (stdout, last_number, verified)."""
    rc, out, err = run(code)
    if rc != 0 or not out:
        raise Fail("generated code crashed: " + (err[-400:] or "no output"))
    rc2, out2, _ = run(code)
    nums = re.findall(r"-?\d+(?:\.\d+)?", out.splitlines()[-1].replace(",", ""))
    val = float(nums[-1]) if nums else None
    return out, val, (rc2 == 0 and out2 == out)


def T(t, var):
    return f"{var} = load({t['path']!r}, {t['sheet']!r})\n"


def res(status, answer, why, code=None, value=None, evidence=(), missing=(), assumptions=(), verified=None):
    return dict(status=status, answer=answer, reasoning=why, code=code, value=value,
                evidence=list(evidence), missing_or_ambiguous=list(missing),
                assumptions=list(assumptions), verified=verified)


# ---------------------------------------------------------------------------
# Helpers for the question functions
# ---------------------------------------------------------------------------
def tbl(d, M, g):
    return d.tables.get(M.get(g + "_t", ""))


def need(d, M, *roles):
    """Return a refusal result if a needed table/column is not mapped, else None."""
    lacking = []
    for r in roles:
        g, label = ROLES[r]
        t = tbl(d, M, g)
        if t is None:
            lacking.append(f"{GROUPS[g]} is missing (not uploaded or not selected)")
        elif not M.get(r) or M[r] not in t["df"].columns:
            lacking.append(f"{label} column is missing in {t['name']}")
    if lacking:
        return res("CANNOT_DETERMINE", "This question cannot be answered: " + "; ".join(lacking) + ".",
                   "The question needs data that is not available. " + " ".join(lacking),
                   missing=lacking)
    return None


def dup_note(t):
    n = t["raw"] - len(t["df"])
    return f"{t['name']}: {t['raw']} raw rows, {n} exact duplicate rows removed, {len(t['df'])} rows used."


def conflicts(df, idc):
    dd = df[df.duplicated(idc, keep=False) & df[idc].notna()]
    return sorted(dd[idc].astype(str).unique())


def order_issues(d, M, cur=None):
    """Data-quality issues for order amounts (optionally within one currency)."""
    o = tbl(d, M, "o")["df"]
    idc, amt = M["o_id"], M["o_amt"]
    cf = conflicts(o, idc)
    one = o.drop_duplicates(idc)
    a = clean(one[amt])
    scope = pd.Series(True, index=one.index)
    miss_cur = []
    if M["o_cur"]:
        cu = one[M["o_cur"]].astype(str).str.strip().str.upper()
        miss_cur = one.loc[one[M["o_cur"]].isna(), idc].astype(str).tolist()
        if cur:
            scope = cu == cur
    miss_amt = one.loc[scope & a.isna(), idc].astype(str).tolist()
    return cf, miss_amt, miss_cur, one, a


def judge_amounts(status_ok_text, cf, miss_amt, miss_cur, base_missing=()):
    """Turn quality issues into (status, missing_list)."""
    missing = list(base_missing)
    status = "ANSWERED"
    if miss_amt:
        status = "CANNOT_DETERMINE"
        missing.append(f"missing amount for orders {miss_amt[:10]} - cannot treat missing values as zero")
    if cf:
        if status == "ANSWERED":
            status = "AMBIGUOUS"
        missing.append(f"same order id with different content for {cf[:10]} - sources conflict")
    if miss_cur:
        if status == "ANSWERED":
            status = "AMBIGUOUS"
        missing.append(f"missing currency for orders {miss_cur[:10]}")
    return status, missing


# ---------------------------------------------------------------------------
# The 12 questions
# ---------------------------------------------------------------------------
def q_unique_orders(d, M, a, b):
    r = need(d, M, "o_id")
    if r:
        return r
    o = tbl(d, M, "o")
    code = HEADER + T(o, "s") + f"print(s[{M['o_id']!r}].nunique())\n"
    out, val, ver = prove(code)
    n_missing = int(o["df"][M["o_id"]].isna().sum())
    cf = conflicts(o["df"], M["o_id"])
    missing = []
    status = "ANSWERED"
    if n_missing:
        status = "AMBIGUOUS"
        missing.append(f"{n_missing} rows have a missing order id and cannot be counted")
    why = "Counted distinct order ids after removing exact duplicate rows (the export can contain repeated rows)."
    if cf:
        why += f" Ids {cf[:10]} appear with different content; each id is still one order."
    return res(status, f"There are {val:.0f} unique orders.", why, code, val,
               [dup_note(o), "code output: " + out.splitlines()[-1]], missing,
               ["one order = one distinct order id"], ver)


def _sum_or_avg(d, M, cur, avg):
    r = need(d, M, "o_id", "o_amt", "o_cur")
    if r:
        return r
    cur = (cur or "").strip().upper()
    o = tbl(d, M, "o")
    cf, miss_amt, miss_cur, one, a = order_issues(d, M, cur)
    cu = one[M["o_cur"]].astype(str).str.strip().str.upper()
    n = int((cu == cur).sum())
    if n == 0:
        return res("CANNOT_DETERMINE", f"No orders found in currency {cur}.", f"No row has currency {cur}.",
                   missing=[f"no orders in {cur}"])
    agg = "mean()" if avg else "sum()"
    code = (HEADER + T(o, "s") + f"s = s.drop_duplicates({M['o_id']!r})\n"
            f"amt = clean(s[{M['o_amt']!r}])\n"
            f"m = s[{M['o_cur']!r}].astype(str).str.strip().str.upper() == {cur!r}\n"
            f"print(round(amt[m].{agg}, 2))\n")
    out, val, ver = prove(code)
    status, missing = judge_amounts("", cf, miss_amt, miss_cur)
    word = "average" if avg else "total"
    ans = f"The {word} order amount in {cur} is {val:,.2f} (from {n} orders)."
    if status != "ANSWERED":
        ans = f"Not reliable: the {word} in {cur} over the usable rows is {val:,.2f}, but there are data problems."
    return res(status, ans,
               f"Filtered currency == {cur}, kept one row per order id, took the {word} of the amount column "
               "(gross order value; refunds are not deducted).",
               code, val if status == "ANSWERED" else None,
               [dup_note(o), f"{n} distinct orders in {cur}", "code output: " + out.splitlines()[-1]], missing,
               ["exact duplicate rows removed, then one row per order id", "amounts are gross; refunds not deducted"],
               ver)


def q_sum(d, M, cur, b):
    return _sum_or_avg(d, M, cur, False)


def q_avg(d, M, cur, b):
    return _sum_or_avg(d, M, cur, True)


def q_top_customer(d, M, a, b):
    r = need(d, M, "o_id", "o_cust")
    if r:
        return r
    o = tbl(d, M, "o")
    code = (HEADER + T(o, "s") + f"t = s.groupby({M['o_cust']!r})[{M['o_id']!r}].nunique().sort_values(ascending=False)\n"
            "print(t.index[0], int(t.iloc[0]))\n")
    out, val, ver = prove(code)
    one = o["df"].drop_duplicates(M["o_id"])
    t = one.groupby(M["o_cust"])[M["o_id"]].nunique().sort_values(ascending=False)
    top = t[t == t.iloc[0]].index.astype(str).tolist()
    miss_c = int(one[M["o_cust"]].isna().sum())
    missing = []
    status = "ANSWERED"
    if len(top) > 1:
        status = "AMBIGUOUS"
        missing.append(f"tie: customers {top} all have {int(t.iloc[0])} orders")
    if miss_c:
        status = "AMBIGUOUS"
        missing.append(f"{miss_c} orders have a missing customer id and cannot be attributed")
    ans = f"Customer {top[0]} has the most orders ({int(t.iloc[0])})." if len(top) == 1 else \
        f"Tie: customers {', '.join(top)} each have {int(t.iloc[0])} orders."
    return res(status, ans, "Counted distinct order ids per customer after removing duplicate rows.",
               code, float(t.iloc[0]), [dup_note(o), "code output: " + out.splitlines()[-1]], missing,
               ["duplicates removed first"], ver)


def q_refund_count(d, M, a, b):
    r = need(d, M, "r_id")
    if r:
        return r
    t = tbl(d, M, "r")
    code = HEADER + T(t, "r") + f"print(r[{M['r_id']!r}].nunique())\n"
    out, val, ver = prove(code)
    return res("ANSWERED", f"There are {val:.0f} refunds in the data.",
               "Counted distinct refund ids after removing exact duplicate rows.", code, val,
               [dup_note(t), "code output: " + out.splitlines()[-1]], [], [], ver)


def q_ship_country(d, M, country, b):
    r = need(d, M, "o_id", "o_ship")
    if r:
        return r
    country = (country or "").strip()
    o = tbl(d, M, "o")
    code = (HEADER + T(o, "s") + f"s = s.drop_duplicates({M['o_id']!r})\n"
            f"m = s[{M['o_ship']!r}].astype(str).str.strip().str.lower() == {country.lower()!r}\n"
            f"print(s[m][{M['o_id']!r}].nunique())\n")
    out, val, ver = prove(code)
    one = o["df"].drop_duplicates(M["o_id"])
    miss = one.loc[one[M["o_ship"]].isna(), M["o_id"]].astype(str).tolist()
    missing, status = [], "ANSWERED"
    if miss:
        status = "AMBIGUOUS"
        missing.append(f"orders {miss[:10]} have a missing ship country - they might also have gone to {country}")
    return res(status, f"{val:.0f} orders were shipped to {country}." if not miss else
               f"At least {val:.0f} orders were shipped to {country}; {len(miss)} orders have no ship country.",
               f"Counted distinct orders whose ship country is {country} (case/space-insensitive).",
               code, val, [dup_note(o), "code output: " + out.splitlines()[-1]], missing,
               ["'shipped to' means the order's ship country column, not the customer's account country"], ver)


MONTH_CODE = r'''
import re
s = s.drop_duplicates(ID)
sure = maybe = unread = miss = 0
amb_ids = []
for oid, v in zip(s[ID], s[DATE]):
    if pd.isna(v):
        miss += 1
        continue
    v = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", v)
    if m:
        opts = {(int(m[1]), int(m[2]))}
    else:
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", v)
        if m:
            a, b, y = int(m[1]), int(m[2]), int(m[3])
            opts = {(y, b)} if a > 12 else ({(y, a)} if b > 12 else {(y, a), (y, b)})
        else:
            t = pd.to_datetime(v, errors="coerce")
            opts = {(t.year, t.month)} if pd.notna(t) else set()
    if not opts:
        unread += 1
    elif (YEAR, MONTH) in opts:
        if len(opts) == 1:
            sure += 1
        else:
            maybe += 1
            amb_ids.append(str(oid))
print("maybe=%d unreadable=%d missing=%d ambiguous_ids=%s" % (maybe, unread, miss, amb_ids))
print("certain orders:", sure)
'''


def q_month(d, M, month, year):
    r = need(d, M, "o_id", "o_date")
    if r:
        return r
    try:
        mo = MONTHS.index(month) + 1
        yr = int(str(year).strip())
    except Exception:
        return res("CANNOT_DETERMINE", "Please choose a month and type a 4-digit year.", "Invalid month/year input.",
                   missing=["valid month and year"])
    o = tbl(d, M, "o")
    code = (HEADER + T(o, "s") + f"ID = {M['o_id']!r}\nDATE = {M['o_date']!r}\nYEAR = {yr}\nMONTH = {mo}\n" + MONTH_CODE)
    out, val, ver = prove(code)
    st = re.search(r"maybe=(\d+) unreadable=(\d+) missing=(\d+) ambiguous_ids=(.*)", out)
    maybe, unread, miss = int(st[1]), int(st[2]), int(st[3])
    sure = int(val)
    ev = [dup_note(o)] + out.splitlines()
    assum = ["slash dates like 03/04/2025 can be DD/MM or MM/DD; ISO dates (YYYY-MM-DD) are unambiguous"]
    if not (maybe or unread or miss):
        return res("ANSWERED", f"{sure} orders were placed in {month} {yr}.",
                   "All order dates are unambiguous, counted orders dated in that month.", code, val, ev, [], assum, ver)
    missing = []
    if maybe:
        missing.append(f"{maybe} orders {st[4]} have slash dates that fall in {month} {yr} under one reading only "
                       "(DD/MM vs MM/DD)")
    if unread:
        missing.append(f"{unread} orders have unreadable dates")
    if miss:
        missing.append(f"{miss} orders have a missing date")
    hi = sure + maybe + unread + miss
    return res("AMBIGUOUS", f"Between {sure} and {hi} orders, depending on how the ambiguous/missing dates are read.",
               f"{sure} orders are certainly in {month} {yr}. The count changes with the date-format reading "
               "of the ambiguous dates, so one single number would be a guess.",
               code, val, ev, missing, assum, ver)


def q_profit(d, M, a, b):
    cost, rate = concepts(d)
    if not cost:
        return res("CANNOT_DETERMINE", "Profit cannot be computed: the data contains no cost information.",
                   "Profit = revenue - cost. No column in any table (or note) mentions cost, COGS, expense or margin, "
                   "so any profit number would be invented.",
                   evidence=["concept check: cost / expense / margin NOT FOUND in any table or note"],
                   missing=["cost data (cost, COGS, expenses) is missing from every table"])
    return res("CANNOT_DETERMINE", "Profit cannot be computed automatically.",
               f"Possible cost fields were found ({cost}), but this tool cannot verify how they link to orders, "
               "their currency, or their period, so it will not guess.",
               evidence=[f"possible cost fields: {cost}"],
               missing=["cost link to orders, currency and period must be confirmed manually"])


def q_total_revenue(d, M, a, b):
    r = need(d, M, "o_id", "o_amt")
    if r:
        return r
    o = tbl(d, M, "o")
    cf, miss_amt, miss_cur, one, am = order_issues(d, M)
    cost, rate = concepts(d)
    if M["o_cur"]:
        cu = one[M["o_cur"]].astype(str).str.strip().str.upper().where(one[M["o_cur"]].notna())
        curs = sorted(cu.dropna().unique())
        if len(curs) > 1:
            parts = {c: round(float(am[cu == c].sum()), 2) for c in curs}
            why = (f"Sales are in several currencies {curs}. Adding them as one number mixes units. "
                   + ("Exchange-rate fields exist (" + ", ".join(rate) + ") but this tool does not apply them."
                      if rate else "No exchange rate exists in any table or note, and outside rates are not used."))
            return res("AMBIGUOUS", "No single total can be given: orders are in " + " and ".join(curs) + " and cannot be added without an exchange rate.",
                       why, evidence=[dup_note(o)] + [f"known total in {c}: {v:,.2f}" for c, v in parts.items()],
                       missing=["exchange rate between " + "/".join(curs) + " is missing from the data"
                                if not rate else "exchange rate must be applied manually"])
    code = (HEADER + T(o, "s") + f"s = s.drop_duplicates({M['o_id']!r})\n"
            f"print(round(clean(s[{M['o_amt']!r}]).sum(), 2))\n")
    out, val, ver = prove(code)
    status, missing = judge_amounts("", cf, miss_amt, miss_cur)
    unit = ""
    assum = ["amounts are gross order value; refunds are not deducted", "duplicates removed"]
    if not M["o_cur"]:
        assum.append("no currency column exists, so all amounts are assumed to be in the same unit")
    elif miss_cur == [] and len(curs) == 1:
        unit = " " + curs[0]
    ans = f"Total revenue is {val:,.2f}{unit}." if status == "ANSWERED" else \
        f"Not reliable: the sum of known amounts is {val:,.2f}{unit}, but there are data problems."
    return res(status, ans, "Summed the amount column, one row per order id.", code,
               val if status == "ANSWERED" else None, [dup_note(o), "code output: " + out.splitlines()[-1]],
               missing, assum, ver)


def q_total_refunded(d, M, a, b):
    r = need(d, M, "r_amt")
    if r:
        return r
    rt = tbl(d, M, "r")
    rdf = rt["df"]
    amt = clean(rdf[M["r_amt"]])
    idc = M["r_id"] if M["r_id"] in rdf.columns else rdf.columns[0]
    miss_ids = rdf.loc[amt.isna(), idc].astype(str).tolist()
    cur_of = pd.Series(dtype=object)
    cur_set, unknown = set(), []
    ot = tbl(d, M, "o")
    if ot is not None and M["o_id"] and M["o_cur"] and M["r_order"] and M["r_order"] in rdf.columns:
        od = ot["df"].drop_duplicates(M["o_id"]).set_index(M["o_id"])[M["o_cur"]].astype(str).str.strip().str.upper()
        cur_of = rdf[M["r_order"]].map(od)
        known = cur_of[amt.notna()]
        cur_set = set(known.dropna())
        unknown = rdf.loc[amt.notna() & cur_of.isna(), idc].astype(str).tolist()
    _, rate = concepts(d)
    code = HEADER + T(rt, "r") + f"print(round(clean(r[{M['r_amt']!r}]).sum(), 2))\n"
    out, val, ver = prove(code)
    ev = [dup_note(rt), "code output (sum of known amounts): " + out.splitlines()[-1]]
    if len(cur_of):
        by = {c: round(float(amt[(cur_of == c) & amt.notna()].sum()), 2) for c in sorted(cur_set)}
        ev += [f"known refunds in {c}: {v:,.2f}" for c, v in by.items()]
    missing = []
    status = "ANSWERED"
    if len(cur_set) > 1:
        status = "AMBIGUOUS"
        missing.append("refunds are in " + " and ".join(sorted(cur_set)) + " (currency of the refunded order); "
                       + ("exchange rate fields exist but are not applied" if rate else "no exchange rate in the data"))
    if miss_ids:
        if status == "ANSWERED":
            status = "CANNOT_DETERMINE"
        missing.append(f"missing refund amount for {miss_ids[:10]} - not treated as zero")
    if unknown:
        if status == "ANSWERED":
            status = "AMBIGUOUS"
        missing.append(f"refunds {unknown[:10]} point to orders with unknown currency")
    if status == "ANSWERED":
        unit = (" " + next(iter(cur_set))) if len(cur_set) == 1 else ""
        return res(status, f"The total refunded is {val:,.2f}{unit}.", "Summed refund_amount over distinct refund rows.",
                   code, val, ev, [], ["refund amounts are in the currency of the refunded order"], ver)
    return res(status, "A reliable total refunded cannot be given. " + " ".join(missing),
               "The sum of known amounts is not trustworthy because of: " + "; ".join(missing),
               evidence=ev, missing=missing)


def q_customer_country(d, M, cid, b):
    r = need(d, M, "c_id", "c_country")
    if r:
        return r
    cid = (cid or "").strip()
    ct = tbl(d, M, "c")
    c = ct["df"]
    row = c[c[M["c_id"]].astype(str).str.strip() == cid]
    if row.empty:
        return res("CANNOT_DETERMINE", f"Customer {cid} was not found.", "No such id in the customers table.",
                   missing=[f"customer {cid} missing from customers table"])
    code = (HEADER + T(ct, "c") + f"x = c[c[{M['c_id']!r}].astype(str).str.strip() == {cid!r}]\n"
            f"print(sorted(set(x[{M['c_country']!r}].dropna().astype(str))))\n")
    out, _, ver = prove(code)
    homes = sorted(set(row[M["c_country"]].dropna().astype(str)))
    ev = [dup_note(ct), "code output: " + out.splitlines()[-1]]
    if not homes:
        return res("CANNOT_DETERMINE", f"The country of {cid} is missing in the customers table.",
                   "The country value is missing for this customer.", evidence=ev, missing=[f"country missing for {cid}"])
    if len(homes) > 1:
        return res("AMBIGUOUS", f"The customers table lists several countries for {cid}: {homes}.",
                   "The same customer has conflicting countries.", code, None, ev,
                   [f"conflicting countries {homes}"], [], ver)
    home = homes[0]
    ships = []
    ot = tbl(d, M, "o")
    if ot is not None and M["o_cust"] and M["o_ship"]:
        o = ot["df"]
        ships = sorted(set(o.loc[o[M["o_cust"]].astype(str).str.strip() == cid, M["o_ship"]].dropna().astype(str)))
        ev.append(f"{ot['name']} ship countries for {cid}: {ships}")
    if ships and {s.lower() for s in ships} != {home.lower()}:
        return res("AMBIGUOUS", f"Unclear: the customers table says {cid} is in {home}, but their orders shipped to {', '.join(ships)}.",
                   f"Two sources disagree. The account country is {home}; the orders' delivery country is {ships}. "
                   "They measure different things, so neither can be called 'the' country without a definition.",
                   evidence=ev, missing=[f"{home} (account country) vs {ships} (ship country) - need to know which is meant"])
    return res("ANSWERED", f"Customer {cid} is in {home}.", "The customers table and the order shipments agree.",
               code, None, ev, [], [], ver)


def q_avg_converted(d, M, target, b):
    r = need(d, M, "o_id", "o_amt", "o_cur")
    if r:
        return r
    target = (target or "").strip().upper()
    cf, miss_amt, miss_cur, one, am = order_issues(d, M)
    cu = one[M["o_cur"]].astype(str).str.strip().str.upper().where(one[M["o_cur"]].notna())
    curs = sorted(cu.dropna().unique())
    others = [c for c in curs if c != target]
    cost, rate = concepts(d)
    if others:
        why = (f"Orders exist in {others} besides {target}. Converting needs an exchange rate. "
               + ("Rate fields exist (" + ", ".join(rate) + ") but this tool does not apply them automatically."
                  if rate else "No exchange rate exists in any table or note, and outside rates are not used."))
        return res("CANNOT_DETERMINE" if not rate else "AMBIGUOUS",
                   f"Cannot express the average in {target}: orders in {others} need an exchange rate that is not in the data.",
                   why, evidence=[f"currencies in data: {curs}"],
                   missing=[f"exchange rate {'/'.join(others)} to {target} is missing from the data"])
    return _sum_or_avg(d, M, target, True)


QUESTIONS = [
    ("1. How many unique orders are there?", None, q_unique_orders),
    ("2. What is the total sales in a chosen currency?", "currency", q_sum),
    ("3. What is the average order amount in a chosen currency?", "currency", q_avg),
    ("4. Which customer has the most orders?", None, q_top_customer),
    ("5. How many refunds are in the data?", None, q_refund_count),
    ("6. How many orders were shipped to a chosen country?", "country", q_ship_country),
    ("7. How many orders were placed in a chosen month and year?", "month", q_month),
    ("8. What was the profit?", None, q_profit),
    ("9. What is the total revenue in one number?", None, q_total_revenue),
    ("10. What is the total amount refunded?", None, q_total_refunded),
    ("11. What country is a chosen customer in?", "customer", q_customer_country),
    ("12. What is the average order amount expressed in a chosen currency (converted)?", "currency", q_avg_converted),
]


def solve(qi, d, M, p1, p2):
    try:
        return QUESTIONS[qi][2](d, M, p1, p2)
    except Fail as e:
        return res("CANNOT_DETERMINE", "The analysis could not run.", str(e), missing=[str(e)])
    except Exception as e:  # never crash the interface
        return res("CANNOT_DETERMINE", "The analysis could not run on this data.",
                   f"Unexpected problem: {e!r}. Check the column selections.", missing=[f"error: {e!r}"])


def format_result(r):
    lines = [f"STATUS: {r['status']}   |   verified by re-running the code: "
             + {True: "YES", False: "NO (runs differed)", None: "n/a (no number to verify)"}[r["verified"]],
             "", "ANSWER: " + r["answer"], "", "WHY: " + r["reasoning"]]
    for title, key in (("EVIDENCE", "evidence"), ("MISSING / AMBIGUOUS", "missing_or_ambiguous"),
                       ("ASSUMPTIONS", "assumptions")):
        if r[key]:
            lines += ["", title + ":"] + ["  - " + x for x in r[key]]
    if r["code"]:
        lines += ["", "CODE (save as check.py and run it - it prints the number on the last line):", "", r["code"]]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
BG, CARD, ACCENT, ACCENT_DARK = "#f4f6fa", "#ffffff", "#2f6fed", "#2458c8"
TEXT, MUTED, BORDER, CODE_BG = "#1f2937", "#6b7280", "#dfe3ea", "#f1f3f8"
STATUS_COLORS = {"ANSWERED": "#1f9d55", "AMBIGUOUS": "#d97706", "CANNOT_DETERMINE": "#dc2626"}
FONT = "Segoe UI"


def setup_style(root):
    root.configure(bg=BG)
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    st.configure(".", font=(FONT, 10), background=BG, foreground=TEXT)
    st.configure("TFrame", background=BG)
    st.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1, bordercolor=BORDER)
    st.configure("Plain.TFrame", background=CARD)
    st.configure("Card.TLabel", background=CARD, foreground=TEXT)
    st.configure("Muted.TLabel", background=CARD, foreground=MUTED)
    st.configure("Head.TLabel", background=CARD, foreground=TEXT, font=(FONT, 11, "bold"))
    st.configure("Sub.TLabel", background=CARD, foreground=ACCENT, font=(FONT, 9, "bold"))
    st.configure("TButton", background="#e9edf5", foreground=TEXT, borderwidth=0, padding=(12, 6))
    st.map("TButton", background=[("active", "#dce2ee"), ("pressed", "#dce2ee")])
    st.configure("Accent.TButton", background=ACCENT, foreground="white", font=(FONT, 10, "bold"),
                 borderwidth=0, padding=(18, 7))
    st.map("Accent.TButton", background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK)])
    st.configure("TCombobox", fieldbackground="white", background="#e9edf5", bordercolor=BORDER,
                 arrowcolor=TEXT, padding=3)
    st.map("TCombobox", fieldbackground=[("readonly", "white")], bordercolor=[("focus", ACCENT)])
    st.configure("TEntry", fieldbackground="white", bordercolor=BORDER, padding=3)
    st.configure("Vertical.TScrollbar", background="#e9edf5", troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)


def card(parent, step, title):
    """White bordered card with a numbered header that folds away when clicked.
    Returns (outer, body, set_open)."""
    outer = ttk.Frame(parent, style="Card.TFrame", padding=(14, 10))
    head = ttk.Frame(outer, style="Plain.TFrame")
    head.pack(fill="x")
    badge = tk.Label(head, text=str(step), bg=ACCENT, fg="white", font=(FONT, 9, "bold"), width=2)
    badge.pack(side="left")
    name = ttk.Label(head, text="  " + title, style="Head.TLabel")
    name.pack(side="left")
    hint = ttk.Label(head, text="", style="Muted.TLabel")
    hint.pack(side="left")
    arrow = tk.Label(head, text="\u25BE", bg=CARD, fg=MUTED, font=(FONT, 12))
    arrow.pack(side="right")
    body = ttk.Frame(outer, style="Plain.TFrame")
    state = {"open": True}

    def set_open(is_open):
        state["open"] = is_open
        if is_open:
            body.pack(fill="both", expand=True, pady=(8, 0))
            arrow.config(text="\u25BE")
            hint.config(text="")
        else:
            body.pack_forget()
            arrow.config(text="\u25B8")
            hint.config(text="   (click to expand)")

    for w in (head, badge, name, hint, arrow):
        w.bind("<Button-1>", lambda e: set_open(not state["open"]))
        w.configure(cursor="hand2")
    set_open(True)
    return outer, body, set_open


class App:
    def __init__(self, root):
        self.root = root
        root.title("Proof-Carrying Data Analyst")
        root.geometry("1100x900")
        root.minsize(900, 700)
        setup_style(root)
        root.option_add("*TCombobox*Listbox.font", (FONT, 10))
        self.data = Data()
        self.history = []
        self.last_code = ""
        self.vars = {}   # role/table -> StringVar
        self.boxes = {}  # role/table -> Combobox
        self._build()

    def _build(self):
        r = self.root
        r.columnconfigure(0, weight=1)
        r.rowconfigure(2, weight=1)

        # header bar
        bar = tk.Frame(r, bg=ACCENT)
        bar.grid(row=0, column=0, sticky="ew")
        tk.Label(bar, text="Proof-Carrying Data Analyst", bg=ACCENT, fg="white",
                 font=(FONT, 16, "bold")).pack(anchor="w", padx=18, pady=(10, 0))
        tk.Label(bar, text="Every number comes with code you can re-run. Unreliable questions are refused, not guessed.",
                 bg=ACCENT, fg="#dbe6ff", font=(FONT, 9)).pack(anchor="w", padx=18, pady=(0, 10))

        top = ttk.Frame(r)
        top.grid(row=1, column=0, sticky="ew", padx=14, pady=(12, 0))
        top.columnconfigure(0, weight=1)

        # step 1
        c1, b1, self.set_step1 = card(top, 1, "Upload your tables")
        c1.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        b1.columnconfigure(2, weight=1)
        ttk.Button(b1, text="Add files...", style="Accent.TButton", command=self.add_files).grid(row=0, column=0)
        ttk.Button(b1, text="Clear", command=self.clear).grid(row=0, column=1, padx=8)
        self.loaded = tk.StringVar(value="CSV or Excel files (optional .txt notes). No files loaded yet.")
        ttk.Label(b1, textvariable=self.loaded, style="Muted.TLabel", wraplength=800,
                  justify="left").grid(row=0, column=2, sticky="w", padx=6)

        # step 2
        c2, b2, self.set_step2 = card(top, 2, "Check the detected columns")
        c2.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(b2, text="The app guesses these automatically. Change a dropdown only if something is wrong.",
                  style="Muted.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        col = 0
        for g, title in GROUPS.items():
            fr = ttk.Frame(b2, style="Plain.TFrame")
            fr.grid(row=1, column=col, padx=(0, 22), sticky="n")
            col += 1
            ttk.Label(fr, text=title.upper(), style="Sub.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 3))
            self._combo(fr, g + "_t", "Table", 1, self.on_table)
            row = 2
            for k, (gg, label) in ROLES.items():
                if gg == g:
                    self._combo(fr, k, label, row, None)
                    row += 1

        # step 3
        c3, b3, _ = card(top, 3, "Choose a question")
        c3.grid(row=2, column=0, sticky="ew")
        b3.columnconfigure(0, weight=1)
        self.qvar = tk.StringVar()
        self.qbox = ttk.Combobox(b3, textvariable=self.qvar, state="readonly", values=[q[0] for q in QUESTIONS])
        self.qbox.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.qbox.current(0)
        self.qbox.bind("<<ComboboxSelected>>", lambda e: self.refresh_params())
        ttk.Button(b3, text="Get answer", style="Accent.TButton", command=self.answer).grid(row=0, column=1)

        prm = ttk.Frame(b3, style="Plain.TFrame")
        prm.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        prm.columnconfigure(5, weight=1)
        self.p1_label = ttk.Label(prm, text="", style="Card.TLabel")
        self.p1 = ttk.Combobox(prm, width=26)
        self.p2_label = ttk.Label(prm, text="Year", style="Card.TLabel")
        self.p2 = ttk.Entry(prm, width=8)
        self.p1_label.grid(row=0, column=0, padx=(0, 6))
        self.p1.grid(row=0, column=1)
        self.p2_label.grid(row=0, column=2, padx=(14, 6))
        self.p2.grid(row=0, column=3)
        ttk.Button(prm, text="Save report...", command=self.save).grid(row=0, column=6, padx=(6, 0))
        ttk.Button(prm, text="Copy code", command=self.copy_code).grid(row=0, column=5, padx=6, sticky="e")

        # output
        fo = ttk.Frame(r)
        fo.grid(row=2, column=0, sticky="nsew", padx=14, pady=(12, 16))
        fo.rowconfigure(0, weight=1)
        fo.columnconfigure(0, weight=1)
        self.text = tk.Text(fo, wrap="word", font=(FONT, 10), bg=CARD, fg=TEXT, relief="flat", padx=30, pady=24,
                            highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER,
                            spacing2=4, spacing3=6)
        sb = ttk.Scrollbar(fo, command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._tags()
        self.show_hint()
        self.refresh_params()

    def _tags(self):
        t = self.text
        t.tag_configure("q", foreground=MUTED, font=(FONT, 10), spacing3=12)
        t.tag_configure("ver", foreground=MUTED, font=(FONT, 9), spacing3=10)
        t.tag_configure("h", foreground=ACCENT, font=(FONT, 9, "bold"), spacing1=26, spacing3=8)
        t.tag_configure("answer", font=(FONT, 14, "bold"), foreground=TEXT, spacing1=14, spacing2=5, spacing3=6)
        t.tag_configure("body", font=(FONT, 10), foreground=TEXT, lmargin1=0, lmargin2=0, spacing2=5, spacing3=6)
        t.tag_configure("bullet", font=(FONT, 10), foreground=TEXT, lmargin1=14, lmargin2=30, spacing2=5, spacing3=8)
        t.tag_configure("code", font=("Consolas", 10), background=CODE_BG, foreground="#111827",
                        lmargin1=14, lmargin2=14, spacing1=3, spacing2=3, spacing3=3)
        t.tag_configure("hint_title", font=(FONT, 13, "bold"), foreground=TEXT, spacing3=6)
        t.tag_configure("hint", font=(FONT, 10), foreground=MUTED, spacing3=4)
        for k, c in STATUS_COLORS.items():
            t.tag_configure("badge_" + k, background=c, foreground="white", font=(FONT, 10, "bold"))

    def show_hint(self):
        t = self.text
        t.delete("1.0", "end")
        t.insert("end", "Welcome\n", "hint_title")
        t.insert("end", "1.  Click 'Add files...' and pick your CSV or Excel tables.\n", "hint")
        t.insert("end", "2.  Check that the detected columns look right.\n", "hint")
        t.insert("end", "3.  Choose a question and click 'Get answer'.\n\n", "hint")
        t.insert("end", "Your answer, the reasoning and the re-runnable code will appear here.", "hint")

    def _combo(self, parent, key, label, row, cb):
        ttk.Label(parent, text=label, style="Card.TLabel").grid(row=row, column=0, sticky="w", pady=2, padx=(0, 8))
        v = tk.StringVar(value=NONE)
        c = ttk.Combobox(parent, textvariable=v, state="readonly", width=21, values=[NONE])
        c.grid(row=row, column=1, pady=2)
        c.bind("<<ComboboxSelected>>", lambda e: (cb() if cb else None, self.refresh_params()))
        self.vars[key], self.boxes[key] = v, c

    # --- file handling
    def add_files(self):
        paths = filedialog.askopenfilenames(title="Choose tables", filetypes=[
            ("Tables and notes", "*.csv *.xlsx *.xlsm *.xls *.txt *.md"), ("All files", "*.*")])
        for p in paths:
            try:
                self.data.add(p)
            except Exception as e:
                messagebox.showerror("Could not read file", f"{os.path.basename(p)}\n\n{e}")
        self.after_load()

    def clear(self):
        self.data.clear()
        self.after_load()

    def after_load(self):
        names = [f"{n} ({len(t['df'])} rows)" for n, t in self.data.tables.items()] + \
                [f"{n} (notes)" for n in self.data.docs]
        self.loaded.set("   \u2022   ".join(names) if names else
                        "CSV or Excel files (optional .txt notes). No files loaded yet.")
        M = guess(self.data) if self.data.tables else {}
        tnames = [NONE] + list(self.data.tables)
        for g in GROUPS:
            self.boxes[g + "_t"]["values"] = tnames
            self.vars[g + "_t"].set(M.get(g + "_t") or NONE)
        self.on_table(M)
        self.refresh_params()

    def on_table(self, M=None):
        for k, (g, _) in ROLES.items():
            t = self.data.tables.get(self.vars[g + "_t"].get())
            cols = [NONE] + (list(t["df"].columns) if t else [])
            self.boxes[k]["values"] = cols
            cur = (M or {}).get(k) or self.vars[k].get()
            self.vars[k].set(cur if cur in cols else NONE)

    def mapping(self):
        M = {}
        for k in list(ROLES) + [g + "_t" for g in GROUPS]:
            v = self.vars[k].get()
            M[k] = "" if v == NONE else v
        return M

    # --- parameters
    def refresh_params(self):
        qi = self.qbox.current()
        kind = QUESTIONS[qi][1]
        M = self.mapping()
        d = self.data
        opts, label = [], ""
        try:
            o = tbl(d, M, "o")
            if kind == "currency" and o is not None and M["o_cur"]:
                opts = sorted(o["df"][M["o_cur"]].dropna().astype(str).str.strip().str.upper().unique())
                label = "Currency" if qi != 11 else "Express in currency"
            elif kind == "country" and o is not None and M["o_ship"]:
                opts = sorted(o["df"][M["o_ship"]].dropna().astype(str).str.strip().unique())
                label = "Country"
            elif kind == "customer":
                c = tbl(d, M, "c")
                if c is not None and M["c_id"]:
                    opts = sorted(c["df"][M["c_id"]].dropna().astype(str).str.strip().unique())
                label = "Customer ID"
            elif kind == "month":
                opts, label = MONTHS, "Month"
        except Exception:
            pass
        if kind in ("currency", "country", "customer", "month"):
            self.p1_label.config(text=label or "Value")
            self.p1["values"] = opts
            if self.p1.get() not in opts:
                self.p1.set(opts[0] if opts else "")
            self.p1_label.grid()
            self.p1.grid()
        else:
            self.p1_label.grid_remove()
            self.p1.grid_remove()
        if kind == "month":
            if not self.p2.get():
                self.p2.insert(0, "2025")
            self.p2_label.grid()
            self.p2.grid()
        else:
            self.p2_label.grid_remove()
            self.p2.grid_remove()

    # --- answer
    def _section(self, title, items, bullets=False):
        t = self.text
        t.insert("end", title.upper() + "\n", "h")
        for x in items:
            t.insert("end", ("\u2022  " if bullets else "") + x + "\n", "bullet" if bullets else "body")

    def render(self, qtext, r):
        t = self.text
        t.delete("1.0", "end")
        t.insert("end", qtext + "\n", "q")
        t.insert("end", " " + r["status"].replace("_", " ") + " ", "badge_" + r["status"])
        ver = {True: "verified: code re-run twice, same result", False: "NOT verified: runs differed",
               None: "no number to verify"}[r["verified"]]
        t.insert("end", "    " + ver + "\n", "ver")
        t.insert("end", "\n" + r["answer"] + "\n", "answer")
        self._section("Why", [r["reasoning"]])
        for title, key in (("Missing / ambiguous", "missing_or_ambiguous"), ("Evidence", "evidence"),
                           ("Assumptions", "assumptions")):
            if r[key]:
                self._section(title, r[key], bullets=True)
        if r["code"]:
            t.insert("end", "RE-RUNNABLE CODE  (save as check.py and run it; the last line is the number)\n", "h")
            t.insert("end", "\n" + r["code"].strip() + "\n", "code")

    def answer(self):
        if not self.data.tables:
            messagebox.showinfo("No data", "Please add at least one CSV or Excel file first.")
            return
        qi = self.qbox.current()
        self.root.config(cursor="watch")
        self.root.update()
        try:
            r = solve(qi, self.data, self.mapping(), self.p1.get(), self.p2.get())
        finally:
            self.root.config(cursor="")
        kind = QUESTIONS[qi][1]
        extra = f"  [{self.p1.get()}{' ' + self.p2.get() if kind == 'month' else ''}]" if kind else ""
        qtext = QUESTIONS[qi][0] + extra
        self.history.append((qtext, r))
        self.last_code = r["code"] or ""
        self.render(qtext, r)
        self.set_step1(False)  # fold the setup cards away so the answer gets the space
        self.set_step2(False)

    def copy_code(self):
        if not self.last_code:
            messagebox.showinfo("No code", "The current answer has no code (the question was refused or has no number).")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_code)
        messagebox.showinfo("Copied", "Code copied. Paste it into a new file named check.py and run it.")

    def save(self):
        if not self.history:
            messagebox.showinfo("Nothing to save", "Answer at least one question first.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".md", initialfile="report.md",
                                         filetypes=[("Markdown", "*.md"), ("Text", "*.txt")])
        if not p:
            return
        with open(p, "w", encoding="utf-8") as f:
            f.write("# Proof-Carrying Data Analyst - report\n")
            for q, r in self.history:
                f.write(f"\n## {q}\n\n```\n{format_result(r)}\n```\n")
        messagebox.showinfo("Saved", "Report saved to " + p)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
