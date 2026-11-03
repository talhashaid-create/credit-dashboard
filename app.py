"""Credit Recovery Intelligence - vehicle installment (EMI) portfolio.
Run:  streamlit run app.py
"""
import hmac
import io
import itertools
import json
import math
import os
import re
import warnings
from datetime import date

import numpy as np
import pandas as pd
import pdfplumber
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

warnings.filterwarnings("ignore", message="Could not infer format")
st.set_page_config(page_title="Credit Recovery Intelligence", page_icon="💳", layout="wide")


# ======================================================================
# Secrets / access control
# ======================================================================
def get_secret(name, default=None):
    """Read from Streamlit secrets, then environment variables. Never hard-code keys."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


_pw = get_secret("APP_PASSWORD")
if _pw:
    entered = st.sidebar.text_input("Password", type="password")
    if not hmac.compare_digest(entered.encode(), str(_pw).encode()):
        st.warning("Enter the password to continue.")
        st.stop()

# ======================================================================
# Configuration
# ======================================================================
FIELDS = {
    "serial": "Serial Number", "customer": "Customer Name", "branch": "Branch Name",
    "sales_person": "Sales Personnel", "account": "Account Id", "collector": "Collection Personnel",
    "vehicle": "Vehicle registration number", "vehicle_model": "Vehicle Model",
    "sales_price": "Sales price", "down_payment": "Down Payment", "interest": "Interest",
    "bill_amount": "Bill amount", "balance_receivable": "Balance receivable",
    "no_emi": "No of EMI", "emi_date": "EMI Date", "emi": "EMI",
    "inst_due_no": "Installment due number", "emi_coll_no": "EMI Collection number",
    "due_emi_no": "Due EMI Number", "total_due": "Total due amount",
    "total_coll": "Total Collection", "overdue_amt": "Overdue amount",
    "last_month_coll": "Last month collection", "net_balance": "Net balance",
}
TEXT_FIELDS = ["serial", "customer", "branch", "sales_person", "account", "collector",
               "vehicle", "vehicle_model"]
NUM_FIELDS = [k for k in FIELDS if k not in TEXT_FIELDS + ["emi_date"]]
ALIASES = {
    "inst_due_no": ["instalmentduenumber"],
    "vehicle": ["vehicleregno", "vehicleregistration", "regno", "vehicleno"],
    "vehicle_model": ["model", "vehiclemodel", "vehicletype"],
    "branch": ["branch"], "customer": ["customer"],
    "account": ["accountno", "accountnumber", "accid"],
    "sales_person": ["salesperson", "salesman"],
    "collector": ["collectionperson", "collector"],
    "total_coll": ["totalcollected"], "last_month_coll": ["lastmonthcollected"],
}
NONE = "— none —"

BUCKETS = ["Current", "1 EMI", "2 EMI", "3 EMI", "4–5 EMI", "6+ EMI"]
RISK = {"Current": "Regular", "1 EMI": "Low risk", "2 EMI": "Medium risk",
        "3 EMI": "High risk", "4–5 EMI": "High risk", "6+ EMI": "Critical"}
RISK_ORDER = ["Regular", "Low risk", "Medium risk", "High risk", "Critical"]
RISK_COLORS = {"Regular": "#4C9F70", "Low risk": "#9BC53D", "Medium risk": "#F2C14E",
               "High risk": "#E8833A", "Critical": "#C0392B"}
PROV_GROUPS = ["1 EMI due", "2 EMI due", "3–5 EMI due", "6+ EMI due"]
PROV_MAP = {"1 EMI": "1 EMI due", "2 EMI": "2 EMI due", "3 EMI": "3–5 EMI due",
            "4–5 EMI": "3–5 EMI due", "6+ EMI": "6+ EMI due"}
PROV_DEFAULT = [5.0, 25.0, 50.0, 100.0]
DRIVER_LABEL = {"overdue": "Overdue amount", "due_emi": "Due EMI count",
                "down_payment": "Low down payment", "outstanding": "Outstanding balance"}
FMT = {"outstanding": "{:,.0f}", "total_due": "{:,.0f}", "collected": "{:,.0f}",
       "overdue": "{:,.0f}", "last_month": "{:,.0f}", "efficiency": "{:.1%}",
       "overdue_pct": "{:.1%}", "provision": "{:,.0f}", "rate_pct": "{:.1f}%",
       "overdue_amt": "{:,.0f}", "score": "{:.1f}"}


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def pct(x, d=1):
    return "n/a" if x is None or pd.isna(x) else f"{x:.{d}%}"


def div(a, b):
    return a / b if b else np.nan


def auto_map(columns):
    nm = {norm(c): c for c in columns}
    out = {}
    for key, label in FIELDS.items():
        for cand in [norm(label)] + ALIASES.get(key, []):
            if cand in nm:
                out[key] = nm[cand]
                break
    return out


# ======================================================================
# Loading
# ======================================================================
@st.cache_data(show_spinner=False)
def get_sheets(fb):
    return pd.ExcelFile(io.BytesIO(fb)).sheet_names


@st.cache_data(show_spinner=False)
def detect_header_row(fb, sheet):
    """Row (1-based) that contains the most known column names."""
    raw = pd.read_excel(io.BytesIO(fb), sheet_name=sheet, header=None, nrows=40)
    known = set()
    for k, label in FIELDS.items():
        known.add(norm(label))
        known.update(ALIASES.get(k, []))
    best, best_score = 1, 0
    for i in range(len(raw)):
        score = sum(norm(v) in known for v in raw.iloc[i].tolist() if pd.notna(v))
        if score > best_score:
            best, best_score = i + 1, score
    return best if best_score >= 3 else 1


@st.cache_data(show_spinner=False)
def load_sheet(fb, sheet, header_row):
    df = pd.read_excel(io.BytesIO(fb), sheet_name=sheet, header=header_row - 1)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    junk = [c for c in df.columns if c.startswith("Unnamed") and df[c].isna().all()]
    return df.drop(columns=junk)


# ======================================================================
# Cleaning (every change is reported; source values are never edited)
# ======================================================================
def to_text(s):
    t = s.astype("string").str.strip().fillna("Unknown").replace("", "Unknown").astype(str)
    return t.replace({"nan": "Unknown", "None": "Unknown", "NaT": "Unknown"})


def to_num(s):
    blank = s.isna() | (s.astype(str).str.strip() == "")
    t = (s.astype(str).str.replace(r"[,\s]", "", regex=True)
         .str.replace(r"^\((.*)\)$", r"-\1", regex=True))
    out = pd.to_numeric(t, errors="coerce")
    out[blank] = np.nan
    return out, blank, out.isna() & ~blank


def parse_dates(s):
    num = pd.to_numeric(s, errors="coerce")
    if num.notna().mean() > 0.8:
        v = num.dropna()
        if v.between(1, 31).all():
            return pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]"), "day_of_month"
        if v.between(20000, 80000).all():
            return pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce"), "excel_serial"
    return pd.to_datetime(s, errors="coerce", dayfirst=True), "date"


@st.cache_data(show_spinner="Cleaning data…")
def prepare(raw, m, basis, rounding, dedupe):
    notes, exc, removed, derived = [], {}, [], []
    mapped = {k: bool(m.get(k)) for k in FIELDS}
    df = pd.DataFrame(index=raw.index)

    for k in TEXT_FIELDS:
        df[k] = to_text(raw[m[k]]) if mapped[k] else "Unknown"
    for k in ("account", "serial"):
        df[k] = df[k].str.replace(r"\.0$", "", regex=True)
    if not mapped["vehicle_model"]:
        df["vehicle_model"] = "Not available"

    blanks, invalid = {}, {}
    for k in NUM_FIELDS:
        if mapped[k]:
            df[k], blanks[k], invalid[k] = to_num(raw[m[k]])
        else:
            df[k] = np.nan

    date_kind = None
    if mapped["emi_date"]:
        df["emi_date_dt"], date_kind = parse_dates(raw[m["emi_date"]])
    else:
        df["emi_date_dt"] = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")

    def drop(mask, reason):
        nonlocal df
        mask = mask.to_numpy()
        if mask.any():
            r = raw.loc[df.index[mask]].copy()
            r.insert(0, "Removal reason", reason)
            removed.append(r)
            notes.append(("info", f"{int(mask.sum())} row(s) removed: {reason}."))
            df = df[~mask]

    # --- invalid / total rows ---
    pat = r"(?i)^\s*(?:grand\s*|sub\s*-?\s*)?totals?\s*:?\s*$"
    is_total = (df.customer.str.contains(pat) | df.serial.str.contains(pat)
                | df.branch.str.contains(pat))
    is_total |= df.customer.str.contains(r"(?i)\btotal\b") & (df.account == "Unknown")
    drop(is_total, "Sub-total / grand-total row")
    drop((df.customer == "Unknown") & (df.account == "Unknown"),
         "Blank row (no customer name and no account id)")

    # --- duplicates ---
    known = df.account != "Unknown"
    drop(known & df.drop(columns=["serial"]).duplicated(keep="first"),
         "Exact duplicate row (same Account Id, identical values)")
    known = df.account != "Unknown"
    dup_acc = known & df.account.duplicated(keep="first")
    if dedupe:
        drop(dup_acc, "Duplicate Account Id with different values (first occurrence kept)")
    elif dup_acc.any():
        notes.append(("warning", f"{int(dup_acc.sum())} duplicate Account Ids kept because duplicate "
                                 "removal is off; totals and account matching may be overstated."))

    # --- blanks / non-numeric values (read as 0, reported) ---
    for k in NUM_FIELDS:
        if mapped[k]:
            ni = int(invalid[k].loc[df.index].sum())
            nb = int(blanks[k].loc[df.index].sum())
            if ni:
                notes.append(("warning", f"'{m[k]}': {ni} non-numeric value(s) were read as 0."))
            if nb:
                notes.append(("info", f"'{m[k]}': {nb} blank cell(s) were read as 0 in totals."))
    df[NUM_FIELDS] = df[NUM_FIELDS].fillna(0)

    # --- outstanding basis ---
    df["outstanding"] = df["net_balance" if basis == "Net balance" else "balance_receivable"]

    # --- overdue amount ---
    if mapped["overdue_amt"]:
        df["overdue_raw"] = df.overdue_amt
    else:
        df["overdue_raw"] = df.total_due - df.total_coll
        derived.append("Overdue amount = Total due amount − Total Collection (column was not mapped).")
    df["overdue_amt"] = df.overdue_raw.clip(lower=0)

    # --- due EMI number ---
    if mapped["due_emi_no"]:
        df["due_raw"] = df.due_emi_no
    elif mapped["inst_due_no"] and mapped["emi_coll_no"]:
        df["due_raw"] = df.inst_due_no - df.emi_coll_no
        derived.append("Due EMI Number = Installment due number − EMI Collection number (column not mapped).")
    elif mapped["emi"]:
        df["due_raw"] = np.where(df.emi > 0, df.overdue_raw / df.emi.replace(0, np.nan), 0)
        derived.append("Due EMI Number = Overdue amount ÷ EMI (column not mapped; estimate).")
    else:
        df["due_raw"] = 0.0
        notes.append(("error", "Due EMI Number cannot be determined - bucket analysis is unreliable."))
    df["due_emi_no"] = df.due_raw.clip(lower=0)
    for d in derived:
        notes.append(("warning", "Derived field: " + d))

    # --- buckets ---
    fn = np.ceil if rounding.startswith("Round up") else np.floor
    df["due_emi_bkt"] = fn(df.due_emi_no.round(6)).astype(int)
    df["bucket"] = pd.cut(df.due_emi_bkt, [-1, 0, 1, 2, 3, 5, np.inf], labels=BUCKETS)
    df["bidx"] = df.bucket.cat.codes
    df["risk"] = pd.Categorical(df.bucket.astype(str).map(RISK), categories=RISK_ORDER, ordered=True)
    df["prov_group"] = pd.Categorical(df.bucket.astype(str).map(PROV_MAP),
                                      categories=PROV_GROUPS, ordered=True)
    if mapped["sales_price"] and mapped["down_payment"]:
        df["dp_pct"] = np.where(df.sales_price > 0,
                                df.down_payment / df.sales_price.replace(0, np.nan), np.nan)
    else:
        df["dp_pct"] = np.nan

    # --- validation checks ---
    show_cols = ["account", "customer", "branch", "collector", "outstanding",
                 "total_due", "total_coll", "overdue_raw", "due_raw"]

    def flag(mask, msg, level="warning"):
        n = int(mask.sum())
        if n:
            notes.append((level, f"{n} account(s): {msg}"))
            exc[msg] = df.loc[mask, show_cols]

    if mapped["overdue_amt"] and mapped["total_due"] and mapped["total_coll"]:
        flag(((df.total_due - df.total_coll) - df.overdue_raw).abs() > 1.0,
             "Overdue amount differs from Total due − Total Collection (tolerance 1.00)")
    if mapped["due_emi_no"] and mapped["inst_due_no"] and mapped["emi_coll_no"]:
        flag(((df.inst_due_no - df.emi_coll_no) - df.due_raw).abs() > 0.01,
             "Due EMI Number differs from Installment due − EMI Collection number")
    flag(df.outstanding < 0, "Negative outstanding balance")
    flag(df.overdue_raw < 0, "Negative overdue (advance payment?) - excluded from overdue totals")
    flag(df.due_raw < 0, "Negative Due EMI Number - treated as 0 in buckets")
    flag((df.overdue_amt > 0) & (df.due_raw <= 0), "Overdue amount > 0 but Due EMI Number is 0", "info")
    flag((df.overdue_amt == 0) & (df.due_raw > 0), "Due EMI Number > 0 but Overdue amount is 0", "info")
    flag(df.overdue_amt > df.outstanding + 1.0, "Overdue amount exceeds outstanding balance")
    if mapped["sales_price"] and mapped["down_payment"]:
        flag(df.down_payment > df.sales_price, "Down payment exceeds sales price")
    for col, nm_ in (("branch", "branch"), ("collector", "collection personnel")):
        n = int((df[col] == "Unknown").sum())
        if n:
            notes.append(("info", f"{n} account(s) have an unknown {nm_}."))
    if date_kind == "date":
        nbad = int((df.emi_date_dt.isna() & raw.loc[df.index, m["emi_date"]].notna()).sum())
        if nbad:
            notes.append(("warning", f"{nbad} EMI Date value(s) could not be read as dates."))
    elif date_kind == "day_of_month":
        notes.append(("info", "EMI Date holds a day-of-month (1-31); the reporting-period filter is disabled."))
    if not mapped["vehicle_model"]:
        notes.append(("info", "No Vehicle Model column mapped - vehicle-model analysis unavailable."))

    quality = {"notes": notes, "exceptions": exc, "mapped": mapped, "derived": derived,
               "removed": pd.concat(removed) if removed else pd.DataFrame(),
               "rows_in": len(raw), "rows_out": len(df)}
    return df, quality


# ======================================================================
# Analytics
# ======================================================================
def kpis(df):
    out, due = df.outstanding.sum(), df.total_due.sum()
    coll, od = df.total_coll.sum(), df.overdue_amt.sum()

    def par(n):
        return df.loc[df.due_emi_bkt >= n, "outstanding"].sum()

    demand, lm = df.loc[df.outstanding > 0, "emi"].sum(), df.last_month_coll.sum()
    return dict(bill=df.bill_amount.sum(), outstanding=out, total_due=due, collected=coll,
                overdue=od, efficiency=div(coll, due), overdue_pct=div(od, out),
                accounts=len(df), od_accounts=int((df.overdue_amt > 0).sum()),
                par1=div(par(1), out), par2=div(par(2), out), par3=div(par(3), out),
                par3_amt=par(3), par3_accts=int((df.due_emi_bkt >= 3).sum()),
                last_month=lm, demand=demand, lm_ratio=div(lm, demand), interest=df.interest.sum())


def perf(df, by, sort=True):
    g = df.groupby(by, observed=False).agg(
        accounts=("account", "size"), outstanding=("outstanding", "sum"),
        total_due=("total_due", "sum"), collected=("total_coll", "sum"),
        overdue=("overdue_amt", "sum"), last_month=("last_month_coll", "sum"),
        accts_overdue=("overdue_amt", lambda s: int((s > 0).sum())),
        accts_3plus=("due_emi_bkt", lambda s: int((s >= 3).sum())))
    g["efficiency"] = g.collected / g.total_due.replace(0, np.nan)
    g["overdue_pct"] = g.overdue / g.outstanding.replace(0, np.nan)
    return g.sort_values("overdue", ascending=False) if sort else g


MCOLS = ["account", "customer", "branch", "collector", "outstanding", "overdue_amt",
         "total_coll", "bucket", "bidx"]


def mom_accounts(cur, prev):
    """Match accounts on Account Id between two months."""
    c = cur.loc[cur.account != "Unknown", MCOLS]
    p = prev.loc[prev.account != "Unknown", MCOLS]
    mg = c.merge(p, on="account", how="outer", suffixes=("_cur", "_prev"), indicator=True)
    for col in ("customer", "branch", "collector"):
        mg[col] = mg[f"{col}_cur"].fillna(mg[f"{col}_prev"])
    mg["movement"] = np.select(
        [mg._merge == "left_only", mg._merge == "right_only",
         mg.bidx_cur > mg.bidx_prev, mg.bidx_cur < mg.bidx_prev],
        ["New account", "Closed / not in current file", "Moved to higher bucket",
         "Moved to lower bucket"], "No change")
    mg["overdue_change"] = mg.overdue_amt_cur.fillna(0) - mg.overdue_amt_prev.fillna(0)
    mg["collection_in_period"] = mg.total_coll_cur - mg.total_coll_prev
    mg["bucket_cur"] = mg.bucket_cur.astype(object)
    mg["bucket_prev"] = mg.bucket_prev.astype(object)
    return mg


def compare_group(cur, prev, by):
    cols = ["accounts", "outstanding", "overdue", "collected", "efficiency"]
    a, b = perf(cur, by, False)[cols], perf(prev, by, False)[cols]
    j = a.join(b, how="outer", lsuffix="_cur", rsuffix="_prev")
    for c in ("accounts", "outstanding", "overdue", "collected"):
        j[[f"{c}_cur", f"{c}_prev"]] = j[[f"{c}_cur", f"{c}_prev"]].fillna(0)
    j["overdue_change"] = j.overdue_cur - j.overdue_prev
    j["collected_change"] = j.collected_cur - j.collected_prev
    j["efficiency_change_pts"] = (j.efficiency_cur - j.efficiency_prev) * 100
    return j.sort_values("overdue_cur", ascending=False)


def roll_matrix(cur, prev):
    """Opening (previous file) -> closing (current file) bucket movement."""
    c = cur.loc[cur.account != "Unknown", ["account", "bucket", "outstanding"]]
    p = prev.loc[prev.account != "Unknown", ["account", "bucket", "outstanding"]]
    mg = p.merge(c, on="account", how="outer", suffixes=("_open", "_close"), indicator=True)
    new_lbl, closed_lbl = "New (not in opening)", "Closed / not in closing"
    o = mg.bucket_open.astype(object).where(mg._merge != "right_only", new_lbl)
    cl = mg.bucket_close.astype(object).where(mg._merge != "left_only", closed_lbl)
    rows, cols = BUCKETS + [new_lbl], BUCKETS + [closed_lbl]
    cnt = pd.crosstab(o, cl).reindex(index=rows, columns=cols, fill_value=0)
    amt = pd.crosstab(o, cl, values=mg.outstanding_open.fillna(mg.outstanding_close),
                      aggfunc="sum").reindex(index=rows, columns=cols).fillna(0)
    base = cnt.loc[BUCKETS]
    rate = base.div(base.sum(axis=1).replace(0, np.nan), axis=0)
    summ = []
    for i, b in enumerate(BUCKETS):
        tot = int(base.loc[b].sum())
        hi, lo = int(base.loc[b, BUCKETS[i + 1:]].sum()), int(base.loc[b, BUCKETS[:i]].sum())
        summ.append({"Opening bucket": b, "Opening accounts": tot,
                     "Improved (lower bucket)": lo, "Same bucket": int(base.loc[b, b]),
                     "Rolled to higher bucket": hi, "Closed / not in closing": int(base.loc[b, closed_lbl]),
                     "Roll-forward rate": div(hi, tot)})
    return {"cnt": cnt, "amt": amt, "rate": rate, "summary": pd.DataFrame(summ).set_index("Opening bucket")}


def provision_table(df, rates):
    g = df.groupby("prov_group", observed=False).agg(
        accounts=("account", "size"), outstanding=("outstanding", "sum"), overdue=("overdue_amt", "sum"))
    g["rate_pct"] = [rates[i] for i in g.index]
    g["provision"] = g.outstanding * g.rate_pct / 100
    return g


def priority(df, w, dp_cap, cur, top_n):
    d = df[df.overdue_amt > 0].copy()
    if d.empty:
        return d
    d["p_overdue"] = d.overdue_amt.rank(pct=True)
    d["p_outstanding"] = d.outstanding.rank(pct=True)
    c = pd.DataFrame(index=d.index)
    c["overdue"] = d.p_overdue * 100
    c["due_emi"] = d.due_emi_no.clip(upper=6) / 6 * 100
    c["down_payment"] = (1 - d.dp_pct.clip(0, dp_cap) / dp_cap) * 100  # NaN stays NaN (not invented)
    c["outstanding"] = d.p_outstanding * 100
    wt = pd.Series(w)[c.columns]
    num = c.mul(wt).sum(axis=1, skipna=True)
    den = c.notna().mul(wt).sum(axis=1)
    d["score"] = num / den.replace(0, np.nan)
    d["main_driver"] = c.mul(wt).div(den.replace(0, np.nan), axis=0).fillna(-1).idxmax(axis=1).map(DRIVER_LABEL)
    d = d.sort_values("score", ascending=False).head(top_n)
    d["rank"] = range(1, len(d) + 1)

    def topx(p):
        return max(1, int(math.ceil((1 - p) * 100)))

    d["explanation"] = [
        f"Overdue {cur}{r.overdue_amt:,.0f} (top {topx(r.p_overdue)}% of overdue accounts); "
        f"{r.due_emi_no:g} EMI due; "
        + (f"down payment {r.dp_pct:.0%} of sales price; " if pd.notna(r.dp_pct)
           else "down payment % not available (excluded from score); ")
        + f"outstanding {cur}{r.outstanding:,.0f} (top {topx(r.p_outstanding)}%). "
          f"Main driver: {r.main_driver}."
        for r in d.itertuples()]
    return d


def read_targets(file):
    t = pd.read_excel(file)
    cols = {norm(c): c for c in t.columns}
    ccol = next((cols[k] for k in cols if "collect" in k or k in ("name", "staff", "personnel")), t.columns[0])
    tcol = next((cols[k] for k in cols if "target" in k), None)
    if tcol is None:
        raise ValueError("No column containing 'target' found.")
    s = pd.Series(pd.to_numeric(t[tcol], errors="coerce").fillna(0).values,
                  index=t[ccol].astype(str).str.strip().str.lower())
    return s.groupby(level=0).sum()


# ======================================================================
# Rule-based commentary
# ======================================================================
def make_insights(k, br, co, cur, top10_share, mom):
    m = lambda x: f"{cur}{x:,.0f}"
    ins = [
        f"Outstanding is {m(k['outstanding'])} across {k['accounts']:,} accounts; {m(k['overdue'])} "
        f"({pct(k['overdue_pct'])}) is overdue, spread over {k['od_accounts']:,} accounts.",
        f"Collection efficiency is {pct(k['efficiency'])}: {m(k['collected'])} collected against {m(k['total_due'])} due.",
        f"PAR1 {pct(k['par1'])}, PAR2 {pct(k['par2'])}, PAR3 {pct(k['par3'])}; {k['par3_accts']:,} accounts "
        f"have 3+ EMIs due with {m(k['par3_amt'])} outstanding.",
        f"Last-month collection was {m(k['last_month'])} against a monthly EMI demand of {m(k['demand'])} "
        f"({pct(k['lm_ratio'], 0)}).",
    ]
    if pd.notna(top10_share):
        ins.append(f"The 10 largest overdue accounts hold {pct(top10_share, 0)} of total overdue.")
    b = br[br.outstanding > 0]
    if len(b) > 1:
        ins.append(f"Highest overdue % by branch: {b.overdue_pct.idxmax()} ({pct(b.overdue_pct.max())}); "
                   f"best collection efficiency: {b.efficiency.idxmax()} ({pct(b.efficiency.max())}).")
    c = co[co.total_due > 0]
    if len(c) > 1:
        ins.append(f"Lowest collector efficiency: {c.efficiency.idxmin()} ({pct(c.efficiency.min())}).")
    if mom:
        ins.append(f"Versus the previous file: outstanding {mom['d_out']:+,.0f}, overdue {mom['d_od']:+,.0f}, "
                   f"efficiency {mom['d_eff'] * 100:+.1f} pts; {mom['higher']:,} accounts moved to a higher "
                   f"bucket and {mom['lower']:,} improved.")
    return ins


def make_recs(k, br, co, cur, mom):
    recs = []
    if k["par3_accts"]:
        recs.append((f"Issue final notices and start vehicle-recovery review for {k['par3_accts']:,} accounts with "
                     f"3+ EMIs due ({cur}{k['par3_amt']:,.0f} outstanding).", "Within 7 days"))
    recs.append(("Contact every 1-EMI-due customer to stop roll-over into higher buckets.", "Within 7 days"))
    if pd.notna(k["efficiency"]) and k["efficiency"] < 0.90:
        recs.append(("Collection efficiency is below 90%: set weekly collector targets and daily follow-up reviews.",
                     "This month"))
    b = br[br.outstanding > 0]
    if len(b) > 1:
        recs.append((f"Hold a recovery review with the branch manager of {b.overdue_pct.idxmax()} "
                     "(highest overdue %).", "Within 14 days"))
    c = co[co.total_due > 0]
    if len(c) > 1:
        recs.append((f"Review workload and give field support to {c.efficiency.idxmin()} (lowest efficiency).",
                     "Within 14 days"))
    if mom and mom["higher"]:
        recs.append((f"Investigate the {mom['higher']:,} accounts that moved to a higher bucket since the last file.",
                     "Within 14 days"))
    recs.append(("Tighten approval for low down-payment sales; track overdue by sales personnel monthly.",
                 "Next quarter"))
    return recs


# ======================================================================
# Gemini (free tier) - only aggregated figures are sent
# ======================================================================
def _jd(o):
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.floating):
        return None if np.isnan(o) else float(o)
    return str(o)


def gemini_commentary(payload, key, model):
    prompt = (
        "You are a senior credit-risk analyst writing for the management of a vehicle installment-sale company.\n"
        "Rules: use ONLY the figures in the JSON; never invent numbers; if something is not provided, say so; "
        "be concise and use plain executive language.\n"
        "Write these sections with headings: 1. Overall performance and major changes, 2. Collection efficiency, "
        "3. High-risk accounts, 4. Branch and collector performance, 5. Recovery recommendations (5 specific actions).\n\n"
        f"DATA:\n{json.dumps(payload, default=_jd)}")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    r = requests.post(url, headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                      json={"contents": [{"parts": [{"text": prompt}]}],
                            "generationConfig": {"temperature": 0.2, "maxOutputTokens": 4096}}, timeout=90)
    r.raise_for_status()
    return r.json()["candidates"][0]["content"]["parts"][0]["text"]


# ======================================================================
# PDF
# ======================================================================
@st.cache_data(show_spinner=False)
def read_pdf(b):
    text, tables = "", []
    with pdfplumber.open(io.BytesIO(b)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
            for t in page.extract_tables():
                if t and len(t) > 1:
                    tables.append(pd.DataFrame(t).fillna("").astype(str))
    return text, tables


# ======================================================================
# PowerPoint (native, editable charts)
# ======================================================================
def _rgb(h):
    return RGBColor.from_string(h.lstrip("#").upper())


def para(tf, text, size=14, bold=False, color=None, first=False, space=6, align=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    r = p.add_run()
    r.text = text
    r.font.size, r.font.bold = Pt(size), bold
    if color is not None:
        r.font.color.rgb = color
    p.space_after = Pt(space)
    if align:
        p.alignment = align
    return p


def build_pptx(c):
    brand = _rgb(c["brand"])
    palette = [brand, _rgb("#E8833A"), _rgb("#4C9F70"), _rgb("#C0392B"), _rgb("#7F8C8D")]
    white, grey = RGBColor(255, 255, 255), RGBColor(0x6B, 0x75, 0x80)
    cur, k = c["cur"], c["k"]
    m = lambda x: f"{cur}{x:,.0f}"
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    counter = [0]

    def new_slide(title):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        counter[0] += 1
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.95))
        bar.fill.solid()
        bar.fill.fore_color.rgb = brand
        bar.line.fill.background()
        tb = s.shapes.add_textbox(Inches(0.5), Inches(0.12), Inches(12.3), Inches(0.75))
        para(tb.text_frame, title, 28, True, white, first=True)
        ft = s.shapes.add_textbox(Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.35))
        para(ft.text_frame, f"{c['company']}  |  Credit Recovery Review  |  As of {c['as_of']}  |  {counter[0]}",
             10, False, grey, first=True)
        return s

    def kpi_card(s, label, value, x, y, w=3.9, h=1.2):
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        card.fill.solid()
        card.fill.fore_color.rgb = _rgb("#F2F5F9")
        card.line.color.rgb = brand
        tf = card.text_frame
        tf.word_wrap = True
        para(tf, value, 26, True, brand, first=True, space=0, align=PP_ALIGN.CENTER)
        para(tf, label, 12, False, grey, space=0, align=PP_ALIGN.CENTER)

    def bullets(s, items, x, y, w, h, size=14):
        tf = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
        tf.word_wrap = True
        for i, t in enumerate(items):
            para(tf, "• " + t, size, first=(i == 0))

    def chart(s, cats, series, x, y, w, h, title, fmt="#,##0"):
        cd = CategoryChartData()
        cd.categories = [str(i) for i in cats]
        for name, vals in series.items():
            cd.add_series(name, [None if pd.isna(v) else float(v) for v in vals])
        ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y),
                                Inches(w), Inches(h), cd).chart
        ch.has_title = True
        ch.chart_title.text_frame.text = title
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
        ch.has_legend = len(series) > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.number_format = fmt
        plot.data_labels.number_format_is_linked = False
        plot.data_labels.font.size = Pt(9)
        for i, ser in enumerate(plot.series):
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = palette[i % len(palette)]
        ch.category_axis.tick_labels.font.size = Pt(10)
        ch.value_axis.tick_labels.font.size = Pt(10)
        ch.value_axis.tick_labels.number_format = fmt
        ch.value_axis.tick_labels.number_format_is_linked = False

    def ppt_table(s, header, rows, x, y, w, widths, size=11, row_h=0.36):
        t = s.shapes.add_table(len(rows) + 1, len(header), Inches(x), Inches(y), Inches(w),
                               Inches(row_h * (len(rows) + 1))).table
        for j, fr in enumerate(widths):
            t.columns[j].width = Inches(w * fr)
        for j, h_ in enumerate(header):
            cell = t.cell(0, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = brand
            cell.text_frame.clear()
            para(cell.text_frame, str(h_), size, True, white, first=True, space=0)
        for i, row in enumerate(rows, 1):
            for j, v in enumerate(row):
                cell = t.cell(i, j)
                cell.text_frame.clear()
                para(cell.text_frame, str(v), size, False, None, first=True, space=0)

    aging, br, co = c["aging"], c["br"], c["co"]

    # 1 Executive summary
    s = new_slide("Executive Summary")
    cards = [("Total Outstanding", m(k["outstanding"])), ("Total Overdue", m(k["overdue"])),
             ("Total Collection", m(k["collected"])), ("Collection Efficiency", pct(k["efficiency"])),
             ("PAR3 (3+ EMI due)", pct(k["par3"])), ("Accounts / with overdue", f"{k['accounts']:,} / {k['od_accounts']:,}")]
    for i, (lbl, val) in enumerate(cards):
        r_, c_ = divmod(i, 3)
        kpi_card(s, lbl, val, 0.5 + c_ * 4.2, 1.2 + r_ * 1.4)
    bullets(s, c["insights"][:5], 0.5, 4.1, 12.3, 2.9, 13)
    s.notes_slide.notes_text_frame.text = c["ai_text"] or "Rule-based insights shown on slide."

    # 2 Portfolio overview
    s = new_slide("Portfolio Overview")
    chart(s, aging.index, {"Accounts": aging.accounts}, 0.5, 1.2, 6.4, 5.7, "Accounts by Due-EMI bucket")
    rows = [("Bill amount (portfolio)", m(k["bill"])), ("Total outstanding", m(k["outstanding"])),
            ("Total due to date", m(k["total_due"])), ("Total collection", m(k["collected"])),
            ("Overdue amount", m(k["overdue"])), ("Accounts", f"{k['accounts']:,}"),
            ("Accounts with overdue", f"{k['od_accounts']:,}"), ("Interest", m(k["interest"])),
            ("Last-month collection", m(k["last_month"])), ("Monthly EMI demand", m(k["demand"]))]
    ppt_table(s, ["Metric", "Value"], rows, 7.2, 1.3, 5.6, [0.6, 0.4], 12, 0.42)

    # 3 Outstanding & overdue
    s = new_slide("Outstanding & Overdue Analysis")
    chart(s, aging.index, {"Outstanding": aging.outstanding, "Overdue": aging.overdue},
          0.5, 1.2, 7.6, 5.7, "Outstanding and overdue by Due-EMI bucket")
    b3 = [f"Overdue is {pct(k['overdue_pct'])} of outstanding ({m(k['overdue'])}).",
          f"PAR1 {pct(k['par1'])} | PAR2 {pct(k['par2'])} | PAR3 {pct(k['par3'])}."]
    od_b = aging.loc[BUCKETS[1:], "overdue"]
    if od_b.sum() > 0:
        b3.append(f"Largest overdue balance sits in the '{od_b.idxmax()}' bucket ({m(od_b.max())}).")
    if c["mom"]:
        mo = c["mom"]
        b3 += [f"Outstanding moved {mo['d_out']:+,.0f} and overdue {mo['d_od']:+,.0f} versus the previous file.",
               f"{mo['higher']:,} accounts moved to a higher bucket; {mo['lower']:,} improved."]
    bullets(s, b3, 8.3, 1.4, 4.6, 5.0, 14)

    # 4 Collection efficiency
    s = new_slide("Collection Efficiency")
    chart(s, ["Total due", "Total collection", "Overdue"],
          {"Amount": [k["total_due"], k["collected"], k["overdue"]]}, 0.5, 1.2, 5.8, 4.6, "Due vs collected")
    chart(s, aging.index, {"Efficiency": aging.efficiency}, 6.6, 1.2, 6.2, 4.6,
          "Collection efficiency by Due-EMI bucket", "0.0%")
    bullets(s, [f"Collection efficiency = Total Collection ÷ Total due amount = {pct(k['efficiency'])}.",
                f"Last-month collection {m(k['last_month'])} is {pct(k['lm_ratio'], 0)} of the monthly EMI demand."],
            0.5, 6.0, 12.3, 1.0, 13)

    # 5 Branch
    s = new_slide("Branch Performance")
    b = br.head(8)
    chart(s, b.index, {"Overdue": b.overdue}, 0.5, 1.2, 6.1, 5.7, "Overdue amount by branch (top 8)")
    chart(s, b.index, {"Efficiency": b.efficiency}, 6.8, 1.2, 6.1, 5.7, "Collection efficiency by branch", "0.0%")

    # 6 Collector
    s = new_slide("Collector Performance")
    cc = co.head(8)
    chart(s, cc.index, {"Overdue": cc.overdue}, 0.5, 1.2, 6.1, 5.7, "Overdue amount by collector (top 8)")
    tg = c["targets"]
    if tg is not None:
        t8 = tg[tg.Target > 0].sort_values("Achievement %", ascending=False).head(10)
        chart(s, t8.Collector, {"Achievement": t8["Achievement %"]}, 6.8, 1.2, 6.1, 5.7,
              "Target achievement by collector", "0%")
    else:
        chart(s, cc.index, {"Efficiency": cc.efficiency}, 6.8, 1.2, 6.1, 5.7,
              "Collection efficiency by collector", "0.0%")

    # 7 Vehicle model
    s = new_slide("Vehicle Model Analysis")
    if c["vm"] is not None:
        v = c["vm"].head(8)
        chart(s, v.index, {"Overdue": v.overdue}, 0.5, 1.2, 6.1, 5.7, "Overdue by vehicle model (top 8)")
        chart(s, v.index, {"Efficiency": v.efficiency}, 6.8, 1.2, 6.1, 5.7, "Collection efficiency by model", "0.0%")
    else:
        bullets(s, ["Vehicle-model analysis is unavailable: the uploaded Excel has no Vehicle Model column.",
                    "Add a model column (or map an existing one in the app) and regenerate this deck."],
                0.8, 2.0, 11.5, 3.0, 18)

    # 8 Roll-rate
    s = new_slide("Roll-Rate Analysis")
    if c["roll"] is not None:
        rate = c["roll"]["rate"]
        hdr = ["Opening \\ Closing"] + [str(x) for x in rate.columns]
        rows = [[i] + ["–" if pd.isna(v) else f"{v:.0%}" for v in rate.loc[i]] for i in rate.index]
        ppt_table(s, hdr, rows, 0.5, 1.3, 12.3, [0.17] + [0.1186] * 7, 12, 0.5)
        sm = c["roll"]["summary"]
        tot_open, tot_hi = sm["Opening accounts"].sum(), sm["Rolled to higher bucket"].sum()
        bullets(s, [f"Overall, {pct(div(tot_hi, tot_open))} of opening accounts rolled into a higher bucket.",
                    "Rows = opening bucket (previous file); columns = closing bucket (current file); "
                    "accounts matched on Account Id."], 0.5, 5.2, 12.3, 1.6, 13)
    else:
        bullets(s, ["Roll-rate analysis needs both the current and previous month files. "
                    "Upload the previous month Excel to populate this slide."], 0.8, 2.0, 11.5, 2.0, 18)

    # 9 High-risk & provisioning
    s = new_slide("High-Risk Accounts & Provisioning")
    pr = c["prio"]
    if len(pr):
        rows = [(r.rank, str(r.customer)[:22], str(r.branch)[:14], f"{r.due_emi_no:g}",
                 m(r.overdue_amt), f"{r.score:.0f}") for r in pr.head(8).itertuples()]
        ppt_table(s, ["#", "Customer", "Branch", "Due EMI", "Overdue", "Score"], rows, 0.4, 1.3, 7.5,
                  [0.07, 0.31, 0.21, 0.13, 0.18, 0.1], 11, 0.42)
    pv = c["prov"]
    rows = [(i, m(r.outstanding), f"{r.rate_pct:.0f}%", m(r.provision)) for i, r in pv.iterrows()]
    rows.append(("Total", m(pv.outstanding.sum()), "", m(pv.provision.sum())))
    ppt_table(s, ["Bucket", "Outstanding", "Rate", "Provision"], rows, 8.1, 1.3, 4.8, [0.28, 0.28, 0.14, 0.3], 11, 0.42)
    w = c["weights"]
    bullets(s, ["Provisions are a MANAGEMENT ESTIMATE, not an official accounting provision.",
                f"Priority score weights - overdue {w['overdue']:.0%}, due EMI {w['due_emi']:.0%}, "
                f"down payment {w['down_payment']:.0%}, outstanding {w['outstanding']:.0%}."],
            0.4, 5.2, 12.5, 1.6, 12)

    # 10 Recommendations
    s = new_slide("Recommendations & Action Plan")
    rows = [(i, a, t) for i, (a, t) in enumerate(c["recs"][:7], 1)]
    ppt_table(s, ["#", "Recommended action", "Suggested timing"], rows, 0.5, 1.3, 12.3, [0.05, 0.75, 0.2], 12, 0.62)
    notes = "\n\n".join(x for x in (c["ai_text"], c["pdf_notes"]) if x)
    if notes:
        s.notes_slide.notes_text_frame.text = notes

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_excel(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, (d, idx) in sheets.items():
            d.to_excel(w, sheet_name=name[:31], index=idx)
        for ws in w.book.worksheets:
            for col in ws.columns:
                width = max((len(str(x.value)) for x in col[:200] if x.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 45)
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
    return buf.getvalue()


# ======================================================================
# UI helpers
# ======================================================================
_ids = itertools.count()


def plot(fig):
    key = f"pc{next(_ids)}"
    try:
        st.plotly_chart(fig, use_container_width=True, key=key)
    except TypeError:  # future Streamlit versions
        st.plotly_chart(fig, width="stretch", key=key)


def show_df(d, fmt=None, style=True, **kw):
    try:
        if style:
            f = {c: s for c, s in {**FMT, **(fmt or {})}.items() if c in d.columns}
            d = d.style.format(f, na_rep="–")
        st.dataframe(d, use_container_width=True, **kw)
    except TypeError:
        st.dataframe(d, width="stretch", **kw)


def style_fig(fig, title=None, h=380):
    fig.update_layout(template="plotly_white", colorway=[brand, "#E8833A", "#4C9F70", "#C0392B", "#7F8C8D"],
                      title=title, height=h, margin=dict(l=10, r=10, t=55, b=10),
                      legend=dict(orientation="h", y=1.12, x=0))
    return fig


def grouped_bar(tbl, cols, labels, title, top=15):
    t = tbl.head(top)
    fig = go.Figure([go.Bar(name=lb, x=t.index.astype(str), y=t[c_]) for c_, lb in zip(cols, labels)])
    fig.update_layout(barmode="group")
    return style_fig(fig, title)


# ======================================================================
# SIDEBAR
# ======================================================================
st.sidebar.markdown("## 💳 Credit Recovery Intelligence")
company = st.sidebar.text_input("Company / report name", "Your Company")
brand = st.sidebar.color_picker("Brand colour", "#1F3A5F")
cur = st.sidebar.text_input("Currency symbol (optional)", "")
as_of = st.sidebar.date_input("Report date", date.today())
m_ = lambda x: f"{cur}{x:,.0f}"

st.markdown(f"""<style>
.block-container{{padding-top:1.2rem}}
.banner{{background:{brand};color:#fff;padding:16px 24px;border-radius:10px;margin-bottom:14px}}
.banner h2{{margin:0;color:#fff}} .banner p{{margin:2px 0 0;opacity:.85}}
div[data-testid="stMetric"]{{background:#fff;border:1px solid #e3e7ec;border-left:5px solid {brand};
padding:10px 14px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.05)}}
</style>""", unsafe_allow_html=True)

st.sidebar.subheader("1 · Upload")
f_cur = st.sidebar.file_uploader("Current month Excel", type=["xlsx", "xls"])
f_prev = st.sidebar.file_uploader("Previous month Excel (optional)", type=["xlsx", "xls"])
f_pdf = st.sidebar.file_uploader("PDF report / policy (optional)", type=["pdf"])

if not f_cur:
    st.markdown(f"<div class='banner'><h2>{company} · Credit Recovery Dashboard</h2>"
                "<p>Upload the current month Excel file in the sidebar to begin.</p></div>", unsafe_allow_html=True)
    st.write("**Expected columns:**", ", ".join(v for k_, v in FIELDS.items() if k_ != "vehicle_model"),
             "(optional: Vehicle Model)")
    st.stop()


def sheet_controls(label, file, key):
    fb = file.getvalue()
    sheet = st.sidebar.selectbox(f"{label} – sheet", get_sheets(fb), key=f"sheet_{key}")
    auto = detect_header_row(fb, sheet)
    hdr = st.sidebar.number_input(f"{label} – header row (auto-detected)", 1, 50, auto,
                                  key=f"hdr_{key}_{file.name}_{sheet}")
    return fb, sheet, int(hdr)


fb, sheet, hdr = sheet_controls("Current", f_cur, "cur")
try:
    raw = load_sheet(fb, sheet, hdr)
except Exception as e:
    st.error(f"Could not read the sheet: {e}")
    st.stop()

detected = auto_map(raw.columns)
st.sidebar.caption(f"Auto-detected {len(detected)} of {len(FIELDS)} columns.")
mapping = {}
with st.sidebar.expander("Column mapping (check / change)"):
    opts = [NONE] + list(raw.columns)
    for key, label in FIELDS.items():
        idx = opts.index(detected[key]) if key in detected else 0
        sel = st.selectbox(label, opts, index=idx, key=f"map_{f_cur.name}_{sheet}_{hdr}_{key}")
        mapping[key] = None if sel == NONE else sel


def missing_fields(mp):
    miss = []
    if not mp.get("customer"):
        miss.append("Customer Name")
    if not mp.get("account"):
        miss.append("Account Id")
    if not (mp.get("net_balance") or mp.get("balance_receivable")):
        miss.append("Net balance or Balance receivable")
    if not (mp.get("overdue_amt") or (mp.get("total_due") and mp.get("total_coll"))):
        miss.append("Overdue amount (or Total due amount + Total Collection)")
    if not (mp.get("due_emi_no") or (mp.get("inst_due_no") and mp.get("emi_coll_no")) or mp.get("emi")):
        miss.append("Due EMI Number (or Installment due + EMI Collection number, or EMI)")
    return miss


miss = missing_fields(mapping)
if miss:
    st.error("Please map these columns in the sidebar: " + "; ".join(miss))
    st.stop()

st.sidebar.subheader("2 · Settings")
basis_opts = [n for n, k_ in [("Net balance", "net_balance"), ("Balance receivable", "balance_receivable")]
              if mapping[k_]]
basis = st.sidebar.radio("Outstanding =", basis_opts)
rounding = st.sidebar.radio("Fractional Due EMI", ["Round up (conservative)", "Round down"],
                            help="Used only to place accounts in buckets. Raw values are kept.")
dedupe = st.sidebar.checkbox("Remove duplicate Account Ids (keep first)", True)

cur_all, Q = prepare(raw, mapping, basis, rounding, dedupe)
if cur_all.empty:
    st.error("No valid rows remain after cleaning. Check the header row and column mapping.")
    st.stop()

# ---- previous month ----
prev_all = None
if f_prev:
    fbp, shp, hp = sheet_controls("Previous", f_prev, "prev")
    try:
        raw_p = load_sheet(fbp, shp, hp)
        mp = {k_: v for k_, v in auto_map(raw_p.columns).items()}
        for k_, v in mapping.items():  # reuse the current mapping when the header exists
            if v and v in raw_p.columns:
                mp[k_] = v
        mp = {k_: mp.get(k_) for k_ in FIELDS}
        pm = missing_fields(mp)
        if pm or not mp.get("net_balance" if basis == "Net balance" else "balance_receivable"):
            st.sidebar.error("Previous file: missing " + "; ".join(pm or [basis]))
        else:
            prev_all, Qp = prepare(raw_p, mp, basis, rounding, dedupe)
    except Exception as e:
        st.sidebar.error(f"Previous file could not be read: {e}")

# ---- filters ----
st.sidebar.subheader("3 · Filters")
branches = sorted(cur_all.branch.unique())
sel_b = st.sidebar.multiselect("Branch", branches, default=branches)
collectors = sorted(cur_all[cur_all.branch.isin(sel_b)].collector.unique())
sel_c = st.sidebar.multiselect("Collector", collectors, default=collectors)
mask = cur_all.branch.isin(sel_b) & cur_all.collector.isin(sel_c)
has_model = Q["mapped"]["vehicle_model"]
if has_model:
    models = sorted(cur_all.vehicle_model.unique())
    sel_m = st.sidebar.multiselect("Vehicle model", models, default=models)
    mask &= cur_all.vehicle_model.isin(sel_m)
sel_bk = st.sidebar.multiselect("Due EMI bucket", BUCKETS, default=BUCKETS)
mask &= cur_all.bucket.isin(sel_bk)
period_txt = "all EMI dates"
dts = cur_all.emi_date_dt.dropna()
if len(dts) > 0.5 * len(cur_all):
    if st.sidebar.checkbox("Reporting period (filter by EMI Date)"):
        rng = st.sidebar.date_input("EMI Date range", (dts.min().date(), dts.max().date()))
        if isinstance(rng, tuple) and len(rng) == 2:
            mask &= cur_all.emi_date_dt.between(pd.Timestamp(rng[0]), pd.Timestamp(rng[1]))
            period_txt = f"{rng[0]} to {rng[1]}"
df = cur_all[mask]
if df.empty:
    st.warning("No accounts match the selected filters.")
    st.stop()

# ======================================================================
# COMPUTE
# ======================================================================
k = kpis(df)
aging = perf(df, "bucket", False)
br, co, sp = perf(df, "branch"), perf(df, "collector"), perf(df, "sales_person")
vm = perf(df, "vehicle_model") if has_model else None
risk_t = perf(df, "risk", False)
top_df = df[df.overdue_amt > 0].nlargest(15, "overdue_amt")
top10_share = div(df.nlargest(10, "overdue_amt").overdue_amt.sum(), k["overdue"])

mom_sum = mv = roll = None
if prev_all is not None:
    kc, kp = kpis(cur_all), kpis(prev_all)
    mv = mom_accounts(cur_all, prev_all)
    roll = roll_matrix(cur_all, prev_all)
    vc = mv.movement.value_counts()
    mom_sum = dict(d_out=kc["outstanding"] - kp["outstanding"], d_od=kc["overdue"] - kp["overdue"],
                   d_eff=(kc["efficiency"] - kp["efficiency"]) if pd.notna(kc["efficiency"]) and pd.notna(kp["efficiency"]) else 0,
                   higher=int(vc.get("Moved to higher bucket", 0)), lower=int(vc.get("Moved to lower bucket", 0)),
                   new=int(vc.get("New account", 0)), closed=int(vc.get("Closed / not in current file", 0)))

insights = make_insights(k, br, co, cur, top10_share, mom_sum)
recs = make_recs(k, br, co, cur, mom_sum)

st.markdown(f"<div class='banner'><h2>{company} · Credit Recovery Dashboard</h2>"
            f"<p>As of {as_of} · {len(df):,} of {len(cur_all):,} accounts shown · "
            f"Outstanding basis: {basis} · Period: {period_txt}</p></div>", unsafe_allow_html=True)

tabs = st.tabs(["📊 Dashboard", "📈 Month-over-Month", "🔁 Roll-Rate", "🎯 Collector Targets",
                "🧮 Provisioning", "🚨 Recovery Priority", "✅ Data Quality", "📄 PDF", "📑 Reports & AI"])
t_dash, t_mom, t_roll, t_tar, t_prov, t_pri, t_dq, t_pdf, t_rep = tabs

# ======================================================================
# DASHBOARD
# ======================================================================
with t_dash:
    c1 = st.columns(6)
    c1[0].metric("Total Outstanding", m_(k["outstanding"]))
    c1[1].metric("Total Overdue", m_(k["overdue"]), pct(k["overdue_pct"]) + " of outstanding", delta_color="inverse")
    c1[2].metric("Total Collection", m_(k["collected"]))
    c1[3].metric("Collection Efficiency", pct(k["efficiency"]))
    c1[4].metric("Total Accounts", f"{k['accounts']:,}")
    c1[5].metric("Accounts with Overdue", f"{k['od_accounts']:,}",
                 pct(div(k["od_accounts"], k["accounts"]), 0) + " of accounts", delta_color="inverse")
    c2 = st.columns(5)
    c2[0].metric("PAR1 (1+ EMI)", pct(k["par1"]))
    c2[1].metric("PAR2 (2+ EMI)", pct(k["par2"]))
    c2[2].metric("PAR3 (3+ EMI)", pct(k["par3"]), f"{k['par3_accts']:,} accounts", delta_color="inverse")
    c2[3].metric("Last-month Collection", m_(k["last_month"]), pct(k["lm_ratio"], 0) + " of EMI demand", delta_color="off")
    c2[4].metric("Interest", m_(k["interest"]))

    t = aging.reset_index()
    t["bucket"] = t.bucket.astype(str)
    a, b = st.columns(2)
    with a:
        plot(style_fig(px.bar(t, x="bucket", y="accounts", text_auto=True), "Due EMI distribution - accounts"))
    with b:
        plot(style_fig(px.bar(t, x="bucket", y=["outstanding", "overdue"], barmode="group", text_auto=".3s"),
                       "Outstanding and overdue by Due EMI bucket"))

    st.subheader("Branch-wise collection")
    a, b = st.columns(2)
    with a:
        plot(grouped_bar(br, ["total_due", "collected"], ["Total due", "Collected"], "Due vs collected (top 15 by overdue)"))
    with b:
        e = br[br.total_due > 0].sort_values("efficiency")
        plot(style_fig(px.bar(e.reset_index(), x="efficiency", y="branch", orientation="h", text_auto=".1%"),
                       "Collection efficiency by branch"))
    show_df(br)

    st.subheader("Collector-wise performance")
    a, b = st.columns(2)
    with a:
        e = co[co.total_due > 0].sort_values("efficiency").tail(25)
        plot(style_fig(px.bar(e.reset_index(), x="efficiency", y="collector", orientation="h", text_auto=".1%"),
                       "Collection efficiency by collector", 460))
    with b:
        plot(grouped_bar(co, ["overdue", "collected"], ["Overdue", "Collected"], "Overdue vs collected (top 15)"))
    show_df(co)

    st.subheader("Vehicle model-wise recovery")
    if vm is not None:
        a, b = st.columns(2)
        with a:
            plot(grouped_bar(vm, ["total_due", "collected"], ["Total due", "Collected"], "Due vs collected by model"))
        with b:
            plot(style_fig(px.bar(vm.reset_index(), x="vehicle_model", y="efficiency", text_auto=".1%"),
                           "Collection efficiency by model"))
        show_df(vm)
    else:
        st.info("Your file has no Vehicle Model column, so this analysis is unavailable. "
                "Map a model column under 'Column mapping' if one exists.")

    st.subheader("Risk status (overdue-wise)")
    st.caption("Management classification by Due EMI count: Regular = 0, Low = 1, Medium = 2, High = 3-5, Critical = 6+.")
    a, b = st.columns(2)
    rt = risk_t.reset_index()
    rt["risk"] = rt.risk.astype(str)
    with a:
        plot(style_fig(px.bar(rt, x="risk", y="overdue", color="risk", color_discrete_map=RISK_COLORS,
                              text_auto=".3s"), "Overdue amount by risk status"))
    with b:
        plot(style_fig(px.pie(rt, names="risk", values="outstanding", color="risk",
                              color_discrete_map=RISK_COLORS, hole=0.45), "Outstanding by risk status"))
    show_df(risk_t)

    st.subheader("Key insights")
    for i in insights:
        st.write("•", i)
    st.subheader("Recommended actions")
    for a_, t_ in recs:
        st.write(f"• {a_} *({t_})*")
    with st.expander("How the metrics are calculated"):
        st.markdown(f"""
- **Outstanding** = {basis}. **Overdue** = *Overdue amount* column (negative values excluded and reported).
- **Collection efficiency** = Total Collection ÷ Total due amount.
- **PARn** = outstanding of accounts with ≥ n Due EMIs ÷ total outstanding.
- **Buckets** use *Due EMI Number* ({rounding.lower()} for fractions): Current 0, 1, 2, 3, 4-5, 6+.
- **Monthly EMI demand** = sum of EMI for accounts with outstanding > 0.
- Blank / non-numeric cells are read as 0 and listed in the **Data Quality** tab. Source values are never edited.""")

# ======================================================================
# MONTH OVER MONTH
# ======================================================================
with t_mom:
    if prev_all is None:
        st.info("Upload the previous month Excel file in the sidebar to enable this tab.")
    else:
        st.caption("Uses the full uploaded files (sidebar filters not applied) so accounts can be matched on Account Id.")
        a = st.columns(5)
        a[0].metric("Matched accounts", f"{int((mv._merge == 'both').sum()):,}")
        a[1].metric("Moved to higher bucket", f"{mom_sum['higher']:,}")
        a[2].metric("Moved to lower bucket", f"{mom_sum['lower']:,}")
        a[3].metric("New accounts", f"{mom_sum['new']:,}")
        a[4].metric("Closed / missing", f"{mom_sum['closed']:,}")
        spec = [("Total outstanding", "outstanding", "m"), ("Total overdue", "overdue", "m"),
                ("Total collection (cumulative)", "collected", "m"), ("Last-month collection", "last_month", "m"),
                ("Collection efficiency", "efficiency", "p"), ("Accounts", "accounts", "i"),
                ("Accounts with overdue", "od_accounts", "i"), ("PAR3", "par3", "p")]
        rows = []
        for lbl, key, kind in spec:
            cv, pv_ = kc[key], kp[key]
            f = {"m": lambda x: m_(x), "p": lambda x: pct(x), "i": lambda x: f"{x:,.0f}"}[kind]
            if kind == "p":
                chg, chg_p = (f"{(cv - pv_) * 100:+.1f} pts" if pd.notna(cv) and pd.notna(pv_) else "n/a"), ""
            else:
                chg, chg_p = f"{cv - pv_:+,.0f}", pct(div(cv - pv_, abs(pv_)))
            rows.append({"Metric": lbl, "Previous": f(pv_), "Current": f(cv), "Change": chg, "Change %": chg_p})
        st.dataframe(pd.DataFrame(rows).set_index("Metric"), use_container_width=True)
        st.caption("Collection growth: compare 'Total collection (cumulative)' and 'Last-month collection' rows above.")

        mvc = mv.movement.value_counts().reset_index()
        mvc.columns = ["movement", "accounts"]
        plot(style_fig(px.bar(mvc, x="movement", y="accounts", text_auto=True), "Account movement between buckets"))

        st.subheader("Accounts that moved into a higher overdue bucket")
        hi = mv[mv.movement == "Moved to higher bucket"].sort_values("overdue_change", ascending=False)
        show_df(hi[["account", "customer", "branch", "collector", "bucket_prev", "bucket_cur",
                    "overdue_amt_prev", "overdue_amt_cur", "overdue_change"]]
                .rename(columns={"overdue_amt_prev": "overdue_prev", "overdue_amt_cur": "overdue_cur"}),
                fmt={"overdue_prev": "{:,.0f}", "overdue_cur": "{:,.0f}", "overdue_change": "{:+,.0f}"})

        for by, nm in (("branch", "Branch"), ("collector", "Collector")):
            st.subheader(f"{nm}-wise comparison")
            cg = compare_group(cur_all, prev_all, by)
            a, b = st.columns(2)
            with a:
                plot(grouped_bar(cg, ["overdue_prev", "overdue_cur"], ["Previous", "Current"], f"Overdue by {by.lower()}"))
            with b:
                plot(grouped_bar(cg, ["collected_prev", "collected_cur"], ["Previous", "Current"], f"Total collection by {by.lower()}"))
            show_df(cg, fmt={c_: "{:,.0f}" for c_ in cg.columns if c_.startswith(("outstanding", "overdue", "collected", "accounts"))}
                    | {"efficiency_cur": "{:.1%}", "efficiency_prev": "{:.1%}", "efficiency_change_pts": "{:+.1f}"})

# ======================================================================
# ROLL RATE
# ======================================================================
with t_roll:
    if roll is None:
        st.info("Roll-rate needs opening (previous) and closing (current) data. Upload the previous month Excel.")
    else:
        st.caption("Rows = opening bucket (previous file). Columns = closing bucket (current file). "
                   "Accounts matched on Account Id. Roll-rate % = share of opening accounts in each closing bucket.")
        rate = roll["rate"] * 100
        fig = px.imshow(rate, text_auto=".1f", aspect="auto", color_continuous_scale="Blues",
                        labels=dict(x="Closing bucket", y="Opening bucket", color="%"))
        plot(style_fig(fig, "Roll-rate matrix (% of opening accounts)", 430))
        st.subheader("Summary by opening bucket")
        show_df(roll["summary"], fmt={"Roll-forward rate": "{:.1%}"})
        a, b = st.columns(2)
        with a:
            st.markdown("**Account counts**")
            st.dataframe(roll["cnt"], use_container_width=True)
        with b:
            st.markdown("**Outstanding amount** (opening balance; closing balance for new accounts)")
            show_df(roll["amt"], fmt={c_: "{:,.0f}" for c_ in roll["amt"].columns})

# ======================================================================
# COLLECTOR TARGETS
# ======================================================================
targets_result = None
with t_tar:
    st.caption("Targets apply to the collectors currently selected in the sidebar filters.")
    bopts = ["Last month collection (column)", "Total Collection (cumulative)"]
    if mv is not None:
        bopts.append("Increase in Total Collection vs previous file (matched accounts)")
    abasis = st.radio("Actual collection =", bopts)
    up = st.file_uploader("Upload targets (Excel with a collector-name column and a 'Target' column)",
                          type=["xlsx", "xls"], key="tgt_up")
    colls = sorted(df.collector.unique())
    base = pd.DataFrame({"Collector": colls, "Target": 0.0})
    if up:
        try:
            tu = read_targets(up)
            low = base.Collector.str.strip().str.lower()
            base["Target"] = low.map(tu).fillna(0.0)
            unmatched = [n for n in tu.index if n not in set(low)]
            if unmatched:
                st.warning("Target file names not found in the data: " + ", ".join(unmatched[:15]))
        except Exception as e:
            st.error(f"Could not read targets: {e}")
    edited = st.data_editor(base, disabled=["Collector"], hide_index=True, use_container_width=True,
                            key=f"tgt_{up.name if up else 'manual'}_{len(base)}",
                            column_config={"Target": st.column_config.NumberColumn(min_value=0, format="%.0f")})
    if abasis.startswith("Last"):
        actual = df.groupby("collector").last_month_coll.sum()
    elif abasis.startswith("Total"):
        actual = df.groupby("collector").total_coll.sum()
    else:
        actual = mv[mv._merge == "both"].groupby("collector").collection_in_period.sum()
    res = edited.set_index("Collector")
    res["Actual"] = actual.reindex(res.index).fillna(0)
    res["Achievement %"] = np.where(res.Target > 0, res.Actual / res.Target.replace(0, np.nan), np.nan)
    res["Shortfall"] = np.where(res.Target > 0, (res.Target - res.Actual).clip(lower=0), 0)
    res["Rank"] = res["Achievement %"].rank(ascending=False, method="min")
    res = res.sort_values("Rank")
    if (res.Target > 0).any():
        targets_result = res.reset_index()
        tt, aa = res.Target.sum(), res.Actual[res.Target > 0].sum()
        x = st.columns(4)
        x[0].metric("Total target", m_(tt))
        x[1].metric("Actual (collectors with target)", m_(aa))
        x[2].metric("Achievement", pct(div(aa, tt)))
        x[3].metric("Total shortfall", m_(res.Shortfall.sum()))
        show_df(res, fmt={"Target": "{:,.0f}", "Actual": "{:,.0f}", "Achievement %": "{:.1%}",
                          "Shortfall": "{:,.0f}", "Rank": "{:.0f}"})
        r2 = res[res.Target > 0].reset_index()
        plot(grouped_bar(r2.set_index("Collector"), ["Target", "Actual"], ["Target", "Actual"], "Target vs actual", 25))
    else:
        st.info("Enter targets in the table (or upload an Excel) to see achievement, shortfall and ranking.")

# ======================================================================
# PROVISIONING
# ======================================================================
with t_prov:
    st.warning("**Management estimate only - not an official accounting or regulatory provision.** "
               "Provision = outstanding balance of accounts in each bucket × rate. No deduction is made for "
               "vehicle (collateral) value. Rates are editable assumptions.")
    rate_df = pd.DataFrame({"Bucket group": PROV_GROUPS, "Rate %": PROV_DEFAULT})
    ed = st.data_editor(rate_df, disabled=["Bucket group"], hide_index=True, key="prov_rates",
                        column_config={"Rate %": st.column_config.NumberColumn(min_value=0.0, max_value=100.0,
                                                                                step=1.0, format="%.1f")})
    rates = dict(zip(ed["Bucket group"], ed["Rate %"].fillna(0.0)))
    prov = provision_table(df, rates)
    x = st.columns(3)
    x[0].metric("Estimated provision", m_(prov.provision.sum()))
    x[1].metric("As % of total outstanding", pct(div(prov.provision.sum(), k["outstanding"])))
    x[2].metric("Outstanding in provisioned buckets", m_(prov.outstanding.sum()))
    show_df(prov)
    plot(style_fig(px.bar(prov.reset_index(), x="prov_group", y="provision", text_auto=".3s"),
                   "Estimated provision by bucket"))
    dfp = df.assign(prov_amt=df.prov_group.astype(object).map(rates).fillna(0) * df.outstanding / 100)
    st.subheader("Estimated provision by branch")
    show_df(dfp.groupby("branch").agg(accounts=("account", "size"), outstanding=("outstanding", "sum"),
                                      provision=("prov_amt", "sum")).sort_values("provision", ascending=False))

# ======================================================================
# RECOVERY PRIORITY
# ======================================================================
with t_pri:
    st.caption("Transparent score from 0-100. Adjust the weights (they are normalised to 100%).")
    wc = st.columns(5)
    w_raw = {"overdue": wc[0].slider("Overdue amount", 0, 100, 40),
             "due_emi": wc[1].slider("Due EMI count", 0, 100, 30),
             "down_payment": wc[2].slider("Low down payment", 0, 100, 15),
             "outstanding": wc[3].slider("Outstanding balance", 0, 100, 15)}
    dp_cap = wc[4].slider("Down payment treated as safe at", 10, 50, 30, format="%d%%") / 100
    tot_w = sum(w_raw.values())
    weights = {k_: (v / tot_w if tot_w else 0.25) for k_, v in w_raw.items()}
    top_n = st.number_input("Accounts to list", 10, 1000, 100, step=10)
    prio = priority(df, weights, dp_cap, cur, int(top_n))
    with st.expander("How the score works"):
        st.markdown(f"""
Score = weighted average of four components (each 0-100), over overdue accounts only:
1. **Overdue amount** - percentile rank among overdue accounts.
2. **Due EMI count** - Due EMIs ÷ 6 (capped at 100).
3. **Down payment** - 100 when down payment is 0% of sales price, falling to 0 at {dp_cap:.0%} or more.
   If sales price / down payment is unavailable the component is **left out** and the other weights are rescaled.
4. **Outstanding balance** - percentile rank among overdue accounts.

The score is a ranking aid for follow-up order, not a probability of default.""")
    if len(prio):
        show_df(prio[["rank", "account", "customer", "vehicle", "branch", "collector", "due_emi_no",
                      "overdue_amt", "outstanding", "dp_pct", "score", "main_driver", "explanation"]],
                fmt={"dp_pct": "{:.0%}", "due_emi_no": "{:g}", "outstanding": "{:,.0f}"}, hide_index=True)
    else:
        st.info("No overdue accounts in the current selection.")

# ======================================================================
# DATA QUALITY
# ======================================================================
with t_dq:
    st.write(f"Rows read: **{Q['rows_in']:,}** · rows kept: **{Q['rows_out']:,}** · "
             f"removed: **{Q['rows_in'] - Q['rows_out']:,}**")
    if not Q["notes"]:
        st.success("No data-quality issues detected.")
    for lvl, msg in Q["notes"]:
        {"error": st.error, "warning": st.warning}.get(lvl, st.info)(msg)
    for name, d in Q["exceptions"].items():
        with st.expander(f"{name} ({len(d)})"):
            show_df(d, fmt={c_: "{:,.2f}" for c_ in ("outstanding", "total_due", "total_coll", "overdue_raw", "due_raw")})
    if len(Q["removed"]):
        with st.expander(f"Removed rows ({len(Q['removed'])})"):
            st.dataframe(Q["removed"].astype(str), use_container_width=True)
    if prev_all is not None:
        with st.expander("Previous file - quality notes"):
            for lvl, msg in Qp["notes"]:
                st.write(f"{lvl.upper()}: {msg}")

# ======================================================================
# PDF
# ======================================================================
with t_pdf:
    if f_pdf:
        text, tables = read_pdf(f_pdf.getvalue())
        st.text_area("Extracted text", text, height=240)
        kw = st.text_input("Search keyword (e.g. repossession, provision, legal)")
        if kw:
            for line in [l for l in text.splitlines() if kw.lower() in l.lower()][:30]:
                st.write("•", line)
        for i, tb in enumerate(tables):
            st.write(f"Table {i + 1}")
            st.dataframe(tb)
        st.session_state["pdf_notes"] = st.text_area("Notes to include in the presentation (speaker notes, slide 10)",
                                                     value=text[:800], height=140)
    else:
        st.info("Upload a PDF in the sidebar (optional). Scanned PDFs need OCR and are not supported.")

# ======================================================================
# REPORTS & AI
# ======================================================================
with t_rep:
    st.subheader("AI management commentary (Gemini)")
    gkey = get_secret("GEMINI_API_KEY")
    if not gkey:
        st.info("Add `GEMINI_API_KEY` to Streamlit secrets or an environment variable to enable this section.")
    else:
        model = st.text_input("Gemini model", get_secret("GEMINI_MODEL", "gemini-2.5-flash"))
        incl = st.checkbox("Send collector names to Gemini (otherwise anonymised as Collector 1, 2…)", False)
        st.caption("Only aggregated figures are sent. Customer names, account IDs and vehicle numbers never leave "
                   "this app. Free-tier prompts may be used by Google to improve its products - do not enable if "
                   "your policy forbids sharing portfolio figures externally.")
        if st.button("Generate commentary"):
            cot = co.head(8).copy()
            if not incl:
                cot.index = [f"Collector {i + 1}" for i in range(len(cot))]
            cols_ = ["accounts", "outstanding", "overdue", "efficiency", "overdue_pct"]
            payload = {"currency": cur or "n/a", "as_of": str(as_of), "kpis": k,
                       "buckets": aging[cols_].reset_index().astype({"bucket": str}).to_dict("records"),
                       "branches_top8": br.head(8)[cols_].reset_index().to_dict("records"),
                       "collectors_top8": cot[cols_].reset_index().to_dict("records"),
                       "risk_status": risk_t[cols_].reset_index().astype({"risk": str}).to_dict("records"),
                       "provision_estimate_total": float(provision_table(df, rates).provision.sum()),
                       "month_over_month": mom_sum or "previous file not uploaded",
                       "roll_forward_rate_by_bucket": (roll["summary"]["Roll-forward rate"].to_dict()
                                                       if roll else "not available"),
                       "top_priority_accounts_count": int(len(prio))}
            try:
                with st.spinner("Asking Gemini…"):
                    st.session_state["ai_text"] = gemini_commentary(payload, gkey, model)
            except requests.HTTPError as e:
                st.error(f"Gemini request failed (HTTP {e.response.status_code}). Check the key, model name and quota.")
            except Exception as e:
                st.error(f"Gemini request failed: {type(e).__name__}")
    ai_text = st.session_state.get("ai_text", "")
    if ai_text:
        st.markdown(ai_text)
        st.caption("AI-generated draft - review against the figures before presenting.")

    st.divider()
    st.subheader("Exports")
    st.caption("Exports reflect the current sidebar filters. Click generate again after changing filters.")
    ctx = dict(company=company, brand=brand, as_of=str(as_of), cur=cur, k=k, aging=aging, br=br, co=co, vm=vm,
               roll=roll, prov=prov, prio=prio, targets=targets_result, insights=insights, recs=recs,
               mom=mom_sum, ai_text=ai_text, pdf_notes=st.session_state.get("pdf_notes", ""), weights=weights)
    a, b = st.columns(2)
    if a.button("Build PowerPoint (10 slides)"):
        try:
            st.session_state["ppt"] = build_pptx(ctx)
        except Exception as e:
            st.error(f"PowerPoint build failed: {type(e).__name__}: {e}")
    if "ppt" in st.session_state:
        a.download_button("⬇️ Download PPTX", st.session_state["ppt"], "credit_recovery_review.pptx",
                          "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    if b.button("Build Excel report"):
        kdf = pd.DataFrame([("Total outstanding", k["outstanding"]), ("Total overdue", k["overdue"]),
                            ("Total collection", k["collected"]), ("Collection efficiency", k["efficiency"]),
                            ("PAR1", k["par1"]), ("PAR2", k["par2"]), ("PAR3", k["par3"]),
                            ("Accounts", k["accounts"]), ("Accounts with overdue", k["od_accounts"]),
                            ("Outstanding basis", basis), ("Fractional Due EMI", rounding),
                            ("Provision note", "Management estimate - not an official accounting provision")],
                           columns=["Metric", "Value"])
        sheets = {"Summary": (kdf, False), "By Bucket": (aging, True), "By Branch": (br, True),
                  "By Collector": (co, True), "By Sales Person": (sp, True), "By Risk Status": (risk_t, True),
                  "Provisioning (estimate)": (prov.reset_index(), False),
                  "Priority List": (prio.drop(columns=["p_overdue", "p_outstanding"], errors="ignore")
                                    if len(prio) else pd.DataFrame(), False),
                  "Data Quality": (pd.DataFrame(Q["notes"], columns=["Level", "Message"]), False)}
        if vm is not None:
            sheets["By Vehicle Model"] = (vm, True)
        if targets_result is not None:
            sheets["Collector Targets"] = (targets_result, False)
        if roll is not None:
            sheets["Roll-Rate %"] = (roll["rate"], True)
            sheets["Roll-Rate Counts"] = (roll["cnt"], True)
            sheets["Roll-Rate Summary"] = (roll["summary"], True)
            sheets["MoM Accounts"] = (mv.drop(columns=["_merge"]), False)
        sheets["Processed Data"] = (df, False)
        try:
            st.session_state["xlsx"] = build_excel(sheets)
        except Exception as e:
            st.error(f"Excel build failed: {type(e).__name__}: {e}")
    if "xlsx" in st.session_state:
        b.download_button("⬇️ Download Excel report", st.session_state["xlsx"], "credit_recovery_report.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
