```python
"""AI Business Analytics, Messy Data Auto-Cleaner & Credit Recovery Studio.
Run locally: streamlit run app.py
"""
import hmac
import io
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

warnings.filterwarnings("ignore")
st.set_page_config(
    page_title="AI Analytics & Credit Intelligence Studio",
    page_icon="🚀",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    .hero-box {
        padding: 1.3rem 1.6rem; border-radius: 12px; color: white;
        margin-bottom: 1rem; box-shadow: 0 4px 14px rgba(0,0,0,0.08);
    }
    .kpi-card {
        background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;
        padding: 0.9rem 1rem; box-shadow: 0 1px 3px rgba(0,0,0,0.03);
        margin-bottom: 0.5rem;
    }
    .kpi-title {font-size: 0.78rem; color: #64748B; text-transform: uppercase; font-weight: 600;}
    .kpi-val {font-size: 1.35rem; color: #0F172A; font-weight: 700; margin-top: 0.2rem;}
    .kpi-sub {font-size: 0.75rem; color: #475569; margin-top: 0.15rem;}
    .insight-box {
        background: #F8FAFC; border-left: 4px solid #2563EB;
        padding: 0.85rem 1rem; border-radius: 6px; margin-bottom: 0.6rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def get_secret(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.environ.get(key, default)


def check_password() -> bool:
    pw = get_secret("APP_PASSWORD", "").strip()
    if not pw:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.sidebar.markdown("### 🔒 Access Verification")
    entered = st.sidebar.text_input("Password", type="password")
    if entered:
        if hmac.compare_digest(entered.strip(), pw):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.sidebar.error("Incorrect password.")
    st.info("Enter the access password in the left sidebar to unlock the studio.")
    return False


if not check_password():
    st.stop()


def hex_to_rgb(h: str):
    h = h.lstrip("#")
    if len(h) != 6:
        return (31, 58, 95)
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def fmt_num(v, cur: str = "", decimals: int = 0) -> str:
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return "-"
    prefix = f"{cur} " if cur else ""
    if decimals == 0:
        return f"{prefix}{v:,.0f}"
    return f"{prefix}{v:,.{decimals}f}"


def show_table(df: pd.DataFrame, safe_mode: bool = False, max_rows: int = 100):
    if df is None or df.empty:
        st.caption("No records to display.")
        return
    view = df.head(max_rows)
    if safe_mode:
        st.markdown(
            f'<div style="overflow-x:auto; max-height:420px; border:1px solid #e2e8f0; border-radius:6px;">'
            f'{view.to_html(index=False, classes="table", border=0)}</div>',
            unsafe_allow_html=True,
        )
    else:
        try:
            st.dataframe(view, use_container_width=True)
        except Exception:
            st.markdown(view.to_html(index=False), unsafe_allow_html=True)


def make_sample_credit_df() -> pd.DataFrame:
    rng = np.random.default_rng(42)
    branches = ["Dhaka North", "Chattogram", "Bogura", "Rangpur", "Jashore", "Rajshahi"]
    collectors = ["Azmir", "Aminul", "Zahid", "Mannan", "Hasen", "Zamil", "Akter"]
    sales_p = ["Rahim", "Karim", "Tariq", "Noman", "Sajid"]
    models = ["Cargo Truck 3T", "Pickup 1.5T", "Mini Covered Van", "Passenger Bus", "Prime Mover"]
    rows = []
    for i in range(1, 85):
        sp = int(rng.choice([1250000, 1650000, 2100000, 2850000, 3400000]))
        dp_pct = float(rng.choice([0.0, 0.08, 0.12, 0.15, 0.20, 0.25, 0.30]))
        dp = round(sp * dp_pct)
        interest = round((sp - dp) * 0.18)
        bill = sp + interest
        no_emi = int(rng.choice([24, 36, 48, 60]))
        emi = round((bill - dp) / no_emi)
        inst_due = int(rng.integers(4, no_emi))
        due_emi = int(rng.choice([0, 0, 0, 1, 1, 2, 3, 4, 6, 8, 12, 15]))
        due_emi = min(due_emi, inst_due)
        coll_no = max(0, inst_due - due_emi)
        tot_due = inst_due * emi
        tot_coll = coll_no * emi
        overdue = max(0, tot_due - tot_coll)
        last_m = emi if due_emi <= 1 else int(rng.choice([0, 0, round(emi * 0.5)]))
        net_bal = max(0, bill - dp - tot_coll)
        rows.append(
            {
                "Serial Number": i,
                "Customer Name": f"Customer {i:03d}",
                "Branch Name": rng.choice(branches),
                "Sales Personnel": rng.choice(sales_p),
                "Account Id": f"ACC-{340480000 + i}",
                "Collection Personnel": rng.choice(collectors),
                "Vehicle registration number": f"DM-NA-{rng.integers(11,25)}-{1000+i}",
                "Vehicle Model": rng.choice(models),
                "Sales price": sp,
                "Down Payment": dp,
                "Interest": interest,
                "Bill amount": bill,
                "Balance receivable": net_bal,
                "No of EMI": no_emi,
                "EMI Date": f"2025-{int(rng.integers(1,13)):02d}-15",
                "EMI": emi,
                "Installment due number": inst_due,
                "EMI Collection number": coll_no,
                "Due EMI Number": due_emi,
                "Total due amount": tot_due,
                "Total Collection": tot_coll,
                "Overdue amount": overdue,
                "Last month collection": last_m,
                "Net balance": net_bal,
            }
        )
    return pd.DataFrame(rows)


def make_sample_messy_business_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Order_ID": "ORD-1001", "Order Date": "2025-01-05", "Region": "  dhaka ", "Category": "Electronics", "Sales_Rep": "Tanvir", "Revenue": "BDT 125,000/-", "Cost": "92,000", "Profit": "33,000", "Status": "Completed"},
            {"Order_ID": "ORD-1002", "Order Date": "2025/01/12", "Region": "Chattogram", "Category": "Furniture", "Sales_Rep": "Nadia", "Revenue": "84,500", "Cost": "61,000", "Profit": "23,500", "Status": "Completed"},
            {"Order_ID": "ORD-1002", "Order Date": "2025/01/12", "Region": "Chattogram", "Category": "Furniture", "Sales_Rep": "Nadia", "Revenue": "84,500", "Cost": "61,000", "Profit": "23,500", "Status": "Completed"},
            {"Order_ID": "ORD-1003", "Order Date": "18-01-2025", "Region": "DHAKA", "Category": "Electronics", "Sales_Rep": "Tanvir", "Revenue": "$210,000", "Cost": "145,000", "Profit": "65,000", "Status": "Pending"},
            {"Order_ID": "ORD-1004", "Order Date": "2025-02-02", "Region": "Sylhet", "Category": "Industrial", "Sales_Rep": "Farhan", "Revenue": "340,000", "Cost": "230,000", "Profit": "110,000", "Status": "Completed"},
            {"Order_ID": "ORD-1005", "Order Date": "2025-02-14", "Region": "Rajshahi", "Category": "Machinery", "Sales_Rep": "Mehedi", "Revenue": "520,000", "Cost": "390,000", "Profit": "130,000", "Status": "Completed"},
            {"Order_ID": "ORD-1006", "Order Date": "2025-02-21", "Region": "  chattogram", "Category": "Industrial", "Sales_Rep": "Nadia", "Revenue": "175,000", "Cost": "132,000", "Profit": "43,000", "Status": "Overdue"},
            {"Order_ID": "ORD-1007", "Order Date": "2025-03-04", "Region": "Dhaka", "Category": "Machinery", "Sales_Rep": "Tanvir", "Revenue": "890,000", "Cost": "610,000", "Profit": "280,000", "Status": "Completed"},
            {"Order_ID": "ORD-1008", "Order Date": "2025-03-15", "Region": "Khulna", "Category": "Furniture", "Sales_Rep": "Sadia", "Revenue": "64,000", "Cost": "49,000", "Profit": "15,000", "Status": "Pending"},
            {"Order_ID": "Grand Total", "Order Date": "", "Region": "Total", "Category": "", "Sales_Rep": "", "Revenue": "2,493,000", "Cost": "1,770,000", "Profit": "723,000", "Status": ""},
        ]
    )


TOTAL_PAT = re.compile(r"^\s*(grand\s*total|sub\s*total|total|sum|সর্বমোট|মোট)\b", re.I)


def detect_best_header_row(raw_df: pd.DataFrame) -> int:
    best_row, best_score = 0, -1e9
    limit = min(35, len(raw_df))
    for r in range(limit):
        row = raw_df.iloc[r]
        non_null = row.dropna()
        if len(non_null) < 2:
            continue
        str_cells = [str(x).strip() for x in non_null if str(x).strip() != ""]
        if len(str_cells) < 2:
            continue
        alpha_cnt = sum(1 for s in str_cells if any(c.isalpha() for c in s))
        uniq_ratio = len(set(s.lower() for s in str_cells)) / max(1, len(str_cells))
        score = alpha_cnt * 3 + uniq_ratio * 5 - r * 0.15
        if score > best_score:
            best_score, best_row = score, r
    return best_row


def parse_messy_number_series(s: pd.Series):
    if pd.api.types.is_numeric_dtype(s):
        return s, 0
    non_na = s.dropna().astype(str).str.strip()
    non_na = non_na[non_na != ""]
    if len(non_na) == 0:
        return s, 0
    cleaned = non_na.copy()
    cleaned = cleaned.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    cleaned = cleaned.str.replace(
        r"(?i)(bdt|tk|taka|usd|eur|inr|rs|৳|\$|€|£|/-|,|\s)", "", regex=True
    )
    cleaned = cleaned.str.replace(r"%$", "", regex=True)
    num = pd.to_numeric(cleaned, errors="coerce")
    success_rate = num.notna().mean()
    if success_rate >= 0.75:
        full_clean = s.astype(str).str.strip()
        full_clean = full_clean.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
        full_clean = full_clean.str.replace(
            r"(?i)(bdt|tk|taka|usd|eur|inr|rs|৳|\$|€|£|/-|,|\s)", "", regex=True
        )
        full_clean = full_clean.str.replace(r"%$", "", regex=True)
        out = pd.to_numeric(full_clean, errors="coerce")
        fixed_cnt = int((out.notna() & ~s.apply(lambda x: isinstance(x, (int, float)))).sum())
        return out, fixed_cnt
    return s, 0


def auto_clean_dataframe(
    df_in: pd.DataFrame,
    drop_totals: bool = True,
    drop_dups: bool = True,
    standardize_text: bool = True,
    missing_num_policy: str = "Keep as NaN (transparent)",
    missing_txt_policy: str = "Fill with 'Unknown'",
):
    df = df_in.copy()
    audit = []

    new_cols = []
    seen = {}
    for i, c in enumerate(df.columns):
        name = re.sub(r"\s+", " ", str(c).strip())
        if not name or name.lower().startswith("unnamed"):
            name = f"Column_{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        new_cols.append(name)
    df.columns = new_cols

    r0, c0 = df.shape
    df = df.replace(r"^\s*$", np.nan, regex=True)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    if r0 - len(df) > 0:
        audit.append(f"Removed {r0 - len(df)} completely blank row(s).")
    if c0 - df.shape[1] > 0:
        audit.append(f"Removed {c0 - df.shape[1]} completely empty column(s).")

    if drop_totals and len(df) > 0:
        check_cols = df.columns[: min(5, len(df.columns))]
        mask_total = pd.Series(False, index=df.index)
        for col in check_cols:
            mask_total = mask_total | df[col].astype(str).str.match(TOTAL_PAT, na=False)
        tot_cnt = int(mask_total.sum())
        if tot_cnt > 0:
            df = df.loc[~mask_total].copy()
            audit.append(f"Removed {tot_cnt} embedded Subtotal / Grand Total row(s).")

    txt_fixed = 0
    for col in df.columns:
        if df[col].dtype == object:
            orig = df[col].copy()
            s = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
            s = s.replace({"nan": np.nan, "None": np.nan, "NULL": np.nan, "null": np.nan, "": np.nan})
            non_na = s.dropna()
            if standardize_text and 0 < non_na.nunique() <= 40:
                lower_nunique = non_na.str.lower().nunique()
                if lower_nunique < non_na.nunique():
                    s = s.str.title()
            changed = int((orig.fillna("") != s.fillna("")).sum())
            txt_fixed += changed
            df[col] = s
    if txt_fixed > 0:
        audit.append(f"Trimmed extra spaces and standardized text in {txt_fixed} cell(s).")

    num_cols_converted = []
    for col in df.columns:
        col_low = col.lower()
        if any(k in col_low for k in ["id", "serial", "account", "phone", "mobile", "reg", "vehicle", "code"]):
            continue
        converted, cnt = parse_messy_number_series(df[col])
        if cnt > 0:
            df[col] = converted
            num_cols_converted.append(f"{col} ({cnt} cells)")
    if num_cols_converted:
        audit.append("Converted messy number/currency text to numeric: " + ", ".join(num_cols_converted))

    if drop_dups and len(df) > 0:
        before = len(df)
        df = df.drop_duplicates().copy()
        if before - len(df) > 0:
            audit.append(f"Removed {before - len(df)} exact duplicate row(s).")

    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    txt_cols = [c for c in df.columns if c not in num_cols]

    if missing_num_policy == "Fill with 0":
        na_cnt = int(df[num_cols].isna().sum().sum()) if num_cols else 0
        if na_cnt > 0:
            df[num_cols] = df[num_cols].fillna(0)
            audit.append(f"Filled {na_cnt} missing numeric cell(s) with 0 (per user setting).")
    elif missing_num_policy == "Fill with Column Median":
        for c in num_cols:
            na_c = int(df[c].isna().sum())
            if na_c > 0 and df[c].notna().any():
                df[c] = df[c].fillna(df[c].median())
                audit.append(f"Filled {na_c} missing value(s) in '{c}' with median ({df[c].median():,.2f}).")

    if missing_txt_policy == "Fill with 'Unknown'":
        for c in txt_cols:
            na_c = int(df[c].isna().sum())
            if na_c > 0:
                df[c] = df[c].fillna("Unknown")
        audit.append("Filled blank text/category cells with 'Unknown'.")

    if not audit:
        audit.append("Dataset was already clean; no structural fixes were needed.")

    return df.reset_index(drop=True), audit


CREDIT_FIELDS = [
    ("serial", "Serial Number", ["serial number", "sl", "sl no", "sl.", "s/n", "serial"]),
    ("customer", "Customer Name", ["customer name", "customer", "client name", "party name", "borrower"]),
    ("branch", "Branch Name", ["branch name", "branch", "zone", "office", "location", "showroom"]),
    ("sales_person", "Sales Personnel", ["sales personnel", "sales person", "sales officer", "executive", "rm"]),
    ("account_id", "Account Id", ["account id", "account no", "acc id", "loan id", "agreement no", "code", "id"]),
    ("collector", "Collection Personnel", ["collection personnel", "collector", "recovery officer", "collection officer"]),
    ("vehicle_reg", "Vehicle registration number", ["vehicle registration number", "registration no", "reg no", "vehicle no", "chassis"]),
    ("vehicle_model", "Vehicle Model (optional)", ["vehicle model", "model", "product", "asset model"]),
    ("sales_price", "Sales price", ["sales price", "sale price", "vehicle price", "cash price", "invoice price"]),
    ("down_payment", "Down Payment", ["down payment", "dp", "advance", "down pay"]),
    ("interest", "Interest", ["interest", "interest amount", "markup", "profit"]),
    ("bill_amount", "Bill amount", ["bill amount", "total bill", "contract value", "hire price"]),
    ("balance_rec", "Balance receivable", ["balance receivable", "receivable", "bal receivable", "Closing Balance"]),
    ("no_of_emi", "No of EMI", ["no of emi", "total emi", "tenure", "installments", "no. of emi"]),
    ("emi_date", "EMI Date", ["emi date", "installment date", "due date", "disbursement date"]),
    ("emi", "EMI", ["emi", "emi amount", "installment amount", "monthly emi"]),
    ("inst_due_no", "Installment due number", ["installment due number", "inst due no", "due installment no", "installment due no"]),
    ("emi_coll_no", "EMI Collection number", ["emi collection number", "emi coll no", "collected emi no", "paid emi"]),
    ("due_emi_no", "Due EMI Number", ["due emi number", "due emi no", "due emi", "overdue emi", "arrear emi"]),
    ("total_due", "Total due amount", ["total due amount", "total due amt/", "total due amt", "total due", "cumulative due"]),
    ("total_coll", "Total Collection", ["total collection", "total coll", "cumulative collection", "collected amount"]),
    ("overdue_amt", "Overdue amount", ["overdue amount", "over due amt", "overdue amt", "overdue", "arrear amount"]),
    ("last_m_coll", "Last month collection", ["last month collection", "last month coll", "lm collection", "current month collection"]),
    ("net_balance", "Net balance", ["net balance", "net bal", "outstanding", "outstanding balance", "principal outstanding"]),
]


def auto_map_credit_columns(cols):
    norm_map = {re.sub(r"[^a-z0-9]+", " ", str(c).lower()).strip(): c for c in cols}
    mapping = {}
    used = set()
    for key, _, aliases in CREDIT_FIELDS:
        matched = None
        for al in aliases:
            al_norm = re.sub(r"[^a-z0-9]+", " ", al.lower()).strip()
            if al_norm in norm_map and norm_map[al_norm] not in used:
                matched = norm_map[al_norm]
                break
        if not matched:
            for al in aliases:
                al_norm = re.sub(r"[^a-z0-9]+", " ", al.lower()).strip()
                for cn, orig_c in norm_map.items():
                    if orig_c not in used and (al_norm in cn or cn in al_norm) and len(al_norm) >= 3:
                        if "remark" in cn or "note" in cn or "comment" in cn:
                            continue
                        matched = orig_c
                        break
                if matched:
                    break
        mapping[key] = matched
        if matched:
            used.add(matched)
    return mapping


def to_clean_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    cleaned = (
        s.astype(str)
        .str.strip()
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
        .str.replace(r"(?i)(bdt|tk|usd|৳|\$|,|\s|/-)", "", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def prepare_credit_data(raw: pd.DataFrame, mapping: dict, out_basis: str):
    df = pd.DataFrame(index=raw.index)
    notes = []

    mask_tot = pd.Series(False, index=raw.index)
    for c in raw.columns[: min(5, len(raw.columns))]:
        mask_tot = mask_tot | raw[c].astype(str).str.match(TOTAL_PAT, na=False)
    if mask_tot.any():
        notes.append(("Info", f"Excluded {int(mask_tot.sum())} total/subtotal row(s)."))
        raw = raw.loc[~mask_tot].copy()
        df = pd.DataFrame(index=raw.index)

    txt_keys = {"serial", "customer", "branch", "sales_person", "account_id", "collector", "vehicle_reg", "vehicle_model", "emi_date"}
    for key, label, _ in CREDIT_FIELDS:
        col = mapping.get(key)
        if col and col in raw.columns:
            if key in txt_keys:
                df[key] = raw[col].astype(str).str.strip().replace({"nan": "Unknown", "None": "Unknown", "": "Unknown"})
            else:
                df[key] = to_clean_num(raw[col])
        else:
            df[key] = "Unknown" if key in txt_keys else 0.0

    if mapping.get("overdue_amt") is None and (mapping.get("total_due") or mapping.get("total_coll")):
        df["overdue_amt"] = (df["total_due"] - df["total_coll"]).clip(lower=0.0)
        notes.append(("Warning", "Overdue amount derived as Total Due - Total Collection."))
    else:
        df["overdue_amt"] = df["overdue_amt"].clip(lower=0.0)

    if mapping.get("due_emi_no") is None:
        if mapping.get("inst_due_no") and mapping.get("emi_coll_no"):
            df["due_emi_no"] = (df["inst_due_no"] - df["emi_coll_no"]).clip(lower=0.0)
            notes.append(("Warning", "Due EMI Number derived as Installment Due - EMI Collection Number."))
        elif mapping.get("emi"):
            df["due_emi_no"] = np.where(df["emi"] > 0, df["overdue_amt"] / df["emi"], 0.0)
            notes.append(("Warning", "Due EMI Number estimated as Overdue Amount / EMI."))

    df["outstanding"] = df["net_balance"] if out_basis == "Net balance" else df["balance_rec"]
    if (df["outstanding"] == 0).all() and (df["net_balance"] > 0).any():
        df["outstanding"] = df["net_balance"]
    elif (df["outstanding"] == 0).all() and (df["balance_rec"] > 0).any():
        df["outstanding"] = df["balance_rec"]

    emi_int = np.ceil(df["due_emi_no"].clip(lower=0)).astype(int)
    emi_int = np.where((emi_int == 0) & (df["overdue_amt"] > 0), 1, emi_int)

    df["bucket"] = np.select(
        [emi_int <= 0, emi_int == 1, emi_int == 2, emi_int == 3, emi_int <= 5],
        ["Current (0 EMI)", "1 EMI", "2 EMI", "3 EMI", "4-5 EMI"],
        default="6+ EMI",
    )
    df["risk_status"] = np.select(
        [emi_int <= 0, emi_int == 1, emi_int == 2, emi_int <= 5],
        ["Regular", "Low Risk", "Medium Risk", "High Risk"],
        default="Critical Risk",
    )
    df["dp_pct"] = np.where(df["sales_price"] > 0, (df["down_payment"] / df["sales_price"]).clip(0, 1), 0.0)
    return df.reset_index(drop=True), notes


def score_recovery_priority(df: pd.DataFrame) -> pd.DataFrame:
    od = df[df["overdue_amt"] > 0].copy()
    if od.empty:
        return pd.DataFrame()
    od["p_overdue"] = od["overdue_amt"].rank(pct=True) * 100
    od["p_out"] = od["outstanding"].rank(pct=True) * 100
    od["p_emi"] = (od["due_emi_no"].clip(0, 6) / 6.0) * 100
    od["p_dp"] = (1.0 - (od["dp_pct"].clip(0, 0.30) / 0.30)) * 100
    od["score"] = (0.40 * od["p_overdue"] + 0.30 * od["p_emi"] + 0.15 * od["p_dp"] + 0.15 * od["p_out"]).round(1)
    od = od.sort_values("score", ascending=False).reset_index(drop=True)
    od["rank"] = np.arange(1, len(od) + 1)
    od["main_driver"] = np.where(
        od["p_overdue"] >= od["p_emi"], "Overdue amount", "High Due EMI count"
    )
    od["explanation"] = od.apply(
        lambda r: f"Overdue {r['overdue_amt']:,.0f}; {r['due_emi_no']:.0f} EMI due; DP {r['dp_pct']*100:.0f}%",
        axis=1,
    )
    return od


def generate_ai_insights(summary_text: str) -> str:
    api_key = get_secret("GEMINI_API_KEY", "").strip()
    model = get_secret("GEMINI_MODEL", "gemini-2.5-flash").strip()
    if not api_key:
        return (
            "### 🤖 Automated Executive Briefing (Built-in Engine)\n"
            f"{summary_text}\n\n"
            "**Strategic Recommendations:**\n"
            "1. **Immediate High-Risk Focus:** Prioritize the top 20% accounts/categories driving the majority of overdue or revenue variance.\n"
            "2. **Process & Data Standardization:** Use the cleaned Excel export as the single source of truth for management review.\n"
            "3. **Weekly Tracking:** Monitor collection efficiency and category conversion weekly to prevent bucket roll-forward."
        )
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    prompt = (
        "You are an Executive Management Consultant. Based strictly on the following verified dataset summary "
        "(do not invent any numbers), write a concise, board-ready executive summary with: "
        "1) Key Findings, 2) Risk & Performance Highlights, and 3) 4 Actionable Recommendations.\n\n"
        f"{summary_text}"
    )
    try:
        resp = requests.post(
            url,
            headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
            json={"contents": [{"parts": [{"text": prompt}]}]},
            timeout=25,
        )
        if resp.status_code == 200:
            data = resp.json()
            return data["candidates"][0]["content"]["parts"][0]["text"]
        return f"Gemini API returned status {resp.status_code}. Built-in summary:\n\n{summary_text}"
    except Exception as e:
        return f"AI service unreachable ({e}). Built-in summary:\n\n{summary_text}"


def _add_slide_header(slide, title: str, subtitle: str, rgb):
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.95))
    banner.fill.solid()
    banner.fill.fore_color.rgb = RGBColor(*rgb)
    banner.line.fill.background()
    tf = banner.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(11)
        p2.font.color.rgb = RGBColor(220, 230, 242)


def _add_table_to_slide(slide, df: pd.DataFrame, left=0.5, top=1.3, width=12.3, max_rows=10):
    view = df.head(max_rows).copy()
    rows, cols = len(view) + 1, len(view.columns)
    if cols == 0:
        return
    shape = slide.shapes.add_table(rows, cols, Inches(left), Inches(top), Inches(width), Inches(0.38 * rows))
    tbl = shape.table
    for j, c in enumerate(view.columns):
        cell = tbl.cell(0, j)
        cell.text = str(c)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(10)
            p.font.bold = True
    for i in range(len(view)):
        for j, c in enumerate(view.columns):
            val = view.iloc[i, j]
            cell = tbl.cell(i + 1, j)
            if isinstance(val, (int, float, np.number)) and not pd.isna(val):
                cell.text = f"{val:,.1f}" if isinstance(val, float) and not val.is_integer() else f"{val:,.0f}"
            else:
                cell.text = str(val) if not pd.isna(val) else "-"
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(9)


def _add_bar_chart_to_slide(slide, cats, series_list, left=0.6, top=1.35, width=12.0, height=5.4):
    cd = CategoryChartData()
    cd.categories = [str(c)[:24] for c in cats]
    for name, vals in series_list:
        clean_vals = [0.0 if (v is None or pd.isna(v) or math.isinf(float(v))) else float(v) for v in vals]
        cd.add_series(str(name), clean_vals)
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(left), Inches(top), Inches(width), Inches(height), cd
    )
    chart = chart_shape.chart
    chart.has_legend = len(series_list) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False


def build_credit_ppt(df: pd.DataFrame, prio: pd.DataFrame, company: str, cur: str, rgb, ai_text: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    tot_out = df["outstanding"].sum()
    tot_od = df["overdue_amt"].sum()
    tot_due = df["total_due"].sum()
    tot_coll = df["total_coll"].sum()
    eff = (tot_coll / tot_due * 100) if tot_due > 0 else 0.0
    od_acc = int((df["overdue_amt"] > 0).sum())

    s1 = prs.slides.add_slide(blank)
    _add_slide_header(s1, f"{company} — Credit Recovery Executive Summary", f"Report Date: {date.today()}", rgb)
    summary_df = pd.DataFrame(
        [
            {"Metric": "Total Accounts", "Value": f"{len(df):,}"},
            {"Metric": "Accounts with Overdue", "Value": f"{od_acc:,} ({od_acc/max(1,len(df))*100:.1f}%)"},
            {"Metric": "Total Outstanding Portfolio", "Value": fmt_num(tot_out, cur)},
            {"Metric": "Total Overdue Amount", "Value": fmt_num(tot_od, cur)},
            {"Metric": "Total Cumulative Collection", "Value": fmt_num(tot_coll, cur)},
            {"Metric": "Collection Efficiency", "Value": f"{eff:.1f}%"},
        ]
    )
    _add_table_to_slide(s1, summary_df, top=1.4, max_rows=8)

    s2 = prs.slides.add_slide(blank)
    _add_slide_header(s2, "Portfolio Aging by Due EMI Bucket", "Outstanding & Overdue by Bucket", rgb)
    ag = (
        df.groupby("bucket", as_index=False)
        .agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"))
    )
    _add_bar_chart_to_slide(s2, ag["bucket"], [("Outstanding", ag["Outstanding"]), ("Overdue", ag["Overdue"])])

    s3 = prs.slides.add_slide(blank)
    _add_slide_header(s3, "Branch-Wise Recovery Performance", "Top Branches by Overdue & Collection", rgb)
    br = (
        df.groupby("branch", as_index=False)
        .agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum"))
        .sort_values("Overdue", ascending=False)
        .head(10)
    )
    _add_bar_chart_to_slide(s3, br["branch"], [("Overdue", br["Overdue"]), ("Collected", br["Collected"])])

    s4 = prs.slides.add_slide(blank)
    _add_slide_header(s4, "Collector-Wise Recovery Performance", "Overdue vs Collected (Top 10 Collectors)", rgb)
    co = (
        df.groupby("collector", as_index=False)
        .agg(Accounts=("account_id", "count"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum"), Due=("total_due", "sum"))
        .sort_values("Overdue", ascending=False)
        .head(10)
    )
    _add_bar_chart_to_slide(s4, co["collector"], [("Overdue", co["Overdue"]), ("Collected", co["Collected"])])

    s5 = prs.slides.add_slide(blank)
    _add_slide_header(s5, "Portfolio Risk Status Classification", "Accounts & Exposure by Risk Tier", rgb)
    rk = df.groupby("risk_status", as_index=False).agg(
        Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum")
    )
    _add_table_to_slide(s5, rk, top=1.5, max_rows=10)

    s6 = prs.slides.add_slide(blank)
    _add_slide_header(s6, "Top Priority Accounts for Immediate Recovery", "Ranked by Multi-Factor Risk Score (0-100)", rgb)
    if not prio.empty:
        cols_show = ["rank", "account_id", "customer", "vehicle_reg", "branch", "collector", "due_emi_no", "overdue_amt", "score"]
        _add_table_to_slide(s6, prio[cols_show], top=1.3, max_rows=10)

    s7 = prs.slides.add_slide(blank)
    _add_slide_header(s7, "Executive Insights & Action Plan", "AI-Assisted Management Commentary", rgb)
    tb = s7.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.6))
    tf = tb.text_frame
    tf.word_wrap = True
    for idx, line in enumerate(ai_text.splitlines()[:18]):
        clean_line = line.replace("###", "").replace("**", "").strip()
        if not clean_line:
            continue
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = clean_line
        p.font.size = Pt(12)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_universal_ppt(df: pd.DataFrame, company: str, cat_col: str, num_col: str, rgb, ai_text: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    _add_slide_header(s1, f"{company} — Executive Data Analysis Report", f"Generated on {date.today()}", rgb)
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    kpi_rows = [{"Metric": "Total Clean Records", "Value": f"{len(df):,}"}, {"Metric": "Total Columns Analyzed", "Value": f"{df.shape[1]}"}]
    for nc in num_cols[:5]:
        kpi_rows.append({"Metric": f"Total {nc}", "Value": f"{df[nc].sum():,.2f}"})
        kpi_rows.append({"Metric": f"Average {nc}", "Value": f"{df[nc].mean():,.2f}"})
    _add_table_to_slide(s1, pd.DataFrame(kpi_rows), top=1.4, max_rows=10)

    if cat_col and num_col and cat_col in df.columns and num_col in df.columns:
        s2 = prs.slides.add_slide(blank)
        _add_slide_header(s2, f"{num_col} by {cat_col}", "Top 10 Breakdown", rgb)
        grp = (
            df.groupby(cat_col, as_index=False)[num_col]
            .sum()
            .sort_values(num_col, ascending=False)
            .head(10)
        )
        _add_bar_chart_to_slide(s2, grp[cat_col], [(num_col, grp[num_col])])

        s3 = prs.slides.add_slide(blank)
        _add_slide_header(s3, f"Detailed Summary by {cat_col}", f"Count, Sum & Average of {num_col}", rgb)
        grp_tbl = (
            df.groupby(cat_col, as_index=False)
            .agg(Records=(num_col, "count"), Total=(num_col, "sum"), Average=(num_col, "mean"))
            .sort_values("Total", ascending=False)
        )
        _add_table_to_slide(s3, grp_tbl, top=1.4, max_rows=11)

    s4 = prs.slides.add_slide(blank)
    _add_slide_header(s4, "Executive Commentary & Recommendations", "Automated Intelligence", rgb)
    tb = s4.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.6))
    tf = tb.text_frame
    tf.word_wrap = True
    for idx, line in enumerate(ai_text.splitlines()[:18]):
        clean_line = line.replace("###", "").replace("**", "").strip()
        if not clean_line:
            continue
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = clean_line
        p.font.size = Pt(12)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_multi_sheet_excel(sheets_dict: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, d in sheets_dict.items():
            safe_name = re.sub(r"[\[\]\*:/\?\\]", "", str(name))[:31] or "Sheet"
            d.to_excel(writer, sheet_name=safe_name, index=False)
    return buf.getvalue()


st.sidebar.markdown("## 🚀 AI Analytics Studio")
service_mode = st.sidebar.radio(
    "Select Service Mode",
    [
        "1️⃣ Universal Messy Data Cleaner & AI Analytics",
        "2️⃣ Credit Recovery Intelligence Pro",
    ],
)

company_name = st.sidebar.text_input("Company / Client Name", value="Uttara Motors Limited")
brand_color = st.sidebar.color_picker("Brand Theme Colour", value="#1F3A5F")
currency_sym = st.sidebar.text_input("Currency Symbol", value="BDT")
safe_table_mode = st.sidebar.checkbox(
    "Firewall-safe table mode (prevents TypeError)",
    value=True,
    help="Keep checked on office networks to avoid red DataFrame.js TypeError boxes.",
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📂 1. Load Data")
demo_choice = st.sidebar.selectbox(
    "Quick Demo Dataset (for testing / Fiverr portfolio)",
    ["-- Upload My Own File --", "Demo 1: Vehicle Credit Recovery Data", "Demo 2: Messy Business Sales Data"],
)

uploaded_file = st.sidebar.file_uploader("Upload Excel, CSV, or PDF", type=["xlsx", "xls", "csv", "pdf"])

raw_df = None
pdf_text = ""

if demo_choice == "Demo 1: Vehicle Credit Recovery Data":
    raw_df = make_sample_credit_df()
elif demo_choice == "Demo 2: Messy Business Sales Data":
    raw_df = make_sample_messy_business_df()
elif uploaded_file is not None:
    fname = uploaded_file.name.lower()
    if fname.endswith(".csv"):
        raw_df = pd.read_csv(uploaded_file)
    elif fname.endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(uploaded_file)
        sheet = st.sidebar.selectbox("Excel Sheet", xl.sheet_names)
        preview = xl.parse(sheet, header=None)
        auto_hdr = detect_best_header_row(preview)
        hdr_row = st.sidebar.number_input("Header Row (1-based)", min_value=1, max_value=max(1, len(preview)), value=auto_hdr + 1)
        raw_df = xl.parse(sheet, header=int(hdr_row - 1))
    elif fname.endswith(".pdf"):
        tables = []
        with pdfplumber.open(uploaded_file) as pdf:
            for p in pdf.pages:
                pdf_text += (p.extract_text() or "") + "\n"
                for t in p.extract_tables() or []:
                    if len(t) > 1:
                        tables.append(pd.DataFrame(t[1:], columns=t[0]))
        if tables:
            raw_df = pd.concat(tables, ignore_index=True)

st.markdown(
    f"""
    <div class="hero-box" style="background: linear-gradient(135deg, {brand_color}, #0F172A);">
        <h2 style="margin:0; color:white;">{company_name} — AI Data & Credit Intelligence Studio</h2>
        <p style="margin:0.3rem 0 0 0; opacity:0.9;">
            Automated Messy Data Cleaning • Executive Business Analytics • Credit Recovery Intelligence • 1-Click PPT & Excel Builder
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

if raw_df is None:
    st.info(
        "👈 **Get Started:** Upload any Excel/CSV/PDF file in the left sidebar, or select a **Quick Demo Dataset** to test the platform immediately."
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### 🧹 Auto Messy Data Cleaner\nFixes broken headers, removes subtotal/duplicate rows, strips currency text (`BDT 12,000/-`), and exports clean Excel.")
    with c2:
        st.markdown("#### 📊 AI Business Analytics\nAutomatically detects categories, metrics, and dates to generate interactive charts, Pareto tables, and PPT slides.")
    with c3:
        st.markdown("#### 💳 Credit Recovery Pro\nEMI aging buckets, PAR analysis, collector efficiency, 0–100 recovery priority scoring, and executive presentations.")
    st.stop()


if service_mode.startswith("1️⃣"):
    st.sidebar.markdown("### 🧹 2. Cleaning Rules")
    drop_tot = st.sidebar.checkbox("Remove embedded Total/Subtotal rows", value=True)
    drop_dup = st.sidebar.checkbox("Remove exact duplicate rows", value=True)
    std_txt = st.sidebar.checkbox("Trim spaces & standardize category case", value=True)
    num_pol = st.sidebar.selectbox("Missing Numeric Values", ["Keep as NaN (transparent)", "Fill with 0", "Fill with Column Median"])
    txt_pol = st.sidebar.selectbox("Missing Text Values", ["Fill with 'Unknown'", "Keep as blank"])

    clean_df, audit_log = auto_clean_dataframe(raw_df, drop_tot, drop_dup, std_txt, num_pol, txt_pol)

    num_cols = clean_df.select_dtypes(include=[np.number]).columns.tolist()
    cat_cols = [c for c in clean_df.columns if c not in num_cols]

    t_clean, t_dash, t_anom, t_export = st.tabs(
        ["🧹 1. Auto Data Cleaner (Before vs After)", "📊 2. Executive Business Dashboard", "🔍 3. Pivot & Anomaly Finder", "📥 4. AI Report, PPT & Excel Export"]
    )

    with t_clean:
        m1, m2, m3, m4 = st.columns(4)
        m1.markdown(f'<div class="kpi-card"><div class="kpi-title">Raw Rows</div><div class="kpi-val">{len(raw_df):,}</div></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="kpi-card"><div class="kpi-title">Clean Rows</div><div class="kpi-val">{len(clean_df):,}</div></div>', unsafe_allow_html=True)
        m3.markdown(f'<div class="kpi-card"><div class="kpi-title">Numeric Columns</div><div class="kpi-val">{len(num_cols)}</div></div>', unsafe_allow_html=True)
        m4.markdown(f'<div class="kpi-card"><div class="kpi-title">Category/Text Columns</div><div class="kpi-val">{len(cat_cols)}</div></div>', unsafe_allow_html=True)

        st.markdown("#### ✅ Automated Cleaning Audit Log")
        for item in audit_log:
            st.markdown(f'<div class="insight-box">✔️ {item}</div>', unsafe_allow_html=True)

        col_b, col_a = st.columns(2)
        with col_b:
            st.markdown("##### Before Cleaning (Raw Preview)")
            show_table(raw_df, safe_table_mode, 25)
        with col_a:
            st.markdown("##### After Cleaning (Fixed & Standardized)")
            show_table(clean_df, safe_table_mode, 25)

    with t_dash:
        if not num_cols:
            st.warning("No numeric columns detected yet. Check your header row in the sidebar.")
        else:
            c_sel1, c_sel2 = st.columns(2)
            primary_cat = c_sel1.selectbox("Select Primary Category (e.g., Branch, Region, Product, Collector)", cat_cols if cat_cols else clean_df.columns)
            primary_num = c_sel2.selectbox("Select Primary Metric (e.g., Revenue, Overdue, Balance, Profit)", num_cols)

            total_val = clean_df[primary_num].sum()
            avg_val = clean_df[primary_num].mean()
            med_val = clean_df[primary_num].median()
            max_val = clean_df[primary_num].max()

            k1, k2, k3, k4 = st.columns(4)
            k1.markdown(f'<div class="kpi-card"><div class="kpi-title">Total {primary_num}</div><div class="kpi-val">{fmt_num(total_val, currency_sym)}</div></div>', unsafe_allow_html=True)
            k2.markdown(f'<div class="kpi-card"><div class="kpi-title">Average {primary_num}</div><div class="kpi-val">{fmt_num(avg_val, currency_sym, 1)}</div></div>', unsafe_allow_html=True)
            k3.markdown(f'<div class="kpi-card"><div class="kpi-title">Median {primary_num}</div><div class="kpi-val">{fmt_num(med_val, currency_sym, 1)}</div></div>', unsafe_allow_html=True)
            k4.markdown(f'<div class="kpi-card"><div class="kpi-title">Maximum {primary_num}</div><div class="kpi-val">{fmt_num(max_val, currency_sym)}</div></div>', unsafe_allow_html=True)

            grp = (
                clean_df.groupby(primary_cat, as_index=False)
                .agg(Records=(primary_num, "count"), Total=(primary_num, "sum"), Average=(primary_num, "mean"))
                .sort_values("Total", ascending=False)
            )
            grp["Share %"] = (grp["Total"] / max(1e-9, grp["Total"].sum()) * 100).round(1)

            ch1, ch2 = st.columns(2)
            with ch1:
                fig_bar = px.bar(grp.head(15), x=primary_cat, y="Total", text="Share %", title=f"Top {primary_cat} by {primary_num}", color_discrete_sequence=[brand_color])
                st.plotly_chart(fig_bar, use_container_width=True)
            with ch2:
                fig_pie = px.pie(grp.head(10), names=primary_cat, values="Total", hole=0.45, title=f"Share Distribution ({primary_cat})")
                st.plotly_chart(fig_pie, use_container_width=True)

            st.markdown(f"#### 📋 Summary Table by {primary_cat}")
            show_table(grp, safe_table_mode, 50)

    with t_anom:
        st.markdown("#### 🔍 Statistical Outlier & High-Value Record Finder")
        if num_cols:
            anom_col = st.selectbox("Select Metric to Scan for Outliers (Z-Score > 2.0)", num_cols, key="anom_c")
            s_num = clean_df[anom_col].dropna()
            if len(s_num) > 3 and s_num.std() > 0:
                z = (clean_df[anom_col] - s_num.mean()) / s_num.std()
                outliers = clean_df[z.abs() >= 2.0].sort_values(anom_col, ascending=False)
                st.write(f"Found **{len(outliers)}** statistical outlier record(s) where `{anom_col}` is unusually high or low:")
                show_table(outliers, safe_table_mode, 50)
            else:
                st.info("Not enough variance to compute statistical outliers.")

    with t_export:
        st.markdown("#### 🤖 AI Executive Summary & Client Deliverables")
        p_cat = cat_cols[0] if cat_cols else clean_df.columns[0]
        p_num = num_cols[0] if num_cols else clean_df.columns[-1]
        summary_lines = [
            f"- **Dataset Size:** {len(clean_df):,} cleaned records across {clean_df.shape[1]} columns.",
            f"- **Cleaning Actions Completed:** {'; '.join(audit_log)}",
        ]
        if num_cols:
            top_grp = clean_df.groupby(p_cat)[p_num].sum().sort_values(ascending=False)
            if len(top_grp) > 0:
                top_name = top_grp.index[0]
                top_val = top_grp.iloc[0]
                share = top_val / max(1e-9, clean_df[p_num].sum()) * 100
                summary_lines.append(
                    f"- **Primary Metric ({p_num}):** Total = {fmt_num(clean_df[p_num].sum(), currency_sym)}, Average = {fmt_num(clean_df[p_num].mean(), currency_sym, 1)}."
                )
                summary_lines.append(
                    f"- **Top Performer ({p_cat}):** '{top_name}' leads with {fmt_num(top_val, currency_sym)} ({share:.1f}% of total)."
                )
        auto_summary_text = "\n".join(summary_lines)

        if st.button("✨ Generate AI Executive Commentary"):
            st.session_state["uni_ai_text"] = generate_ai_insights(auto_summary_text)
        ai_txt = st.session_state.get("uni_ai_text", generate_ai_insights(auto_summary_text))
        st.markdown(ai_txt)

        st.markdown("---")
        b1, b2 = st.columns(2)
        with b1:
            if st.button("📊 Build Executive PowerPoint (.pptx)"):
                st.session_state["uni_ppt"] = build_universal_ppt(
                    clean_df, company_name, p_cat, p_num if num_cols else None, hex_to_rgb(brand_color), ai_txt
                )
            if "uni_ppt" in st.session_state:
                st.download_button(
                    "⬇️ Download PowerPoint Presentation (.pptx)",
                    st.session_state["uni_ppt"],
                    file_name="Executive_Analytics_Presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        with b2:
            if st.button("📗 Build Cleaned & Analyzed Excel (.xlsx)"):
                sheets = {
                    "Cleaned_Data": clean_df,
                    "Cleaning_Audit_Log": pd.DataFrame({"Step": range(1, len(audit_log) + 1), "Action": audit_log}),
                    "Numeric_Summary": clean_df.describe().T.reset_index(),
                }
                if num_cols and cat_cols:
                    sheets["Category_Breakdown"] = (
                        clean_df.groupby(p_cat, as_index=False)
                        .agg(Records=(p_num, "count"), Total=(p_num, "sum"), Average=(p_num, "mean"))
                        .sort_values("Total", ascending=False)
                    )
                st.session_state["uni_xlsx"] = build_multi_sheet_excel(sheets)
            if "uni_xlsx" in st.session_state:
                st.download_button(
                    "⬇️ Download Clean Excel Report (.xlsx)",
                    st.session_state["uni_xlsx"],
                    file_name="Cleaned_Business_Analytics_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

else:
    auto_map = auto_map_credit_columns(raw_df.columns)
    with st.sidebar.expander("🔧 Column Mapping (Auto-Detected)", expanded=False):
        opts = [None] + list(raw_df.columns)
        user_map = {}
        for key, label, _ in CREDIT_FIELDS:
            default_col = auto_map.get(key)
            idx = opts.index(default_col) if default_col in opts else 0
            user_map[key] = st.selectbox(label, opts, index=idx, key=f"map_{key}")

    out_basis = st.sidebar.radio("Outstanding Basis", ["Net balance", "Balance receivable"])
    df_cred, cred_notes = prepare_credit_data(raw_df, user_map, out_basis)

    st.sidebar.markdown("### 🔍 Filters")
    branches = sorted(df_cred["branch"].unique())
    sel_br = st.sidebar.multiselect("Branch", branches, default=branches)
    collectors = sorted(df_cred["collector"].unique())
    sel_co = st.sidebar.multiselect("Collector", collectors, default=collectors)
    df_f = df_cred[df_cred["branch"].isin(sel_br) & df_cred["collector"].isin(sel_co)].copy()

    prio_df = score_recovery_priority(df_f)

    tot_out = df_f["outstanding"].sum()
    tot_od = df_f["overdue_amt"].sum()
    tot_due = df_f["total_due"].sum()
    tot_coll = df_f["total_coll"].sum()
    eff = (tot_coll / tot_due * 100) if tot_due > 0 else 0.0
    par1 = (df_f.loc[df_f["due_emi_no"] >= 1, "outstanding"].sum() / tot_out * 100) if tot_out > 0 else 0.0
    od_cnt = int((df_f["overdue_amt"] > 0).sum())

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Outstanding</div><div class="kpi-val">{fmt_num(tot_out, currency_sym)}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Overdue</div><div class="kpi-val">{fmt_num(tot_od, currency_sym)}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Collection</div><div class="kpi-val">{fmt_num(tot_coll, currency_sym)}</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div class="kpi-title">Collection Efficiency</div><div class="kpi-val">{eff:.1f}%</div></div>', unsafe_allow_html=True)
    k5.markdown(f'<div class="kpi-card"><div class="kpi-title">Overdue Accounts</div><div class="kpi-val">{od_cnt:,} / {len(df_f):,}</div></div>', unsafe_allow_html=True)
    k6.markdown(f'<div class="kpi-card"><div class="kpi-title">PAR 1+ (At Risk)</div><div class="kpi-val">{par1:.1f}%</div></div>', unsafe_allow_html=True)

    t1, t2, t3, t4 = st.tabs(
        ["📊 1. Executive Dashboard", "🚨 2. Recovery Priority List (0-100 Score)", "🛡️ 3. Provisioning & Data Quality", "📥 4. PowerPoint & Excel Export"]
    )

    with t1:
        c1, c2 = st.columns(2)
        with c1:
            ag = df_f.groupby("bucket", as_index=False).agg(Accounts=("account_id", "count"), Overdue=("overdue_amt", "sum"), Outstanding=("outstanding", "sum"))
            fig1 = px.bar(ag, x="bucket", y=["Outstanding", "Overdue"], barmode="group", title="Portfolio by Due EMI Bucket")
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            br_df = df_f.groupby("branch", as_index=False).agg(Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum")).sort_values("Overdue", ascending=False).head(12)
            fig2 = px.bar(br_df, x="branch", y=["Overdue", "Collected"], barmode="group", title="Branch-Wise Overdue vs Collection")
            st.plotly_chart(fig2, use_container_width=True)

        c3, c4 = st.columns(2)
        with c3:
            co_df = df_f.groupby("collector", as_index=False).agg(Due=("total_due", "sum"), Collected=("total_coll", "sum"), Overdue=("overdue_amt", "sum"))
            co_df["Efficiency %"] = np.where(co_df["Due"] > 0, (co_df["Collected"] / co_df["Due"] * 100).round(1), 0.0)
            co_df = co_df.sort_values("Efficiency %", ascending=True)
            fig3 = px.bar(co_df, x="Efficiency %", y="collector", orientation="h", text="Efficiency %", title="Collection Efficiency by Collector (%)")
            st.plotly_chart(fig3, use_container_width=True)
        with c4:
            rk_df = df_f.groupby("risk_status", as_index=False).agg(Accounts=("account_id", "count"), Overdue=("overdue_amt", "sum"))
            fig4 = px.pie(rk_df, names="risk_status", values="Overdue", hole=0.45, title="Overdue Share by Risk Tier")
            st.plotly_chart(fig4, use_container_width=True)

    with t2:
        st.markdown("#### 🚨 Vehicle / Loan Recovery Priority Ranking")
        if prio_df.empty:
            st.success("No overdue accounts found in the current filter.")
        else:
            cols_p = ["rank", "account_id", "customer", "vehicle_reg", "branch", "collector", "due_emi_no", "overdue_amt", "outstanding", "score", "main_driver", "explanation"]
            show_table(prio_df[cols_p], safe_table_mode, 100)

    with t3:
        st.markdown("#### 🛡️ Management Provisioning Estimate (Non-Accounting Estimate)")
        prov = (
            df_f.groupby("bucket", as_index=False)
            .agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"))
        )
        rate_map = {"Current (0 EMI)": 0.0, "1 EMI": 0.05, "2 EMI": 0.25, "3 EMI": 0.50, "4-5 EMI": 0.50, "6+ EMI": 1.00}
        prov["Provision Rate %"] = prov["bucket"].map(rate_map).fillna(0.0) * 100
        prov["Estimated Provision"] = (prov["Outstanding"] * prov["Provision Rate %"] / 100.0).round(0)
        show_table(prov, safe_table_mode, 20)

        if cred_notes:
            st.markdown("#### ⚠️ Data Quality & Derivation Notes")
            show_table(pd.DataFrame(cred_notes, columns=["Level", "Note"]), safe_table_mode, 20)

    with t4:
        cred_summary_txt = (
            f"- **Total Accounts:** {len(df_f):,} ({od_cnt:,} with overdue).\n"
            f"- **Total Outstanding ({out_basis}):** {fmt_num(tot_out, currency_sym)}.\n"
            f"- **Total Overdue Amount:** {fmt_num(tot_od, currency_sym)}.\n"
            f"- **Collection Efficiency:** {eff:.1f}% (PAR 1+: {par1:.1f}%)."
        )
        ai_cred = generate_ai_insights(cred_summary_txt)
        st.markdown(ai_cred)

        b1, b2 = st.columns(2)
        with b1:
            if st.button("📊 Build Credit Recovery PowerPoint"):
                st.session_state["cred_ppt"] = build_credit_ppt(
                    df_f, prio_df, company_name, currency_sym, hex_to_rgb(brand_color), ai_cred
                )
            if "cred_ppt" in st.session_state:
                st.download_button(
                    "⬇️ Download Credit Recovery PPT (.pptx)",
                    st.session_state["cred_ppt"],
                    file_name="Credit_Recovery_Executive_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        with b2:
            if st.button("📗 Build Full Credit Excel Report"):
                sheets_c = {
                    "Branch_Summary": df_f.groupby("branch", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum")),
                    "Collector_Summary": df_f.groupby("collector", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum")),
                    "Priority_List": prio_df,
                    "Cleaned_Portfolio": df_f,
                }
                st.session_state["cred_xlsx"] = build_multi_sheet_excel(sheets_c)
            if "cred_xlsx" in st.session_state:
                st.download_button(
                    "⬇️ Download Credit Recovery Excel (.xlsx)",
                    st.session_state["cred_xlsx"],
                    file_name="Credit_Recovery_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
```
