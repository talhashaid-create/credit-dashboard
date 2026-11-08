"""Enterprise v3.0: Automated AI Data Cleaner, Business Analytics & Credit Recovery Studio.
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
from pptx.util import Inches, Pt

warnings.filterwarnings("ignore")
st.set_page_config(
    page_title="AI Data & Credit Intelligence Studio v3.0",
    page_icon="🚀",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.1rem; padding-bottom: 2.2rem;}
    .hero-banner {
        padding: 1.4rem 1.8rem; border-radius: 14px; color: white;
        margin-bottom: 1.1rem; box-shadow: 0 6px 18px rgba(15,23,42,0.12);
    }
    .kpi-card {
        background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
        padding: 0.9rem 1rem; box-shadow: 0 2px 5px rgba(0,0,0,0.03);
        margin-bottom: 0.6rem; height: 100%;
    }
    .kpi-title {font-size: 0.75rem; color: #64748B; text-transform: uppercase; font-weight: 700; letter-spacing: 0.03em;}
    .kpi-val {font-size: 1.32rem; color: #0F172A; font-weight: 800; margin-top: 0.2rem;}
    .kpi-sub {font-size: 0.74rem; color: #059669; margin-top: 0.15rem; font-weight: 600;}
    .audit-item {
        background: #F8FAFC; border-left: 4px solid #2563EB;
        padding: 0.65rem 0.95rem; border-radius: 6px; margin-bottom: 0.45rem; font-size: 0.9rem;
    }
    .badge-ok {
        display: inline-block; background: #DCFCE7; color: #166534;
        padding: 0.2rem 0.6rem; border-radius: 999px; font-size: 0.78rem; font-weight: 700;
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
    st.sidebar.markdown("### 🔒 Studio Access")
    entered = st.sidebar.text_input("Enter Password", type="password")
    if entered:
        if hmac.compare_digest(entered.strip(), pw):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.sidebar.error("Incorrect password.")
    st.info("👈 Please enter your studio password in the left sidebar to continue.")
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


def show_table(df: pd.DataFrame, safe_mode: bool = True, max_rows: int = 100):
    if df is None or df.empty:
        st.caption("No records to display.")
        return
    view = df.head(max_rows).copy()
    for c in view.select_dtypes(include=[np.number]).columns:
        view[c] = view[c].apply(
            lambda x: "-" if pd.isna(x) else (f"{x:,.0f}" if float(x).is_integer() else f"{x:,.2f}")
        )
    if safe_mode:
        st.markdown(
            f'<div style="overflow-x:auto; max-height:440px; border:1px solid #e2e8f0; border-radius:8px; background:white; padding:4px;">'
            f'{view.to_html(index=False, border=0)}</div>',
            unsafe_allow_html=True,
        )
    else:
        try:
            st.dataframe(df.head(max_rows), use_container_width=True)
        except Exception:
            st.markdown(view.to_html(index=False), unsafe_allow_html=True)


def make_demo_credit() -> pd.DataFrame:
    rng = np.random.default_rng(42)
    branches = ["Dhaka North", "Chattogram", "Bogura", "Rangpur", "Jashore", "Rajshahi", "Sylhet"]
    collectors = ["Azmir", "Aminul", "Zahid", "Mannan", "Hasen", "Zamil", "Akter"]
    sales_p = ["Rahim", "Karim", "Tariq", "Noman", "Sajid"]
    models = ["Cargo Truck 3T", "Pickup 1.5T", "Covered Van", "Passenger Bus", "Prime Mover"]
    rows = []
    for i in range(1, 95):
        sp = int(rng.choice([1250000, 1650000, 2100000, 2850000, 3400000]))
        dp_pct = float(rng.choice([0.0, 0.08, 0.12, 0.15, 0.20, 0.25, 0.30]))
        dp = round(sp * dp_pct)
        interest = round((sp - dp) * 0.18)
        bill = sp + interest
        no_emi = int(rng.choice([24, 36, 48, 60]))
        emi = round((bill - dp) / no_emi)
        inst_due = int(rng.integers(5, no_emi))
        due_emi = int(rng.choice([0, 0, 0, 1, 1, 2, 3, 4, 6, 8, 11, 15]))
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
                "Account Id": f"34048{i:04d}",
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


def make_demo_messy_sales() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Invoice_No": "INV-101", "Date": "2025-01-05", "Branch": "  dhaka ", "Product_Line": "Electronics", "Manager": "Tanvir", "Gross_Sales": "BDT 225,000/-", "Cost": "162,000", "Net_Profit": "63,000", "Units": "15"},
            {"Invoice_No": "INV-102", "Date": "2025/01/14", "Branch": "Chattogram", "Product_Line": "Home Appliances", "Manager": "Nadia", "Gross_Sales": "184,500", "Cost": "131,000", "Net_Profit": "53,500", "Units": "12"},
            {"Invoice_No": "INV-102", "Date": "2025/01/14", "Branch": "Chattogram", "Product_Line": "Home Appliances", "Manager": "Nadia", "Gross_Sales": "184,500", "Cost": "131,000", "Net_Profit": "53,500", "Units": "12"},
            {"Invoice_No": "INV-103", "Date": "2025-02-02", "Branch": "DHAKA", "Product_Line": "Electronics", "Manager": "Tanvir", "Gross_Sales": "$310,000", "Cost": "215,000", "Net_Profit": "95,000", "Units": "22"},
            {"Invoice_No": "INV-104", "Date": "2025-02-18", "Branch": "Sylhet", "Product_Line": "Industrial Gear", "Manager": "Farhan", "Gross_Sales": "440,000", "Cost": "310,000", "Net_Profit": "130,000", "Units": "8"},
            {"Invoice_No": "INV-105", "Date": "2025-03-04", "Branch": "Rajshahi", "Product_Line": "Heavy Machinery", "Manager": "Mehedi", "Gross_Sales": "620,000", "Cost": "460,000", "Net_Profit": "160,000", "Units": "5"},
            {"Invoice_No": "INV-106", "Date": "2025-03-19", "Branch": "  chattogram", "Product_Line": "Industrial Gear", "Manager": "Nadia", "Gross_Sales": "275,000", "Cost": "198,000", "Net_Profit": "77,000", "Units": "9"},
            {"Invoice_No": "INV-107", "Date": "2025-04-06", "Branch": "Dhaka", "Product_Line": "Heavy Machinery", "Manager": "Tanvir", "Gross_Sales": "990,000", "Cost": "680,000", "Net_Profit": "310,000", "Units": "14"},
            {"Invoice_No": "INV-108", "Date": "2025-04-22", "Branch": "Khulna", "Product_Line": "Home Appliances", "Manager": "Sadia", "Gross_Sales": "114,000", "Cost": "84,000", "Net_Profit": "30,000", "Units": "7"},
            {"Invoice_No": "INV-109", "Date": "2025-05-10", "Branch": "Bogura", "Product_Line": "Electronics", "Manager": "Farhan", "Gross_Sales": "305,000", "Cost": "220,000", "Net_Profit": "85,000", "Units": "19"},
            {"Invoice_No": "Grand Total", "Date": "", "Branch": "Total", "Product_Line": "", "Manager": "", "Gross_Sales": "3,648,000", "Cost": "2,591,000", "Net_Profit": "1,057,000", "Units": "123"},
        ]
    )


def make_demo_finance() -> pd.DataFrame:
    rng = np.random.default_rng(7)
    depts = ["Operations", "Logistics", "Marketing", "IT & Cloud", "Sales & Distribution", "Administration"]
    vendors = ["Global Tech", "Prime Fuels", "Metro Logistics", "Apex Media", "Delta Supplies"]
    rows = []
    for i in range(1, 65):
        budget = int(rng.choice([150000, 250000, 400000, 600000]))
        actual = int(budget * rng.uniform(0.75, 1.35))
        variance = actual - budget
        rows.append(
            {
                "Voucher_ID": f"VCH-{2000+i}",
                "Expense_Date": f"2025-{int(rng.integers(1,7)):02d}-{int(rng.integers(1,28)):02d}",
                "Department": rng.choice(depts),
                "Vendor_Name": rng.choice(vendors),
                "Approved_Budget": f"BDT {budget:,}",
                "Actual_Spend": f"BDT {actual:,}",
                "Variance_Amount": variance,
            }
        )
    return pd.DataFrame(rows)


def make_demo_hr() -> pd.DataFrame:
    rng = np.random.default_rng(99)
    depts = ["Sales", "Collection", "Service Workshop", "Finance", "HR & Admin", "Supply Chain"]
    locations = ["Dhaka HQ", "Chattogram", "Bogura", "Rangpur", "Jashore"]
    rows = []
    for i in range(1, 75):
        sal = int(rng.choice([28000, 38000, 52000, 75000, 110000]))
        kpi = round(float(rng.uniform(58, 99)), 1)
        rows.append(
            {
                "Emp_ID": f"EMP-{100+i}",
                "Department": rng.choice(depts),
                "Location": rng.choice(locations),
                "Monthly_Salary": sal,
                "KPI_Score": kpi,
                "Overtime_Hours": int(rng.integers(0, 35)),
                "Training_Hours": int(rng.integers(2, 24)),
            }
        )
    return pd.DataFrame(rows)


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
    if num.notna().mean() >= 0.72:
        full_clean = (
            s.astype(str)
            .str.strip()
            .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
            .str.replace(r"(?i)(bdt|tk|taka|usd|eur|inr|rs|৳|\$|€|£|/-|,|\s)", "", regex=True)
            .str.replace(r"%$", "", regex=True)
        )
        out = pd.to_numeric(full_clean, errors="coerce")
        fixed_cnt = int((out.notna() & ~s.apply(lambda x: isinstance(x, (int, float)))).sum())
        return out, fixed_cnt
    return s, 0


def compute_health_score(df: pd.DataFrame) -> float:
    if df is None or df.empty:
        return 0.0
    total_cells = max(1, df.shape[0] * df.shape[1])
    missing_ratio = df.isna().sum().sum() / total_cells
    dup_ratio = df.duplicated().mean() if len(df) > 1 else 0.0
    unnamed_ratio = sum(1 for c in df.columns if str(c).lower().startswith("unnamed")) / max(1, len(df.columns))
    score = 100.0 - (missing_ratio * 45.0) - (dup_ratio * 30.0) - (unnamed_ratio * 25.0)
    return round(float(np.clip(score, 10.0, 100.0)), 1)


def auto_clean_dataframe(
    df_in: pd.DataFrame,
    drop_totals: bool = True,
    drop_dups: bool = True,
    standardize_text: bool = True,
    missing_num_policy: str = "Fill with 0",
    missing_txt_policy: str = "Fill with 'Unknown'",
):
    raw_health = compute_health_score(df_in)
    df = df_in.copy()
    audit = []

    new_cols, seen = [], {}
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
        audit.append(f"Removed {c0 - df.shape[1]} empty column(s).")

    if drop_totals and len(df) > 0:
        check_cols = df.columns[: min(5, len(df.columns))]
        mask_total = pd.Series(False, index=df.index)
        for col in check_cols:
            mask_total = mask_total | df[col].astype(str).str.match(TOTAL_PAT, na=False)
        tot_cnt = int(mask_total.sum())
        if tot_cnt > 0:
            df = df.loc[~mask_total].copy()
            audit.append(f"Removed {tot_cnt} Subtotal / Grand Total summary row(s).")

    txt_fixed = 0
    for col in df.columns:
        if df[col].dtype == object:
            orig = df[col].copy()
            s = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
            s = s.replace({"nan": np.nan, "None": np.nan, "NULL": np.nan, "null": np.nan, "": np.nan})
            non_na = s.dropna()
            if standardize_text and 0 < non_na.nunique() <= 45:
                if non_na.str.lower().nunique() < non_na.nunique():
                    s = s.str.title()
            txt_fixed += int((orig.fillna("") != s.fillna("")).sum())
            df[col] = s
    if txt_fixed > 0:
        audit.append(f"Trimmed spaces & standardized text casing in {txt_fixed} cell(s).")

    num_cols_converted = []
    for col in df.columns:
        col_low = col.lower()
        if any(k in col_low for k in ["id", "serial", "account", "phone", "mobile", "reg", "vehicle", "code", "invoice", "voucher", "emp"]):
            continue
        converted, cnt = parse_messy_number_series(df[col])
        if cnt > 0:
            df[col] = converted
            num_cols_converted.append(f"{col} ({cnt})")
    if num_cols_converted:
        audit.append("Converted formatted currency/number strings to numeric: " + ", ".join(num_cols_converted))

    date_cols = []
    for col in df.columns:
        if df[col].dtype == object and any(k in col.lower() for k in ["date", "month", "day", "period", "time"]):
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().mean() >= 0.6:
                df[col] = parsed.dt.strftime("%Y-%m-%d")
                date_cols.append(col)
    if date_cols:
        audit.append("Standardized date formatting (YYYY-MM-DD) in: " + ", ".join(date_cols))

    if drop_dups and len(df) > 0:
        before = len(df)
        df = df.drop_duplicates().copy()
        if before - len(df) > 0:
            audit.append(f"Removed {before - len(df)} exact duplicate row(s).")

    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    txt_cols = [c for c in df.columns if c not in num_cols]

    if missing_num_policy == "Fill with 0" and num_cols:
        na_cnt = int(df[num_cols].isna().sum().sum())
        if na_cnt > 0:
            df[num_cols] = df[num_cols].fillna(0)
            audit.append(f"Filled {na_cnt} missing numeric cell(s) with 0.")
    elif missing_num_policy == "Fill with Column Median" and num_cols:
        for c in num_cols:
            na_c = int(df[c].isna().sum())
            if na_c > 0 and df[c].notna().any():
                med = df[c].median()
                df[c] = df[c].fillna(med)
                audit.append(f"Filled {na_c} missing value(s) in '{c}' with median ({med:,.1f}).")

    if missing_txt_policy == "Fill with 'Unknown'" and txt_cols:
        na_txt = int(df[txt_cols].isna().sum().sum())
        if na_txt > 0:
            df[txt_cols] = df[txt_cols].fillna("Unknown")
            audit.append(f"Replaced {na_txt} blank text cell(s) with 'Unknown'.")

    if not audit:
        audit.append("Dataset passed all automated checks with zero structural issues.")

    clean_health = compute_health_score(df)
    return df.reset_index(drop=True), audit, raw_health, clean_health, date_cols


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
    ("balance_rec", "Balance receivable", ["balance receivable", "receivable", "bal receivable", "closing balance"]),
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
    mapping, used = {}, set()
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
                        if any(w in cn for w in ["remark", "note", "comment", "status"]):
                            continue
                        matched = orig_c
                        break
                if matched:
                    break
        mapping[key] = matched
        if matched:
            used.add(matched)
    return mapping


def is_credit_dataset(cols) -> bool:
    mp = auto_map_credit_columns(cols)
    core_hits = sum(
        1 for k in ["overdue_amt", "due_emi_no", "net_balance", "balance_rec", "total_coll", "emi", "collector"]
        if mp.get(k) is not None
    )
    return core_hits >= 3


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
    mask_tot = pd.Series(False, index=raw.index)
    for c in raw.columns[: min(5, len(raw.columns))]:
        mask_tot = mask_tot | raw[c].astype(str).str.match(TOTAL_PAT, na=False)
    if mask_tot.any():
        raw = raw.loc[~mask_tot].copy()

    df = pd.DataFrame(index=raw.index)
    txt_keys = {"serial", "customer", "branch", "sales_person", "account_id", "collector", "vehicle_reg", "vehicle_model", "emi_date"}
    for key, _, _ in CREDIT_FIELDS:
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
    else:
        df["overdue_amt"] = df["overdue_amt"].clip(lower=0.0)

    if mapping.get("due_emi_no") is None:
        if mapping.get("inst_due_no") and mapping.get("emi_coll_no"):
            df["due_emi_no"] = (df["inst_due_no"] - df["emi_coll_no"]).clip(lower=0.0)
        elif mapping.get("emi"):
            df["due_emi_no"] = np.where(df["emi"] > 0, df["overdue_amt"] / df["emi"], 0.0)

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
    return df.reset_index(drop=True)


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
    od["main_driver"] = np.where(od["p_overdue"] >= od["p_emi"], "Overdue amount", "High Due EMI count")
    od["recommended_action"] = np.select(
        [od["due_emi_no"] >= 6, od["due_emi_no"] >= 3, od["due_emi_no"] >= 2],
        [
            "Legal Notice & Immediate Vehicle Seizure Team",
            "Field Visit by Branch Manager + Guarantor Call",
            "Collector Physical Visit within 48 Hours",
        ],
        default="Automated SMS & Phone Follow-up",
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
            "### 🤖 Automated Executive Briefing\n"
            f"{summary_text}\n\n"
            "**Strategic Action Plan:**\n"
            "1. **Focus on Top 20% Drivers (Pareto 80/20):** Direct immediate management attention to the top branches/categories contributing the largest share of overdue or variance.\n"
            "2. **Field Recovery & Follow-up:** Assign the top-ranked high-risk accounts to senior collectors with daily tracking.\n"
            "3. **Data Governance:** Use the auto-cleaned Excel workbook for board reporting to avoid duplicate or subtotal distortion."
        )
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    prompt = (
        "You are an Executive Management Consultant. Based strictly on the verified summary below "
        "(never invent numbers), write a concise executive briefing with: "
        "1) Executive Summary, 2) Key Risk & Performance Insights, and 3) 4 Actionable Recommendations.\n\n"
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
            return resp.json()["candidates"][0]["content"]["parts"][0]["text"]
        return f"### 🤖 Automated Executive Briefing\n{summary_text}"
    except Exception:
        return f"### 🤖 Automated Executive Briefing\n{summary_text}"


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
            {"Metric": "Total Active Accounts", "Value": f"{len(df):,}"},
            {"Metric": "Accounts with Overdue", "Value": f"{od_acc:,} ({od_acc/max(1,len(df))*100:.1f}%)"},
            {"Metric": "Total Outstanding Portfolio", "Value": fmt_num(tot_out, cur)},
            {"Metric": "Total Overdue Amount", "Value": fmt_num(tot_od, cur)},
            {"Metric": "Total Cumulative Collection", "Value": fmt_num(tot_coll, cur)},
            {"Metric": "Overall Collection Efficiency", "Value": f"{eff:.1f}%"},
        ]
    )
    _add_table_to_slide(s1, summary_df, top=1.4, max_rows=8)

    s2 = prs.slides.add_slide(blank)
    _add_slide_header(s2, "Portfolio Aging by Due EMI Bucket", "Outstanding & Overdue Exposure", rgb)
    ag = df.groupby("bucket", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"))
    _add_bar_chart_to_slide(s2, ag["bucket"], [("Outstanding", ag["Outstanding"]), ("Overdue", ag["Overdue"])])

    s3 = prs.slides.add_slide(blank)
    _add_slide_header(s3, "Branch-Wise Recovery Performance", "Top Branches by Overdue & Collection", rgb)
    br = df.groupby("branch", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum")).sort_values("Overdue", ascending=False).head(10)
    _add_bar_chart_to_slide(s3, br["branch"], [("Overdue", br["Overdue"]), ("Collected", br["Collected"])])

    s4 = prs.slides.add_slide(blank)
    _add_slide_header(s4, "Collector-Wise Performance", "Overdue vs Collected (Top 10 Collectors)", rgb)
    co = df.groupby("collector", as_index=False).agg(Accounts=("account_id", "count"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum"), Due=("total_due", "sum")).sort_values("Overdue", ascending=False).head(10)
    _add_bar_chart_to_slide(s4, co["collector"], [("Overdue", co["Overdue"]), ("Collected", co["Collected"])])

    s5 = prs.slides.add_slide(blank)
    _add_slide_header(s5, "Risk Status & Provisioning Estimate", "Management Risk Classification", rgb)
    rk = df.groupby("risk_status", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"))
    _add_table_to_slide(s5, rk, top=1.5, max_rows=10)

    s6 = prs.slides.add_slide(blank)
    _add_slide_header(s6, "Top Priority Accounts for Immediate Action", "Ranked by Multi-Factor Risk Score (0-100)", rgb)
    if not prio.empty:
        cols_show = ["rank", "account_id", "customer", "vehicle_reg", "branch", "collector", "due_emi_no", "overdue_amt", "score", "recommended_action"]
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
        grp = df.groupby(cat_col, as_index=False)[num_col].sum().sort_values(num_col, ascending=False).head(10)
        _add_bar_chart_to_slide(s2, grp[cat_col], [(num_col, grp[num_col])])

        s3 = prs.slides.add_slide(blank)
        _add_slide_header(s3, f"Detailed Summary by {cat_col}", f"Count, Sum & Average of {num_col}", rgb)
        grp_tbl = df.groupby(cat_col, as_index=False).agg(Records=(num_col, "count"), Total=(num_col, "sum"), Average=(num_col, "mean")).sort_values("Total", ascending=False)
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


st.sidebar.markdown("## 🚀 AI Analytics Studio v3.0")
company_name = st.sidebar.text_input("Company / Client Name", value="Abu Shaid Md. Talha")
brand_color = st.sidebar.color_picker("Brand Theme Colour", value="#1F3A5F")
currency_sym = st.sidebar.text_input("Currency Symbol", value="BDT")

st.sidebar.markdown("### 📂 1. Load Data or Portfolio Demo")
demo_choice = st.sidebar.selectbox(
    "Instant Demo Datasets (1-Click Test)",
    [
        "-- Upload My Own File --",
        "Demo 1: Vehicle Credit & EMI Recovery",
        "Demo 2: Messy Retail & E-Commerce Sales",
        "Demo 3: Corporate Finance & Budget Audit",
        "Demo 4: HR & Employee KPI Analytics",
    ],
)

uploaded_file = st.sidebar.file_uploader("Upload Any Excel, CSV, or PDF", type=["xlsx", "xls", "csv", "pdf"])
prev_file = st.sidebar.file_uploader("Optional: Previous Month File (for MoM Comparison)", type=["xlsx", "xls", "csv"])

raw_df, prev_raw_df = None, None

if demo_choice == "Demo 1: Vehicle Credit & EMI Recovery":
    raw_df = make_demo_credit()
elif demo_choice == "Demo 2: Messy Retail & E-Commerce Sales":
    raw_df = make_demo_messy_sales()
elif demo_choice == "Demo 3: Corporate Finance & Budget Audit":
    raw_df = make_demo_finance()
elif demo_choice == "Demo 4: HR & Employee KPI Analytics":
    raw_df = make_demo_hr()
elif uploaded_file is not None:
    fname = uploaded_file.name.lower()
    if fname.endswith(".csv"):
        raw_df = pd.read_csv(uploaded_file)
    elif fname.endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(uploaded_file)
        sheet = st.sidebar.selectbox("Select Sheet", xl.sheet_names)
        preview = xl.parse(sheet, header=None)
        auto_hdr = detect_best_header_row(preview)
        hdr_row = st.sidebar.number_input("Header Row (Auto-Detected)", min_value=1, max_value=max(1, len(preview)), value=auto_hdr + 1)
        raw_df = xl.parse(sheet, header=int(hdr_row - 1))
    elif fname.endswith(".pdf"):
        tables = []
        with pdfplumber.open(uploaded_file) as pdf:
            for p in pdf.pages:
                for t in p.extract_tables() or []:
                    if len(t) > 1:
                        tables.append(pd.DataFrame(t[1:], columns=t[0]))
        if tables:
            raw_df = pd.concat(tables, ignore_index=True)

if prev_file is not None:
    pfname = prev_file.name.lower()
    if pfname.endswith(".csv"):
        prev_raw_df = pd.read_csv(prev_file)
    elif pfname.endswith((".xlsx", ".xls")):
        pxl = pd.ExcelFile(prev_file)
        p_preview = pxl.parse(pxl.sheet_names[0], header=None)
        p_hdr = detect_best_header_row(p_preview)
        prev_raw_df = pxl.parse(pxl.sheet_names[0], header=int(p_hdr))

st.sidebar.markdown("---")
safe_table_mode = st.sidebar.checkbox("Firewall-Safe Table Mode (Recommended)", value=True)

st.markdown(
    f"""
    <div class="hero-banner" style="background: linear-gradient(135deg, {brand_color}, #0F172A);">
        <h2 style="margin:0; color:white;">{company_name} — Automated AI Data & Credit Intelligence Studio</h2>
        <p style="margin:0.35rem 0 0 0; opacity:0.92; font-size:0.95rem;">
            🤖 Zero-Config Auto-Detection • 🧹 Messy Data Auto-Fixer • 📊 Pareto & Time-Series BI • 💳 Credit Recovery Pro • 📥 1-Click PPT & Excel
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

if raw_df is None:
    st.info("👈 **Welcome!** Select any **Instant Demo Dataset** in the left sidebar or upload your own Excel/CSV/PDF file. The studio will automatically detect your data type, clean messy records, and build your dashboard & PowerPoint!")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown('<div class="kpi-card"><div class="kpi-title">Step 1</div><div class="kpi-val">Upload Any Data</div><div class="kpi-sub">Excel, CSV, or PDF tables</div></div>', unsafe_allow_html=True)
    c2.markdown('<div class="kpi-card"><div class="kpi-title">Step 2</div><div class="kpi-val">Auto-Clean & Fix</div><div class="kpi-sub">Removes totals, dups & currency text</div></div>', unsafe_allow_html=True)
    c3.markdown('<div class="kpi-card"><div class="kpi-title">Step 3</div><div class="kpi-val">AI Visual Analytics</div><div class="kpi-sub">Pareto 80/20, trends & risk scores</div></div>', unsafe_allow_html=True)
    c4.markdown('<div class="kpi-card"><div class="kpi-title">Step 4</div><div class="kpi-val">1-Click Deliverables</div><div class="kpi-sub">Editable PPTX, Excel & Fiverr report</div></div>', unsafe_allow_html=True)
    st.stop()

detected_credit = is_credit_dataset(raw_df.columns)
auto_label = "💳 Credit Recovery Intelligence Pro" if detected_credit else "📊 Universal Business Analytics & Cleaner"

mode_choice = st.sidebar.radio(
    "🧠 Studio Workflow Mode",
    [
        f"🤖 Smart Auto-Detect ({auto_label})",
        "📊 Force Universal Business Analytics",
        "💳 Force Credit Recovery Pro",
    ],
)

run_credit_mode = (mode_choice.startswith("🤖") and detected_credit) or mode_choice.startswith("💳")

clean_df, audit_log, raw_health, clean_health, detected_dates = auto_clean_dataframe(raw_df)

if run_credit_mode:
    st.markdown(
        f'<span class="badge-ok">✨ Auto-Detected Mode: Credit Recovery & EMI Intelligence Portfolio</span> '
        f'&nbsp; <span class="badge-ok">🧹 Data Health Score: {raw_health}% → {clean_health}%</span>',
        unsafe_allow_html=True,
    )

    auto_map = auto_map_credit_columns(raw_df.columns)
    with st.sidebar.expander("🔧 Column Mapping (Auto-Matched)", expanded=False):
        opts = [None] + list(raw_df.columns)
        user_map = {}
        for key, label, _ in CREDIT_FIELDS:
            default_col = auto_map.get(key)
            idx = opts.index(default_col) if default_col in opts else 0
            user_map[key] = st.selectbox(label, opts, index=idx, key=f"map_{key}")

    out_basis = st.sidebar.radio("Outstanding Basis", ["Net balance", "Balance receivable"])
    df_cred = prepare_credit_data(raw_df, user_map, out_basis)

    st.sidebar.markdown("### 🔍 Quick Filters")
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
    last_m_tot = df_f["last_m_coll"].sum()
    eff = (tot_coll / tot_due * 100) if tot_due > 0 else 0.0
    par1 = (df_f.loc[df_f["due_emi_no"] >= 1, "outstanding"].sum() / tot_out * 100) if tot_out > 0 else 0.0
    par3 = (df_f.loc[df_f["due_emi_no"] >= 3, "outstanding"].sum() / tot_out * 100) if tot_out > 0 else 0.0
    od_cnt = int((df_f["overdue_amt"] > 0).sum())

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Outstanding</div><div class="kpi-val">{fmt_num(tot_out, currency_sym)}</div><div class="kpi-sub">Basis: {out_basis}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Overdue</div><div class="kpi-val">{fmt_num(tot_od, currency_sym)}</div><div class="kpi-sub">{od_cnt:,} overdue accounts</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Collection</div><div class="kpi-val">{fmt_num(tot_coll, currency_sym)}</div><div class="kpi-sub">Last Month: {fmt_num(last_m_tot, currency_sym)}</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div class="kpi-title">Collection Efficiency</div><div class="kpi-val">{eff:.1f}%</div><div class="kpi-sub">Collected vs Total Due</div></div>', unsafe_allow_html=True)
    k5.markdown(f'<div class="kpi-card"><div class="kpi-title">PAR 1+ Exposure</div><div class="kpi-val">{par1:.1f}%</div><div class="kpi-sub">PAR 3+: {par3:.1f}%</div></div>', unsafe_allow_html=True)
    k6.markdown(f'<div class="kpi-card"><div class="kpi-title">Active Portfolio</div><div class="kpi-val">{len(df_f):,}</div><div class="kpi-sub">{len(branches)} Branches</div></div>', unsafe_allow_html=True)

    t1, t2, t3, t4, t5 = st.tabs(
        [
            "📊 1. Executive Dashboard",
            "🚨 2. Priority List & Auto Call/SMS Scripts",
            "🎯 3. Collector Grading & Provisioning",
            "🧹 4. Data Cleaning & MoM Comparison",
            "📥 5. 1-Click PPT, Excel & Client Certificate",
        ]
    )

    with t1:
        c1, c2 = st.columns(2)
        with c1:
            ag = df_f.groupby("bucket", as_index=False).agg(Accounts=("account_id", "count"), Overdue=("overdue_amt", "sum"), Outstanding=("outstanding", "sum"))
            fig1 = px.bar(ag, x="bucket", y=["Outstanding", "Overdue"], barmode="group", title="Portfolio Aging by Due EMI Bucket")
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            br_df = df_f.groupby("branch", as_index=False).agg(Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum"), Outstanding=("outstanding", "sum")).sort_values("Overdue", ascending=False).head(12)
            fig2 = px.bar(br_df, x="branch", y=["Overdue", "Collected"], barmode="group", title="Branch-Wise Overdue vs Total Collection")
            st.plotly_chart(fig2, use_container_width=True)

        c3, c4 = st.columns(2)
        with c3:
            co_df = df_f.groupby("collector", as_index=False).agg(Due=("total_due", "sum"), Collected=("total_coll", "sum"), Overdue=("overdue_amt", "sum"))
            co_df["Efficiency %"] = np.where(co_df["Due"] > 0, (co_df["Collected"] / co_df["Due"] * 100).round(1), 0.0)
            co_df = co_df.sort_values("Efficiency %", ascending=True)
            fig3 = px.bar(co_df, x="Efficiency %", y="collector", orientation="h", text="Efficiency %", title="Collector Efficiency Ranking (%)")
            st.plotly_chart(fig3, use_container_width=True)
        with c4:
            rk_df = df_f.groupby("risk_status", as_index=False).agg(Accounts=("account_id", "count"), Overdue=("overdue_amt", "sum"))
            fig4 = px.pie(rk_df, names="risk_status", values="Overdue", hole=0.45, title="Overdue Amount Share by Risk Tier")
            st.plotly_chart(fig4, use_container_width=True)

    with t2:
        st.markdown("#### 🚨 Automated Vehicle / Loan Recovery Priority List (0–100 Risk Score)")
        if prio_df.empty:
            st.success("🎉 Zero overdue accounts in the selected filter!")
        else:
            cols_p = ["rank", "account_id", "customer", "vehicle_reg", "branch", "collector", "due_emi_no", "overdue_amt", "outstanding", "score", "recommended_action"]
            show_table(prio_df[cols_p], safe_table_mode, 100)

            st.markdown("---")
            st.markdown("#### 📱 Automated Follow-Up SMS / WhatsApp Script Generator (For Top Overdue Customers)")
            top_acc = prio_df.iloc[0]
            s_col1, s_col2 = st.columns(2)
            with s_col1:
                st.markdown("**🇧🇩 বাংলা তাগাদা মেসেজ (Ready to Copy):**")
                st.code(
                    f"প্রিয় {top_acc['customer']} (A/C: {top_acc['account_id']}, গাড়ি: {top_acc['vehicle_reg']}),\n"
                    f"{company_name}-এ আপনার {top_acc['due_emi_no']:.0f}টি কিস্তি বাবদ মোট {currency_sym} {top_acc['overdue_amt']:,.0f} টাকা বকেয়া রয়েছে। "
                    f"আইনি ব্যবস্থা এড়াতে আগামী ৪৮ ঘণ্টার মধ্যে বকেয়া পরিশোধ করার জন্য বিশেষভাবে অনুরোধ করা হলো।",
                    language="text",
                )
            with s_col2:
                st.markdown("**🇬🇧 English Follow-Up Notice (Ready to Copy):**")
                st.code(
                    f"Dear {top_acc['customer']} (Account: {top_acc['account_id']}, Vehicle: {top_acc['vehicle_reg']}),\n"
                    f"Your account currently has {top_acc['due_emi_no']:.0f} overdue installment(s) totaling {currency_sym} {top_acc['overdue_amt']:,.0f}. "
                    f"Please settle the overdue balance within 48 hours to avoid recovery escalation. — {company_name}",
                    language="text",
                )

    with t3:
        st.markdown("#### 🎯 Automated Collector Performance Grading")
        cg = df_f.groupby("collector", as_index=False).agg(
            Accounts=("account_id", "count"),
            Monthly_EMI_Target=("emi", "sum"),
            Last_Month_Collection=("last_m_coll", "sum"),
            Cumulative_Due=("total_due", "sum"),
            Cumulative_Collected=("total_coll", "sum"),
            Current_Overdue=("overdue_amt", "sum"),
        )
        cg["Overall_Efficiency_%"] = np.where(cg["Cumulative_Due"] > 0, (cg["Cumulative_Collected"] / cg["Cumulative_Due"] * 100).round(1), 0.0)
        cg["Grade"] = np.select(
            [cg["Overall_Efficiency_%"] >= 95, cg["Overall_Efficiency_%"] >= 88, cg["Overall_Efficiency_%"] >= 75],
            ["A+ (Excellent)", "A (Good)", "B (Satisfactory)"],
            default="C (Needs Action)",
        )
        cg = cg.sort_values("Overall_Efficiency_%", ascending=False)
        show_table(cg, safe_table_mode, 30)

        st.markdown("#### 🛡️ Automated Management Provisioning Table (Estimate)")
        prov = df_f.groupby("bucket", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"))
        rate_map = {"Current (0 EMI)": 0.0, "1 EMI": 5.0, "2 EMI": 25.0, "3 EMI": 50.0, "4-5 EMI": 50.0, "6+ EMI": 100.0}
        prov["Provision_Rate_%"] = prov["bucket"].map(rate_map).fillna(0.0)
        prov["Estimated_Provision"] = (prov["Outstanding"] * prov["Provision_Rate_%"] / 100.0).round(0)
        show_table(prov, safe_table_mode, 15)

    with t4:
        st.markdown("#### 🧹 Automated Data Cleaning Audit Log")
        for item in audit_log:
            st.markdown(f'<div class="audit-item">✔️ {item}</div>', unsafe_allow_html=True)

        if prev_raw_df is not None:
            st.markdown("#### 🔄 Automated Month-over-Month (MoM) Portfolio Comparison")
            prev_map = auto_map_credit_columns(prev_raw_df.columns)
            df_prev = prepare_credit_data(prev_raw_df, prev_map, out_basis)
            mom_df = pd.DataFrame(
                [
                    {"Metric": "Total Accounts", "Previous Month": len(df_prev), "Current Month": len(df_cred), "Change": len(df_cred) - len(df_prev)},
                    {"Metric": "Total Outstanding", "Previous Month": df_prev["outstanding"].sum(), "Current Month": df_cred["outstanding"].sum(), "Change": df_cred["outstanding"].sum() - df_prev["outstanding"].sum()},
                    {"Metric": "Total Overdue", "Previous Month": df_prev["overdue_amt"].sum(), "Current Month": df_cred["overdue_amt"].sum(), "Change": df_cred["overdue_amt"].sum() - df_prev["overdue_amt"].sum()},
                ]
            )
            show_table(mom_df, safe_table_mode, 10)
        else:
            st.info("💡 Tip: Upload a 'Previous Month File' in the sidebar anytime to unlock automated Month-over-Month & Roll-Rate comparisons.")

    with t5:
        cred_summary_txt = (
            f"- **Total Active Portfolio:** {len(df_f):,} accounts across {len(branches)} branches.\n"
            f"- **Total Outstanding ({out_basis}):** {fmt_num(tot_out, currency_sym)}.\n"
            f"- **Total Overdue Exposure:** {fmt_num(tot_od, currency_sym)} across {od_cnt:,} accounts.\n"
            f"- **Collection Efficiency:** {eff:.1f}% (PAR 1+: {par1:.1f}%, PAR 3+: {par3:.1f}%)."
        )
        ai_cred = generate_ai_insights(cred_summary_txt)
        st.markdown(ai_cred)

        st.markdown("---")
        b1, b2 = st.columns(2)
        with b1:
            if st.button("📊 1-Click Build Executive PowerPoint (.pptx)"):
                st.session_state["cred_ppt"] = build_credit_ppt(
                    df_f, prio_df, company_name, currency_sym, hex_to_rgb(brand_color), ai_cred
                )
            if "cred_ppt" in st.session_state:
                st.download_button(
                    "⬇️ Download Credit Recovery PPT (.pptx)",
                    st.session_state["cred_ppt"],
                    file_name="Credit_Recovery_Executive_Presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        with b2:
            if st.button("📗 1-Click Build Master Excel Report (.xlsx)"):
                sheets_c = {
                    "Executive_Summary": pd.DataFrame(
                        [
                            {"Metric": "Total Accounts", "Value": len(df_f)},
                            {"Metric": "Overdue Accounts", "Value": od_cnt},
                            {"Metric": "Total Outstanding", "Value": tot_out},
                            {"Metric": "Total Overdue", "Value": tot_od},
                            {"Metric": "Total Collection", "Value": tot_coll},
                            {"Metric": "Collection Efficiency %", "Value": round(eff, 2)},
                        ]
                    ),
                    "Recovery_Priority_List": prio_df,
                    "Branch_Performance": df_f.groupby("branch", as_index=False).agg(Accounts=("account_id", "count"), Outstanding=("outstanding", "sum"), Overdue=("overdue_amt", "sum"), Collected=("total_coll", "sum")),
                    "Collector_Grading": cg,
                    "Provisioning_Estimate": prov,
                    "Cleaned_Portfolio_Data": df_f,
                }
                st.session_state["cred_xlsx"] = build_multi_sheet_excel(sheets_c)
            if "cred_xlsx" in st.session_state:
                st.download_button(
                    "⬇️ Download Multi-Sheet Excel Report (.xlsx)",
                    st.session_state["cred_xlsx"],
                    file_name="Credit_Recovery_Master_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        st.markdown("---")
        st.markdown("#### 💼 Ready-to-Send Fiverr / Upwork Client Delivery Note")
        st.code(
            f"Hi! I have completed the full Credit Recovery & Portfolio Intelligence audit for {company_name}.\n\n"
            f"✅ Data Quality & Health Score: Improved from {raw_health}% to {clean_health}%\n"
            f"✅ Total Active Accounts Analyzed: {len(df_f):,}\n"
            f"✅ Total Portfolio Outstanding: {fmt_num(tot_out, currency_sym)}\n"
            f"✅ Total Overdue Identified: {fmt_num(tot_od, currency_sym)} ({od_cnt:,} accounts)\n"
            f"✅ Overall Collection Efficiency: {eff:.1f}%\n\n"
            "Attached Deliverables:\n"
            "1. Executive PowerPoint Presentation (.pptx) with native editable charts\n"
            "2. Multi-Sheet Master Excel Report (.xlsx) including 0-100 Recovery Priority List, Collector Grading & Provisioning",
            language="text",
        )

else:
    st.markdown(
        f'<span class="badge-ok">✨ Auto-Detected Mode: Universal Business Analytics & Data Cleaner</span> '
        f'&nbsp; <span class="badge-ok">🧹 Data Health Score: {raw_health}% → {clean_health}%</span>',
        unsafe_allow_html=True,
    )

    num_cols = clean_df.select_dtypes(include=[np.number]).columns.tolist()
    cat_cols = [c for c in clean_df.columns if c not in num_cols]

    primary_cat = cat_cols[0] if cat_cols else clean_df.columns[0]
    primary_num = num_cols[0] if num_cols else None

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-title">Raw → Clean Rows</div><div class="kpi-val">{len(raw_df):,} → {len(clean_df):,}</div><div class="kpi-sub">Auto-verified</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-title">Data Health Score</div><div class="kpi-val">{clean_health}%</div><div class="kpi-sub">Up from {raw_health}%</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div class="kpi-title">Numeric Metrics</div><div class="kpi-val">{len(num_cols)}</div><div class="kpi-sub">{len(cat_cols)} Category cols</div></div>', unsafe_allow_html=True)
    if primary_num:
        k4.markdown(f'<div class="kpi-card"><div class="kpi-title">Total {primary_num}</div><div class="kpi-val">{fmt_num(clean_df[primary_num].sum(), currency_sym)}</div><div class="kpi-sub">Primary Metric</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="kpi-card"><div class="kpi-title">Average {primary_num}</div><div class="kpi-val">{fmt_num(clean_df[primary_num].mean(), currency_sym, 1)}</div><div class="kpi-sub">Per Record</div></div>', unsafe_allow_html=True)
    else:
        k4.markdown('<div class="kpi-card"><div class="kpi-title">Total Columns</div><div class="kpi-val">' + str(clean_df.shape[1]) + '</div></div>', unsafe_allow_html=True)
        k5.markdown('<div class="kpi-card"><div class="kpi-title">Status</div><div class="kpi-val">Cleaned</div></div>', unsafe_allow_html=True)

    u1, u2, u3, u4 = st.tabs(
        [
            "📊 1. Automated Business & Pareto (80/20) Dashboard",
            "🧹 2. Messy Data Auto-Cleaner (Before vs After)",
            "🔍 3. Time-Series, Correlations & Anomalies",
            "📥 4. 1-Click PPT, Clean Excel & Fiverr Certificate",
        ]
    )

    with u1:
        if not num_cols:
            st.warning("No numeric columns found. Check the 'Messy Data Auto-Cleaner' tab to inspect your data.")
        else:
            sel_c1, sel_c2 = st.columns(2)
            primary_cat = sel_c1.selectbox("Category Dimension (Auto-Selected)", cat_cols if cat_cols else clean_df.columns)
            primary_num = sel_c2.selectbox("Numeric Metric (Auto-Selected)", num_cols)

            grp = (
                clean_df.groupby(primary_cat, as_index=False)
                .agg(Records=(primary_num, "count"), Total=(primary_num, "sum"), Average=(primary_num, "mean"))
                .sort_values("Total", ascending=False)
            )
            tot_sum = max(1e-9, grp["Total"].sum())
            grp["Share_%"] = (grp["Total"] / tot_sum * 100).round(1)
            grp["Cumulative_Share_%"] = grp["Share_%"].cumsum().round(1)

            ch1, ch2 = st.columns(2)
            with ch1:
                fig_b = px.bar(grp.head(15), x=primary_cat, y="Total", text="Share_%", title=f"Top {primary_cat} by {primary_num} (Pareto Ranking)", color_discrete_sequence=[brand_color])
                st.plotly_chart(fig_b, use_container_width=True)
            with ch2:
                fig_p = px.pie(grp.head(10), names=primary_cat, values="Total", hole=0.45, title=f"Share Breakdown by {primary_cat}")
                st.plotly_chart(fig_p, use_container_width=True)

            st.markdown(f"#### 📋 Automated Pareto (80/20) Breakdown Table by `{primary_cat}`")
            show_table(grp, safe_table_mode, 50)

    with u2:
        st.markdown("#### ✅ Automated Cleaning & Standardization Log")
        for item in audit_log:
            st.markdown(f'<div class="audit-item">✔️ {item}</div>', unsafe_allow_html=True)

        cb, ca = st.columns(2)
        with cb:
            st.markdown("##### ❌ Before Cleaning (Raw Messy Data)")
            show_table(raw_df, safe_table_mode, 25)
        with ca:
            st.markdown("##### ✅ After Cleaning (100% Standardized Data)")
            show_table(clean_df, safe_table_mode, 25)

    with u3:
        if detected_dates and num_cols:
            st.markdown("#### 📅 Automated Time-Series Trend")
            d_col = detected_dates[0]
            ts_df = clean_df.groupby(d_col, as_index=False)[primary_num].sum().sort_values(d_col)
            fig_ts = px.line(ts_df, x=d_col, y=primary_num, markers=True, title=f"{primary_num} Trend Over Time ({d_col})")
            st.plotly_chart(fig_ts, use_container_width=True)

        if len(num_cols) >= 2:
            st.markdown("#### 🔗 Automated Correlation Matrix")
            corr = clean_df[num_cols].corr().round(2).reset_index()
            show_table(corr, safe_table_mode, 20)

        if num_cols:
            st.markdown("#### 🚨 Statistical Outlier / Anomaly Detector (Z-Score > 2.0)")
            s_num = clean_df[primary_num].dropna()
            if len(s_num) > 3 and s_num.std() > 0:
                z = (clean_df[primary_num] - s_num.mean()) / s_num.std()
                outliers = clean_df[z.abs() >= 2.0].sort_values(primary_num, ascending=False)
                st.write(f"Detected **{len(outliers)}** unusual/outlier record(s) in `{primary_num}`:")
                show_table(outliers, safe_table_mode, 30)
            else:
                st.info("No extreme statistical outliers detected in the primary metric.")

    with u4:
        summary_lines = [
            f"- **Dataset Overview:** {len(clean_df):,} clean records across {clean_df.shape[1]} columns (Data Health Score improved from {raw_health}% to {clean_health}%).",
            f"- **Automated Cleaning Actions:** {'; '.join(audit_log)}",
        ]
        if primary_num:
            top_g = clean_df.groupby(primary_cat)[primary_num].sum().sort_values(ascending=False)
            if len(top_g) > 0:
                t_name, t_val = top_g.index[0], top_g.iloc[0]
                t_share = t_val / max(1e-9, clean_df[primary_num].sum()) * 100
                summary_lines.append(f"- **Key Metric ({primary_num}):** Total = {fmt_num(clean_df[primary_num].sum(), currency_sym)}, Average = {fmt_num(clean_df[primary_num].mean(), currency_sym, 1)}.")
                summary_lines.append(f"- **Top Performer ({primary_cat}):** '{t_name}' leads with {fmt_num(t_val, currency_sym)} ({t_share:.1f}% of total).")

        ai_uni = generate_ai_insights("\n".join(summary_lines))
        st.markdown(ai_uni)

        st.markdown("---")
        b1, b2 = st.columns(2)
        with b1:
            if st.button("📊 1-Click Build Executive PowerPoint (.pptx)"):
                st.session_state["uni_ppt"] = build_universal_ppt(
                    clean_df, company_name, primary_cat, primary_num, hex_to_rgb(brand_color), ai_uni
                )
            if "uni_ppt" in st.session_state:
                st.download_button(
                    "⬇️ Download Executive PowerPoint (.pptx)",
                    st.session_state["uni_ppt"],
                    file_name="Executive_Business_Analytics_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        with b2:
            if st.button("📗 1-Click Build Clean Excel & Audit Report (.xlsx)"):
                sheets_u = {
                    "Cleaned_Master_Data": clean_df,
                    "Data_Cleaning_Audit_Log": pd.DataFrame({"Step": range(1, len(audit_log) + 1), "Automated_Action": audit_log}),
                    "Statistical_Summary": clean_df.describe().T.reset_index(),
                }
                if primary_num:
                    sheets_u["Pareto_Category_Summary"] = grp
                st.session_state["uni_xlsx"] = build_multi_sheet_excel(sheets_u)
            if "uni_xlsx" in st.session_state:
                st.download_button(
                    "⬇️ Download Cleaned & Analyzed Excel (.xlsx)",
                    st.session_state["uni_xlsx"],
                    file_name="Cleaned_Business_Analytics_Master.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        st.markdown("---")
        st.markdown("#### 💼 Ready-to-Send Fiverr / Upwork Client Delivery Certificate")
        st.code(
            f"Hi! I have completed the automated data cleaning, validation, and executive business analysis for {company_name}.\n\n"
            f"✅ Raw Data Health Score: {raw_health}% → Cleaned Data Health Score: {clean_health}%\n"
            f"✅ Total Clean Records Delivered: {len(clean_df):,} rows across {clean_df.shape[1]} standardized columns\n"
            f"✅ Audit Actions Completed: {'; '.join(audit_log)}\n\n"
            "Attached Files:\n"
            "1. Cleaned_Business_Analytics_Master.xlsx (100% standardized data + Pareto summary + Audit Log)\n"
            "2. Executive_Business_Analytics_Deck.pptx (Board-ready presentation with native editable charts)",
            language="text",
        )
