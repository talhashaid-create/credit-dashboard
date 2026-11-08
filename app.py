"""Enterprise v3.1: ArrowTypeError-Proof AI Data Cleaner, Business Analytics & Credit Studio.
Run locally: streamlit run app.py
"""
import hmac
import io
import math
import os
import re
import warnings
from datetime import date, datetime, time

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
try:
    pd.set_option("future.infer_string", False)
except Exception:
    pass

st.set_page_config(
    page_title="AI Data & Credit Intelligence Studio v3.1",
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


def safe_str_val(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    if isinstance(x, (datetime, date)):
        return x.strftime("%Y-%m-%d")
    if isinstance(x, time):
        return x.strftime("%H:%M:%S")
    return str(x).strip()


def sanitize_df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            out[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            out[c] = [safe_str_val(v) for v in df[c].tolist()]
    return out


def show_table(df: pd.DataFrame, safe_mode: bool = True, max_rows: int = 100):
    if df is None or df.empty:
        st.caption("No records to display.")
        return
    view = sanitize_df_for_display(df.head(max_rows))
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
            st.dataframe(view, use_container_width=True)
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
HEADER_KEYWORDS = {
    "serial", "sl", "customer", "branch", "sales", "account", "code", "collector",
    "collection", "vehicle", "reg", "price", "down", "interest", "bill", "balance",
    "receivable", "emi", "installment", "due", "overdue", "net", "remarks", "date",
}


def detect_best_header_row(raw_df: pd.DataFrame) -> int:
    best_row, best_score = 0, -1e9
    limit = min(40, len(raw_df))
    for r in range(limit):
        vals = [safe_str_val(x) for x in raw_df.iloc[r].tolist()]
        str_cells = [s for s in vals if s != ""]
        if len(str_cells) < 2:
            continue
        alpha_cnt = sum(1 for s in str_cells if any(c.isalpha() for c in s))
        kw_hits = sum(
            1 for s in str_cells
            if any(k in re.sub(r"[^a-z0-9]+", " ", s.lower()).split() for k in HEADER_KEYWORDS)
        )
        uniq_ratio = len(set(s.lower() for s in str_cells)) / max(1, len(str_cells))
        score = alpha_cnt * 3 + kw_hits * 8 + uniq_ratio * 5 - r * 0.1
        if score > best_score:
            best_score, best_row = score, r
    return best_row


def parse_messy_number_list(str_list: list):
    non_empty = [s for s in str_list if s != ""]
    if not non_empty:
        return None, 0
    cleaned_all = []
    for s in str_list:
        if s == "":
            cleaned_all.append(np.nan)
            continue
        c = re.sub(r"^\((.*)\)$", r"-\1", s)
        c = re.sub(r"(?i)(bdt|tk|taka|usd|eur|inr|rs|৳|\$|€|£|/-|,|\s)", "", c)
        c = re.sub(r"%$", "", c)
        cleaned_all.append(c)
    num = pd.to_numeric(pd.Series(cleaned_all), errors="coerce")
    valid_cnt = int(num.notna().sum())
    if valid_cnt / max(1, len(non_empty)) >= 0.72:
        return num, valid_cnt
    return None, 0


def compute_health_score_safe(df: pd.DataFrame) -> float:
    if df is None or df.empty:
        return 0.0
    total_cells = max(1, df.shape[0] * df.shape[1])
    empty_cnt = 0
    for c in df.columns:
        vals = [safe_str_val(v) for v in df[c].tolist()]
        empty_cnt += sum(1 for v in vals if v == "" or v.lower() in {"nan", "none", "null"})
    missing_ratio = empty_cnt / total_cells
    unnamed_ratio = sum(1 for c in df.columns if str(c).lower().startswith("unnamed")) / max(1, len(df.columns))
    score = 100.0 - (missing_ratio * 45.0) - (unnamed_ratio * 25.0)
    return round(float(np.clip(score, 15.0, 100.0)), 1)


def auto_clean_dataframe(
    df_in: pd.DataFrame,
    drop_totals: bool = True,
    drop_dups: bool = True,
    standardize_text: bool = True,
    missing_num_policy: str = "Fill with 0",
    missing_txt_policy: str = "Fill with 'Unknown'",
):
    raw_health = compute_health_score_safe(df_in)
    audit = []

    # 1. Clean column headers safely
    new_cols, seen = [], {}
    for i, c in enumerate(df_in.columns):
        name = re.sub(r"\s+", " ", safe_str_val(c))
        if not name or name.lower().startswith("unnamed"):
            name = f"Column_{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        new_cols.append(name)

    # 2. Build a type-safe intermediate DataFrame (immune to mixed-type ArrowTypeError)
    clean_cols_dict = {}
    empty_col_cnt = 0
    for idx, col_name in enumerate(new_cols):
        raw_series = df_in.iloc[:, idx]
        if pd.api.types.is_numeric_dtype(raw_series):
            num_s = pd.to_numeric(raw_series, errors="coerce")
            if num_s.isna().all():
                empty_col_cnt += 1
                continue
            clean_cols_dict[col_name] = num_s
        else:
            str_vals = [safe_str_val(v) for v in raw_series.tolist()]
            str_vals = ["" if v.lower() in {"nan", "none", "null", "nat"} else v for v in str_vals]
            if all(v == "" for v in str_vals):
                empty_col_cnt += 1
                continue
            clean_cols_dict[col_name] = pd.Series(str_vals, index=df_in.index, dtype="object")

    if empty_col_cnt > 0:
        audit.append(f"Removed {empty_col_cnt} completely empty column(s).")

    df = pd.DataFrame(clean_cols_dict, index=df_in.index)
    if df.empty:
        return pd.DataFrame(), ["Dataset is empty."], raw_health, raw_health, []

    # 3. Drop completely blank rows
    non_empty_row_mask = pd.Series(False, index=df.index)
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            non_empty_row_mask = non_empty_row_mask | df[c].notna()
        else:
            non_empty_row_mask = non_empty_row_mask | (df[c] != "")
    blank_rows = int((~non_empty_row_mask).sum())
    if blank_rows > 0:
        df = df.loc[non_empty_row_mask].copy()
        audit.append(f"Removed {blank_rows} completely blank row(s).")

    # 4. Remove embedded Subtotal / Grand Total rows
    if drop_totals and len(df) > 0:
        check_cols = df.columns[: min(6, len(df.columns))]
        mask_total = pd.Series(False, index=df.index)
        for col in check_cols:
            col_strs = [safe_str_val(v) for v in df[col].tolist()]
            matches = [bool(TOTAL_PAT.match(s)) for s in col_strs]
            mask_total = mask_total | pd.Series(matches, index=df.index)
        tot_cnt = int(mask_total.sum())
        if tot_cnt > 0:
            df = df.loc[~mask_total].copy()
            audit.append(f"Removed {tot_cnt} Subtotal / Grand Total summary row(s).")

    # 5. Trim & standardize text categories + convert messy numbers
    txt_fixed = 0
    num_cols_converted = []
    date_cols = []

    for col in list(df.columns):
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        orig_list = df[col].tolist()
        clean_list = [re.sub(r"\s+", " ", s.strip()) for s in orig_list]

        # Check if messy numeric column
        col_low = col.lower()
        is_id_like = any(k in col_low for k in ["id", "serial", "account", "phone", "mobile", "reg", "vehicle", "code", "invoice", "voucher", "emp", "name", "branch", "collector", "personnel", "remark"])
        if not is_id_like:
            num_series, conv_cnt = parse_messy_number_list(clean_list)
            if num_series is not None:
                df[col] = num_series.values
                num_cols_converted.append(f"{col} ({conv_cnt})")
                continue

        # Check if date column
        if any(k in col_low for k in ["date", "month", "period"]):
            parsed = pd.to_datetime(pd.Series(clean_list).replace("", np.nan), errors="coerce")
            if parsed.notna().mean() >= 0.6:
                df[col] = parsed.dt.strftime("%Y-%m-%d").fillna("Unknown").values
                date_cols.append(col)
                continue

        # Standardize low-cardinality category casing
        non_empty = [s for s in clean_list if s != ""]
        if standardize_text and 0 < len(set(non_empty)) <= 45:
            if len(set(s.lower() for s in non_empty)) < len(set(non_empty)):
                clean_list = [s.title() if s != "" else "" for s in clean_list]

        txt_fixed += sum(1 for a, b in zip(orig_list, clean_list) if a != b)
        df[col] = pd.Series(clean_list, index=df.index, dtype="object")

    if txt_fixed > 0:
        audit.append(f"Trimmed extra spaces & standardized text in {txt_fixed} cell(s).")
    if num_cols_converted:
        audit.append("Converted formatted currency/number strings to numeric: " + ", ".join(num_cols_converted))
    if date_cols:
        audit.append("Standardized date formatting (YYYY-MM-DD) in: " + ", ".join(date_cols))

    # 6. Drop exact duplicates safely
    if drop_dups and len(df) > 1:
        before = len(df)
        df = df.drop_duplicates().copy()
        if before - len(df) > 0:
            audit.append(f"Removed {before - len(df)} exact duplicate row(s).")

    # 7. Missing value policies
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    txt_cols = [c for c in df.columns if c not in num_cols]

    if missing_num_policy == "Fill with 0" and num_cols:
        na_cnt = int(df[num_cols].isna().sum().sum())
        if na_cnt > 0:
            df[num_cols] = df[num_cols].fillna(0.0)
            audit.append(f"Filled {na_cnt} missing numeric cell(s) with 0.")
    elif missing_num_policy == "Fill with Column Median" and num_cols:
        for c in num_cols:
            na_c = int(df[c].isna().sum())
            if na_c > 0 and df[c].notna().any():
                med = float(df[c].median())
                df[c] = df[c].fillna(med)
                audit.append(f"Filled {na_c} missing 
