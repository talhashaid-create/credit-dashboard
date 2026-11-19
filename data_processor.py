import re
import difflib
import numpy as np
import pandas as pd
from typing import Dict, List, Tuple
from datetime import datetime

from config import COLUMN_PATTERNS, REQUIRED_FIELDS, NUMERIC_FIELDS, DATE_FIELDS

def clean_numeric_value(value) -> float:
    if pd.isna(value): return np.nan
    s = str(value).strip()
    if s == "" or s.lower() in ("nil", "na", "n/a", "-", "none", "null", "#n/a"): return np.nan
    s = re.sub(r"[₹$€£,\s]", "", s)
    if s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    s = re.sub(r"[^\d.\-]", "", s)
    try: return float(s)
    except: return np.nan

def clean_date_value(value) -> pd.Timestamp:
    if pd.isna(value): return pd.NaT
    s = str(value).strip()
    if s == "" or s.lower() in ("nil", "na", "n/a", "-", "none"): return pd.NaT
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%d %b %Y", "%b %d, %Y", "%Y/%m/%d"):
        try: return pd.Timestamp(datetime.strptime(s, fmt))
        except: continue
    try: return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except: return pd.NaT

def read_excel_file(uploaded_file) -> Dict[str, pd.DataFrame]:
    try:
        sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=str, engine="openpyxl")
        # Drop fully empty rows/cols
        sheets = {name: df.dropna(how="all").dropna(axis=1, how="all") for name, df in sheets.items() if df is not None and not df.empty}
        return sheets
    except Exception as exc:
        raise ValueError(f"Could not read Excel file: {exc}")

def detect_best_sheet(sheets: Dict[str, pd.DataFrame]) -> str:
    scores = {}
    for name, df in sheets.items():
        score = 0.0
        cols_lower = [str(c).lower().strip() for c in df.columns]
        for _field, patterns in COLUMN_PATTERNS.items():
            for pat in patterns:
                for cl in cols_lower:
                    if pat in cl or cl in pat:
                        score += 2
                        break
        score += min(len(df) / 20, 5)
        scores[name] = score
    return max(scores, key=scores.get) if scores else list(sheets)[0]

def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    mapping = {}
    used_cols = set()
    cols_lower = {str(c).lower().strip(): c for c in df.columns}

    for field, patterns in COLUMN_PATTERNS.items():
        best_col, best_score = None, 0.0
        for cl_lower, original in cols_lower.items():
            if original in used_cols: continue
            for pat in patterns:
                score = difflib.SequenceMatcher(None, cl_lower, pat).ratio()
                if pat in cl_lower or cl_lower in pat: score = max(score, 0.85)
                if score > best_score and score >= 0.60:
                    best_score, best_col = score, original
        if best_col:
            mapping[field] = best_col
            used_cols.add(best_col)
    return mapping

def process_data(raw_df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    df = raw_df.copy()
    rename = {v: k for k, v in mapping.items()}
    df = df.rename(columns=rename)
    
    keep = [c for c in rename.values() if c in df.columns]
    df = df[keep].copy()

    for field in NUMERIC_FIELDS:
        if field in df.columns: df[field] = df[field].map(clean_numeric_value)
    for field in DATE_FIELDS:
        if field in df.columns: df[field] = df[field].map(clean_date_value)

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().replace(["", "nan", "NaN", "None"], np.nan)

    # শুধুমাত্র হুবহু একই রকম ডুপ্লিকেট রো ডিলিট করবে, কাস্টমারের অন্য কিস্তি ডিলিট করবে না
    df = df.drop_duplicates(keep="first")

    for field in NUMERIC_FIELDS:
        if field in df.columns: df[field] = df[field].fillna(0)

    if "outstanding_amount" in df.columns and "collection_amount" in df.columns:
        df["collection_efficiency"] = np.where(df["outstanding_amount"] > 0, (df["collection_amount"] / df["outstanding_amount"] * 100).round(2), 0)

    if "overdue_days" not in df.columns and "due_date" in df.columns:
        df["overdue_days"] = np.where(df["due_date"].notna(), (pd.Timestamp.today() - df["due_date"]).dt.days.clip(lower=0), 0)

    def _assign_risk(row):
        od, ce, ov = row.get("overdue_days", 0) or 0, row.get("collection_efficiency", 100) or 0, row.get("overdue_amount", 0) or 0
        if od >= 180 or (ov > 0 and ce < 10): return "🔴 Critical"
        if od >= 90 or (ov > 0 and ce < 30): return "🟠 High"
        if od >= 30 or (ov > 0 and ce < 60): return "🟡 Medium"
        return "🟢 Low"

    df["risk_category"] = df.apply(_assign_risk, axis=1)
    return df.reset_index(drop=True)

def data_quality_report(raw_df, clean_df, mapping):
    return {
        "total_rows_source": len(raw_df),
        "total_rows_cleaned": len(clean_df),
        "duplicates_removed": max(len(raw_df) - len(clean_df), 0),
        "columns_mapped": len(mapping),
    }

def generate_sample_data(n=200):
    np.random.seed(42)
    models = ["Toyota Fortuner", "Hyundai Creta", "Mahindra XUV700", "Tata Nexon", "Kia Seltos"]
    collectors = ["Rajesh Kumar", "Priya Sharma", "Amit Patel", "Sneha Reddy"]
    rows = []
    for i in range(1, n + 1):
        total = np.random.randint(500000, 3000000)
        collection = np.random.randint(0, int(total * 0.95))
        rows.append({
            "customer_id": f"CR-{i:05d}", "customer_name": f"Customer {i:04d}",
            "vehicle_model": np.random.choice(models), "collector_name": np.random.choice(collectors),
            "total_amount": total, "collection_amount": collection, "outstanding_amount": total - collection,
            "overdue_amount": np.random.randint(0, (total - collection) + 1) if (total - collection) > 0 else 0,
            "overdue_days": int(np.random.choice([0, 15, 30, 45, 60, 90, 120, 180, 240])),
            "due_date": pd.Timestamp("2024-01-01") + pd.Timedelta(days=np.random.randint(0, 365)),
            "region": np.random.choice(["North", "South", "East", "West"])
        })
    return pd.DataFrame(rows)
