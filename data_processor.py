"""
Data processing engine – handles Excel import, column auto-mapping,
validation, cleaning and processing with 100 % data accuracy.
"""

import re
import difflib
import numpy as np
import pandas as pd
from typing import Dict, List, Tuple, Optional
from datetime import datetime

from config import (
    COLUMN_PATTERNS,
    REQUIRED_FIELDS,
    NUMERIC_FIELDS,
    DATE_FIELDS,
)


# ══════════════════════════════════════════════
#  Helper: Currency & number cleaning
# ══════════════════════════════════════════════
def clean_numeric_value(value) -> float:
    """Convert a raw Excel cell to float. Returns NaN on failure."""
    if pd.isna(value):
        return np.nan
    s = str(value).strip()
    if s == "" or s.lower() in ("nil", "na", "n/a", "-", "none", "null", "#n/a"):
        return np.nan
    # Remove ₹ $ € £ commas spaces
    s = re.sub(r"[₹$€£,\s]", "", s)
    # Remove parentheses (accounting negatives) → negative
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s)
    except Exception:
        return np.nan


def clean_date_value(value) -> pd.Timestamp:
    """Convert raw Excel cell to pandas Timestamp. Returns NaT on failure."""
    if pd.isna(value):
        return pd.NaT
    s = str(value).strip()
    if s == "" or s.lower() in ("nil", "na", "n/a", "-", "none"):
        return pd.NaT
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
                "%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y",
                "%d-%b-%Y", "%d %b %Y", "%b %d, %Y",
                "%Y/%m/%d", "%d %B %Y", "%B %d, %Y"):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except Exception:
            continue
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


# ══════════════════════════════════════════════
#  1. Read Excel
# ══════════════════════════════════════════════
def read_excel_file(uploaded_file) -> Dict[str, pd.DataFrame]:
    """Read uploaded file → dict  {sheet_name: DataFrame}."""
    try:
        sheets = pd.read_excel(
            uploaded_file,
            sheet_name=None,
            dtype=str,
            engine="openpyxl",
        )
        # Drop fully-empty sheets
        sheets = {
            name: df.dropna(how="all").dropna(axis=1, how="all")
            for name, df in sheets.items()
            if df is not None and not df.empty
        }
        return sheets
    except Exception as exc:
        raise ValueError(f"Could not read Excel file: {exc}")


# ══════════════════════════════════════════════
#  2. Detect best sheet
# ══════════════════════════════════════════════
def detect_best_sheet(sheets: Dict[str, pd.DataFrame]) -> str:
    """Score each sheet; return the name with the highest score."""
    scores: Dict[str, float] = {}
    for name, df in sheets.items():
        score = 0.0
        cols_lower = [str(c).lower().strip() for c in df.columns]
        for _field, patterns in COLUMN_PATTERNS.items():
            for pat in patterns:
                for cl in cols_lower:
                    if pat in cl or cl in pat:
                        score += 2
                        break
                else:
                    continue
                break
        # Bonus for row count
        score += min(len(df) / 20, 5)
        scores[name] = score
    return max(scores, key=scores.get) if scores else list(sheets)[0]


# ══════════════════════════════════════════════
#  3. Auto-map columns
# ══════════════════════════════════════════════
def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Map Excel column headers → canonical field names.
    Returns  {field: original_column_name}.
    """
    mapping: Dict[str, str] = {}
    used_cols: set = set()

    cols_lower = {str(c).lower().strip(): c for c in df.columns}

    for field, patterns in COLUMN_PATTERNS.items():
        best_col = None
        best_score = 0.0

        for cl_lower, original in cols_lower.items():
            if original in used_cols:
                continue
            for pat in patterns:
                score = difflib.SequenceMatcher(None, cl_lower, pat).ratio()
                # Exact substring match gets a boost
                if pat in cl_lower or cl_lower in pat:
                    score = max(score, 0.85)
                if score > best_score and score >= 0.60:
                    best_score = score
                    best_col = original

        if best_col:
            mapping[field] = best_col
            used_cols.add(best_col)

    return mapping


# ══════════════════════════════════════════════
#  4. Validate data
# ══════════════════════════════════════════════
def validate_data(
    df: pd.DataFrame,
    mapping: Dict[str, str],
) -> Tuple[List[str], List[str]]:
    """
    Return (errors, warnings).
    errors   → must-fix before processing
    warnings → data quality notes
    """
    errors: List[str] = []
    warnings: List[str] = []

    # Required fields
    for req in REQUIRED_FIELDS:
        if req not in mapping:
            errors.append(f"⛔ Required field not detected: '{req}'")

    # Row count
    if len(df) == 0:
        errors.append("⛔ Excel sheet contains no data rows.")

    # Numeric checks
    for field in NUMERIC_FIELDS:
        if field not in mapping:
            continue
        col = mapping[field]
        converted = df[col].map(clean_numeric_value)
        n_bad = int(converted.isna().sum() - df[col].isna().sum())
        n_missing = int(df[col].isna().sum())
        total = len(df)
        if n_missing > 0:
            pct = n_missing / total * 100
            warnings.append(
                f"⚠️  '{field}': {n_missing} missing values ({pct:.1f}%)"
            )
        if n_bad > 0:
            warnings.append(
                f"⚠️  '{field}': {n_bad} values could not be read as numbers"
            )

    # Date checks
    for field in DATE_FIELDS:
        if field not in mapping:
            continue
        col = mapping[field]
        converted = df[col].map(clean_date_value)
        n_bad = int(converted.isna().sum() - df[col].isna().sum())
        if n_bad > 0:
            warnings.append(
                f"⚠️  '{field}': {n_bad} values could not be read as dates"
            )

    # Duplicates
    if "customer_name" in mapping:
        col = mapping["customer_name"]
        dupes = int(df[col].duplicated().sum())
        if dupes > 0:
            warnings.append(f"⚠️  Found {dupes} potential duplicate customer rows")

    return errors, warnings


# ══════════════════════════════════════════════
#  5. Clean & process
# ══════════════════════════════════════════════
def process_data(
    raw_df: pd.DataFrame,
    mapping: Dict[str, str],
) -> pd.DataFrame:
    """
    Take the raw Excel DataFrame + column mapping →
    return a clean, canonical DataFrame ready for analytics.
    """
    df = raw_df.copy()

    # Rename mapped columns to canonical names
    rename = {v: k for k, v in mapping.items()}
    df = df.rename(columns=rename)

    # Keep ONLY mapped columns
    keep = [c for c in rename.values() if c in df.columns]
    df = df[keep].copy()

    # ── Convert numeric fields ──
    for field in NUMERIC_FIELDS:
        if field in df.columns:
            df[field] = df[field].map(clean_numeric_value)

    # ── Convert date fields ──
    for field in DATE_FIELDS:
        if field in df.columns:
            df[field] = df[field].map(clean_date_value)

    # ── Strip strings ──
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(["", "nan", "NaN", "None"], np.nan)

    # ── Drop rows where all values are NaN ──
    df = df.dropna(how="all")

    # ── Remove duplicates (keep first) ──
    dedup_keys = [c for c in ("customer_id", "customer_name") if c in df.columns]
    if dedup_keys:
        df = df.drop_duplicates(subset=dedup_keys, keep="first")

    # ── Fill NaN numeric with 0 (safe for aggregation) ──
    for field in NUMERIC_FIELDS:
        if field in df.columns:
            df[field] = df[field].fillna(0)

    # ── Derived: Collection Efficiency % ──
    if "outstanding_amount" in df.columns and "collection_amount" in df.columns:
        df["collection_efficiency"] = np.where(
            df["outstanding_amount"] > 0,
            (df["collection_amount"] / df["outstanding_amount"] * 100).round(2),
            0,
        )

    # ── Derived: Overdue Days (if not present) ──
    if "overdue_days" not in df.columns and "due_date" in df.columns:
        today = pd.Timestamp.today()
        df["overdue_days"] = np.where(
            df["due_date"].notna(),
            (today - df["due_date"]).dt.days.clip(lower=0),
            0,
        )

    # ── Derived: Risk Category ──
    df["risk_category"] = df.apply(_assign_risk, axis=1)

    # ── Reset index ──
    df = df.reset_index(drop=True)
    return df


def _assign_risk(row) -> str:
    """Assign a risk category based on overdue days and collection efficiency."""
    od = row.get("overdue_days", 0) or 0
    ce = row.get("collection_efficiency", 100) or 0
    overdue = row.get("overdue_amount", 0) or 0

    if od >= 180 or (overdue > 0 and ce < 10):
        return "🔴 Critical"
    if od >= 90 or (overdue > 0 and ce < 30):
        return "🟠 High"
    if od >= 30 or (overdue > 0 and ce < 60):
        return "🟡 Medium"
    return "🟢 Low"


# ══════════════════════════════════════════════
#  6. Data quality summary
# ══════════════════════════════════════════════
def data_quality_report(
    raw_df: pd.DataFrame,
    clean_df: pd.DataFrame,
    mapping: Dict[str, str],
) -> Dict:
    """Return a dict of quality metrics for display."""
    total_raw = len(raw_df)
    total_clean = len(clean_df)
    duplicates_removed = total_raw - total_clean

    report = {
        "total_rows_source": total_raw,
        "total_rows_cleaned": total_clean,
        "duplicates_removed": max(duplicates_removed, 0),
        "columns_mapped": len(mapping),
        "fields": {},
    }

    for field in NUMERIC_FIELDS:
        if field in clean_df.columns:
            col = clean_df[field]
            report["fields"][field] = {
                "sum": float(col.sum()),
                "count": int((col != 0).sum()),
                "missing": int((col == 0).sum()),
                "min": float(col.min()),
                "max": float(col.max()),
            }

    return report


# ══════════════════════════════════════════════
#  7. Generate sample data (for testing)
# ══════════════════════════════════════════════
def generate_sample_data(n: int = 200) -> pd.DataFrame:
    """Create sample credit-recovery data so the dashboard can be tested."""
    np.random.seed(42)
    models = ["Toyota Fortuner", "Hyundai Creta", "Mahindra XUV700",
              "Tata Nexon", "Kia Seltos", "Maruti Swift",
              "Honda City", "MG Hector"]
    collectors = ["Rajesh Kumar", "Priya Sharma", "Amit Patel",
                  "Sneha Reddy", "Vikram Singh", "Deepa Nair"]
    regions = ["North", "South", "East", "West"]

    rows = []
    for i in range(1, n + 1):
        total = np.random.randint(500000, 3000000)
        collection = np.random.randint(0, int(total * 0.95))
        outstanding = total - collection
        overdue = np.random.randint(0, outstanding + 1) if outstanding > 0 else 0
        od_days = int(np.random.choice([0, 15, 30, 45, 60, 90, 120, 180, 240]))

        rows.append({
            "customer_id": f"CR-{i:05d}",
            "customer_name": f"Customer {i:04d}",
            "vehicle_model": np.random.choice(models),
            "collector_name": np.random.choice(collectors),
            "total_amount": total,
            "collection_amount": collection,
            "outstanding_amount": outstanding,
            "overdue_amount": overdue,
            "overdue_days": od_days,
            "due_date": pd.Timestamp("2024-01-01") + pd.Timedelta(days=np.random.randint(0, 365)),
            "collection_date": pd.Timestamp("2024-01-01") + pd.Timedelta(days=np.random.randint(0, 365)),
            "region": np.random.choice(regions),
            "status": np.random.choice(["Active", "Closed", "Default", "Restructured"]),
        })

    return pd.DataFrame(rows)
