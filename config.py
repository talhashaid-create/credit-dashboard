"""
Configuration module – Credit Recovery & Analytics Management System
Contains all constants, column patterns, colour palette & CSS.
"""

# ──────────────────────────────────────────────
#  COLOUR PALETTE  –  Executive Corporate Theme
# ──────────────────────────────────────────────
COLOR_PRIMARY        = "#1B2A4A"
COLOR_SECONDARY      = "#2E86AB"
COLOR_ACCENT         = "#E63946"
COLOR_SUCCESS        = "#2D6A4F"
COLOR_WARNING        = "#E09F3E"
COLOR_DANGER         = "#C1121F"
COLOR_BG             = "#F1F3F8"
COLOR_CARD           = "#FFFFFF"
COLOR_TEXT_DARK      = "#1B2A4A"
COLOR_TEXT_MEDIUM    = "#4A5568"
COLOR_TEXT_LIGHT     = "#A0AEC0"
COLOR_BORDER         = "#E2E8F0"

CHART_COLORS = [
    "#2E86AB", "#E63946", "#2D6A4F", "#E09F3E",
    "#7B2D8B", "#FF6B6B", "#4ECDC4", "#45B7D1",
    "#96CEB4", "#F39C12", "#DDA0DD", "#98D8C8",
]

PIE_COLORS = ["#2E86AB", "#E63946", "#2D6A4F", "#E09F3E",
              "#7B2D8B", "#FF6B6B", "#4ECDC4", "#45B7D1"]

# ──────────────────────────────────────────────
#  COLUMN MAPPING PATTERNS
#  (intelligent fuzzy detection from Excel headers)
# ──────────────────────────────────────────────
COLUMN_PATTERNS = {
    "customer_id": [
        "a/c no", "ac no", "account no", "account number",
        "a/c number", "customer id", "customer no", "customer number",
        "client id", "client no", "cif", "cif no", "cif number",
        "borrower id", "loan id", "loan no", "agreement no",
        "agreement number", "contract no", "contract number",
        "file no", "file number", "case no", "case number",
        "customer code", "client code", "account id", "ac id",
    ],
    "customer_name": [
        "customer name", "customer", "name", "client name",
        "client", "borrower name", "borrower", "account holder",
        "holder name", "customer full name", "full name",
        "party name", "party", "debtor name", "debtor",
        "customer details", "client details",
    ],
    "vehicle_model": [
        "vehicle model", "vehicle", "model", "vehicle name",
        "vehicle type", "vehicle details", "car model", "car",
        "make", "make model", "vehicle make", "vehicle desc",
        "vehicle description", "product", "product name",
        "asset", "asset details", "vehicle no", "vehicle number",
    ],
    "collector_name": [
        "collector", "collector name", "recovery officer",
        "ro name", "field officer", "fo name", "recovery agent",
        "agent name", "agent", "representative", "executive",
        "relationship manager", "rm name", "collection officer",
        "co name", "recovery executive", "rex name",
        "assigned to", "handled by", "collector code",
    ],
    "outstanding_amount": [
        "outstanding", "outstanding amount", "total outstanding",
        "balance", "balance amount", "due amount", "amount due",
        "total due", "remaining", "remaining amount",
        "outstanding balance", "total balance", "net balance",
        "net outstanding", "unpaid", "unpaid amount",
        "pending", "pending amount", "os amount", "os",
        "outstanding dues", "loan outstanding", "total os",
        "outstanding balance", "remaining balance",
    ],
    "collection_amount": [
        "collection", "collected amount", "recovery",
        "recovery amount", "payment", "payment amount",
        "amount collected", "amount recovered", "amount paid",
        "paid", "paid amount", "total collected",
        "total recovered", "total paid", "collection amount",
        "recovered", "received", "received amount",
        "amount received", "actual collection",
        "month collection", "monthly collection",
    ],
    "overdue_amount": [
        "overdue", "overdue amount", "arrears", "arrear amount",
        "past due", "past due amount", "overdue balance",
        "arrear", "overdue dues", "dpd amount", "delinquent",
        "delinquent amount", "overdue pending",
        "pending overdue", "total overdue", "aged amount",
        "aged balance", "overdue os", "overdue total",
    ],
    "total_amount": [
        "total", "total amount", "invoice amount", "bill amount",
        "contract amount", "agreement amount", "loan amount",
        "financed amount", "finance amount", "principal",
        "principal amount", "original amount", "original balance",
        "sanctioned amount", "disbursement amount",
        "vehicle price", "vehicle value", "total cost",
        "net amount", "gross amount", "total invoice",
    ],
    "overdue_days": [
        "overdue days", "dpd", "days overdue", "ageing",
        "ageing days", "days past due", "past due days", "age",
        "age days", "arrear days", "delinquent days",
        "days arrear", "ageing bucket", "overdue period",
    ],
    "due_date": [
        "due date", "payment due date", "maturity date", "due",
        "due dt", "payment due", "next due", "installment due",
        "installment date", "emi due date", "scheduled date",
        "due month", "due period", "due date",
    ],
    "collection_date": [
        "collection date", "payment date", "recovery date",
        "received date", "paid date", "recovered date",
        "transaction date", "received on", "paid on",
        "collection dt", "payment dt", "date of collection",
        "date of payment", "recovery dt",
    ],
    "region": [
        "region", "area", "zone", "branch", "location", "state",
        "city", "district", "territory", "division", "office",
        "branch name", "branch code", "region name",
    ],
    "status": [
        "status", "payment status", "recovery status",
        "account status", "collection status", "loan status",
        "account type", "category", "priority", "risk",
        "risk category", "segment", "portfolio",
    ],
    "invoice_number": [
        "invoice no", "invoice number", "invoice",
        "bill no", "bill number", "bill", "receipt no",
        "receipt number", "transaction no", "transaction number",
        "voucher no", "voucher number", "ref no",
        "reference no", "doc no", "document no",
    ],
}

# Fields that MUST be present
REQUIRED_FIELDS = ["customer_name", "outstanding_amount"]

# Numeric fields for conversion
NUMERIC_FIELDS = [
    "outstanding_amount", "collection_amount",
    "overdue_amount", "total_amount", "overdue_days",
]

# Date fields
DATE_FIELDS = ["due_date", "collection_date"]

# ──────────────────────────────────────────────
#  CUSTOM CSS
# ──────────────────────────────────────────────
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ── Global ── */
* { font-family: 'Inter', sans-serif !important; }
.stApp { background-color: #F1F3F8; }

/* ── Header ── */
.main-header {
    background: linear-gradient(135deg, #1B2A4A 0%, #2E86AB 50%, #1B2A4A 100%);
    padding: 28px 36px;
    border-radius: 0 0 18px 18px;
    margin-bottom: 28px;
    color: white;
    box-shadow: 0 6px 24px rgba(27,42,74,0.25);
}
.main-header h1 {
    margin: 0;
    font-size: 26px;
    font-weight: 800;
    letter-spacing: -0.5px;
}
.main-header p {
    margin: 4px 0 0;
    font-size: 13px;
    opacity: 0.82;
    font-weight: 400;
    letter-spacing: 0.3px;
}

/* ── KPI Cards ── */
.kpi-row {
    display: flex;
    gap: 18px;
    margin: 18px 0 26px 0;
    flex-wrap: wrap;
}
.kpi-card {
    flex: 1;
    min-width: 210px;
    background: #fff;
    border-radius: 14px;
    padding: 22px 24px;
    box-shadow: 0 2px 14px rgba(27,42,74,0.07);
    border-left: 5px solid #2E86AB;
    transition: transform 0.2s;
}
.kpi-card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(27,42,74,0.12); }
.kpi-card.green  { border-left-color: #2D6A4F; }
.kpi-card.red    { border-left-color: #E63946; }
.kpi-card.orange { border-left-color: #E09F3E; }
.kpi-card.purple { border-left-color: #7B2D8B; }
.kpi-card.teal   { border-left-color: #4ECDC4; }

.kpi-label {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    color: #A0AEC0;
    margin-bottom: 6px;
    font-weight: 600;
}
.kpi-value {
    font-size: 26px;
    font-weight: 800;
    color: #1B2A4A;
    margin: 4px 0 6px;
    line-height: 1.15;
}
.kpi-sub {
    font-size: 12px;
    color: #718096;
    font-weight: 500;
}

/* ── Section Titles ── */
.section-title {
    font-size: 18px;
    font-weight: 700;
    color: #1B2A4A;
    margin: 32px 0 14px 0;
    padding-bottom: 8px;
    border-bottom: 2px solid #E2E8F0;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #fff;
    border-radius: 12px;
    padding: 6px;
    box-shadow: 0 2px 8px rgba(27,42,74,0.06);
}
.stTabs [data-baseweb="tab"] {
    padding: 10px 22px;
    font-weight: 600;
    font-size: 13px;
    color: #4A5568;
    border-radius: 8px;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #2E86AB, #1B2A4A);
    color: #fff !important;
}

/* ── Tables ── */
.dataframe {
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 2px 10px rgba(27,42,74,0.06);
}
.dataframe thead th {
    background: #1B2A4A;
    color: white;
    font-weight: 600;
    padding: 12px 14px;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.dataframe tbody td {
    padding: 10px 14px;
    font-size: 13px;
}
.dataframe tbody tr:nth-child(even) { background: #F7F9FC; }
.dataframe tbody tr:hover { background: #EDF2F7; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: #fff;
    border-right: 1px solid #E2E8F0;
    box-shadow: 2px 0 12px rgba(27,42,74,0.05);
}
section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
    color: #1B2A4A;
}

/* ── Info / Warning / Error boxes ── */
.stAlert { border-radius: 10px; }

/* ── Metric ── */
[data-testid="stMetricValue"] {
    font-size: 22px !important;
    font-weight: 700 !important;
}
[data-testid="stMetricLabel"] {
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
}

/* ── Buttons ── */
.stButton > button {
    border-radius: 10px;
    font-weight: 600;
    padding: 8px 24px;
    transition: all 0.2s;
}
.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 14px rgba(27,42,74,0.15);
}

/* ── Footer ── */
.footer {
    text-align: center;
    padding: 24px;
    color: #A0AEC0;
    font-size: 12px;
    margin-top: 40px;
    border-top: 1px solid #E2E8F0;
}
</style>
"""
