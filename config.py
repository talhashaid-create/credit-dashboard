COLOR_PRIMARY        = "#5A3A2A" # Deep Brown
COLOR_SECONDARY      = "#A67C52" # Copper
COLOR_ACCENT         = "#E8D1B5" # Beige
COLOR_SUCCESS        = "#4A7C59"
COLOR_WARNING        = "#D68C45"
COLOR_DANGER         = "#C1121F"
COLOR_BG             = "#F8F5F2" # Light background
COLOR_TEXT_DARK      = "#3E2723"
COLOR_TEXT_LIGHT     = "#FFFFFF"

CHART_COLORS = ["#5A3A2A", "#A67C52", "#E8D1B5", "#D68C45", "#8D6E63"]
PIE_COLORS = ["#5A3A2A", "#A67C52", "#E8D1B5", "#D68C45", "#8D6E63"]

COLUMN_PATTERNS = {
    "customer_id": ["sl", "account id", "regd. no", "customer id"],
    "customer_name": ["customer name", "customer", "name", "client"],
    "vehicle_model": ["model", "vehicle model", "vehicle"],
    "collector_name": ["collection person", "sales person", "collector", "agent"],
    "region": ["branch name", "branch", "region", "area"],
    "outstanding_amount": ["balance receivable", "net balance", "outstanding", "balance"],
    "collection_amount": ["total collection amount", "last month collection", "collection"],
    "overdue_amount": ["over due amount", "total installment due amount", "overdue", "arrears"],
    "total_amount": ["sale price", "bill amount", "total amount", "loan amount"],
    "overdue_days": ["overdue days", "dpd", "days overdue", "ageing days"],
    "due_date": ["emi date", "due date", "payment due date"],
    "collection_date": ["collection date", "payment date", "recovery date"],
}

REQUIRED_FIELDS = ["customer_name", "outstanding_amount"]
NUMERIC_FIELDS = ["outstanding_amount", "collection_amount", "overdue_amount", "total_amount", "overdue_days"]
DATE_FIELDS = ["due_date", "collection_date"]

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
* { font-family: 'Inter', sans-serif !important; }
.stApp { background-color: #F8F5F2; }

.main-header { background: linear-gradient(135deg, #5A3A2A 0%, #3E2723 100%); padding: 20px 30px; border-radius: 10px; margin-bottom: 20px; color: white; box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
.main-header h1 { margin: 0; font-size: 24px; font-weight: 700; }
.main-header p { margin: 4px 0 0; font-size: 13px; opacity: 0.8; }

/* Filters / Sidebar */
section[data-testid="stSidebar"] { background-color: #5A3A2A; padding-top: 20px;}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] > div { background-color: #A67C52 !important; border-color: #8D6E63 !important; }
section[data-testid="stSidebar"] .stButton button { background-color: #A67C52; color: white; border: none; }

/* KPI Cards */
div[data-testid="metric-container"] {
    background-color: white;
    border: 1px solid #E8D1B5;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    border-top: 4px solid #A67C52;
}
div[data-testid="metric-container"] label { color: #5A3A2A !important; font-weight: 600; }
div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #3E2723 !important; font-size: 24px; font-weight: 700;}
</style>
"""
