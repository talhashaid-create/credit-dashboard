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

CHART_COLORS = ["#2E86AB", "#E63946", "#2D6A4F", "#E09F3E", "#7B2D8B", "#FF6B6B", "#4ECDC4", "#45B7D1"]
PIE_COLORS = ["#2E86AB", "#E63946", "#2D6A4F", "#E09F3E", "#7B2D8B", "#FF6B6B", "#4ECDC4", "#45B7D1"]

COLUMN_PATTERNS = {
    "customer_id": ["sl", "account id", "regd. no", "customer id", "a/c no"],
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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
* { font-family: 'Inter', sans-serif !important; }
.stApp { background-color: #F1F3F8; }
.main-header { background: linear-gradient(135deg, #1B2A4A 0%, #2E86AB 50%, #1B2A4A 100%); padding: 28px 36px; border-radius: 0 0 18px 18px; margin-bottom: 28px; color: white; box-shadow: 0 6px 24px rgba(27,42,74,0.25); }
.main-header h1 { margin: 0; font-size: 26px; font-weight: 800; }
.main-header p { margin: 4px 0 0; font-size: 13px; opacity: 0.82; }
.kpi-row { display: flex; gap: 18px; margin: 18px 0 26px 0; flex-wrap: wrap; }
.kpi-card { flex: 1; min-width: 210px; background: #fff; border-radius: 14px; padding: 22px 24px; box-shadow: 0 2px 14px rgba(27,42,74,0.07); border-left: 5px solid #2E86AB; }
.kpi-card.green  { border-left-color: #2D6A4F; }
.kpi-card.red    { border-left-color: #E63946; }
.kpi-card.orange { border-left-color: #E09F3E; }
.kpi-card.purple { border-left-color: #7B2D8B; }
.kpi-card.teal   { border-left-color: #4ECDC4; }
.kpi-label { font-size: 11px; text-transform: uppercase; letter-spacing: 1.2px; color: #A0AEC0; margin-bottom: 6px; font-weight: 600; }
.kpi-value { font-size: 26px; font-weight: 800; color: #1B2A4A; margin: 4px 0 6px; line-height: 1.15; }
.kpi-sub { font-size: 12px; color: #718096; font-weight: 500; }
.section-title { font-size: 18px; font-weight: 700; color: #1B2A4A; margin: 32px 0 14px 0; padding-bottom: 8px; border-bottom: 2px solid #E2E8F0; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #fff; border-radius: 12px; padding: 6px; }
.stTabs [data-baseweb="tab"] { padding: 10px 22px; font-weight: 600; font-size: 13px; color: #4A5568; border-radius: 8px; }
.stTabs [aria-selected="true"] { background: linear-gradient(135deg, #2E86AB, #1B2A4A); color: #fff !important; }
.dataframe { border-radius: 10px; overflow: hidden; }
.footer { text-align: center; padding: 24px; color: #A0AEC0; font-size: 12px; margin-top: 40px; border-top: 1px solid #E2E8F0; }
</style>
"""
