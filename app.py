"""
╔══════════════════════════════════════════════════════╗
║   CREDIT RECOVERY & BUSINESS INTELLIGENCE DASHBOARD   ║
║   Executive-Level Analytics Management System         ║
╚══════════════════════════════════════════════════════╝
"""

import io
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

from config import CUSTOM_CSS, COLOR_PRIMARY
from data_processor import (
    read_excel_file, detect_best_sheet, map_columns,
    validate_data, process_data, data_quality_report,
    generate_sample_data,
)
from analytics import (
    calculate_kpis, get_overdue_analysis, get_customer_risk,
    get_collector_performance, get_vehicle_analysis,
    get_monthly_trends, get_high_risk_customers,
    verify_calculations,
)
from visualizations import (
    render_kpi_cards, render_kpi_cards_secondary,
    create_overdue_pie, create_overdue_bar,
    create_collector_chart, create_collector_efficiency_gauge,
    create_vehicle_chart, create_vehicle_pie,
    create_monthly_trend,
    create_risk_distribution, create_risk_scatter,
    create_region_chart,
    format_currency, format_number,
)
from reports import export_excel, export_pdf


# ──────────────────────────────────────────────
#  PAGE CONFIG
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Credit Recovery Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Apply custom CSS
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ──────────────────────────────────────────────
#  SESSION STATE INITIALISATION
# ──────────────────────────────────────────────
for key, default in [
    ("processed_df", None),
    ("raw_df", None),
    ("column_mapping", None),
    ("data_uploaded", False),
    ("use_sample", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ──────────────────────────────────────────────
#  HEADER
# ──────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊  Credit Recovery &amp; Business Intelligence Dashboard</h1>
    <p>Executive-Level Analytics Management System &nbsp;|&nbsp; Data-Driven Decision Making</p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  SIDEBAR — Upload & Filters
# ══════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁  Data Source")

    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=["xlsx", "xls", "csv"],
        help="Upload your credit recovery Excel file. "
             "The system will auto-detect columns and validate data.",
    )

    st.divider()

    st.markdown("### 🧪  Demo Mode")
    if st.button("▶ Load Sample Data (for testing)", use_container_width=True):
        sample = generate_sample_data(200)
        st.session_state["raw_df"] = sample
        st.session_state["column_mapping"] = {
            "customer_id": "customer_id",
            "customer_name": "customer_name",
            "vehicle_model": "vehicle_model",
            "collector_name": "collector_name",
            "outstanding_amount": "outstanding_amount",
            "collection_amount": "collection_amount",
            "overdue_amount": "overdue_amount",
            "total_amount": "total_amount",
            "overdue_days": "overdue_days",
            "due_date": "due_date",
            "collection_date": "collection_date",
            "region": "region",
            "status": "status",
        }
        st.session_state["processed_df"] = process_data(sample, st.session_state["column_mapping"])
        st.session_state["data_uploaded"] = True
        st.session_state["use_sample"] = True
        st.rerun()

    st.divider()

    # ── Filters (only when data is loaded) ──
    if st.session_state["processed_df"] is not None:
        st.markdown("### 🔍  Filters")
        df = st.session_state["processed_df"]

        filter_collector = "All"
        filter_model = "All"
        filter_region = "All"
        filter_status = "All"
        filter_date = None

        if "collector_name" in df.columns:
            opts = ["All"] + sorted(df["collector_name"].dropna().unique().tolist())
            filter_collector = st.selectbox("Collector", opts)

        if "vehicle_model" in df.columns:
            opts = ["All"] + sorted(df["vehicle_model"].dropna().unique().tolist())
            filter_model = st.selectbox("Vehicle Model", opts)

        if "region" in df.columns:
            opts = ["All"] + sorted(df["region"].dropna().unique().tolist())
            filter_region = st.selectbox("Region", opts)

        if "status" in df.columns:
            opts = ["All"] + sorted(df["status"].dropna().unique().tolist())
            filter_status = st.selectbox("Status", opts)

        if "due_date" in df.columns and df["due_date"].notna().any():
            valid_dates = df["due_date"].dropna()
            min_d, max_d = valid_dates.min().date(), valid_dates.max().date()
            filter_date = st.date_input(
                "Date Range (Due Date)",
                value=(min_d, max_d),
                min_value=min_d,
                max_value=max_d,
            )

        st.divider()
        st.markdown("### ℹ️  About")
        st.caption(
            "This dashboard processes your Excel credit recovery data "
            "with full accuracy validation. All figures are computed "
            "directly from source – no estimates."
        )
        st.caption(f"📅  {datetime.now().strftime('%d %b %Y')}")


# ══════════════════════════════════════════════
#  FILE UPLOAD PROCESSING
# ══════════════════════════════════════════════
if uploaded_file is not None and not st.session_state["data_uploaded"]:
    try:
        with st.spinner("🔄  Reading Excel file …"):
            sheets = read_excel_file(uploaded_file)

        if not sheets:
            st.error("❌ The uploaded file contains no readable data sheets.")
            st.stop()

        # ── Sheet selection ──
        if len(sheets) > 1:
            best = detect_best_sheet(sheets)
            st.info(f"📁  Found {len(sheets)} sheet(s). "
                    f"Auto-selected: **{best}**")
            sheet_name = st.selectbox(
                "Choose the sheet with credit recovery data:",
                list(sheets.keys()),
                index=list(sheets.keys()).index(best),
            )
        else:
            sheet_name = list(sheets.keys())[0]

        raw_df = sheets[sheet_name]
        st.session_state["raw_df"] = raw_df

        # ── Column mapping ──
        mapping = map_columns(raw_df)
        st.session_state["column_mapping"] = mapping

        # ── Show mapping ──
        with st.expander("🔍  Column Auto-Mapping Results", expanded=True):
            if mapping:
                map_df = pd.DataFrame([
                    {"Detected Field": k, "Excel Column": v}
                    for k, v in mapping.items()
                ])
                st.dataframe(map_df, use_container_width=True, hide_index=True)
            else:
                st.warning("⚠️  No columns could be auto-mapped. "
                           "Please check your Excel headers.")
                st.stop()

        # ── Validation ──
        errors, warnings = validate_data(raw_df, mapping)

        if errors:
            st.error("**Data Validation Errors:**")
            for e in errors:
                st.markdown(f"- {e}")
            st.stop()

        if warnings:
            with st.expander("⚠️  Data Quality Warnings", expanded=True):
                for w in warnings:
                    st.markdown(f"- {w}")

        # ── Process ──
        with st.spinner("⚙️  Processing data …"):
            processed = process_data(raw_df, mapping)

        st.session_state["processed_df"] = processed
        st.session_state["data_uploaded"] = True

        # ── Quality report ──
        quality = data_quality_report(raw_df, processed, mapping)
        with st.expander("✅  Data Quality Report", expanded=True):
            c1, c2, c3 = st.columns(3)
            c1.metric("Source Rows", format_number(quality["total_rows_source"]))
            c2.metric("Clean Rows", format_number(quality["total_rows_cleaned"]))
            c3.metric("Duplicates Removed", format_number(quality["duplicates_removed"]))

        st.success("✅  Data processed successfully! View the dashboard below.")
        st.rerun()

    except Exception as exc:
        st.error(f"❌  Error processing file: {exc}")
        st.stop()


# ══════════════════════════════════════════════
#  NO DATA — Show upload prompt
# ══════════════════════════════════════════════
if st.session_state["processed_df"] is None:
    st.markdown("""
    <div style="text-align:center; padding:60px 20px;">
        <h2 style="color:#1B2A4A;">Welcome to the Credit Recovery Dashboard</h2>
        <p style="color:#718096; font-size:16px; max-width:600px; margin:20px auto;">
            Upload your Excel file from the sidebar or click
            <b>"Load Sample Data"</b> to explore the dashboard with demo data.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Feature overview
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        #### 🎯  Data Accuracy
        - Intelligent column auto-mapping
        - Full validation & cleaning
        - No fabricated figures
        - Source data verification
        """)
    with c2:
        st.markdown("""
        #### 📊  Executive Analytics
        - KPI dashboards
        - Overdue analysis
        - Customer risk scoring
        - Collector rankings
        """)
    with c3:
        st.markdown("""
        #### 📄  Reporting
        - Multi-sheet Excel export
        - Professional PDF reports
        - Dynamic filtering
        - Consistent figures
        """)
    st.stop()


# ══════════════════════════════════════════════
#  APPLY FILTERS
# ══════════════════════════════════════════════
df = st.session_state["processed_df"].copy()

# Collector
if filter_collector != "All" and "collector_name" in df.columns:
    df = df[df["collector_name"] == filter_collector]

# Vehicle Model
if filter_model != "All" and "vehicle_model" in df.columns:
    df = df[df["vehicle_model"] == filter_model]

# Region
if filter_region != "All" and "region" in df.columns:
    df = df[df["region"] == filter_region]

# Status
if filter_status != "All" and "status" in df.columns:
    df = df[df["status"] == filter_status]

# Date Range
if filter_date and len(filter_date) == 2 and "due_date" in df.columns:
    start, end = filter_date
    df = df[
        (df["due_date"] >= pd.Timestamp(start)) &
        (df["due_date"] <= pd.Timestamp(end))
    ]

if df.empty:
    st.warning("⚠️  No records match the current filters. "
               "Please adjust filter criteria.")
    st.stop()


# ══════════════════════════════════════════════
#  COMPUTE ANALYTICS
# ══════════════════════════════════════════════
kpis = calculate_kpis(df)
overdue_df = get_overdue_analysis(df)
collector_df = get_collector_performance(df)
vehicle_df = get_vehicle_analysis(df)
risk_df = get_customer_risk(df)
trend_df = get_monthly_trends(df)
high_risk_df = get_high_risk_customers(df, top_n=20)
verify_df = verify_calculations(df)


# ══════════════════════════════════════════════
#  MAIN DASHBOARD TABS
# ══════════════════════════════════════════════
tabs = st.tabs([
    "🏠  Executive Dashboard",
    "⏰  Overdue Analysis",
    "👥  Customer Risk",
    "🏆  Collector Performance",
    "🚗  Vehicle Model Analysis",
    "📈  Monthly Trends",
    "📋  Source Data & Verification",
    "📤  Reports & Export",
])


# ──────────────────────────────────────────────
#  TAB 1: Executive Dashboard
# ──────────────────────────────────────────────
with tabs[0]:
    st.markdown(render_kpi_cards(kpis), unsafe_allow_html=True)
    st.markdown(render_kpi_cards_secondary(kpis), unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(create_overdue_pie(overdue_df), use_container_width=True)
    with c2:
        st.plotly_chart(create_risk_distribution(risk_df), use_container_width=True)

    st.markdown('<div class="section-title">Regional Overview</div>',
                unsafe_allow_html=True)
    if "region" in df.columns:
        st.plotly_chart(create_region_chart(df), use_container_width=True)

    # Top tables
    c3, c4 = st.columns(2)
    with c3:
        st.markdown("#### 🏆  Top 5 Collectors")
        if not collector_df.empty:
            st.dataframe(
                collector_df.head(5)[
                    ["rank", "collector_name", "total_collection",
                     "collection_efficiency", "accounts_handled"]
                ],
                use_container_width=True, hide_index=True,
            )
    with c4:
        st.markdown("#### 🚨  Top 5 Risk Customers")
        if not high_risk_df.empty:
            st.dataframe(
                high_risk_df.head(5)[
                    ["customer_name", "risk_score", "risk_label",
                     "outstanding", "overdue", "collection_efficiency"]
                ],
                use_container_width=True, hide_index=True,
            )


# ──────────────────────────────────────────────
#  TAB 2: Overdue Analysis
# ──────────────────────────────────────────────
with tabs[1]:
    st.markdown("### ⏰  Overdue Analysis by Ageing Bucket")

    if not overdue_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_overdue_pie(overdue_df), use_container_width=True)
        with c2:
            st.plotly_chart(create_overdue_bar(overdue_df), use_container_width=True)

        st.markdown("#### Detailed Overdue Table")
        display = overdue_df.copy()
        for col in ("total_outstanding", "total_overdue", "total_collection"):
            if col in display.columns:
                display[col] = display[col].apply(format_currency)
        st.dataframe(display, use_container_width=True, hide_index=True)

        # Summary
        st.markdown("#### Summary")
        total_overdue = overdue_df["total_overdue"].sum()
        total_records = overdue_df["count"].sum()
        c3, c4, c5 = st.columns(3)
        c3.metric("Total Overdue", format_currency(total_overdue))
        c4.metric("Accounts in Overdue", format_number(total_records))
        c5.metric("Avg Overdue per Account",
                  format_currency(total_overdue / max(total_records, 1)))
    else:
        st.info("No overdue data available.")


# ──────────────────────────────────────────────
#  TAB 3: Customer Risk Analysis
# ──────────────────────────────────────────────
with tabs[2]:
    st.markdown("### 👥  Customer-wise Credit Risk Analysis")

    if not risk_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_risk_distribution(risk_df), use_container_width=True)
        with c2:
            st.plotly_chart(create_risk_scatter(risk_df), use_container_width=True)

        st.markdown("#### 🔴  High-Risk Customers (Top 20)")
        if not high_risk_df.empty:
            display_risk = high_risk_df.copy()
            for col in ("total_amount", "collection", "outstanding", "overdue"):
                if col in display_risk.columns:
                    display_risk[col] = display_risk[col].apply(format_currency)
            st.dataframe(
                display_risk,
                use_container_width=True, hide_index=True,
            )

        # Risk summary
        st.markdown("#### Risk Summary")
        rc = risk_df["risk_label"].value_counts()
        c3, c4, c5, c6 = st.columns(4)
        c3.metric("🔴 Critical", int(rc.get("🔴 Critical", 0)))
        c4.metric("🟠 High", int(rc.get("🟠 High", 0)))
        c5.metric("🟡 Medium", int(rc.get("🟡 Medium", 0)))
        c6.metric("🟢 Low", int(rc.get("🟢 Low", 0)))
    else:
        st.info("No customer data available.")


# ──────────────────────────────────────────────
#  TAB 4: Collector Performance
# ──────────────────────────────────────────────
with tabs[3]:
    st.markdown("### 🏆  Collector-wise Collection Performance & Ranking")

    if not collector_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_collector_chart(collector_df), use_container_width=True)
        with c2:
            st.plotly_chart(create_collector_efficiency_gauge(collector_df), use_container_width=True)

        st.markdown("#### Performance Table")
        display_col = collector_df.copy()
        for col in ("total_assigned", "total_collection", "total_outstanding", "total_overdue"):
            if col in display_col.columns:
                display_col[col] = display_col[col].apply(format_currency)
        st.dataframe(display_col, use_container_width=True, hide_index=True)

        # Top performer highlight
        top = collector_df.iloc[0]
        st.markdown("#### 🌟  Top Performer")
        c3, c4, c5, c6 = st.columns(4)
        c3.metric("Collector", top["collector_name"])
        c4.metric("Collection", format_currency(top["total_collection"]))
        c5.metric("Efficiency", f"{top['collection_efficiency']:.1f}%")
        c6.metric("Accounts", format_number(top["accounts_handled"]))
    else:
        st.info("No collector data available.")


# ──────────────────────────────────────────────
#  TAB 5: Vehicle Model Analysis
# ──────────────────────────────────────────────
with tabs[4]:
    st.markdown("### 🚗  Vehicle Model-wise Outstanding vs Collection")

    if not vehicle_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_vehicle_chart(vehicle_df), use_container_width=True)
        with c2:
            st.plotly_chart(create_vehicle_pie(vehicle_df), use_container_width=True)

        st.markdown("#### Vehicle Model Table")
        display_veh = vehicle_df.copy()
        for col in ("total_amount", "collection", "outstanding", "overdue"):
            if col in display_veh.columns:
                display_veh[col] = display_veh[col].apply(format_currency)
        st.dataframe(display_veh, use_container_width=True, hide_index=True)
    else:
        st.info("No vehicle data available.")


# ──────────────────────────────────────────────
#  TAB 6: Monthly Trends
# ──────────────────────────────────────────────
with tabs[5]:
    st.markdown("### 📈  Monthly Collection Trends")

    if not trend_df.empty:
        st.plotly_chart(create_monthly_trend(trend_df), use_container_width=True)

        st.markdown("#### Monthly Breakdown")
        display_trend = trend_df.copy()
        for col in ("total_collection", "total_outstanding", "total_overdue"):
            if col in display_trend.columns:
                display_trend[col] = display_trend[col].apply(format_currency)
        st.dataframe(display_trend, use_container_width=True, hide_index=True)
    else:
        st.info("No date-based data available for trend analysis. "
                "Ensure your Excel has date columns (Due Date / Collection Date).")


# ──────────────────────────────────────────────
#  TAB 7: Source Data & Verification
# ──────────────────────────────────────────────
with tabs[6]:
    st.markdown("### 📋  Source Data & Calculation Verification")
    st.caption(
        "This tab shows the processed source data and verifies that "
        "all dashboard calculations match the source. "
        "This ensures 100% data accuracy."
    )

    # ── Data Quality ──
    st.markdown("#### Data Quality Summary")
    quality = data_quality_report(
        st.session_state["raw_df"],
        st.session_state["processed_df"],
        st.session_state["column_mapping"],
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Source Rows", format_number(quality["total_rows_source"]))
    c2.metric("Clean Rows", format_number(quality["total_rows_cleaned"]))
    c3.metric("Duplicates Removed", format_number(quality["duplicates_removed"]))
    c4.metric("Columns Mapped", format_number(quality["columns_mapped"]))

    # ── Calculation Verification ──
    st.markdown("#### ✅  Calculation Verification")
    verify_df_check = verify_calculations(df)
    verify_display = pd.DataFrame(verify_df_check)
    if not verify_display.empty:
        st.dataframe(verify_display, use_container_width=True, hide_index=True)

    # ── Column Mapping ──
    st.markdown("#### 🔍  Column Mapping")
    if st.session_state["column_mapping"]:
        map_display = pd.DataFrame([
            {"Field": k, "Excel Column": v}
            for k, v in st.session_state["column_mapping"].items()
        ])
        st.dataframe(map_display, use_container_width=True, hide_index=True)

    # ── Raw/Processed Data ──
    st.markdown("#### 📊  Processed Data (Filtered)")
    st.dataframe(df, use_container_width=True, height=400)

    st.markdown("#### 📊  Source Data (Raw Excel)")
    st.dataframe(
        st.session_state["raw_df"],
        use_container_width=True, height=400,
    )


# ──────────────────────────────────────────────
#  TAB 8: Reports & Export
# ──────────────────────────────────────────────
with tabs[7]:
    st.markdown("### 📤  Reports & Export")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📊  Excel Report")
        st.caption(
            "Multi-sheet Excel workbook with Executive Summary, "
            "Source Data, Overdue Analysis, Collector Performance, "
            "Vehicle Analysis, Customer Risk & Monthly Trends."
        )
        if st.button("📥  Download Excel Report", use_container_width=True, type="primary"):
            with st.spinner("Generating Excel report …"):
                excel_buffer = export_excel(
                    df, kpis, overdue_df, collector_df,
                    vehicle_df, risk_df, trend_df,
                )
            st.download_button(
                label="⬇️  Click to Download .xlsx",
                data=excel_buffer,
                file_name=f"Credit_Recovery_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with c2:
        st.markdown("#### 📄  PDF Report")
        st.caption(
            "Professional PDF summary with Executive Summary, "
            "Overdue Analysis, Collector Performance & Vehicle Analysis."
        )
        if st.button("📥  Download PDF Report", use_container_width=True, type="primary"):
            with st.spinner("Generating PDF report …"):
                pdf_buffer = export_pdf(
                    df, kpis, overdue_df, collector_df, vehicle_df,
                )
            st.download_button(
                label="⬇️  Click to Download .pdf",
                data=pdf_buffer,
                file_name=f"Credit_Recovery_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    # ── Quick Stats ──
    st.markdown("#### 📋  Quick Report Summary")
    st.info(
        f"**Current filter selection includes:**\n\n"
        f"- **{format_number(len(df))}** records\n"
        f"- **Total Outstanding:** {format_currency(kpis.get('total_outstanding', 0))}\n"
        f"- **Total Collection:** {format_currency(kpis.get('total_collection', 0))}\n"
        f"- **Collection Efficiency:** {kpis.get('collection_efficiency', 0):.2f}%\n\n"
        f"All figures above are computed directly from your uploaded source data."
    )


# ══════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════
st.markdown("""
<div class="footer">
    📊  Credit Recovery &amp; Business Intelligence Dashboard &nbsp;|&nbsp;
    Data Accuracy First, Executive Design Second &nbsp;|&nbsp;
    © 2024 &nbsp;|&nbsp; All figures computed from source data – no estimates
</div>
""", unsafe_allow_html=True)
