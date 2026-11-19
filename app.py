"""
Credit Recovery & Business Intelligence Dashboard
Universal Excel Analyzer - Works with ANY credit-related Excel file
Manual Column Mapping - User selects which column maps to which field
"""

import io
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

from config import CUSTOM_CSS
from data_processor import (
    read_excel_file, detect_best_sheet, map_columns,
    validate_data, process_data, data_quality_report,
    generate_sample_data, clean_numeric_value, clean_date_value,
)
from analytics import (
    calculate_kpis, get_overdue_analysis, get_customer_risk,
    get_collector_performance, get_vehicle_analysis,
    get_monthly_trends, get_high_risk_customers,
    verify_calculations,
)
from visualizations import (
    create_overdue_pie, create_overdue_bar,
    create_collector_chart, create_collector_efficiency_gauge,
    create_vehicle_chart, create_vehicle_pie,
    create_monthly_trend,
    create_risk_distribution, create_risk_scatter,
    create_region_chart,
    format_currency, format_number,
)
from reports import export_excel, export_pdf


# ==============================================
#  PAGE CONFIG
# ==============================================
st.set_page_config(
    page_title="Credit Recovery Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================
#  SESSION STATE
# ==============================================
for key, default in [
    ("processed_df", None),
    ("raw_df", None),
    ("column_mapping", None),
    ("data_uploaded", False),
    ("step", "upload"),  # upload -> preview -> mapping -> done
    ("sheet_name", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ==============================================
#  HEADER
# ==============================================
st.markdown("""
<div class="main-header">
    <h1>📊 Credit Recovery &amp; Business Intelligence Dashboard</h1>
    <p>Works with ANY Excel file | Upload, Map Columns, Analyze</p>
</div>
""", unsafe_allow_html=True)


# ==============================================
#  SIDEBAR
# ==============================================
with st.sidebar:
    st.markdown("### 📁 Data Source")

    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=["xlsx", "xls", "csv"],
        help="Upload ANY credit-related Excel file. "
             "You will be able to manually map columns.",
    )

    st.divider()
    st.markdown("### 🧪 Demo Mode")

    if st.button("▶ Load Sample Data", use_container_width=True):
        sample = generate_sample_data(200)
        st.session_state["raw_df"] = sample
        st.session_state["column_mapping"] = {
            "customer_name": "customer_name",
            "outstanding_amount": "outstanding_amount",
            "collection_amount": "collection_amount",
            "overdue_amount": "overdue_amount",
            "total_amount": "total_amount",
            "collector_name": "collector_name",
            "vehicle_model": "vehicle_model",
            "overdue_days": "overdue_days",
            "region": "region",
        }
        st.session_state["processed_df"] = process_data(sample, st.session_state["column_mapping"])
        st.session_state["data_uploaded"] = True
        st.session_state["step"] = "done"
        st.rerun()

    st.divider()

    filter_collector = "All"
    filter_model = "All"
    filter_region = "All"
    filter_status = "All"
    filter_date = None

    if st.session_state["processed_df"] is not None:
        st.markdown("### 🔍 Filters")
        df = st.session_state["processed_df"]

        if "collector_name" in df.columns:
            opts = ["All"] + sorted(df["collector_name"].dropna().astype(str).unique().tolist())
            filter_collector = st.selectbox("Collector", opts, key="f_col")

        if "vehicle_model" in df.columns:
            opts = ["All"] + sorted(df["vehicle_model"].dropna().astype(str).unique().tolist())
            filter_model = st.selectbox("Vehicle Model", opts, key="f_model")

        if "region" in df.columns:
            opts = ["All"] + sorted(df["region"].dropna().astype(str).unique().tolist())
            filter_region = st.selectbox("Branch/Region", opts, key="f_region")

        if "status" in df.columns:
            opts = ["All"] + sorted(df["status"].dropna().astype(str).unique().tolist())
            filter_status = st.selectbox("Status", opts, key="f_status")

        if "due_date" in df.columns and df["due_date"].notna().any():
            valid_dates = df["due_date"].dropna()
            min_d = valid_dates.min().date()
            max_d = valid_dates.max().date()
            if min_d != max_d:
                filter_date = st.date_input(
                    "Date Range", value=(min_d, max_d),
                    min_value=min_d, max_value=max_d, key="f_date",
                )

        st.divider()
        st.caption(f"📅 {datetime.now().strftime('%d %b %Y')}")


# ==============================================
#  PROCESS UPLOADED FILE - STEP 1: Read & Preview
# ==============================================
if uploaded_file is not None and st.session_state["step"] == "upload":
    try:
        with st.spinner("Reading Excel file..."):
            sheets = read_excel_file(uploaded_file)

        if not sheets:
            st.error("No readable data found.")
            st.stop()

        # Sheet selector
        if len(sheets) > 1:
            best = detect_best_sheet(sheets)
            sheet_name = st.selectbox(
                "Choose the correct sheet:",
                list(sheets.keys()),
                index=list(sheets.keys()).index(best),
                key="sheet_pick",
            )
        else:
            sheet_name = list(sheets.keys())[0]

        raw_df = sheets[sheet_name]
        st.session_state["raw_df"] = raw_df
        st.session_state["sheet_name"] = sheet_name

        # Show preview
        st.markdown("### 📋 Preview: First 5 Rows")
        st.dataframe(raw_df.head(5), use_container_width=True)

        st.markdown(f"**Total Rows:** {len(raw_df)} | **Total Columns:** {len(raw_df.columns)}")

        # Show all column names
        st.markdown("### 📌 All Column Names Detected")
        cols_list = list(raw_df.columns)
        col_df = pd.DataFrame({
            "#": range(1, len(cols_list) + 1),
            "Column Name": cols_list,
        })
        st.dataframe(col_df, use_container_width=True, hide_index=True)

        st.divider()

        # Button to proceed to mapping
        if st.button("➡️ Proceed to Column Mapping", type="primary", use_container_width=True):
            st.session_state["step"] = "mapping"
            st.rerun()

    except Exception as exc:
        st.error(f"Error reading file: {exc}")
        st.stop()


# ==============================================
#  STEP 2: MANUAL COLUMN MAPPING
# ==============================================
if st.session_state["step"] == "mapping" and st.session_state["raw_df"] is not None:
    raw_df = st.session_state["raw_df"]

    st.markdown("### 🔧 Step 2: Map Your Excel Columns")

    st.info(
        "**Select which Excel column contains which information.**\n\n"
        "Required: Customer Name and at least one Amount column.\n"
        "Optional: Map other columns for richer analysis."
    )

    # Get all column names from Excel
    excel_columns = ["-- Select Column --"] + list(raw_df.columns)

    # System fields with descriptions and whether required
    system_fields = {
        "customer_name": ("👤 Customer Name", "Required", True),
        "outstanding_amount": ("💰 Outstanding / Balance Amount", "Required", True),
        "collection_amount": ("💵 Collection / Payment Amount", "Recommended", False),
        "overdue_amount": ("⏰ Overdue / Arrears Amount", "Recommended", False),
        "total_amount": ("📊 Total / Bill / Invoice Amount", "Recommended", False),
        "collector_name": ("🏆 Collector / Sales Person", "Recommended", False),
        "vehicle_model": ("🚗 Vehicle / Product / Model", "Recommended", False),
        "region": ("📍 Branch / Region / Area", "Optional", False),
        "due_date": ("📅 Due Date / EMI Date", "Optional", False),
        "collection_date": ("📆 Collection Date / Payment Date", "Optional", False),
        "overdue_days": ("📉 Overdue Days / DPD", "Optional", False),
        "status": ("🏷️ Status / Category", "Optional", False),
        "customer_id": ("🔢 Account ID / Reference", "Optional", False),
    }

    # Auto-detect using map_columns as default suggestions
    auto_mapping = map_columns(raw_df)

    st.markdown("#### Auto-detected mappings (change if wrong):")

    # Create mapping UI
    new_mapping = {}
    auto_found = False

    for field_key, (label, priority, required) in system_fields.items():
        st.markdown(f"**{label}** - *{priority}*")

        # Get default value from auto-detection
        default_col = auto_mapping.get(field_key, "-- Select Column --")
        if default_col in raw_df.columns:
            default_idx = excel_columns.index(default_col)
            auto_found = True
        else:
            default_idx = 0

        selected = st.selectbox(
            f"Which column? ({field_key})",
            excel_columns,
            index=default_idx,
            key=f"map_{field_key}",
        )

        if selected != "-- Select Column --":
            new_mapping[field_key] = selected

    # Validation
    st.divider()

    required_ok = "customer_name" in new_mapping and "outstanding_amount" in new_mapping

    if not required_ok:
        st.error(
            "⛔ **Required fields missing!**\n\n"
            "You MUST map:\n"
            "- Customer Name\n"
            "- Outstanding / Balance Amount"
        )
    else:
        # Show what was mapped
        st.markdown("#### ✅ Your Mapping Summary:")
        map_df = pd.DataFrame([
            {"System Field": k, "Excel Column": v}
            for k, v in new_mapping.items()
        ])
        st.dataframe(map_df, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🚀 Process Data", type="primary", use_container_width=True):
                st.session_state["column_mapping"] = new_mapping
                with st.spinner("Processing..."):
                    processed = process_data(raw_df, new_mapping)

                if len(processed) == 0:
                    st.error("No data could be processed. Check your mappings.")
                    st.stop()

                st.session_state["processed_df"] = processed
                st.session_state["data_uploaded"] = True
                st.session_state["step"] = "done"
                st.success(f"Processed {len(processed)} records!")
                st.rerun()

        with col2:
            if st.button("⬅️ Back to Preview", use_container_width=True):
                st.session_state["step"] = "upload"
                st.rerun()


# ==============================================
#  NO DATA STATE
# ==============================================
if st.session_state["processed_df"] is None and st.session_state["step"] == "upload":
    st.markdown("""
    <div style="text-align:center; padding:60px 20px;">
        <h2 style="color:#1B2A4A;">Universal Credit Recovery Dashboard</h2>
        <p style="color:#718096; font-size:16px; max-width:700px; margin:20px auto;">
            Upload <b>ANY</b> credit-related Excel file.<br>
            The system will show your columns and let you <b>manually map</b> them.
            Then you get a full executive dashboard.
        </p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### 🎯 Any Excel File\n- No specific format needed\n- Manual column mapping\n- Auto-detection with override")
    with c2:
        st.markdown("#### 📊 Executive Analytics\n- KPI dashboards\n- Overdue & risk analysis\n- Collector rankings")
    with c3:
        st.markdown("#### 📄 Export\n- Multi-sheet Excel\n- PDF reports\n- Dynamic filtering")
    st.stop()


# ==============================================
#  APPLY FILTERS
# ==============================================
df = st.session_state["processed_df"].copy()

if filter_collector != "All" and "collector_name" in df.columns:
    df = df[df["collector_name"] == filter_collector]
if filter_model != "All" and "vehicle_model" in df.columns:
    df = df[df["vehicle_model"] == filter_model]
if filter_region != "All" and "region" in df.columns:
    df = df[df["region"] == filter_region]
if filter_status != "All" and "status" in df.columns:
    df = df[df["status"] == filter_status]
if filter_date and len(filter_date) == 2 and "due_date" in df.columns:
    start, end = filter_date
    df = df[(df["due_date"] >= pd.Timestamp(start)) & (df["due_date"] <= pd.Timestamp(end))]

if df.empty:
    st.warning("No records match the current filters.")
    st.stop()


# ==============================================
#  COMPUTE ANALYTICS
# ==============================================
kpis = calculate_kpis(df)
overdue_df = get_overdue_analysis(df)
collector_df = get_collector_performance(df)
vehicle_df = get_vehicle_analysis(df)
risk_df = get_customer_risk(df)
trend_df = get_monthly_trends(df)
high_risk_df = get_high_risk_customers(df, top_n=20)


# ==============================================
#  TABS
# ==============================================
tabs = st.tabs([
    "🏠 Dashboard",
    "⏰ Overdue Analysis",
    "👥 Customer Risk",
    "🏆 Collector Performance",
    "🚗 Vehicle Model",
    "📈 Monthly Trends",
    "📋 Source Data",
    "📤 Export",
])


# ----------------------------------------------
#  TAB 1: Dashboard
# ----------------------------------------------
with tabs[0]:
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total Outstanding",
                  format_currency(kpis.get("total_outstanding", 0)),
                  f"{format_number(kpis.get('total_records', 0))} records")
    with k2:
        st.metric("Total Collection",
                  format_currency(kpis.get("total_collection", 0)),
                  f"Efficiency: {kpis.get('collection_efficiency', 0):.1f}%")
    with k3:
        st.metric("Total Overdue",
                  format_currency(kpis.get("total_overdue", 0)),
                  f"{kpis.get('overdue_pct', 0):.1f}% of total")
    with k4:
        st.metric("Collection Efficiency",
                  f"{kpis.get('collection_efficiency', 0):.1f}%",
                  f"Avg overdue: {kpis.get('avg_overdue_days', 0):.0f} days")

    st.divider()

    k5, k6, k7, k8 = st.columns(4)
    with k5:
        st.metric("Total Portfolio",
                  format_currency(kpis.get("total_amount", 0)),
                  f"{format_number(kpis.get('unique_customers', 0))} customers")
    with k6:
        st.metric("Unique Vehicles",
                  format_number(kpis.get("unique_vehicles", 0)), "Models")
    with k7:
        st.metric("Active Collectors",
                  format_number(kpis.get("unique_collectors", 0)), "In selection")
    with k8:
        st.metric("High Risk Accounts",
                  format_number(kpis.get("critical_count", 0) + kpis.get("high_risk_count", 0)),
                  f"Critical: {kpis.get('critical_count', 0)}")

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(create_overdue_pie(overdue_df), use_container_width=True, key="t1_pie")
    with c2:
        st.plotly_chart(create_risk_distribution(risk_df), use_container_width=True, key="t1_risk")

    if "region" in df.columns and not df["region"].isna().all():
        st.plotly_chart(create_region_chart(df), use_container_width=True, key="t1_region")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("#### 🏆 Top 5 Collectors")
        if not collector_df.empty:
            st.dataframe(
                collector_df.head(5)[["rank", "collector_name", "total_collection",
                                      "collection_efficiency", "accounts_handled"]],
                use_container_width=True, hide_index=True)
    with c4:
        st.markdown("#### 🚨 Top 5 Risk Customers")
        if not high_risk_df.empty:
            cols = ["customer_name", "risk_score", "risk_label"]
            if "outstanding" in high_risk_df.columns:
                cols.append("outstanding")
            if "overdue" in high_risk_df.columns:
                cols.append("overdue")
            st.dataframe(
                high_risk_df.head(5)[cols],
                use_container_width=True, hide_index=True)


# ----------------------------------------------
#  TAB 2: Overdue Analysis
# ----------------------------------------------
with tabs[1]:
    st.markdown("### ⏰ Overdue Analysis")

    if not overdue_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_overdue_pie(overdue_df), use_container_width=True, key="t2_pie")
        with c2:
            st.plotly_chart(create_overdue_bar(overdue_df), use_container_width=True, key="t2_bar")

        display = overdue_df.copy()
        for col in ("total_outstanding", "total_overdue", "total_collection"):
            if col in display.columns:
                display[col] = display[col].apply(format_currency)
        st.dataframe(display, use_container_width=True, hide_index=True)

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Overdue", format_currency(overdue_df["total_overdue"].sum()))
        m2.metric("Accounts", format_number(overdue_df["count"].sum()))
        m3.metric("Avg per Account",
                  format_currency(overdue_df["total_overdue"].sum() / max(overdue_df["count"].sum(), 1)))
    else:
        st.info("No overdue data available.")


# ----------------------------------------------
#  TAB 3: Customer Risk
# ----------------------------------------------
with tabs[2]:
    st.markdown("### 👥 Customer Risk Analysis")

    if not risk_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_risk_distribution(risk_df), use_container_width=True, key="t3_dist")
        with c2:
            st.plotly_chart(create_risk_scatter(risk_df), use_container_width=True, key="t3_scatter")

        st.markdown("#### 🔴 High-Risk Customers (Top 20)")
        display_risk = high_risk_df.copy()
        for col in ("total_amount", "collection", "outstanding", "overdue"):
            if col in display_risk.columns:
                display_risk[col] = display_risk[col].apply(format_currency)
        st.dataframe(display_risk, use_container_width=True, hide_index=True)

        rc = risk_df["risk_label"].value_counts()
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Critical", int(rc.get("Critical", 0)))
        r2.metric("High", int(rc.get("High", 0)))
        r3.metric("Medium", int(rc.get("Medium", 0)))
        r4.metric("Low", int(rc.get("Low", 0)))
    else:
        st.info("No customer data available.")


# ----------------------------------------------
#  TAB 4: Collector Performance
# ----------------------------------------------
with tabs[3]:
    st.markdown("### 🏆 Collector Performance")

    if not collector_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_collector_chart(collector_df), use_container_width=True, key="t4_chart")
        with c2:
            st.plotly_chart(create_collector_efficiency_gauge(collector_df), use_container_width=True, key="t4_gauge")

        display_col = collector_df.copy()
        for col in ("total_assigned", "total_collection", "total_outstanding", "total_overdue"):
            if col in display_col.columns:
                display_col[col] = display_col[col].apply(format_currency)
        st.dataframe(display_col, use_container_width=True, hide_index=True)

        top = collector_df.iloc[0]
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Top Collector", top["collector_name"])
        t2.metric("Collection", format_currency(top["total_collection"]))
        t3.metric("Efficiency", f"{top['collection_efficiency']:.1f}%")
        t4.metric("Accounts", format_number(top["accounts_handled"]))
    else:
        st.info("No collector data available.")


# ----------------------------------------------
#  TAB 5: Vehicle Model
# ----------------------------------------------
with tabs[4]:
    st.markdown("### 🚗 Vehicle Model Analysis")

    if not vehicle_df.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(create_vehicle_chart(vehicle_df), use_container_width=True, key="t5_chart")
        with c2:
            st.plotly_chart(create_vehicle_pie(vehicle_df), use_container_width=True, key="t5_pie")

        display_veh = vehicle_df.copy()
        for col in ("total_amount", "collection", "outstanding", "overdue"):
            if col in display_veh.columns:
                display_veh[col] = display_veh[col].apply(format_currency)
        st.dataframe(display_veh, use_container_width=True, hide_index=True)
    else:
        st.info("No vehicle data available.")


# ----------------------------------------------
#  TAB 6: Monthly Trends
# ----------------------------------------------
with tabs[5]:
    st.markdown("### 📈 Monthly Trends")

    if not trend_df.empty:
        st.plotly_chart(create_monthly_trend(trend_df), use_container_width=True, key="t6_trend")

        display_trend = trend_df.copy()
        for col in ("total_collection", "total_outstanding", "total_overdue"):
            if col in display_trend.columns:
                display_trend[col] = display_trend[col].apply(format_currency)
        st.dataframe(display_trend, use_container_width=True, hide_index=True)
    else:
        st.info("No date data available for trends.")


# ----------------------------------------------
#  TAB 7: Source Data
# ----------------------------------------------
with tabs[6]:
    st.markdown("### 📋 Source Data & Verification")

    quality = data_quality_report(
        st.session_state["raw_df"],
        st.session_state["processed_df"],
        st.session_state["column_mapping"],
    )
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Source Rows", format_number(quality["total_rows_source"]))
    q2.metric("Clean Rows", format_number(quality["total_rows_cleaned"]))
    q3.metric("Duplicates Removed", format_number(quality["duplicates_removed"]))
    q4.metric("Columns Mapped", format_number(quality["columns_mapped"]))

    st.markdown("#### ✅ Verification")
    verify_data = verify_calculations(df)
    if verify_data:
        st.dataframe(pd.DataFrame(verify_data), use_container_width=True, hide_index=True)

    st.markdown("#### 🔍 Your Column Mapping")
    if st.session_state["column_mapping"]:
        map_display = pd.DataFrame([
            {"Field": k, "Excel Column": v}
            for k, v in st.session_state["column_mapping"].items()
        ])
        st.dataframe(map_display, use_container_width=True, hide_index=True)

    st.markdown("#### 📊 Processed Data")
    st.dataframe(df, use_container_width=True, height=400)

    st.markdown("#### 📊 Raw Source Data")
    st.dataframe(st.session_state["raw_df"], use_container_width=True, height=400)


# ----------------------------------------------
#  TAB 8: Export
# ----------------------------------------------
with tabs[7]:
    st.markdown("### 📤 Reports & Export")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📊 Excel Report")
        if st.button("📥 Download Excel", use_container_width=True, type="primary"):
            with st.spinner("Generating..."):
                excel_buffer = export_excel(df, kpis, overdue_df, collector_df,
                                            vehicle_df, risk_df, trend_df)
            st.download_button(
                label="⬇️ Download .xlsx", data=excel_buffer,
                file_name=f"Credit_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    with c2:
        st.markdown("#### 📄 PDF Report")
        if st.button("📥 Download PDF", use_container_width=True, type="primary"):
            with st.spinner("Generating..."):
                pdf_buffer = export_pdf(df, kpis, overdue_df, collector_df, vehicle_df)
            st.download_button(
                label="⬇️ Download .pdf", data=pdf_buffer,
                file_name=f"Credit_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True)

    st.markdown("#### 📋 Quick Summary")
    st.info(
        f"**{format_number(len(df))}** records\n\n"
        f"- Outstanding: {format_currency(kpis.get('total_outstanding', 0))}\n"
        f"- Collection: {format_currency(kpis.get('total_collection', 0))}\n"
        f"- Efficiency: {kpis.get('collection_efficiency', 0):.2f}%"
    )


# ==============================================
#  FOOTER
# ==============================================
st.markdown("""
<div class="footer">
    📊 Universal Credit Recovery Dashboard | Works with ANY Excel file | © 2024
</div>
""", unsafe_allow_html=True)
