import streamlit as st
import pandas as pd
import numpy as np
from config import CUSTOM_CSS
from data_processor import read_excel_file, detect_best_sheet, map_columns, process_data
from analytics import calculate_kpis, get_customer_risk
from visualizations import format_currency, create_gauge_chart, create_donut_chart, create_bar_chart, create_area_chart

st.set_page_config(page_title="Loan Recovery Dashboard", page_icon="🏦", layout="wide", initial_sidebar_state="expanded")
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# Session State
if "df" not in st.session_state: st.session_state["df"] = None

# Sidebar (Filters & Upload)
with st.sidebar:
    st.markdown("<h2 style='text-align: center;'>🏦 Menu</h2>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if uploaded_file:
        sheets = read_excel_file(uploaded_file)
        sheet = detect_best_sheet(sheets) if len(sheets) > 1 else list(sheets.keys())[0]
        raw_df = sheets[sheet]
        mapping = map_columns(raw_df)
        st.session_state["df"] = process_data(raw_df, mapping)
        st.success("Data Loaded Successfully!")

    st.markdown("---")
    st.markdown("### Filters")
    if st.session_state["df"] is not None:
        df = st.session_state["df"]
        
        # আপনার এক্সেলের হেডিং অনুযায়ী ফিল্টার
        year = st.selectbox("Year", ["All"] + [str(x) for x in df.get("due_date", pd.Series()).dt.year.dropna().unique()])
        branch = st.selectbox("Branch Name", ["All"] + sorted(df.get("region", pd.Series()).dropna().unique().tolist()))
        model = st.selectbox("Vehicle Model", ["All"] + sorted(df.get("vehicle_model", pd.Series()).dropna().unique().tolist()))
        collector = st.selectbox("Collection Person", ["All"] + sorted(df.get("collector_name", pd.Series()).dropna().unique().tolist()))

# Main Content
if st.session_state["df"] is None:
    st.markdown('<div class="main-header"><h1>Loan Recovery Services Dashboard</h1></div>', unsafe_allow_html=True)
    st.info("👈 সাইডবার থেকে আপনার এক্সেল ফাইল আপলোড করুন।")
    st.stop()

# Apply Filters
df = st.session_state["df"].copy()
if year != "All" and "due_date" in df.columns: df = df[df["due_date"].dt.year.astype(str) == year]
if branch != "All" and "region" in df.columns: df = df[df["region"] == branch]
if model != "All" and "vehicle_model" in df.columns: df = df[df["vehicle_model"] == model]
if collector != "All" and "collector_name" in df.columns: df = df[df["collector_name"] == collector]

# Header
st.markdown('<div class="main-header"><h1>Loan Recovery Services Dashboard</h1><p>Overview & Analytics</p></div>', unsafe_allow_html=True)

# KPIs Calculation
kpis = calculate_kpis(df)
total_out = kpis.get("total_outstanding", 0)
total_coll = kpis.get("total_collection", 0)
total_over = kpis.get("total_overdue", 0)
total_case = len(df)
eff = kpis.get("collection_efficiency", 0)

# Top Row: 5 KPI Cards
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Case", f"{total_case}")
c2.metric("Total Outstanding", format_currency(total_out))
c3.metric("Recovered Amount", format_currency(total_coll))
c4.metric("Over Due Amount", format_currency(total_over))
c5.metric("Recovered Amount %", f"{eff:.1f}%")

st.markdown("<br>", unsafe_allow_html=True)

# Middle Row: 3 Charts
col1, col2, col3 = st.columns(3)

with col1:
    st.plotly_chart(create_gauge_chart(eff, "Open Case %"), use_container_width=True)

with col2:
    # Model অনুযায়ী ডেটা গ্রুপিং (যেহেতু আপনার ফাইলে Loan Type নেই, আমরা Model ব্যবহার করছি)
    if "vehicle_model" in df.columns:
        donut_data = df.groupby("vehicle_model")["outstanding_amount"].sum().reset_index()
        st.plotly_chart(create_donut_chart(donut_data, "vehicle_model", "outstanding_amount", "Outstanding by Vehicle Model"), use_container_width=True)

with col3:
    # Branch অনুযায়ী ডেটা
    if "region" in df.columns:
        bar_data = df.groupby("region")["collection_amount"].sum().reset_index()
        bar_data["collection_efficiency"] = (bar_data["collection_amount"] / (df["collection_amount"].sum() + 1)) * 100
        st.plotly_chart(create_bar_chart(bar_data, "region", "collection_efficiency", "Recovered Amount % by Branch", is_horizontal=True), use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# Bottom Row: Area Chart
if "collector_name" in df.columns:
    area_data = df.groupby("collector_name")["outstanding_amount"].sum().reset_index()
    st.plotly_chart(create_area_chart(area_data, "collector_name", "outstanding_amount", "Outstanding Amount by Collection Person"), use_container_width=True)
