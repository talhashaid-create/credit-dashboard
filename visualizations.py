"""
Visualization engine – all charts & KPI components.
Uses Plotly for interactive, executive-quality charts.
"""

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from config import (
    CHART_COLORS, PIE_COLORS,
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT,
    COLOR_SUCCESS, COLOR_WARNING, COLOR_DANGER,
    COLOR_TEXT_DARK, COLOR_TEXT_LIGHT, COLOR_BORDER,
)

PLOTLY_LAYOUT = dict(
    plot_bgcolor="white",
    paper_bgcolor="white",
    font=dict(family="Inter, sans-serif", size=12, color=COLOR_TEXT_DARK),
    margin=dict(l=40, r=20, t=50, b=50),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
)


def format_currency(value: float) -> str:
    """Format a number as Indian Rupees with lakh/crore grouping."""
    if pd.isna(value):
        return "₹ 0"
    value = float(value)
    if value >= 10000000:  # 1 crore
        return f"₹ {value/10000000:,.2f} Cr"
    elif value >= 100000:  # 1 lakh
        return f"₹ {value/100000:,.2f} L"
    elif value >= 1000:
        return f"₹ {value:,.0f}"
    else:
        return f"₹ {value:,.2f}"


def format_number(value: float) -> str:
    """Format with Indian-style comma grouping."""
    if pd.isna(value):
        return "0"
    return f"{float(value):,.0f}"


# ══════════════════════════════════════════════
#  KPI Cards  (HTML)
# ══════════════════════════════════════════════
def render_kpi_cards(kpis: dict):
   """Render 4-column KPI card row using HTML."""
    cards = [
        {
            "label": "Total Outstanding",
            "value": format_currency(kpis.get("total_outstanding", 0)),
            "sub": f"{format_number(kpis.get('total_records', 0))} accounts",
            "cls": "",
        },
        {
            "label": "Total Collection",
            "value": format_currency(kpis.get("total_collection", 0)),
            "sub": f"Efficiency: {kpis.get('collection_efficiency', 0):.1f}%",
            "cls": "green",
        },
        {
            "label": "Total Overdue",
            "value": format_currency(kpis.get("total_overdue", 0)),
            "sub": f"{kpis.get('overdue_pct', 0):.1f}% of total",
            "cls": "red",
        },
        {
            "label": "Collection Efficiency",
            "value": f"{kpis.get('collection_efficiency', 0):.1f}%",
            "sub": f"Avg overdue: {kpis.get('avg_overdue_days', 0):.0f} days",
            "cls": "orange",
        },
    ]

    html = '<div class="kpi-row">'
    for c in cards:
        html += f'<div class="kpi-card {c["cls"]}"><div class="kpi-label">{c["label"]}</div><div class="kpi-value">{c["value"]}</div><div class="kpi-sub">{c["sub"]}</div></div>'
    html += "</div>"
    return html


def render_kpi_cards_secondary(kpis: dict):
    """Second row of KPI cards."""
    cards = [
        {
            "label": "Total Portfolio",
            "value": format_currency(kpis.get("total_amount", 0)),
            "sub": f"{format_number(kpis.get('unique_customers', 0))} customers",
            "cls": "purple",
        },
        {
            "label": "Unique Vehicles",
            "value": format_number(kpis.get("unique_vehicles", 0)),
            "sub": "Models tracked",
            "cls": "teal",
        },
        {
            "label": "Active Collectors",
            "value": format_number(kpis.get("unique_collectors", 0)),
            "sub": "In current selection",
            "cls": "green",
        },
        {
            "label": "High Risk Accounts",
            "value": format_number(
                kpis.get("critical_count", 0) + kpis.get("high_risk_count", 0)
            ),
            "sub": f"Critical: {kpis.get('critical_count', 0)} | High: {kpis.get('high_risk_count', 0)}",
            "cls": "red",
        },
    ]

    html = '<div class="kpi-row">'
    for c in cards:
        html += f'<div class="kpi-card {c["cls"]}"><div class="kpi-label">{c["label"]}</div><div class="kpi-value">{c["value"]}</div><div class="kpi-sub">{c["sub"]}</div></div>'
    html += "</div>"
    return html


def render_kpi_cards_secondary(kpis: dict):
    """Second row of KPI cards."""
    cards = [
        {
            "label": "Total Portfolio",
            "value": format_currency(kpis.get("total_amount", 0)),
            "sub": f"{format_number(kpis.get('unique_customers', 0))} customers",
            "cls": "purple",
        },
        {
            "label": "Unique Vehicles",
            "value": format_number(kpis.get("unique_vehicles", 0)),
            "sub": "Models tracked",
            "cls": "teal",
        },
        {
            "label": "Active Collectors",
            "value": format_number(kpis.get("unique_collectors", 0)),
            "sub": "In current selection",
            "cls": "green",
        },
        {
            "label": "High Risk Accounts",
            "value": format_number(
                kpis.get("critical_count", 0) + kpis.get("high_risk_count", 0)
            ),
            "sub": f"Critical: {kpis.get('critical_count', 0)} | High: {kpis.get('high_risk_count', 0)}",
            "cls": "red",
        },
    ]

    html = '<div class="kpi-row">'
    for c in cards:
        html += f"""
        <div class="kpi-card {c['cls']}">
            <div class="kpi-label">{c['label']}</div>
            <div class="kpi-value">{c['value']}</div>
            <div class="kpi-sub">{c['sub']}</div>
        </div>
        """
    html += "</div>"
    return html


# ══════════════════════════════════════════════
#  Chart: Overdue Pie / Donut
# ══════════════════════════════════════════════
def create_overdue_pie(overdue_df: pd.DataFrame) -> go.Figure:
    if overdue_df.empty:
        return go.Figure()
    fig = go.Figure(data=[go.Pie(
        labels=overdue_df["ageing_bucket"],
        values=overdue_df["total_overdue"],
        hole=0.55,
        marker=dict(colors=PIE_COLORS),
        textinfo="label+percent",
        textposition="outside",
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Overdue Amount by Ageing Bucket", font=dict(size=16, color=COLOR_PRIMARY)),
        height=420,
    )
    return fig


# ══════════════════════════════════════════════
#  Chart: Overdue Count Bar
# ══════════════════════════════════════════════
def create_overdue_bar(overdue_df: pd.DataFrame) -> go.Figure:
    if overdue_df.empty:
        return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=overdue_df["ageing_bucket"],
        y=overdue_df["count"],
        marker_color=CHART_COLORS[: len(overdue_df)],
        text=overdue_df["count"],
        textposition="outside",
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Account Count by Ageing Bucket", font=dict(size=16, color=COLOR_PRIMARY)),
        xaxis_title="Ageing Bucket",
        yaxis_title="Number of Accounts",
        height=420,
    )
    return fig


# ══════════════════════════════════════════════
#  Chart: Collector Ranking
# ══════════════════════════════════════════════
def create_collector_chart(collector_df: pd.DataFrame) -> go.Figure:
    if collector_df.empty:
        return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=collector_df["collector_name"],
        x=collector_df["total_collection"],
        name="Collection",
        orientation="h",
        marker_color=COLOR_SUCCESS,
    ))
    fig.add_trace(go.Bar(
        y=collector_df["collector_name"],
        x=collector_df["total_outstanding"],
        name="Outstanding",
        orientation="h",
        marker_color=COLOR_WARNING,
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        barmode="group",
        title=dict(text="Collector Performance: Collection vs Outstanding",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        height=max(350, len(collector_df) * 55),
    )
    return fig


def create_collector_efficiency_gauge(collector_df: pd.DataFrame) -> go.Figure:
    """Horizontal bar showing collection efficiency % per collector."""
    if collector_df.empty:
        return go.Figure()
    fig = go.Figure(go.Bar(
        x=collector_df["collection_efficiency"],
        y=collector_df["collector_name"],
        orientation="h",
        marker_color=[
            COLOR_SUCCESS if v >= 60 else COLOR_WARNING if v >= 30 else COLOR_DANGER
            for v in collector_df["collection_efficiency"]
        ],
        text=collector_df["collection_efficiency"].astype(str) + "%",
        textposition="outside",
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Collection Efficiency by Collector (%)",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        xaxis_title="Efficiency %",
        xaxis_range=[0, 105],
        height=max(350, len(collector_df) * 55),
    )
    return fig


# ══════════════════════════════════════════════
#  Chart: Vehicle Model Analysis
# ══════════════════════════════════════════════
def create_vehicle_chart(vehicle_df: pd.DataFrame) -> go.Figure:
    if vehicle_df.empty:
        return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=vehicle_df["vehicle_model"],
        y=vehicle_df["collection"],
        name="Collection",
        marker_color=COLOR_SUCCESS,
    ))
    fig.add_trace(go.Bar(
        x=vehicle_df["vehicle_model"],
        y=vehicle_df["outstanding"],
        name="Outstanding",
        marker_color=COLOR_WARNING,
    ))
    fig.add_trace(go.Bar(
        x=vehicle_df["vehicle_model"],
        y=vehicle_df["overdue"],
        name="Overdue",
        marker_color=COLOR_DANGER,
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        barmode="group",
        title=dict(text="Vehicle Model: Outstanding vs Collection vs Overdue",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        height=450,
        xaxis_tickangle=-30,
    )
    return fig


def create_vehicle_pie(vehicle_df: pd.DataFrame) -> go.Figure:
    if vehicle_df.empty:
        return go.Figure()
    fig = go.Figure(data=[go.Pie(
        labels=vehicle_df["vehicle_model"],
        values=vehicle_df["total_amount"],
        hole=0.5,
        marker=dict(colors=PIE_COLORS),
        textinfo="label+percent",
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Portfolio Distribution by Vehicle Model",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        height=420,
    )
    return fig


# ══════════════════════════════════════════════
#  Chart: Monthly Trends
# ══════════════════════════════════════════════
def create_monthly_trend(trend_df: pd.DataFrame) -> go.Figure:
    if trend_df.empty:
        return go.Figure()
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Scatter(
        x=trend_df["month"],
        y=trend_df["total_collection"],
        name="Collection",
        mode="lines+markers",
        line=dict(color=COLOR_SUCCESS, width=3),
        marker=dict(size=8),
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        x=trend_df["month"],
        y=trend_df["total_overdue"],
        name="Overdue",
        mode="lines+markers",
        line=dict(color=COLOR_DANGER, width=3, dash="dash"),
        marker=dict(size=8),
    ), secondary_y=False)

    fig.add_trace(go.Bar(
        x=trend_df["month"],
        y=trend_df["record_count"],
        name="Records",
        marker_color=COLOR_SECONDARY,
        opacity=0.3,
    ), secondary_y=True)

    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Monthly Collection & Overdue Trends",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        height=450,
    )
    fig.update_yaxes(title_text="Amount (₹)", secondary_y=False)
    fig.update_yaxes(title_text="Record Count", secondary_y=True)
    return fig


# ══════════════════════════════════════════════
#  Chart: Customer Risk Distribution
# ══════════════════════════════════════════════
def create_risk_distribution(risk_df: pd.DataFrame) -> go.Figure:
    if risk_df.empty:
        return go.Figure()
    counts = risk_df["risk_label"].value_counts().reset_index()
    counts.columns = ["risk_label", "count"]

    colors = {"🔴 Critical": COLOR_DANGER, "🟠 High": COLOR_WARNING,
              "🟡 Medium": "#F1C40F", "🟢 Low": COLOR_SUCCESS}

    fig = go.Figure(data=[go.Pie(
        labels=counts["risk_label"],
        values=counts["count"],
        hole=0.5,
        marker=dict(colors=[colors.get(r, "#ccc") for r in counts["risk_label"]]),
        textinfo="label+value",
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Customer Risk Distribution",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        height=400,
    )
    return fig


def create_risk_scatter(risk_df: pd.DataFrame) -> go.Figure:
    """Scatter: Collection Efficiency vs Overdue Days, bubble = outstanding."""
    if risk_df.empty:
        return go.Figure()
    size = risk_df.get("outstanding", pd.Series([50])).clip(lower=1)
    fig = go.Figure(data=[go.Scatter(
        x=risk_df["collection_efficiency"],
        y=risk_df["avg_overdue_days"],
        mode="markers",
        marker=dict(
            size=np.sqrt(size) / 50 + 5,
            color=risk_df["risk_score"],
            colorscale="RdYlGn_r",
            colorbar=dict(title="Risk Score"),
            opacity=0.75,
            line=dict(width=1, color="white"),
        ),
        text=risk_df["customer_name"],
        hovertemplate=(
            "<b>%{text}</b><br>"
            "Efficiency: %{x:.1f}%<br>"
            "Avg Overdue: %{y:.0f} days<br>"
            "<extra></extra>"
        ),
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Customer Risk Map: Efficiency vs Overdue Days",
                   font=dict(size=16, color=COLOR_PRIMARY)),
        xaxis_title="Collection Efficiency (%)",
        yaxis_title="Avg Overdue Days",
        height=500,
    )
    return fig


# ══════════════════════════════════════════════
#  Chart: Region Analysis
# ══════════════════════════════════════════════
def create_region_chart(df: pd.DataFrame) -> go.Figure:
    if "region" not in df.columns or df["region"].isna().all():
        return go.Figure()
    agg = (
        df.groupby("region")
        .agg(
            outstanding=("outstanding_amount", "sum"),
            collection=("collection_amount", "sum"),
            overdue=("overdue_amount", "sum"),
        )
        .reset_index()
    )
    fig = go.Figure()
    for col, color in [("outstanding", COLOR_WARNING),
                        ("collection", COLOR_SUCCESS),
                        ("overdue", COLOR_DANGER)]:
        fig.add_trace(go.Bar(
            x=agg["region"],
            y=agg[col],
            name=col.title(),
            marker_color=color,
        ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        barmode="group",
        title=dict(text="Regional Performance", font=dict(size=16, color=COLOR_PRIMARY)),
        height=420,
    )
    return fig
