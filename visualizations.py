import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from config import CHART_COLORS, PIE_COLORS, COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT, COLOR_BG

def format_currency(value):
    if pd.isna(value): return "₹ 0"
    value = float(value)
    if value >= 10000000: return f"₹ {value/10000000:,.2f} Cr"
    elif value >= 100000: return f"₹ {value/100000:,.2f} L"
    return f"₹ {value:,.0f}"

def format_number(value):
    if pd.isna(value): return "0"
    return f"{float(value):,.0f}"

def create_gauge_chart(value, title):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = value,
        title = {'text': title, 'font': {'size': 16, 'color': COLOR_PRIMARY}},
        number = {'suffix': "%", 'font': {'color': COLOR_PRIMARY, 'size': 36}},
        gauge = {
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': COLOR_SECONDARY},
            'bar': {'color': COLOR_SECONDARY},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': COLOR_ACCENT,
            'steps': [
                {'range': [0, 50], 'color': COLOR_BG},
                {'range': [50, 100], 'color': "#EFE5DC"}],
        }
    ))
    fig.update_layout(height=250, margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="white")
    return fig

def create_donut_chart(df, names_col, values_col, title):
    if df.empty: return go.Figure()
    fig = px.pie(df, names=names_col, values=values_col, hole=0.6, color_discrete_sequence=PIE_COLORS)
    fig.update_layout(title=dict(text=title, font=dict(size=16, color=COLOR_PRIMARY)), height=300, margin=dict(l=20, r=20, t=50, b=20))
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

def create_bar_chart(df, x_col, y_col, title, is_horizontal=False):
    if df.empty: return go.Figure()
    if is_horizontal:
        fig = px.bar(df, x=y_col, y=x_col, orientation='h', color_discrete_sequence=[COLOR_SECONDARY], text=y_col)
        fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
    else:
        fig = px.bar(df, x=x_col, y=y_col, color_discrete_sequence=[COLOR_SECONDARY], text=y_col)
    fig.update_layout(title=dict(text=title, font=dict(size=16, color=COLOR_PRIMARY)), height=300, margin=dict(l=20, r=20, t=50, b=20), plot_bgcolor="white")
    return fig

def create_area_chart(df, x_col, y_col, title):
    if df.empty: return go.Figure()
    fig = px.area(df, x=x_col, y=y_col, markers=True, color_discrete_sequence=[COLOR_SECONDARY])
    fig.update_layout(title=dict(text=title, font=dict(size=16, color=COLOR_PRIMARY)), height=300, margin=dict(l=20, r=20, t=50, b=20), plot_bgcolor="white")
    fig.update_traces(fill='tozeroy', fillcolor='rgba(166, 124, 82, 0.3)', line=dict(color=COLOR_SECONDARY, width=3))
    return fig

# পুরনো ফাংশনগুলো (রিপোর্টের জন্য দরকারি)
def create_monthly_trend(trend_df):
    if trend_df.empty: return go.Figure()
    fig = px.line(trend_df, x="month", y="total_collection", title="Monthly Collection Trend", markers=True)
    fig.update_layout(height=300)
    return fig
