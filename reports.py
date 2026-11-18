"""
Report generation – Excel & PDF export.
"""

import io
import pandas as pd
import numpy as np
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
)

from visualizations import format_currency, format_number


# ══════════════════════════════════════════════
#  EXCEL EXPORT
# ══════════════════════════════════════════════
def export_excel(
    df: pd.DataFrame,
    kpis: dict,
    overdue_df: pd.DataFrame,
    collector_df: pd.DataFrame,
    vehicle_df: pd.DataFrame,
    risk_df: pd.DataFrame,
    trend_df: pd.DataFrame,
) -> io.BytesIO:
    """Generate a formatted multi-sheet Excel workbook."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── Formats ──
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#1B2A4A", "font_color": "white",
            "border": 1, "text_wrap": True, "valign": "vcenter",
            "align": "center",
        })
        currency_fmt = wb.add_format({"num_format": "₹#,##,##0", "border": 1})
        pct_fmt = wb.add_format({"num_format": "0.00%", "border": 1})
        num_fmt = wb.add_format({"num_format": "#,##0", "border": 1})

        # ── Sheet 1: Executive Summary ──
        summary_data = pd.DataFrame([
            {"Metric": "Total Records", "Value": kpis.get("total_records", 0)},
            {"Metric": "Total Outstanding", "Value": kpis.get("total_outstanding", 0)},
            {"Metric": "Total Collection", "Value": kpis.get("total_collection", 0)},
            {"Metric": "Total Overdue", "Value": kpis.get("total_overdue", 0)},
            {"Metric": "Collection Efficiency (%)", "Value": kpis.get("collection_efficiency", 0)},
            {"Metric": "Total Portfolio Amount", "Value": kpis.get("total_amount", 0)},
            {"Metric": "Unique Customers", "Value": kpis.get("unique_customers", 0)},
            {"Metric": "Unique Collectors", "Value": kpis.get("unique_collectors", 0)},
            {"Metric": "High Risk Accounts", "Value": kpis.get("critical_count", 0) + kpis.get("high_risk_count", 0)},
        ])
        summary_data.to_excel(writer, sheet_name="Executive Summary", index=False, startrow=1)
        ws = writer.sheets["Executive Summary"]
        for i, col in enumerate(summary_data.columns):
            ws.write(0, i, col, header_fmt)
            ws.set_column(i, i, 28)

        # ── Sheet 2: Source Data ──
        df.to_excel(writer, sheet_name="Source Data", index=False)
        ws2 = writer.sheets["Source Data"]
        for i, col in enumerate(df.columns):
            ws2.write(0, i, col, header_fmt)
            ws2.set_column(i, i, 18)

        # ── Sheet 3: Overdue Analysis ──
        if not overdue_df.empty:
            overdue_df.to_excel(writer, sheet_name="Overdue Analysis", index=False, startrow=1)
            ws3 = writer.sheets["Overdue Analysis"]
            for i, col in enumerate(overdue_df.columns):
                ws3.write(0, i, col, header_fmt)
                ws3.set_column(i, i, 20)

        # ── Sheet 4: Collector Performance ──
        if not collector_df.empty:
            collector_df.to_excel(writer, sheet_name="Collector Performance", index=False, startrow=1)
            ws4 = writer.sheets["Collector Performance"]
            for i, col in enumerate(collector_df.columns):
                ws4.write(0, i, col, header_fmt)
                ws4.set_column(i, i, 22)

        # ── Sheet 5: Vehicle Analysis ──
        if not vehicle_df.empty:
            vehicle_df.to_excel(writer, sheet_name="Vehicle Analysis", index=False, startrow=1)
            ws5 = writer.sheets["Vehicle Analysis"]
            for i, col in enumerate(vehicle_df.columns):
                ws5.write(0, i, col, header_fmt)
                ws5.set_column(i, i, 22)

        # ── Sheet 6: Customer Risk ──
        if not risk_df.empty:
            risk_df.to_excel(writer, sheet_name="Customer Risk", index=False, startrow=1)
            ws6 = writer.sheets["Customer Risk"]
            for i, col in enumerate(risk_df.columns):
                ws6.write(0, i, col, header_fmt)
                ws6.set_column(i, i, 20)

        # ── Sheet 7: Monthly Trends ──
        if not trend_df.empty:
            trend_df.to_excel(writer, sheet_name="Monthly Trends", index=False, startrow=1)
            ws7 = writer.sheets["Monthly Trends"]
            for i, col in enumerate(trend_df.columns):
                ws7.write(0, i, col, header_fmt)
                ws7.set_column(i, i, 20)

    buffer.seek(0)
    return buffer


# ══════════════════════════════════════════════
#  PDF EXPORT
# ══════════════════════════════════════════════
def export_pdf(
    df: pd.DataFrame,
    kpis: dict,
    overdue_df: pd.DataFrame,
    collector_df: pd.DataFrame,
    vehicle_df: pd.DataFrame,
) -> io.BytesIO:
    """Generate a professional PDF summary report."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=2 * cm, bottomMargin=2 * cm,
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle", parent=styles["Title"],
        fontSize=20, textColor=colors.HexColor("#1B2A4A"),
        spaceAfter=6,
    )
    heading_style = ParagraphStyle(
        "CustomHeading", parent=styles["Heading2"],
        fontSize=14, textColor=colors.HexColor("#2E86AB"),
        spaceBefore=16, spaceAfter=8,
    )
    body_style = ParagraphStyle(
        "CustomBody", parent=styles["Normal"],
        fontSize=10, textColor=colors.HexColor("#4A5568"),
    )

    story = []

    # Header
    story.append(Paragraph("Credit Recovery & Analytics Report", title_style))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}",
        body_style,
    ))
    story.append(Spacer(1, 12))

    # ── Executive Summary ──
    story.append(Paragraph("Executive Summary", heading_style))

    summary_rows = [
        ["Metric", "Value"],
        ["Total Records", format_number(kpis.get("total_records", 0))],
        ["Total Outstanding", format_currency(kpis.get("total_outstanding", 0))],
        ["Total Collection", format_currency(kpis.get("total_collection", 0))],
        ["Total Overdue", format_currency(kpis.get("total_overdue", 0))],
        ["Collection Efficiency", f"{kpis.get('collection_efficiency', 0):.2f}%"],
        ["Total Portfolio", format_currency(kpis.get("total_amount", 0))],
        ["Unique Customers", format_number(kpis.get("unique_customers", 0))],
        ["High Risk Accounts", format_number(
            kpis.get("critical_count", 0) + kpis.get("high_risk_count", 0))],
    ]
    table = Table(summary_rows, colWidths=[8 * cm, 8 * cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B2A4A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, 0), 11),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(table)
    story.append(Spacer(1, 16))

    # ── Overdue Analysis ──
    if not overdue_df.empty:
        story.append(Paragraph("Overdue Analysis", heading_style))
        rows = [["Ageing Bucket", "Accounts", "Outstanding", "Overdue", "Collection"]]
        for _, r in overdue_df.iterrows():
            rows.append([
                str(r["ageing_bucket"]),
                format_number(r["count"]),
                format_currency(r["total_outstanding"]),
                format_currency(r["total_overdue"]),
                format_currency(r["total_collection"]),
            ])
        t = Table(rows, colWidths=[3.5 * cm, 2.5 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E86AB")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t)
        story.append(Spacer(1, 12))

    # ── Collector Performance ──
    if not collector_df.empty:
        story.append(PageBreak())
        story.append(Paragraph("Collector Performance", heading_style))
        rows = [["Rank", "Collector", "Accounts", "Collection", "Outstanding", "Efficiency %"]]
        for _, r in collector_df.iterrows():
            rows.append([
                str(r.get("rank", "")),
                str(r["collector_name"]),
                format_number(r["accounts_handled"]),
                format_currency(r["total_collection"]),
                format_currency(r["total_outstanding"]),
                f"{r['collection_efficiency']:.1f}%",
            ])
        t = Table(rows, colWidths=[1.5 * cm, 4 * cm, 2.5 * cm, 3 * cm, 3 * cm, 2.5 * cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E86AB")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t)

    # ── Vehicle Analysis ──
    if not vehicle_df.empty:
        story.append(PageBreak())
        story.append(Paragraph("Vehicle Model Analysis", heading_style))
        rows = [["Model", "Units", "Collection", "Outstanding", "Overdue", "Efficiency %"]]
        for _, r in vehicle_df.iterrows():
            rows.append([
                str(r["vehicle_model"]),
                format_number(r["units"]),
                format_currency(r["collection"]),
                format_currency(r["outstanding"]),
                format_currency(r["overdue"]),
                f"{r['collection_efficiency']:.1f}%",
            ])
        t = Table(rows, colWidths=[3.5 * cm, 2 * cm, 3 * cm, 3 * cm, 3 * cm, 2.5 * cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E86AB")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t)

    # Footer note
    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "<i>This report is auto-generated from uploaded source data. "
        "All figures are computed directly from the data – no estimates included.</i>",
        body_style,
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer
