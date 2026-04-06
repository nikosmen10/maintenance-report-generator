import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import io
import sys
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
BRAND_DARK   = colors.HexColor("#1a2a3a")
BRAND_BLUE   = colors.HexColor("#185FA5")
BRAND_TEAL   = colors.HexColor("#1D9E75")
BRAND_AMBER  = colors.HexColor("#BA7517")
BRAND_RED    = colors.HexColor("#A32D2D")
LIGHT_GRAY   = colors.HexColor("#F1EFE8")
MID_GRAY     = colors.HexColor("#D3D1C7")
WHITE        = colors.white

PRIORITY_COLORS = {
    "Critical": "#A32D2D",
    "High":     "#BA7517",
    "Medium":   "#185FA5",
    "Low":      "#1D9E75",
}

# ── Data loading & analysis ───────────────────────────────────────────────────
def load_and_analyze(filepath):
    df = pd.read_excel(filepath)
    df["Date"] = pd.to_datetime(df["Date"])
    df["Month"] = df["Date"].dt.to_period("M")

    stats = {}
    stats["total_events"]    = len(df)
    stats["total_downtime"]  = df["Downtime (hrs)"].sum()
    stats["total_cost"]      = df["Repair Cost (EUR)"].sum()
    stats["avg_downtime"]    = df["Downtime (hrs)"].mean()
    stats["avg_cost"]        = df["Repair Cost (EUR)"].mean()
    stats["pending"]         = (df["Status"] == "Pending").sum()
    stats["critical"]        = (df["Priority"] == "Critical").sum()
    stats["date_range"]      = (df["Date"].min().strftime("%d %b %Y"),
                                df["Date"].max().strftime("%d %b %Y"))

    # Most affected equipment
    eq_summary = (
        df.groupby("Equipment")
        .agg(
            Events=("Equipment", "count"),
            Total_Downtime=("Downtime (hrs)", "sum"),
            Total_Cost=("Repair Cost (EUR)", "sum"),
        )
        .sort_values("Total_Cost", ascending=False)
        .reset_index()
    )

    # Recurring faults (fault type appearing 3+ times on same equipment)
    fault_counts = (
        df.groupby(["Equipment", "Fault Description"])
        .size()
        .reset_index(name="Count")
        .query("Count >= 3")
        .sort_values("Count", ascending=False)
    )

    # MTBF per equipment (days between failures)
    mtbf = {}
    for eq, grp in df.groupby("Equipment"):
        if len(grp) > 1:
            dates = grp["Date"].sort_values()
            gaps = dates.diff().dropna().dt.days
            mtbf[eq] = round(gaps.mean(), 1)

    # Monthly trend
    monthly = (
        df.groupby("Month")
        .agg(Events=("Equipment", "count"), Cost=("Repair Cost (EUR)", "sum"))
        .reset_index()
    )
    monthly["Month_str"] = monthly["Month"].astype(str)

    # Priority breakdown
    priority_counts = df["Priority"].value_counts()

    # Technician workload
    tech_load = (
        df.groupby("Technician")
        .agg(Jobs=("Equipment", "count"), Hours=("Downtime (hrs)", "sum"))
        .sort_values("Jobs", ascending=False)
        .reset_index()
    )

    return df, stats, eq_summary, fault_counts, mtbf, monthly, priority_counts, tech_load


# ── Chart generators ──────────────────────────────────────────────────────────
def chart_to_image(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    buf.seek(0)
    plt.close(fig)
    return buf

def make_monthly_chart(monthly):
    fig, ax1 = plt.subplots(figsize=(10, 3.5))
    months = monthly["Month_str"]
    x = range(len(months))

    ax1.bar(x, monthly["Events"], color="#185FA5", alpha=0.85, label="Events")
    ax1.set_ylabel("Number of Events", fontsize=9, color="#185FA5")
    ax1.tick_params(axis="y", labelcolor="#185FA5")
    ax1.set_xticks(x)
    ax1.set_xticklabels(months, rotation=45, ha="right", fontsize=8)

    ax2 = ax1.twinx()
    ax2.plot(x, monthly["Cost"], color="#BA7517", marker="o",
             linewidth=2, markersize=5, label="Cost (EUR)")
    ax2.set_ylabel("Repair Cost (EUR)", fontsize=9, color="#BA7517")
    ax2.tick_params(axis="y", labelcolor="#BA7517")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=8)

    ax1.set_title("Monthly Maintenance Events & Cost", fontsize=11, pad=10)
    ax1.spines["top"].set_visible(False)
    fig.tight_layout()
    return chart_to_image(fig)

def make_equipment_chart(eq_summary):
    top = eq_summary.head(8)
    fig, ax = plt.subplots(figsize=(10, 3.5))
    bars = ax.barh(top["Equipment"][::-1], top["Total_Cost"][::-1],
                   color="#1D9E75", alpha=0.85)
    ax.set_xlabel("Total Repair Cost (EUR)", fontsize=9)
    ax.set_title("Top Equipment by Repair Cost", fontsize=11, pad=10)
    for bar, val in zip(bars, top["Total_Cost"][::-1]):
        ax.text(bar.get_width() + 20, bar.get_y() + bar.get_height()/2,
                f"€{val:,.0f}", va="center", fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return chart_to_image(fig)

def make_priority_chart(priority_counts):
    fig, ax = plt.subplots(figsize=(4, 3.5))
    clrs = [PRIORITY_COLORS.get(p, "#888780") for p in priority_counts.index]
    wedges, texts, autotexts = ax.pie(
        priority_counts.values, labels=priority_counts.index,
        colors=clrs, autopct="%1.0f%%", startangle=90,
        textprops={"fontsize": 9}
    )
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color("white")
    ax.set_title("Priority Breakdown", fontsize=11, pad=10)
    fig.tight_layout()
    return chart_to_image(fig)


# ── PDF styles ────────────────────────────────────────────────────────────────
def get_styles():
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle("title", fontSize=22, textColor=WHITE,
                                 fontName="Helvetica-Bold", alignment=TA_CENTER, leading=28),
        "subtitle": ParagraphStyle("subtitle", fontSize=10, textColor=colors.HexColor("#9FE1CB"),
                                    fontName="Helvetica", alignment=TA_CENTER),
        "h1": ParagraphStyle("h1", fontSize=14, textColor=BRAND_DARK,
                              fontName="Helvetica-Bold", spaceBefore=16, spaceAfter=6),
        "h2": ParagraphStyle("h2", fontSize=11, textColor=BRAND_BLUE,
                              fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4),
        "body": ParagraphStyle("body", fontSize=9, textColor=BRAND_DARK,
                                fontName="Helvetica", leading=14),
        "small": ParagraphStyle("small", fontSize=8, textColor=colors.HexColor("#5F5E5A"),
                                 fontName="Helvetica"),
        "metric_val": ParagraphStyle("metric_val", fontSize=18, textColor=BRAND_BLUE,
                                      fontName="Helvetica-Bold", alignment=TA_CENTER),
        "metric_lbl": ParagraphStyle("metric_lbl", fontSize=8, textColor=colors.HexColor("#5F5E5A"),
                                      fontName="Helvetica", alignment=TA_CENTER),
        "warning": ParagraphStyle("warning", fontSize=9, textColor=colors.HexColor("#712B13"),
                                   fontName="Helvetica-Bold"),
    }
    return styles


# ── PDF builder ───────────────────────────────────────────────────────────────
def build_pdf(output_path, filepath):
    df, stats, eq_summary, fault_counts, mtbf, monthly, priority_counts, tech_load = \
        load_and_analyze(filepath)

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )
    styles = get_styles()
    story = []

    # ── Cover header ──────────────────────────────────────────────────────────
    header_data = [[
        Paragraph("MAINTENANCE LOG ANALYSIS REPORT", styles["title"]),
    ]]
    header_table = Table(header_data, colWidths=[17*cm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BRAND_DARK),
        ("TOPPADDING",    (0,0), (-1,-1), 18),
        ("BOTTOMPADDING", (0,0), (-1,-1), 18),
        ("ROUNDEDCORNERS", [8]),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.2*cm))

    sub_data = [[
        Paragraph(f"Period: {stats['date_range'][0]} — {stats['date_range'][1]}", styles["small"]),
        Paragraph(f"Generated: {datetime.now().strftime('%d %b %Y')}", styles["small"]),
    ]]
    sub_table = Table(sub_data, colWidths=[8.5*cm, 8.5*cm])
    sub_table.setStyle(TableStyle([("ALIGN", (1,0), (1,0), "RIGHT")]))
    story.append(sub_table)
    story.append(Spacer(1, 0.5*cm))

    # ── KPI cards ─────────────────────────────────────────────────────────────
    story.append(Paragraph("Executive Summary", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    kpis = [
        (f"{stats['total_events']}", "Total Events"),
        (f"{stats['total_downtime']:.0f} hrs", "Total Downtime"),
        (f"€{stats['total_cost']:,.0f}", "Total Repair Cost"),
        (f"{stats['avg_downtime']:.1f} hrs", "Avg Downtime/Event"),
        (f"{stats['pending']}", "Pending Jobs"),
        (f"{stats['critical']}", "Critical Events"),
    ]

    kpi_cells = []
    for val, lbl in kpis:
        cell = [Paragraph(val, styles["metric_val"]), Paragraph(lbl, styles["metric_lbl"])]
        kpi_cells.append(cell)

    kpi_table = Table(
        [kpi_cells[:3], kpi_cells[3:]],
        colWidths=[5.5*cm, 5.5*cm, 5.5*cm]
    )
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), LIGHT_GRAY),
        ("ROUNDEDCORNERS", [6]),
        ("TOPPADDING",    (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("LINEAFTER",     (0,0), (1,-1), 0.5, MID_GRAY),
        ("LINEBELOW",     (0,0), (-1,0), 0.5, MID_GRAY),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 0.6*cm))

    # ── Monthly trend chart ───────────────────────────────────────────────────
    story.append(Paragraph("Maintenance Trends", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    monthly_img = make_monthly_chart(monthly)
    story.append(Image(monthly_img, width=17*cm, height=6*cm))
    story.append(Spacer(1, 0.5*cm))

    # ── Equipment cost + priority charts (side by side) ───────────────────────
    story.append(Paragraph("Equipment & Priority Analysis", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    eq_img  = make_equipment_chart(eq_summary)
    pri_img = make_priority_chart(priority_counts)

    chart_row = [[Image(eq_img, width=11*cm, height=6*cm),
                  Image(pri_img, width=6*cm, height=6*cm)]]
    chart_table = Table(chart_row, colWidths=[11*cm, 6*cm])
    story.append(chart_table)
    story.append(Spacer(1, 0.5*cm))

    story.append(PageBreak())

    # ── Equipment summary table ───────────────────────────────────────────────
    story.append(Paragraph("Equipment Summary", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    eq_header = ["Equipment", "Events", "Downtime (hrs)", "Total Cost (EUR)", "MTBF (days)"]
    eq_rows = [eq_header]
    for _, row in eq_summary.iterrows():
        mtbf_val = mtbf.get(row["Equipment"], "—")
        mtbf_str = str(mtbf_val) if mtbf_val != "—" else "—"
        eq_rows.append([
            row["Equipment"],
            str(int(row["Events"])),
            f"{row['Total_Downtime']:.1f}",
            f"€{row['Total_Cost']:,.0f}",
            mtbf_str,
        ])

    eq_table = Table(eq_rows, colWidths=[5.5*cm, 2.5*cm, 3.5*cm, 3.5*cm, 2.5*cm])
    eq_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), BRAND_DARK),
        ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
        ("ALIGN",         (1,0), (-1,-1), "CENTER"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT_GRAY]),
        ("GRID",          (0,0), (-1,-1), 0.3, MID_GRAY),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story.append(eq_table)
    story.append(Spacer(1, 0.6*cm))

    # ── Recurring faults ──────────────────────────────────────────────────────
    story.append(Paragraph("Recurring Faults (3+ occurrences)", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    if fault_counts.empty:
        story.append(Paragraph("No recurring faults detected.", styles["body"]))
    else:
        story.append(Paragraph(
            "The following fault types have occurred 3 or more times on the same equipment. "
            "These are candidates for root cause analysis or scheduled preventive action.",
            styles["body"]
        ))
        story.append(Spacer(1, 0.3*cm))

        fault_header = ["Equipment", "Fault Description", "Occurrences", "Action Recommended"]
        fault_rows = [fault_header]
        for _, row in fault_counts.iterrows():
            action = "Schedule preventive replacement" if row["Count"] >= 5 else "Investigate root cause"
            fault_rows.append([
                row["Equipment"],
                row["Fault Description"],
                str(int(row["Count"])),
                action,
            ])

        fault_table = Table(fault_rows, colWidths=[4.5*cm, 4.5*cm, 2.5*cm, 5.5*cm])
        fault_table.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), BRAND_RED),
            ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
            ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), 8),
            ("ALIGN",         (2,0), (2,-1), "CENTER"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, colors.HexColor("#FCEBEB")]),
            ("GRID",          (0,0), (-1,-1), 0.3, MID_GRAY),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(fault_table)
    story.append(Spacer(1, 0.6*cm))

    # ── Technician workload ───────────────────────────────────────────────────
    story.append(Paragraph("Technician Workload", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    tech_header = ["Technician", "Jobs Completed", "Total Hours on Site"]
    tech_rows = [tech_header]
    for _, row in tech_load.iterrows():
        tech_rows.append([row["Technician"], str(int(row["Jobs"])), f"{row['Hours']:.1f} hrs"])

    tech_table = Table(tech_rows, colWidths=[6*cm, 5.5*cm, 5.5*cm])
    tech_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), BRAND_TEAL),
        ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
        ("ALIGN",         (1,0), (-1,-1), "CENTER"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT_GRAY]),
        ("GRID",          (0,0), (-1,-1), 0.3, MID_GRAY),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story.append(tech_table)
    story.append(Spacer(1, 0.6*cm))

    # ── Recommendations ───────────────────────────────────────────────────────
    story.append(Paragraph("Recommendations", styles["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY, spaceAfter=8))

    top_eq = eq_summary.iloc[0]["Equipment"]
    top_cost = eq_summary.iloc[0]["Total_Cost"]
    pending_count = stats["pending"]
    critical_count = stats["critical"]

    recommendations = [
        f"<b>High-cost equipment focus:</b> {top_eq} has generated €{top_cost:,.0f} in repair costs "
        f"this period. A detailed reliability review or planned overhaul should be considered.",
        f"<b>Pending jobs:</b> {pending_count} maintenance jobs remain open. Review and prioritise "
        f"to avoid unplanned downtime escalation.",
    ]
    if critical_count > 0:
        recommendations.append(
            f"<b>Critical events:</b> {critical_count} critical-priority events were recorded. "
            f"Ensure post-incident reports have been completed for each."
        )
    if not fault_counts.empty:
        top_fault_eq  = fault_counts.iloc[0]["Equipment"]
        top_fault_typ = fault_counts.iloc[0]["Fault Description"]
        recommendations.append(
            f"<b>Root cause analysis:</b> '{top_fault_typ}' on {top_fault_eq} is a recurring fault. "
            f"Schedule a root cause investigation to prevent further repeat failures."
        )

    rec_rows = [[Paragraph(f"{i+1}. {rec}", styles["body"])] for i, rec in enumerate(recommendations)]
    rec_table = Table(rec_rows, colWidths=[17*cm])
    rec_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), colors.HexColor("#E6F1FB")),
        ("LEFTPADDING",   (0,0), (-1,-1), 12),
        ("RIGHTPADDING",  (0,0), (-1,-1), 12),
        ("TOPPADDING",    (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ("LINEBELOW",     (0,0), (-1,-2), 0.3, MID_GRAY),
    ]))
    story.append(rec_table)
    story.append(Spacer(1, 1*cm))

    # ── Footer ────────────────────────────────────────────────────────────────
    footer_data = [[Paragraph(
        f"Generated by Maintenance Log Analyzer · {datetime.now().strftime('%d %b %Y %H:%M')} · Confidential",
        styles["small"]
    )]]
    footer_table = Table(footer_data, colWidths=[17*cm])
    footer_table.setStyle(TableStyle([
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING",    (0,0), (-1,-1), 8),
        ("LINEABOVE",     (0,0), (-1,-1), 0.5, MID_GRAY),
    ]))
    story.append(footer_table)

    doc.build(story)
    print(f"Report saved to: {output_path}")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    input_file  = sys.argv[1] if len(sys.argv) > 1 else "maintenance_log.xlsx"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "maintenance_report.pdf"
    build_pdf(output_file, input_file)
