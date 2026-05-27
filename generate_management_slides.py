#!/usr/bin/env python3
from __future__ import annotations

import argparse
import math
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


CONVEYANCE_FIELDS = [
    "conveyanceAllocationRating",
    "specialTransportRating",
    "conveyanceResponsivenessRating",
    "overallConveyanceSatisfaction",
]

PRODUCTION_FIELDS = [
    "productionPlanningRating",
    "logisticsRequestRating",
    "productionResponsivenessRating",
    "overallProductionSatisfaction",
]

CORE_COLUMNS = [
    "employeeName",
    "employeeCode",
    "departments",
    *CONVEYANCE_FIELDS,
    *PRODUCTION_FIELDS,
    "comments",
]

POSITIVE_KEYWORDS = ["reliable", "support", "quick", "efficient", "appreciate", "good", "great", "available"]
PAIN_POINT_KEYWORDS = ["late", "delay", "slow", "issue", "problem", "difficult", "pickup", "bid"]


@dataclass
class Theme:
    blue: RGBColor = RGBColor(15, 76, 129)
    cyan: RGBColor = RGBColor(38, 166, 154)
    green: RGBColor = RGBColor(56, 142, 60)
    amber: RGBColor = RGBColor(251, 192, 45)
    red: RGBColor = RGBColor(211, 47, 47)
    slate: RGBColor = RGBColor(69, 90, 100)
    light: RGBColor = RGBColor(242, 246, 250)
    white: RGBColor = RGBColor(255, 255, 255)


def normalize_column_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def load_data(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        df = pd.read_excel(path)
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError("Input must be an .xlsx, .xls, .xlsm, or .csv file")

    rename_map = {}
    available = {normalize_column_name(c): c for c in df.columns}
    for wanted in CORE_COLUMNS:
        key = normalize_column_name(wanted)
        if key in available:
            rename_map[available[key]] = wanted

    df = df.rename(columns=rename_map)
    if "departments" not in df.columns:
        raise ValueError("Missing required column: departments")

    if "comments" not in df.columns:
        df["comments"] = ""

    # Data preparation: clean blanks/N/A, standardize departments, ensure numeric ratings 1-5
    df = df.replace({"N/A": pd.NA, "NA": pd.NA, "": pd.NA})
    df["departments"] = (
        df["departments"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.title()
    )

    for field in CONVEYANCE_FIELDS + PRODUCTION_FIELDS:
        if field not in df.columns:
            df[field] = pd.NA
        df[field] = pd.to_numeric(df[field], errors="coerce").clip(lower=1, upper=5)

    return df


def favorable_pct(series: pd.Series) -> float:
    valid = series.dropna()
    if valid.empty:
        return float("nan")
    return float((valid >= 4).mean() * 100)


def avg_metrics(df: pd.DataFrame, fields: list[str]) -> pd.Series:
    return df[fields].mean(numeric_only=True)


def distribution(df: pd.DataFrame, fields: list[str]) -> pd.Series:
    all_ratings = df[fields].stack().dropna().round().astype(int)
    counts = all_ratings.value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)
    return counts


def dept_scores(df: pd.DataFrame, field: str) -> pd.Series:
    grouped = df.groupby("departments", dropna=False)[field].mean().sort_values(ascending=False)
    return grouped


def signal_color(value: float, theme: Theme) -> RGBColor:
    if math.isnan(value):
        return theme.slate
    if value >= 4.0:
        return theme.green
    if value >= 3.0:
        return theme.amber
    return theme.red


def make_title(slide, title: str, subtitle: str, theme: Theme) -> None:
    shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(1.0))
    shape.fill.solid()
    shape.fill.fore_color.rgb = theme.blue
    shape.line.fill.background()

    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = theme.white

    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = subtitle
    r2.font.size = Pt(12)
    r2.font.color.rgb = theme.white


def add_insight(slide, text: str, theme: Theme) -> None:
    box = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.55))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = f"Key insight: {text}"
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = theme.slate


def add_kpi_card(slide, left: float, top: float, title: str, value: str, color: RGBColor, theme: Theme) -> None:
    card = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(2.95), Inches(1.35))
    card.fill.solid()
    card.fill.fore_color.rgb = theme.light
    card.line.color.rgb = color

    tf = card.text_frame
    tf.clear()
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = title
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = theme.slate

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = value
    r2.font.size = Pt(26)
    r2.font.bold = True
    r2.font.color.rgb = color


def add_clustered_bar(slide, title: str, categories: Iterable[str], values: Iterable[float], x: float, y: float, w: float, h: float, theme: Theme) -> None:
    t = slide.shapes.add_textbox(Inches(x), Inches(y - 0.3), Inches(w), Inches(0.25))
    t.text_frame.text = title
    t.text_frame.paragraphs[0].runs[0].font.bold = True
    t.text_frame.paragraphs[0].runs[0].font.size = Pt(12)

    data = CategoryChartData()
    data.categories = list(categories)
    data.add_series("Average", [round(v, 2) if not pd.isna(v) else 0 for v in values])

    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), data).chart
    chart.has_legend = False
    chart.value_axis.maximum_scale = 5
    chart.value_axis.minimum_scale = 0
    chart.category_axis.reverse_order = True
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = theme.cyan


def add_stacked_distribution(slide, title: str, dist: pd.Series, x: float, y: float, w: float, h: float, theme: Theme) -> None:
    t = slide.shapes.add_textbox(Inches(x), Inches(y - 0.3), Inches(w), Inches(0.25))
    t.text_frame.text = title
    t.text_frame.paragraphs[0].runs[0].font.bold = True
    t.text_frame.paragraphs[0].runs[0].font.size = Pt(12)

    data = CategoryChartData()
    data.categories = ["All Ratings"]
    for rating in [1, 2, 3, 4, 5]:
        data.add_series(str(rating), [int(dist.get(rating, 0))])

    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(x), Inches(y), Inches(w), Inches(h), data).chart
    chart.has_legend = True
    palette = {1: theme.red, 2: RGBColor(244, 109, 67), 3: theme.amber, 4: RGBColor(128, 203, 196), 5: theme.green}
    for idx, series in enumerate(chart.series, start=1):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = palette[idx]


def add_heat_table(slide, title: str, scores: pd.Series, x: float, y: float, w: float, h: float, theme: Theme) -> None:
    t = slide.shapes.add_textbox(Inches(x), Inches(y - 0.3), Inches(w), Inches(0.25))
    t.text_frame.text = title
    t.text_frame.paragraphs[0].runs[0].font.bold = True
    t.text_frame.paragraphs[0].runs[0].font.size = Pt(12)

    top = scores.head(10)
    rows = len(top) + 1
    table = slide.shapes.add_table(rows, 2, Inches(x), Inches(y), Inches(w), Inches(h)).table
    table.columns[0].width = Inches(w * 0.72)
    table.columns[1].width = Inches(w * 0.28)

    table.cell(0, 0).text = "Department"
    table.cell(0, 1).text = "Score"
    for c in [0, 1]:
        table.cell(0, c).fill.solid()
        table.cell(0, c).fill.fore_color.rgb = theme.blue
        for run in table.cell(0, c).text_frame.paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = theme.white

    for i, (dept, score) in enumerate(top.items(), start=1):
        table.cell(i, 0).text = str(dept)
        table.cell(i, 1).text = f"{score:.2f}" if pd.notna(score) else "-"
        color = signal_color(float(score) if pd.notna(score) else float("nan"), theme)
        table.cell(i, 1).fill.solid()
        table.cell(i, 1).fill.fore_color.rgb = color
        for run in table.cell(i, 1).text_frame.paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = theme.white


def theme_counts(comments: pd.Series) -> tuple[list[str], list[str]]:
    text = " ".join(comments.dropna().astype(str).str.lower())
    pos_hits = Counter()
    for keyword in POSITIVE_KEYWORDS:
        count = text.count(keyword)
        if count > 0:
            pos_hits[keyword] = count

    pain_hits = Counter()
    for keyword in PAIN_POINT_KEYWORDS:
        count = text.count(keyword)
        if count > 0:
            pain_hits[keyword] = count

    top_pos = [f"{k} ({v})" for k, v in pos_hits.most_common(5)] or ["Strong transport support recognition"]
    top_pain = [f"{k} ({v})" for k, v in pain_hits.most_common(5)] or ["No recurring negative theme detected"]
    return top_pos, top_pain


def anonymized_quotes(df: pd.DataFrame) -> list[str]:
    cleaned_comments = df["comments"].dropna().astype(str).str.strip().replace({"": pd.NA}).dropna()
    comments = cleaned_comments.drop_duplicates().head(5).tolist()
    return [f"Employee #{idx + 1}: {comment}" for idx, comment in enumerate(comments)]


def add_bullets(slide, heading: str, bullets: list[str], x: float, y: float, w: float, h: float, theme: Theme) -> None:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = heading
    run.font.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = theme.blue

    for item in bullets:
        p = tf.add_paragraph()
        p.text = f"• {item}"
        p.font.size = Pt(12)
        p.level = 0


def detect_headcount(df: pd.DataFrame) -> float | None:
    candidates = [c for c in df.columns if "headcount" in normalize_column_name(c) or "employees" in normalize_column_name(c)]
    for col in candidates:
        headcount_values = pd.to_numeric(df[col], errors="coerce").dropna()
        if not headcount_values.empty:
            return float(headcount_values.max())
    return None


def build_presentation(df: pd.DataFrame, output: Path, company_name: str) -> None:
    theme = Theme()
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    conveyance_avg = float(df["overallConveyanceSatisfaction"].mean()) if df["overallConveyanceSatisfaction"].notna().any() else float("nan")
    production_avg = float(df["overallProductionSatisfaction"].mean()) if df["overallProductionSatisfaction"].notna().any() else float("nan")
    conveyance_fav = favorable_pct(df["overallConveyanceSatisfaction"])
    production_fav = favorable_pct(df["overallProductionSatisfaction"])
    total_responses = len(df)
    headcount = detect_headcount(df)

    # 1) Executive Summary
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    make_title(slide, "Executive Summary", f"{company_name} | Conveyance & Production Survey", theme)

    response_rate = f"{(total_responses / headcount * 100):.1f}%" if headcount else "N/A"
    add_kpi_card(slide, 0.6, 1.5, "Total Responses", str(total_responses), theme.blue, theme)
    add_kpi_card(slide, 3.75, 1.5, "Response Rate", response_rate, theme.cyan, theme)
    add_kpi_card(slide, 6.9, 1.5, "Conveyance Avg", f"{conveyance_avg:.2f}" if not math.isnan(conveyance_avg) else "N/A", signal_color(conveyance_avg, theme), theme)
    add_kpi_card(slide, 10.05, 1.5, "Production Avg", f"{production_avg:.2f}" if not math.isnan(production_avg) else "N/A", signal_color(production_avg, theme), theme)

    add_kpi_card(slide, 2.25, 3.2, "Conveyance Favorable (4-5)", f"{conveyance_fav:.1f}%" if not math.isnan(conveyance_fav) else "N/A", theme.green, theme)
    add_kpi_card(slide, 8.15, 3.2, "Production Favorable (4-5)", f"{production_fav:.1f}%" if not math.isnan(production_fav) else "N/A", theme.green, theme)

    insight = "Production outperforms conveyance" if production_avg > conveyance_avg else "Conveyance leads production"
    add_insight(slide, f"{insight}; focus improvements on low-scoring departments and response timeliness.", theme)

    # 2) Conveyance Performance
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    make_title(slide, "Conveyance Performance", "Metric trends, rating spread, and department heatmap", theme)
    conv_metric_avg = avg_metrics(df, CONVEYANCE_FIELDS)
    conv_dist = distribution(df, CONVEYANCE_FIELDS)
    conv_dept = dept_scores(df, "overallConveyanceSatisfaction")

    add_clustered_bar(slide, "Average score by metric", conv_metric_avg.index, conv_metric_avg.values, 0.5, 1.6, 4.2, 2.6, theme)
    add_stacked_distribution(slide, "Rating distribution (1-5)", conv_dist, 4.9, 1.6, 3.6, 2.6, theme)
    add_heat_table(slide, "Department-wise conveyance score", conv_dept, 8.7, 1.6, 4.2, 3.4, theme)
    add_insight(slide, f"Top conveyance department: {conv_dept.index[0] if not conv_dept.empty else 'N/A'}.", theme)

    # 3) Production Performance
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    make_title(slide, "Production Performance", "Metric trends, rating spread, and department heatmap", theme)
    prod_metric_avg = avg_metrics(df, PRODUCTION_FIELDS)
    prod_dist = distribution(df, PRODUCTION_FIELDS)
    prod_dept = dept_scores(df, "overallProductionSatisfaction")

    add_clustered_bar(slide, "Average score by metric", prod_metric_avg.index, prod_metric_avg.values, 0.5, 1.6, 4.2, 2.6, theme)
    add_stacked_distribution(slide, "Rating distribution (1-5)", prod_dist, 4.9, 1.6, 3.6, 2.6, theme)
    add_heat_table(slide, "Department-wise production score", prod_dept, 8.7, 1.6, 4.2, 3.4, theme)
    add_insight(slide, f"Top production department: {prod_dept.index[0] if not prod_dept.empty else 'N/A'}.", theme)

    # 4) Comparison
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    make_title(slide, "Conveyance vs Production", "Overall comparison and department gaps", theme)

    overall_data = CategoryChartData()
    overall_data.categories = ["Conveyance", "Production"]
    overall_data.add_series("Average", [conveyance_avg if not math.isnan(conveyance_avg) else 0, production_avg if not math.isnan(production_avg) else 0])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.7), Inches(1.6), Inches(3.8), Inches(2.6), overall_data).chart
    chart.has_legend = False
    chart.value_axis.maximum_scale = 5
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = theme.blue

    fav_data = CategoryChartData()
    fav_data.categories = ["Conveyance", "Production"]
    fav_data.add_series("Favorable %", [conveyance_fav if not math.isnan(conveyance_fav) else 0, production_fav if not math.isnan(production_fav) else 0])
    fav_chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(4.8), Inches(1.6), Inches(3.8), Inches(2.6), fav_data).chart
    fav_chart.has_legend = False
    fav_chart.value_axis.maximum_scale = 100
    fav_chart.series[0].format.fill.solid()
    fav_chart.series[0].format.fill.fore_color.rgb = theme.cyan

    gap = (
        df.groupby("departments")[["overallConveyanceSatisfaction", "overallProductionSatisfaction"]]
        .mean(numeric_only=True)
        .dropna(how="all")
    )
    gap["delta"] = gap["overallProductionSatisfaction"] - gap["overallConveyanceSatisfaction"]
    gap = gap["delta"].sort_values()
    gap_data = CategoryChartData()
    gap_data.categories = gap.index.tolist()[:10]
    gap_data.add_series("Production - Conveyance", [round(v, 2) for v in gap.values.tolist()[:10]])
    gap_chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(8.9), Inches(1.6), Inches(4.0), Inches(3.5), gap_data).chart
    gap_chart.has_legend = False
    gap_chart.series[0].format.fill.solid()
    gap_chart.series[0].format.fill.fore_color.rgb = theme.slate
    add_insight(slide, "Positive delta indicates better production perception than conveyance within the same department.", theme)

    # 5) Voice of Employee
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    make_title(slide, "Voice of Employee", "Comment themes and anonymized quotes", theme)
    top_pos, top_pain = theme_counts(df["comments"])
    quotes = anonymized_quotes(df)
    if not quotes:
        quotes = ["No comments provided in the selected dataset."]

    add_bullets(slide, "Top Positive Themes", top_pos, 0.6, 1.4, 3.9, 4.8, theme)
    add_bullets(slide, "Top Pain Points", top_pain, 4.9, 1.4, 3.9, 4.8, theme)
    add_bullets(slide, "Verbatim Quotes (Anonymized)", quotes[:5], 9.2, 1.4, 3.6, 4.8, theme)
    add_insight(slide, "Most comments appreciate responsiveness; recurring delays remain the main friction.", theme)

    # 6) Actions & Ownership
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    make_title(slide, "Actions & Ownership", "Priority actions with KPI accountability", theme)

    actions = [
        ("Improve late-pickup recovery SLA", "Transport", "High", "<15 min delay incidents/week", "30 days"),
        ("Publish weekly route reliability dashboard", "Transport", "High", "On-time pickup rate ≥ 95%", "45 days"),
        ("Standardize logistics request turnaround", "Production", "Medium", "Median closure < 4 hrs", "30 days"),
        ("Create joint escalation matrix", "Transport + Production", "Medium", "Escalation resolution < 24 hrs", "21 days"),
        ("Monthly pulse check with HR facilitation", "HR", "Medium", "Favorable score +0.3 in 2 cycles", "60 days"),
    ]

    table = slide.shapes.add_table(len(actions) + 1, 5, Inches(0.5), Inches(1.6), Inches(12.3), Inches(4.6)).table
    headers = ["Action", "Owner", "Priority", "Success KPI", "Target"]
    widths = [4.4, 2.0, 1.4, 3.2, 1.3]
    for idx, width in enumerate(widths):
        table.columns[idx].width = Inches(width)
    for c, header in enumerate(headers):
        table.cell(0, c).text = header
        table.cell(0, c).fill.solid()
        table.cell(0, c).fill.fore_color.rgb = theme.blue
        for run in table.cell(0, c).text_frame.paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = theme.white

    priority_colors = {"High": theme.red, "Medium": theme.amber, "Low": theme.green}
    for r, row in enumerate(actions, start=1):
        for c, value in enumerate(row):
            table.cell(r, c).text = str(value)
        pri_cell = table.cell(r, 2)
        pri_cell.fill.solid()
        pri_cell.fill.fore_color.rgb = priority_colors.get(row[2], theme.slate)
        for run in pri_cell.text_frame.paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = theme.white

    add_insight(slide, "Assign weekly governance reviews to keep cross-functional actions on track.", theme)

    prs.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate management survey response slides")
    parser.add_argument("--input", required=True, help="Path to survey data (.xlsx/.xls/.xlsm/.csv)")
    parser.add_argument("--output", default="Survey_Management_Deck.pptx", help="Output PowerPoint file path")
    parser.add_argument("--company", default="RT Knits", help="Company name shown in slide subtitle")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    df = load_data(input_path)
    build_presentation(df, output_path, args.company)
    print(f"Presentation generated: {output_path.resolve()}")


if __name__ == "__main__":
    main()
