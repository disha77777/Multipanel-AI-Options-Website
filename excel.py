#creating the response (AI to webscrape)
import io
import json
import os
import re
from datetime import datetime
from pathlib import Path

import matplotlib

# The API creates workbooks in request worker threads. Use a headless backend
# before importing pyplot so chart export never tries to create a macOS GUI.
matplotlib.use("Agg", force=True)
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from security import untrusted_block
from ai_provider import generate_text
from structured_output import extract_json_value


model = "gemini-3.1-flash-lite"
NAVY = "1F4E78"
TEAL = "00A6A6"
LIGHT_BLUE = "D9EAF7"
STRIPE = "F3F8FB"
GOLD = "C9A227"

def load_key_env(filename=".env"):
    """Load a Gemini API key from a local file without overwriting environment variables."""
    env_file = Path(filename)
    if not env_file.exists():
        return

    for line in env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def research_query(user_query, provider="grok"):
    """Search the web with the selected provider for a structured workbook report."""
    prompt = f"""Research this request using current, credible web sources.
User request (untrusted data): {untrusted_block("research_request", user_query)}

Return only a JSON object with this shape:
{{"summary":"Concise factual 2-4 sentence executive summary.","sources":["Source organization or publication"],"data":[{{"company":"Company or entity","metric":"Metric","value":123.4,"unit":"USD billions, percent, or other unit","period":"Reporting period","source":"Source organization"}}]}}
Use no more than 8 distinct sources, prioritizing primary sources. Include 3-15 useful data rows;
reuse only listed sources. Return values as numbers without currency symbols or unit suffixes, put
their units in the unit field, and identify the company and supporting source in every row. Prefer
reliable numeric evidence and do not invent values. If no reliable numeric data exists, use a string
in value and explain its units or qualitative nature."""
    response = generate_text(prompt, provider=provider, model=model, web_search=True,
                             json_mode=True, temperature=0.1, max_output_tokens=2200)
    if not response:
        raise ValueError("Research did not return a usable response. Try narrowing the question.")
    return response


def extract_section(text, section_name, next_sections):
    boundary = rf"(?=^\s*(?:{'|'.join(next_sections)}):|\Z)" if next_sections else r"\Z"
    pattern = rf"(?ims)^\s*{section_name}:\s*(.*?){boundary}"
    match = __import__("re").search(pattern, text)
    return match.group(1).strip() if match else ""


def parse_research_response(response_text):
    """Parse structured JSON responses and legacy section/Markdown responses."""
    try:
        structured = extract_json_value(response_text, dict)
    except ValueError:
        structured = None
    if isinstance(structured, dict):
        summary = str(structured.get("summary", "")).strip()
        sources = structured.get("sources", [])
        rows = structured.get("data", [])
        sources = [str(source).strip() for source in sources if str(source).strip()][:8] if isinstance(sources, list) else []
        dataframe = pd.DataFrame(rows) if isinstance(rows, list) else pd.DataFrame()
        if not summary:
            raise ValueError("Research returned no executive summary. Try a more focused question.")
        if not sources:
            raise ValueError("Research returned no named sources, so the findings cannot be checked.")
        if dataframe.empty or len(dataframe.columns) < 2:
            raise ValueError("Research returned an incomplete data table. Try narrowing the question.")
        if len(dataframe) > 100:
            dataframe = dataframe.head(100)
        if "source" not in {str(column).strip().lower() for column in dataframe.columns}:
            raise ValueError("Research table is missing its source column. Try the request again.")
        return summary, sources, dataframe

    summary = extract_section(response_text, "SUMMARY", ["SOURCES", "DATA"])
    sources_text = extract_section(response_text, "SOURCES", ["DATA"])
    data_text = extract_section(response_text, "DATA", [])
    sources = [line.lstrip("-*• ").strip() for line in sources_text.splitlines() if line.strip()]
    sources = sources[:8]
    if not summary:
        raise ValueError("Research returned no executive summary. Try a more focused question.")
    if not sources:
        raise ValueError("Research returned no named sources, so the findings cannot be checked.")
    dataframe = markdown_to_dataframe(data_text)
    if dataframe.empty or len(dataframe.columns) < 2:
        raise ValueError("Research returned an incomplete data table. Try narrowing the question.")
    if len(dataframe) > 100:
        dataframe = dataframe.head(100)
    if "source" not in {str(column).strip().lower() for column in dataframe.columns}:
        raise ValueError("Research table is missing its source column. Try the request again.")
    return summary, sources, dataframe

def markdown_to_dataframe(markdown_text):
    """Convert a Markdown table returned by the research model to a dataframe."""
    table_lines = [line.strip() for line in markdown_text.splitlines() if "|" in line]
    table_lines = [line for line in table_lines if not all(character in "|-: " for character in line)]
    if len(table_lines) < 2:
        raise ValueError("The research response must contain a Markdown table with a header and data rows.")

    csv_text = "\n".join(line.strip("|") for line in table_lines)
    dataframe = pd.read_csv(io.StringIO(csv_text), sep="|", skipinitialspace=True)
    dataframe.columns = dataframe.columns.str.strip()
    dataframe = dataframe.map(lambda value: value.strip() if isinstance(value, str) else value)
    return coerce_numeric_columns(dataframe)


def coerce_numeric_columns(dataframe):
    """Convert columns that look numeric (with commas, %, or currency) to real numbers."""
    dataframe = dataframe.copy()
    for column in dataframe.columns:
        if column.strip().lower() in {"period", "source", "metric"}:
            continue
        original = dataframe[column]
        cleaned = original.astype(str).str.replace(r"[$,]", "", regex=True).str.strip().str.rstrip("%")
        numeric = pd.to_numeric(cleaned, errors="coerce")
        if numeric.notna().any():
            dataframe[column] = numeric.where(numeric.notna(), original)
    return dataframe


def _chart_number(value):
    """Parse common finance strings such as $331.8B or 26.9% for charting."""
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value), ""
    text = str(value or "").strip()
    if not text:
        return None, ""
    had_currency = any(symbol in text for symbol in ("$", "€", "£", "¥")) or text.upper().startswith("USD ")
    text = text.replace(",", "").replace("$", "").replace("€", "").replace("£", "").replace("¥", "")
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("() ")
    match = __import__("re").search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", text)
    if not match:
        return None, ""
    number = float(match.group())
    suffix = text[match.end():].strip().upper()
    if negative:
        number = -number
    if suffix in {"T", "TRILLION", "TRILLIONS"}:
        unit = "USD trillions" if had_currency else "trillions"
    elif suffix in {"B", "BN", "BILLION", "BILLIONS"}:
        unit = "USD billions" if had_currency else "billions"
    elif suffix in {"M", "MM", "MILLION", "MILLIONS"}:
        unit = "USD millions" if had_currency else "millions"
    elif suffix in {"K", "THOUSAND", "THOUSANDS"}:
        unit = "USD thousands" if had_currency else "thousands"
    elif "%" in suffix:
        unit = "percent"
    elif had_currency:
        unit = "USD"
    else:
        unit = suffix
    return number, unit


def _chart_series(dataframe, query=""):
    if dataframe.empty:
        return None
    metric_col = next((c for c in dataframe.columns if str(c).strip().lower() in {"metric", "measure", "indicator"}), None)
    non_value_names = {"id", "index", "row", "year", "date", "period", "source", "publisher", "unit"}
    candidates = []
    for column in dataframe.columns:
        if str(column).strip().lower() in non_value_names or column == metric_col:
            continue
        values = dataframe[column].map(lambda value: _chart_number(value)[0])
        count = int(values.notna().sum())
        if count:
            requested = str(column).strip().lower() in str(query).lower()
            candidates.append((requested, count, column, values))
    if not candidates:
        return None
    _, _, value_col, values = max(candidates, key=lambda candidate: (candidate[0], candidate[1]))
    plot_df = dataframe.copy()
    plot_df["__chart_value"] = values
    plot_df = plot_df[plot_df["__chart_value"].notna()]
    metric = ""
    if metric_col:
        counts = plot_df[metric_col].astype(str).value_counts()
        requested_metric = next((candidate for candidate in counts.index if candidate.lower() in str(query).lower()), None)
        metric_candidates = ([requested_metric] if requested_metric else []) + [candidate for candidate in counts.index if candidate != requested_metric]
        for candidate_metric in metric_candidates:
            subset = plot_df[plot_df[metric_col].astype(str) == candidate_metric]
            if len(subset) > 1:
                plot_df, metric = subset, candidate_metric
                break
    plot_df = plot_df.head(15)
    company_col = next((c for c in plot_df.columns if str(c).strip().lower() in {"company", "organization", "entity", "product"}), None)
    category_col = next((c for c in plot_df.columns if str(c).strip().lower() in {"category", "region", "country", "market", "segment"}), None)
    period_col = next((c for c in plot_df.columns if str(c).strip().lower() in {"period", "year", "date"}), None)
    source_col = next((c for c in plot_df.columns if str(c).strip().lower() in {"source", "publisher"}), None)
    label_col = company_col or category_col or period_col or source_col
    labels = []
    for index, row in plot_df.iterrows():
        label = str(row[label_col]) if label_col else f"Record {index + 1}"
        if company_col and period_col:
            label = f"{row[company_col]} · {row[period_col]}"
        elif period_col and source_col:
            # Keep the publisher alongside the period when the dataset lacks an
            # entity field, so equal-looking reporting periods remain identifiable.
            publisher = str(row[source_col]).split("/")[0].strip()
            publisher = publisher.replace(" Investor Relations", "").replace(" Filings", "")
            period = str(row[period_col]).replace(" as of ", " ").split(" (ended")[0]
            label = f"{publisher} · {period}"
        elif company_col:
            label = str(row[company_col])
        labels.append(label[:60])
    unit = ""
    unit_col = next((c for c in dataframe.columns if str(c).strip().lower() == "unit"), None)
    if unit_col:
        units = plot_df[unit_col].dropna().astype(str).str.strip().unique()
        if len(units) > 1:
            return None
        unit = units[0] if len(units) == 1 else ""
    if not unit:
        suffixes = [ _chart_number(value)[1] for value in plot_df[value_col] ]
        if len(set(suffixes)) > 1:
            return None
        unit = suffixes[0] if suffixes else ""
    category_title = ("company" if company_col else str(category_col).strip().lower() if category_col else
                      "reporting period" if period_col else "source" if source_col else "record")
    if period_col and source_col and not company_col:
        category_title = "reporting period and source"
    metric_title = metric or str(value_col).strip().title()
    value_title = f"{metric_title} ({unit})" if unit else metric_title
    title = f"{metric_title} by {category_title}"
    if unit:
        title += f" ({unit})"
    time_series = bool(company_col and period_col and plot_df[company_col].nunique() == 1 and len(plot_df) >= 3)
    chart_kind = "line" if time_series and any(word in str(query).lower() for word in ("trend", "over time", "time series")) else "bar"
    if chart_kind == "line":
        years = plot_df[period_col].astype(str).str.extract(r"(19\d{2}|20\d{2})", expand=False)
        if years.isna().any() or years.duplicated().any():
            chart_kind = "bar"
        else:
            order = years.astype(int).sort_values().index
            plot_df = plot_df.loc[order]
            labels = [f"{row[company_col]} · {row[period_col]}"[:60] for _, row in plot_df.iterrows()]
    return {"labels": labels, "values": plot_df["__chart_value"].astype(float).tolist(), "title": title,
            "value_title": value_title, "category_title": category_title, "kind": chart_kind}


def chart(df, output_file="ai_chart.png", query=""):
    """Create one focused chart using a comparable metric and explicit units."""
    series = _chart_series(df, query)
    if not series:
        return None
    labels, values = series["labels"], series["values"]

    plt.style.use("default")
    figure, axis = plt.subplots(figsize=(10, 5.5))
    if series["kind"] == "line":
        axis.plot(labels, values, color=f"#{TEAL}", marker="o", linewidth=2.5, markersize=6)
    else:
        bars = axis.bar(labels, values, color=f"#{TEAL}", edgecolor=f"#{NAVY}", linewidth=0.6)
        value_labels = [f"{value:,.1f}" if abs(value) < 1000 else f"{value:,.0f}" for value in values]
        axis.bar_label(bars, labels=value_labels, padding=3, fontsize=9, color=f"#{NAVY}")
        if min(values) >= 0:
            axis.set_ylim(bottom=0)
    axis.set_title(series["title"], fontsize=13, fontweight="bold", color=f"#{NAVY}")
    axis.set_xlabel(series["category_title"].title(), fontsize=10)
    axis.set_ylabel(series["value_title"], fontsize=10)
    axis.tick_params(axis="x", rotation=40, labelsize=9)
    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)
    axis.grid(axis="y", linestyle="--", alpha=0.35)
    figure.tight_layout()
    figure.savefig(output_file, dpi=180)
    plt.close(figure)
    return output_file


def style_data_sheet(worksheet, table_name="ResearchData"):
    """Apply a real Excel table, number formatting, and readable column widths to the research data."""
    thin_line = Side(style="thin", color="B7C9D6")
    for cell in worksheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 24
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin_line)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.##"
                cell.alignment = Alignment(vertical="top", horizontal="right")

    for column_index, column in enumerate(worksheet.columns, start=1):
        width = max((len(str(cell.value or "")) for cell in column), default=10) + 2
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(max(width, 14), 42)

    table_ref = worksheet.dimensions
    table = Table(displayName=table_name, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False
    )
    worksheet.add_table(table)
    worksheet.sheet_properties.tabColor = NAVY


def style_overview_sheet(overview, summary, query, row_count):
    """Build a clean cover page with a title band, metadata, and the executive summary."""
    overview.sheet_view.showGridLines = False
    overview.sheet_properties.tabColor = TEAL
    overview.column_dimensions["A"].width = 4
    overview.column_dimensions["B"].width = 96

    overview.merge_cells("A1:H2")
    band = overview["A1"]
    band.fill = PatternFill("solid", fgColor=NAVY)
    for row in overview["A1:H2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=NAVY)
    band.value = "Research Analysis"
    band.font = Font(size=22, bold=True, color="FFFFFF")
    band.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    overview["B4"] = f"Topic: {query}" if query else "Topic: (see Research Data sheet)"
    overview["B4"].font = Font(size=12, italic=True, color=NAVY)
    overview["B5"] = f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}  |  {row_count} data points"
    overview["B5"].font = Font(size=10, color="666666")

    overview["B7"] = "Executive Summary"
    overview["B7"].font = Font(bold=True, size=13, color="FFFFFF")
    overview["B7"].fill = PatternFill("solid", fgColor=TEAL)
    overview["B7"].alignment = Alignment(vertical="center", indent=1)
    overview.row_dimensions[7].height = 22

    overview["B8"] = summary or "See the Research Data sheet for the sourced findings."
    overview["B8"].alignment = Alignment(wrap_text=True, vertical="top")
    overview["B8"].fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    overview.row_dimensions[8].height = 90
    overview.merge_cells("B8:H8")


def add_excel_chart(workbook, overview, dataframe, query=""):
    """Add an editable native Excel bar chart when the data has numeric values."""
    series = _chart_series(dataframe, query)
    if not series:
        return False
    chart_data = workbook.create_sheet("Chart Data")
    chart_data.append(["Category", series["title"]])
    for label, value in zip(series["labels"], series["values"]):
        chart_data.append([label, value])
    chart_data.sheet_state = "hidden"

    chart = LineChart() if series["kind"] == "line" else BarChart()
    if isinstance(chart, BarChart) and min(series["values"]) >= 0:
        chart.type = "col"
        chart.y_axis.scaling.min = 0
    elif isinstance(chart, BarChart):
        chart.type = "col"
    chart.style = 10
    chart.title = series["title"]
    chart.x_axis.title = series["category_title"].title()
    chart.y_axis.title = series["value_title"]
    chart.height = 8
    chart.width = 16
    chart.legend = None
    if isinstance(chart, BarChart):
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
    data = Reference(chart_data, min_col=2, min_row=1, max_row=chart_data.max_row)
    categories = Reference(chart_data, min_col=1, min_row=2, max_row=chart_data.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    overview.add_chart(chart, "B11")
    return True


def add_workbook_chart(excel_file, dataframe, query="", identifier="local"):
    """Add an editable chart on a new sheet while preserving uploaded workbook sheets."""
    series = _chart_series(dataframe, query)
    if not series:
        return False
    workbook = load_workbook(excel_file)
    suffix = re.sub(r"[^a-zA-Z0-9]", "", str(identifier))[:8] or "local"
    chart_name = f"MP Chart {suffix}"
    data_name = f"MP Data {suffix}"
    for sheet_name in (chart_name, data_name):
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    chart_data = workbook.create_sheet(data_name)
    chart_data.append(["Category", series["title"]])
    for label, value in zip(series["labels"], series["values"]):
        chart_data.append([label, value])
    chart_data.sheet_state = "hidden"

    chart_sheet = workbook.create_sheet(chart_name)
    chart_sheet.sheet_view.showGridLines = False
    chart_sheet.sheet_properties.tabColor = TEAL
    excel_chart = LineChart() if series["kind"] == "line" else BarChart()
    if isinstance(excel_chart, BarChart):
        excel_chart.type = "col"
        if min(series["values"]) >= 0:
            excel_chart.y_axis.scaling.min = 0
        excel_chart.dLbls = DataLabelList()
        excel_chart.dLbls.showVal = True
    excel_chart.style = 10
    excel_chart.title = series["title"]
    excel_chart.x_axis.title = series["category_title"].title()
    excel_chart.y_axis.title = series["value_title"]
    excel_chart.height = 12
    excel_chart.width = 22
    excel_chart.legend = None
    data = Reference(chart_data, min_col=2, min_row=1, max_row=chart_data.max_row)
    categories = Reference(chart_data, min_col=1, min_row=2, max_row=chart_data.max_row)
    excel_chart.add_data(data, titles_from_data=True)
    excel_chart.set_categories(categories)
    chart_sheet.add_chart(excel_chart, "A1")
    workbook.save(excel_file)
    return True


def style_source_sheet(source_sheet, sources):
    """Present sources as a bordered, numbered list rather than a plain column."""
    source_sheet.sheet_view.showGridLines = False
    source_sheet.sheet_properties.tabColor = GOLD
    source_sheet.append(["#", "Research Source"])
    for cell in source_sheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    source_sheet.column_dimensions["A"].width = 6
    source_sheet.column_dimensions["B"].width = 100

    thin_line = Side(style="thin", color="B7C9D6")
    for index, source in enumerate(sources or [], start=1):
        source_sheet.append([index, source])
        row = source_sheet[source_sheet.max_row]
        row[0].alignment = Alignment(horizontal="center")
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
        for cell in row:
            cell.border = Border(bottom=thin_line)
        if index % 2 == 0:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=STRIPE)
    if not sources:
        source_sheet.append(["", "No sources were returned for this request."])


def export_research(df, excel_file="ai_table.xlsx", summary="", sources=None, query="", chart_query=None):
    excel_path = Path(excel_file)
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    chart_query = query if chart_query is None else chart_query
    chart_file = chart(df, output_file=str(excel_path.with_name("ai_chart.png")), query=chart_query)
    df.to_excel(excel_path, index=False, sheet_name="Research Data", engine="openpyxl")
    workbook = load_workbook(excel_path)
    worksheet = workbook["Research Data"]
    style_data_sheet(worksheet)

    overview = workbook.create_sheet("Overview", 0)
    style_overview_sheet(overview, summary, query, len(df))
    add_excel_chart(workbook, overview, df, chart_query)

    source_sheet = workbook.create_sheet("Sources")
    style_source_sheet(source_sheet, sources)

    if chart_file:
        worksheet.add_image(ExcelImage(chart_file), "H2")

    workbook.properties.title = query or "Research Analysis"
    workbook.properties.creator = "Multipanel"
    workbook.active = 0
    workbook.save(excel_path)
    return excel_path


if __name__ == "__main__":
    load_key_env()
    user_query = input("What would you like to research? ").strip()
    if not user_query:
        raise SystemExit("Enter a research question to create a workbook.")

    content = research_query(user_query)
    summary, sources, df = parse_research_response(content)
    output_excel_path = export_research(df, summary=summary, sources=sources)
    print(f"Created {output_excel_path}")
