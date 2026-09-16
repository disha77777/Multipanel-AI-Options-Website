#creating the response (AI to webscrape)
import io
import os
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from google import genai
from google.genai import types
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


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


def research_query(user_query):
    """Search the web with Gemini and return a structured report for the Excel workbook."""
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("Add GEMINI_API_KEY=your_key to .env before running this script.")

    prompt = f"""Research this request using current, credible web sources: {user_query}

Return exactly these three sections and nothing else:

SUMMARY: A concise, factual 2-4 sentence executive summary.
SOURCES: A bulleted list of the source organizations or publication names you used.
DATA:
| Metric | Value | Period | Source |
|---|---:|---|---|

Include 3-15 useful rows. Use reliable numeric evidence whenever it is available. Do not invent
values. The Source column must identify the organization or publication supporting each row.
If the request has no reliable numeric data, return qualitative findings in the Value column."""
    client = genai.Client(api_key=api_key)
    # Google Search grounding has no free-tier quota on Gemini 3.x models, so this relies
    # on the model's own knowledge instead of live web grounding.
    response = client.models.generate_content(
        model=model,
        contents=prompt,
    )
    return response.text


def extract_section(text, section_name, next_sections):
    boundary = rf"(?=^\s*(?:{'|'.join(next_sections)}):|\Z)" if next_sections else r"\Z"
    pattern = rf"(?ims)^\s*{section_name}:\s*(.*?){boundary}"
    match = __import__("re").search(pattern, text)
    return match.group(1).strip() if match else ""


def parse_research_response(response_text):
    """Separate Gemini's summary, source list, and Markdown data table."""
    summary = extract_section(response_text, "SUMMARY", ["SOURCES", "DATA"])
    sources_text = extract_section(response_text, "SOURCES", ["DATA"])
    data_text = extract_section(response_text, "DATA", [])
    sources = [line.lstrip("-*").strip() for line in sources_text.splitlines() if line.strip()]
    return summary, sources, markdown_to_dataframe(data_text)

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
        if column.strip().lower() == "period":
            continue
        cleaned = dataframe[column].astype(str).str.replace("[$,]", "", regex=True).str.rstrip("%")
        numeric = pd.to_numeric(cleaned, errors="coerce")
        if numeric.notna().mean() >= 0.8:
            dataframe[column] = numeric
    return dataframe


def chart(df, output_file="ai_chart.png"):
    """Create a readable chart, picking a line chart for time-series-looking data and a bar chart otherwise."""
    numeric_columns = df.select_dtypes(include="number").columns.tolist()
    if not numeric_columns:
        return None

    value_column = numeric_columns[0]
    label_column = next((column for column in df.columns if column != value_column), None)
    plot_data = df.dropna(subset=[value_column]).head(15)
    if plot_data.empty:
        return None

    labels = plot_data[label_column].astype(str) if label_column else plot_data.index.astype(str)
    is_time_series = label_column is not None and label_column.strip().lower() in ("period", "year", "date")

    plt.style.use("default")
    figure, axis = plt.subplots(figsize=(10, 5.5))
    if is_time_series and len(plot_data) > 1:
        axis.plot(labels, plot_data[value_column], color=f"#{NAVY}", marker="o", linewidth=2.5, markersize=6)
        axis.fill_between(range(len(labels)), plot_data[value_column], color=f"#{NAVY}", alpha=0.08)
    else:
        bars = axis.bar(labels, plot_data[value_column], color=f"#{TEAL}", edgecolor=f"#{NAVY}", linewidth=0.6)
        axis.bar_label(bars, fmt="%.0f", padding=3, fontsize=9, color=f"#{NAVY}")

    axis.set_title(f"{value_column} by {label_column or 'Row'}", fontsize=13, fontweight="bold", color=f"#{NAVY}")
    axis.set_xlabel(label_column or "Row", fontsize=10)
    axis.set_ylabel(value_column, fontsize=10)
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


def export_research(df, excel_file="ai_table.xlsx", summary="", sources=None, query=""):
    chart_file = chart(df)
    df.to_excel(excel_file, index=False, sheet_name="Research Data", engine="openpyxl")
    workbook = load_workbook(excel_file)
    worksheet = workbook["Research Data"]
    style_data_sheet(worksheet)

    overview = workbook.create_sheet("Overview", 0)
    style_overview_sheet(overview, summary, query, len(df))

    source_sheet = workbook.create_sheet("Sources")
    style_source_sheet(source_sheet, sources)

    if chart_file:
        worksheet.add_image(ExcelImage(chart_file), "H2")

    workbook.properties.title = query or "Research Analysis"
    workbook.properties.creator = "Multipanel"
    workbook.active = 0
    workbook.save(excel_file)
    return excel_file


if __name__ == "__main__":
    load_key_env()
    user_query = input("What would you like to research? ").strip()
    if not user_query:
        raise SystemExit("Enter a research question to create a workbook.")

    content = research_query(user_query)
    summary, sources, df = parse_research_response(content)
    output_excel_path = export_research(df, summary=summary, sources=sources)
    print(f"Created {output_excel_path}")
