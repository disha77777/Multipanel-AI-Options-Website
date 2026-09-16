# Multipanel AI

Multipanel AI is a local Streamlit workspace for AI-assisted research, Excel reports, presentation generation, image-table OCR, and data cleaning.

## Features

- Generate a PowerPoint deck from a topic or thesis statement.
- Research a question and export a styled Excel workbook with data, sources, and charts.
- Extract tables from images into Excel with Gemini OCR.
- Create a research plan or a source-backed verified research report.
- Upload CSV or Excel data and apply natural-language cleaning operations.

## Setup

Use Python 3.10 or later, then install the dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Create a `.env` file in the project directory:

```env
GEMINI_API_KEY=your_gemini_api_key
```

Run the app:

```bash
streamlit run app.py
```

The `.env` file is ignored by Git and must not be committed.