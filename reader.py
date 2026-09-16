import io
import os
import re

import pandas as pd
from google import genai
from PIL import Image


def extract_text_from_image(image_file, model="gemini-3.1-flash-lite"):
    """Use Gemini to extract text and tables from an image file."""
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("Add GEMINI_API_KEY=your_key to .env before using OCR.")

    # Open image and convert to bytes
    image = Image.open(image_file)

    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model,
        contents=[
            image,
            (
                "Act as an advanced OCR detection tool. "
                "Transcribe all the text found in the image exactly how it appears. "
                "If it is formatted in a table, make a table for the values using Markdown format with | separators. "
                "Return the extracted text and tables."
            ),
        ],
    )

    return response.text


def parse_ocr_to_dataframe(markdown_text):
    """Parse markdown tables from OCR output into a pandas DataFrame."""
    # Find markdown tables (lines with | separators)
    lines = markdown_text.split("\n")
    table_lines = [line for line in lines if "|" in line]

    if not table_lines:
        return None

    # Clean table lines: remove formatting characters
    cleaned_lines = []
    for line in table_lines:
        line = line.strip()
        if line and not all(c in "| -:" for c in line.strip()):
            cleaned_lines.append(line)

    if not cleaned_lines:
        return None

    # Parse markdown table into CSV
    cleaned_table = "\n".join(cleaned_lines)
    try:
        df = pd.read_csv(io.StringIO(cleaned_table), sep="|", skipinitialspace=True)
        # Remove unnamed columns and extra whitespace
        df = df.dropna(how="all", axis=1).loc[:, lambda x: ~x.columns.str.contains("^Unnamed")]
        df.columns = df.columns.str.strip()
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
        return df
    except Exception:
        return None


def process_image_ocr(image_file, output_excel_path="ocr_output.xlsx"):
    """Full pipeline: extract text from image, parse table, save to Excel."""
    markdown_text = extract_text_from_image(image_file)
    df = parse_ocr_to_dataframe(markdown_text)

    if df is not None:
        df.to_excel(output_excel_path, index=False)
        return markdown_text, df, output_excel_path
    else:
        return markdown_text, None, None












