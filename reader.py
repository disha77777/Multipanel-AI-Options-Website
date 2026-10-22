import re
from pathlib import Path

import pandas as pd
from google import genai
from google.genai import types
from PIL import Image, ImageEnhance, ImageOps, UnidentifiedImageError
from ai_provider import generate_text, normalize_provider
from security import api_key


OCR_PROMPT = """Extract the user-provided image into data. Read printed and handwritten text, including faint pencil and cursive.
Do not add facts or infer illegible characters from context. Preserve exact spelling, numeric digits, signs, decimals, and row/column order. For uncertain cells, write [unclear: best visible reading?].
Return only JSON using this shape:
{"transcription":"text outside any table, or empty string","table":{"columns":["column name"],"rows":[["cell value"]]} or null,"uncertainties":["row and column needing review"]}
Use null for table when no table is present. Ensure every row has the same number of values as columns."""


def prepare_ocr_images(image_file):
    """Return the original and a contrast-enhanced, enlarged view for difficult scans."""
    try:
        with Image.open(image_file) as source:
            if source.width * source.height > 40_000_000:
                raise ValueError("This image has too many pixels to process safely. Resize it and try again.")
            original = ImageOps.exif_transpose(source).convert("RGB")
            original.thumbnail((4096, 4096), Image.Resampling.LANCZOS)
    except (UnidentifiedImageError, OSError) as error:
        raise ValueError("This image could not be opened. Use a clear JPG, PNG, or WebP image.") from error

    scale = min(2.0, 2048 / max(original.size))
    if scale > 1:
        size = tuple(max(1, round(dimension * scale)) for dimension in original.size)
        original = original.resize(size, Image.Resampling.LANCZOS)

    enhanced = ImageOps.grayscale(original)
    enhanced = ImageOps.autocontrast(enhanced, cutoff=1)
    enhanced = ImageEnhance.Contrast(enhanced).enhance(1.5)
    enhanced = ImageEnhance.Sharpness(enhanced).enhance(1.8).convert("RGB")
    return original, enhanced


def extract_text_from_image(image_file, model="gemini-3.1-flash-lite", provider="grok"):
    """Transcribe printed or handwritten text from the original and enhanced image."""
    original, enhanced = prepare_ocr_images(image_file)
    provider = normalize_provider(provider)
    if provider == "gemini":
        client = genai.Client(api_key=api_key())
        config = types.GenerateContentConfig(response_mime_type="application/json", temperature=0, max_output_tokens=8192)
        response = client.models.generate_content(model=model, contents=[original, OCR_PROMPT], config=config)
        text = _normalize_ocr_response(_response_text(response))
        if text:
            return text
        fallback = client.models.generate_content(model="gemini-3.6-flash", contents=[enhanced, OCR_PROMPT], config=config)
        text = _normalize_ocr_response(_response_text(fallback))
        if text:
            return text
        finish_reason = _finish_reason(fallback) or _finish_reason(response) or "no text candidate"
        raise ValueError(f"The vision service returned no transcription ({finish_reason}). Try uploading the image again.")

    text = _normalize_ocr_response(generate_text(OCR_PROMPT, provider=provider, model=model, json_mode=True,
                                                temperature=0, max_output_tokens=8192, image=original))
    if text:
        return text

    # Retry with a separate enhanced image and a second multimodal model if the
    # first model returns an empty candidate (some valid images get no text part).
    raise ValueError("Grok returned no transcription. Try a clearer image or switch to Gemini.")


def _normalize_ocr_response(text):
    """Convert structured model output to Markdown used by the preview and exporters."""
    if not text:
        return ""
    candidate = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.IGNORECASE)
    try:
        result = __import__("json").loads(candidate)
    except (ValueError, TypeError):
        return text
    if not isinstance(result, dict):
        return text

    lines = []
    table = result.get("table")
    if isinstance(table, dict) and isinstance(table.get("columns"), list) and isinstance(table.get("rows"), list):
        columns = [str(value) for value in table["columns"]]
        if columns:
            lines.append("| " + " | ".join(value.replace("|", r"\|") for value in columns) + " |")
            lines.append("| " + " | ".join("---" for _ in columns) + " |")
            for row in table["rows"]:
                if isinstance(row, list) and len(row) == len(columns):
                    lines.append("| " + " | ".join(str(value).replace("|", r"\|") for value in row) + " |")
    transcription = result.get("transcription")
    if isinstance(transcription, str) and transcription.strip():
        lines.append(transcription.strip())
    uncertainties = result.get("uncertainties")
    if isinstance(uncertainties, list) and uncertainties:
        lines.append("\nNeeds review: " + "; ".join(str(item) for item in uncertainties))
    return "\n".join(lines) or text


def _response_text(response):
    """Extract text from the standard response field or its candidate parts."""
    try:
        text = response.text
    except (AttributeError, ValueError):
        text = None
    if isinstance(text, str) and text.strip():
        return text.strip()
    for candidate in getattr(response, "candidates", None) or []:
        content = getattr(candidate, "content", None)
        for part in getattr(content, "parts", None) or []:
            part_text = getattr(part, "text", None)
            if isinstance(part_text, str) and part_text.strip():
                return part_text.strip()
    return ""


def _finish_reason(response):
    candidates = getattr(response, "candidates", None) or []
    return str(getattr(candidates[0], "finish_reason", "")) if candidates else ""


def parse_ocr_to_dataframe(markdown_text):
    """Parse markdown tables from OCR output into a pandas DataFrame."""
    if not isinstance(markdown_text, str) or "|" not in markdown_text:
        return None

    lines = markdown_text.splitlines()
    table_rows = []
    separator_seen = False
    started = False
    for line in lines:
        if "|" not in line:
            if started and separator_seen:
                break
            continue
        cells = [cell.strip().replace(r"\|", "|") for cell in re.split(r"(?<!\\)\|", line.strip().strip("|"))]
        if cells and all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells):
            if started:
                separator_seen = True
            continue
        if not cells or not any(cells):
            continue
        started = True
        table_rows.append(cells)

    if not separator_seen or len(table_rows) < 2:
        return None
    width = len(table_rows[0])
    if width < 1 or any(len(row) != width for row in table_rows[1:]):
        return None
    headers = [header or f"Column {index + 1}" for index, header in enumerate(table_rows[0])]
    return pd.DataFrame(table_rows[1:], columns=headers)


def process_image_ocr(image_file, output_excel_path="ocr_output.xlsx", provider="grok"):
    """Full pipeline: extract text from image, parse table, save to Excel."""
    markdown_text = extract_text_from_image(image_file, provider=provider)
    df = parse_ocr_to_dataframe(markdown_text)

    if df is not None:
        output_path = Path(output_excel_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(output_path, index=False)
        return markdown_text, df, output_path
    return markdown_text, None, None
