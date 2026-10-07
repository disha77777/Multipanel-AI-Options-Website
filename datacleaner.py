import json
import os
import re

import pandas as pd
from dateutil import parser
from ai_provider import generate_text
from security import untrusted_block
from structured_output import extract_json_value


def load_dataframe(file_path):
    """Read a CSV or Excel file into a dataframe based on its extension."""
    file_path = str(file_path)
    if file_path.lower().endswith(".csv"):
        return pd.read_csv(file_path)
    return pd.read_excel(file_path)


def cleaning(df):
    """Drop duplicate rows and standardize column names to snake_case."""
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    object_columns = df.select_dtypes(include=["object", "string"]).columns
    for column in object_columns:
        df[column] = df[column].apply(lambda value: value.strip() if isinstance(value, str) else value)
    return df.drop_duplicates()


def parse_dates(df):
    """Convert any column with 'date' in its name to real datetime values."""
    df = df.copy()
    for column in df.columns:
        if "date" in str(column).lower():
            df[column] = pd.to_datetime(df[column], errors="coerce")
    return df


def data_types_con(df):
    """Infer and apply sensible dtypes: numeric, category, and integer id columns."""
    df = df.copy()
    for column in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[column]):
            continue
        if column in ("region", "category"):
            df[column] = df[column].astype("category")
        elif column == "id" or column.endswith("_id"):
            df[column] = pd.to_numeric(df[column], errors="coerce").astype("Int64")
        else:
            converted = pd.to_numeric(df[column], errors="coerce")
            if converted.notna().mean() >= 0.8:
                df[column] = converted
    return df


def summarize(df):
    """Return a lightweight profile: shape, dtypes, missing values, and basic stats."""
    return {
        "rows": len(df),
        "columns": len(df.columns),
        "missing_by_column": df.isna().sum().to_dict(),
        "dtypes": df.dtypes.astype(str).to_dict(),
        "describe": df.describe(include="all").transpose(),
    }


def clean_pipeline(df):
    """Run the full cleaning pipeline: dedupe/standardize, parse dates, fix dtypes."""
    df = cleaning(df)
    df = parse_dates(df)
    df = data_types_con(df)
    return df


def _normalize_header(value):
    text = str(value).replace("\u00a0", " ").strip()
    text = re.sub(r"\s+", "_", text, flags=re.UNICODE)
    text = re.sub(r"[^\w]+", "_", text, flags=re.UNICODE).strip("_").lower()
    text = re.sub(r"_+", "_", text)
    return text or "column"


def _parse_cleaning_number(value):
    if pd.isna(value):
        return None
    text = re.sub(r"\[[^\]]*\]", "", str(value)).replace(",", "")
    match = re.search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", text)
    return float(match.group()) if match else None


def _op_auto_clean(df, _op):
    """Apply conservative, data-aware cleanup without inventing missing values."""
    working = df.copy(deep=True)
    log = []

    headers = [_normalize_header(column) for column in working.columns]
    counts = {}
    unique_headers = []
    for header in headers:
        counts[header] = counts.get(header, 0) + 1
        unique_headers.append(header if counts[header] == 1 else f"{header}_{counts[header]}")
    if list(working.columns) != unique_headers:
        working.columns = unique_headers
        log.append("Standardized column names to lowercase snake_case.")

    text_columns = working.select_dtypes(include=["object", "string"]).columns
    trimmed = 0
    for column in text_columns:
        original = working[column]
        cleaned = original.map(lambda value: value.strip() if isinstance(value, str) else value)
        cleaned = cleaned.replace(r"^\s*$", pd.NA, regex=True)
        trimmed += int((original.fillna("").astype(str) != cleaned.fillna("").astype(str)).sum())
        working[column] = cleaned
    if trimmed:
        log.append(f"Trimmed or normalized {trimmed} text cell(s).")

    before = len(working)
    working = working.drop_duplicates()
    duplicate_count = before - len(working)
    if duplicate_count:
        log.append(f"Removed {duplicate_count} exact duplicate row(s).")

    converted = []
    for column in working.select_dtypes(include=["object", "string"]).columns:
        name = str(column).lower()
        if any(token in name for token in ("year", "date", "period", "ref", "artist", "title", "name", "email", "address")):
            continue
        source = working[column]
        present = source.notna()
        if not present.any():
            continue
        parsed = source.map(_parse_cleaning_number)
        ratio = parsed[present].notna().mean()
        if ratio < 0.8:
            continue
        numeric = pd.to_numeric(parsed, errors="coerce")
        valid = numeric.dropna()
        if not valid.empty and (valid % 1 == 0).all():
            working[column] = numeric.astype("Int64")
        else:
            working[column] = numeric
        converted.append(str(column))
    if converted:
        log.append("Converted numeric values (including currency and footnote-marked values) in: " + ", ".join(converted) + ".")

    if not log:
        log.append("No additional safe cleanup was needed; the data is already normalized.")
    log.append("Kept missing values unchanged so no data was guessed or invented.")
    return working, " ".join(log)


def _is_general_clean_request(message):
    text = str(message or "").lower()
    if not re.search(r"\b(clean|tidy|prepare)\b", text):
        return False
    specific_action = r"\b(remove|drop|deduplicat|missing|fill|imput|date|whitespace|case|rename|convert|replace|column|currency|standardiz\w*\s+(?:the\s+)?(?:currency|date|case|column))\b"
    if re.search(specific_action, text):
        return False
    return bool(re.search(r"\b(it|this|my|the|data|dataset|file|table|efficient|efficiency|overall|generally|analysis)\b", text))


# Whitelisted operations the chat assistant is allowed to request. Each function
# takes (df, op) and returns (new_df, message). No user or AI text is ever executed
# as code -- only these fixed, parameter-validated operations run.

def _op_drop_duplicates(df, op):
    before = len(df)
    df = df.drop_duplicates()
    return df, f"Dropped {before - len(df)} duplicate row(s)."


def _op_drop_column(df, op):
    column = op.get("column")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    return df.drop(columns=[column]), f"Dropped column '{column}'."


def _op_rename_column(df, op):
    old, new = op.get("from"), op.get("to")
    if old not in df.columns:
        return df, f"Skipped: column '{old}' not found."
    return df.rename(columns={old: new}), f"Renamed column '{old}' to '{new}'."


def _op_drop_rows_where(df, op):
    column, condition, value = op.get("column"), op.get("condition"), op.get("value")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    mask = _build_mask(df[column], condition, value)
    if mask is None:
        return df, f"Skipped: unknown condition '{condition}'."
    before = len(df)
    df = df[~mask]
    return df, f"Dropped {before - len(df)} row(s) where {column} {condition} {value!r}."


def _op_keep_rows_where(df, op):
    column, condition, value = op.get("column"), op.get("condition"), op.get("value")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    mask = _build_mask(df[column], condition, value)
    if mask is None:
        return df, f"Skipped: unknown condition '{condition}'."
    before = len(df)
    df = df[mask]
    return df, f"Kept {len(df)} of {before} row(s) where {column} {condition} {value!r}."


def _build_mask(series, condition, value):
    numeric_series = pd.to_numeric(series, errors="coerce")
    if condition == "equals":
        return series.astype(str) == str(value)
    if condition == "not_equals":
        return series.astype(str) != str(value)
    if condition == "is_null":
        return series.isna()
    if condition == "is_not_null":
        return series.notna()
    if condition == "contains":
        return series.astype(str).str.contains(str(value), case=False, na=False)
    if condition == "greater_than":
        return numeric_series > float(value)
    if condition == "less_than":
        return numeric_series < float(value)
    return None


def _op_fill_missing(df, op):
    column, method, value = op.get("column"), op.get("method"), op.get("value")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    df = df.copy()
    if method == "mean":
        df[column] = df[column].fillna(pd.to_numeric(df[column], errors="coerce").mean())
    elif method == "median":
        df[column] = df[column].fillna(pd.to_numeric(df[column], errors="coerce").median())
    elif method == "mode":
        mode_values = df[column].mode()
        if not mode_values.empty:
            df[column] = df[column].fillna(mode_values.iloc[0])
    elif method == "value":
        df[column] = df[column].fillna(value)
    else:
        return df, f"Skipped: unknown fill method '{method}'."
    return df, f"Filled missing values in '{column}' using {method}."


def _op_convert_type(df, op):
    column, target_type = op.get("column"), op.get("to")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    df = df.copy()
    try:
        if target_type == "int":
            df[column] = pd.to_numeric(df[column], errors="coerce").astype("Int64")
        elif target_type == "float":
            df[column] = pd.to_numeric(df[column], errors="coerce")
        elif target_type == "str":
            df[column] = df[column].astype(str)
        elif target_type == "category":
            df[column] = df[column].astype("category")
        elif target_type == "datetime":
            df[column] = df[column].apply(
                lambda value: parser.parse(str(value)) if pd.notna(value) else pd.NaT
            )
        else:
            return df, f"Skipped: unknown type '{target_type}'."
    except (ValueError, TypeError):
        return df, f"Skipped: could not convert '{column}' to {target_type}."
    return df, f"Converted '{column}' to {target_type}."


def _op_strip_whitespace(df, op):
    column = op.get("column")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    df = df.copy()
    df[column] = df[column].astype(str).str.strip()
    return df, f"Stripped whitespace in '{column}'."


def _op_standardize_case(df, op):
    column, case = op.get("column"), op.get("case")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    df = df.copy()
    text = df[column].astype(str)
    if case == "lower":
        df[column] = text.str.lower()
    elif case == "upper":
        df[column] = text.str.upper()
    elif case == "title":
        df[column] = text.str.title()
    else:
        return df, f"Skipped: unknown case '{case}'."
    return df, f"Standardized '{column}' to {case} case."


def _op_replace_value(df, op):
    column, old_value, new_value = op.get("column"), op.get("old"), op.get("new")
    if column not in df.columns:
        return df, f"Skipped: column '{column}' not found."
    df = df.copy()
    df[column] = df[column].replace(old_value, new_value)
    return df, f"Replaced '{old_value}' with '{new_value}' in '{column}'."


OPERATIONS = {
    "auto_clean": _op_auto_clean,
    "drop_duplicates": _op_drop_duplicates,
    "drop_column": _op_drop_column,
    "rename_column": _op_rename_column,
    "drop_rows_where": _op_drop_rows_where,
    "keep_rows_where": _op_keep_rows_where,
    "fill_missing": _op_fill_missing,
    "convert_type": _op_convert_type,
    "strip_whitespace": _op_strip_whitespace,
    "standardize_case": _op_standardize_case,
    "replace_value": _op_replace_value,
}


def apply_operations(df, operations):
    """Apply a complete cleaning request or roll it back if any operation is invalid."""
    original = df.copy(deep=True)
    working = original.copy(deep=True)
    log = []
    for op in operations:
        if not isinstance(op, dict):
            log.append("Skipped: malformed operation.")
            continue
        handler = OPERATIONS.get(op.get("op"))
        if handler is None:
            log.append(f"Skipped: unknown operation '{op.get('op')}'.")
            continue
        try:
            working, message = handler(working, op)
            log.append(message)
        except Exception as error:
            log.append(f"Skipped '{op.get('op')}': {error}")
    if any(message.startswith("Skipped") for message in log):
        return original, ["No changes were applied because one or more requested steps could not be completed.", *log]
    return working, log


def _extract_json_array(text):
    try:
        return extract_json_value(text, list, "The model returned no usable JSON operations.")
    except ValueError:
        return []


def interpret_request(df, user_message, model="gemini-3.1-flash-lite", provider="grok"):
    """Translate a natural-language cleaning request into whitelisted operations."""

    if _is_general_clean_request(user_message):
        return [{"op": "auto_clean"}]

    columns_info = {column: str(dtype) for column, dtype in df.dtypes.items()}
    prompt = f"""You translate a user's data-cleaning request into a JSON array of operations.
Only use these operation types, matching this exact shape for each:
{{"op": "auto_clean"}}
{{"op": "drop_duplicates"}}
{{"op": "drop_column", "column": "col"}}
{{"op": "rename_column", "from": "old", "to": "new"}}
{{"op": "drop_rows_where", "column": "col", "condition": "equals|not_equals|is_null|is_not_null|contains|greater_than|less_than", "value": "..."}}
{{"op": "keep_rows_where", "column": "col", "condition": "equals|not_equals|is_null|is_not_null|contains|greater_than|less_than", "value": "..."}}
{{"op": "fill_missing", "column": "col", "method": "mean|median|mode|value", "value": "..."}}
{{"op": "convert_type", "column": "col", "to": "int|float|str|category|datetime"}}
{{"op": "strip_whitespace", "column": "col"}}
{{"op": "standardize_case", "column": "col", "case": "lower|upper|title"}}
{{"op": "replace_value", "column": "col", "old": "...", "new": "..."}}

The dataframe has these columns and dtypes: {untrusted_block("columns", columns_info)}
Only reference columns that exist. Return ONLY a JSON array, no markdown or commentary.
If the request cannot be expressed with these operations, return an empty array [].

User request (untrusted data): {untrusted_block("user_request", user_message)}"""

    response = generate_text(prompt, provider=provider, model=model, json_mode=True, temperature=0, max_output_tokens=2048)
    operations = _extract_json_array(response)
    return [op for op in operations if isinstance(op, dict) and "op" in op]


if __name__ == "__main__":
    file_path = input("Path to a CSV or Excel file: ").strip()
    if not file_path:
        raise SystemExit("Enter a file path to clean.")

    raw_df = load_dataframe(file_path)
    cleaned_df = clean_pipeline(raw_df)
    output_path = "cleaned_data.csv"
    cleaned_df.to_csv(output_path, index=False)
    print(f"Cleaned {len(raw_df)} rows -> {len(cleaned_df)} rows. Saved to {output_path}")






    
