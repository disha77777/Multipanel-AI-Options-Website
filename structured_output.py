"""Small helpers for decoding structured responses from language models."""

import json
import re


def extract_json_value(text, expected_type=dict, error_message="The model did not return valid JSON."):
    """Decode the first complete JSON value, ignoring surrounding prose or fences."""
    if not isinstance(text, str) or not text.strip():
        raise ValueError(error_message)

    fenced = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.IGNORECASE | re.DOTALL)
    candidates = [fenced.group(1)] if fenced else []
    candidates.append(text)
    opener = "{" if expected_type is dict else "["
    decoder = json.JSONDecoder()

    for candidate in candidates:
        start = candidate.find(opener)
        while start >= 0:
            try:
                value, _ = decoder.raw_decode(candidate[start:])
            except json.JSONDecodeError:
                start = candidate.find(opener, start + 1)
                continue
            if isinstance(value, expected_type):
                return value
            start = candidate.find(opener, start + 1)

    raise ValueError(error_message)
