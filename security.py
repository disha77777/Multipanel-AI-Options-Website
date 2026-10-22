"""Small, dependency-free security helpers shared by the app and API."""

import os
import re
from xml.sax.saxutils import escape


DEFAULT_MAX_INPUT_LENGTH = 4_000


def env_flag(name, default=False):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def max_input_length(name, default=DEFAULT_MAX_INPUT_LENGTH):
    try:
        return max(100, min(int(os.getenv(name, default)), 20_000))
    except (TypeError, ValueError):
        return default


def bounded_int(name, default, minimum, maximum):
    try:
        return max(minimum, min(int(os.getenv(name, default)), maximum))
    except (TypeError, ValueError):
        return default


def clean_text(value, limit):
    """Normalize user text and enforce a hard size limit before model use."""
    value = " ".join(str(value or "").split())
    if len(value) > limit:
        raise ValueError(f"Input must be {limit} characters or fewer.")
    return value


def untrusted_block(label, value):
    """Place user-controlled content outside the instruction channel."""
    safe_label = re.sub(r"[^a-zA-Z0-9_]", "_", str(label))
    return f"\n<untrusted_{safe_label}>\n{escape(str(value or ''))}\n</untrusted_{safe_label}>\n"


def allowed_origins():
    return {
        origin.strip().rstrip("/")
        for origin in os.getenv("ALLOWED_ORIGINS", "").split(",")
        if origin.strip()
    }


def origin_allowed(origin):
    origins = allowed_origins()
    return not origin or origin.rstrip("/") in origins


def safe_key_present():
    return bool(os.getenv("GEMINI_API_KEY"))


def api_key():
    """Read the single supported provider credential from the environment."""
    value = os.getenv("GEMINI_API_KEY", "").strip()
    if not value:
        raise RuntimeError("Set GEMINI_API_KEY in the deployment secrets before using AI features.")
    return value


def safe_identifier(value):
    return re.sub(r"[^a-zA-Z0-9_.:-]", "_", str(value))[:120]
