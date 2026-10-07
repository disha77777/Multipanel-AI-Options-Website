"""Small provider adapter for Gemini and xAI Grok."""

import base64
import io
import json
import os
import ssl
import urllib.error
import urllib.request

import certifi
from google import genai
from google.genai import types

from security import api_key


def configured_providers():
    return {
        "gemini": bool(os.getenv("GEMINI_API_KEY", "").strip()),
        "grok": bool(os.getenv("XAI_API_KEY", "").strip()),
    }


def normalize_provider(provider):
    value = str(provider or "grok").strip().lower()
    if value not in {"gemini", "grok"}:
        raise ValueError("Choose Gemini or Grok as the AI provider.")
    if not configured_providers()[value]:
        name = "GEMINI_API_KEY" if value == "gemini" else "XAI_API_KEY"
        raise RuntimeError(f"{value.title()} is not configured. Add {name} to your local .env file and restart the API.")
    return value


def _image_data_url(image):
    if isinstance(image, bytes):
        raw, mime_type = image, "image/jpeg"
    else:
        buffer = io.BytesIO()
        image.convert("RGB").save(buffer, format="JPEG", quality=92)
        raw, mime_type = buffer.getvalue(), "image/jpeg"
    return f"data:{mime_type};base64,{base64.b64encode(raw).decode('ascii')}"


def generate_text(prompt, *, provider="grok", model="gemini-3.1-flash-lite", json_mode=False,
                  temperature=0.2, max_output_tokens=4096, image=None, web_search=False):
    """Return model text, optionally grounding the answer with provider web search."""
    provider = normalize_provider(provider)
    if provider == "gemini":
        contents = [image, prompt] if image is not None else prompt
        config_values = {"temperature": temperature, "max_output_tokens": max_output_tokens}
        if json_mode:
            config_values["response_mime_type"] = "application/json"
        if web_search:
            config_values["tools"] = [types.Tool(google_search=types.GoogleSearch())]
        config = types.GenerateContentConfig(**config_values)
        # Keep the owning Client alive for the entire request. Calling
        # genai.Client(...).models inline can let the temporary owner be
        # garbage-collected while the Models resource is still sending, which
        # closes the underlying httpx client and raises "client has been closed".
        with genai.Client(api_key=api_key()) as client:
            response = client.models.generate_content(
                model=model, contents=contents, config=config,
            )
            return response.text or ""

    if web_search:
        payload = {
            "model": os.getenv("XAI_MODEL", "grok-4.20-0309-non-reasoning"),
            "input": prompt,
            "tools": [{"type": "web_search"}],
            "max_output_tokens": max_output_tokens,
            "store": False,
        }
        if json_mode:
            payload["text"] = {"format": {"type": "json_object"}}
        endpoint = "/responses"
    else:
        content = [{"type": "text", "text": prompt}]
        if image is not None:
            content.append({"type": "image_url", "image_url": {"url": _image_data_url(image)}})
        payload = {
            "model": os.getenv("XAI_MODEL", "grok-4.20-0309-non-reasoning"),
            "messages": [{"role": "user", "content": content if image is not None else prompt}],
            "temperature": temperature,
            "max_tokens": max_output_tokens,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}
        endpoint = "/chat/completions"
    request = urllib.request.Request(
        os.getenv("XAI_BASE_URL", "https://api.x.ai/v1").rstrip("/") + endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {os.environ['XAI_API_KEY'].strip()}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        tls_context = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(request, timeout=75, context=tls_context) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"Grok API request failed ({error.code}): {detail}") from error
    if web_search:
        return "".join(
            content.get("text", "")
            for item in result.get("output", []) if item.get("type") == "message"
            for content in item.get("content", []) if content.get("type") == "output_text"
        )
    choices = result.get("choices") or []
    message = choices[0].get("message", {}) if choices else {}
    text = message.get("content")
    if isinstance(text, list):
        text = "".join(part.get("text", "") for part in text if isinstance(part, dict))
    return text or ""
