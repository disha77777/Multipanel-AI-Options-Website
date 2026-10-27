"""Find openly licensed report charts and attributed Wikimedia images for slides."""

import base64
import html
import ipaddress
import io
import json
import re
import socket
import ssl
import urllib.error
import urllib.parse
import urllib.request

import pypdfium2 as pdfium
from PIL import Image
from pypdf import PdfReader

import certifi
from ai_provider import generate_text
from security import untrusted_block
from structured_output import extract_json_value


COMMONS_API = "https://commons.wikimedia.org/w/api.php"
USER_AGENT = "MultipanelAI/1.0 (presentation image lookup)"
MAX_REPORT_PDF_BYTES = 25_000_000
MAX_REPORT_PAGES = 400
MAX_REPORT_CHARTS = 2
OPEN_LICENSES = {
    "cc by 4.0": "CC BY 4.0",
    "cc-by-4.0": "CC BY 4.0",
    "creative commons attribution 4.0": "CC BY 4.0",
    "cc by-sa 4.0": "CC BY-SA 4.0",
    "cc-by-sa-4.0": "CC BY-SA 4.0",
    "creative commons attribution-sharealike 4.0": "CC BY-SA 4.0",
    "cc0 1.0": "CC0 1.0",
    "cc0-1.0": "CC0 1.0",
    "public domain": "Public domain",
}


class _SafeRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        _validate_public_https_url(new_url)
        return super().redirect_request(request, response, code, message, headers, new_url)


def _validate_public_https_url(url):
    parsed = urllib.parse.urlparse(str(url))
    hostname = parsed.hostname
    if parsed.scheme != "https" or not hostname or parsed.username or parsed.password or parsed.port not in (None, 443):
        raise ValueError("Report sources must use a public HTTPS URL.")
    try:
        addresses = socket.getaddrinfo(hostname, 443, type=socket.SOCK_STREAM)
    except OSError as error:
        raise ValueError("Could not resolve the report source host.") from error
    if not addresses or any(not ipaddress.ip_address(address[4][0]).is_global for address in addresses):
        raise ValueError("Report source host must resolve only to public IP addresses.")
    return parsed


def _open_public_url(url, *, max_bytes):
    _validate_public_https_url(url)
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    opener = urllib.request.build_opener(
        _SafeRedirectHandler,
        urllib.request.HTTPSHandler(context=ssl.create_default_context(cafile=certifi.where())),
    )
    with opener.open(request, timeout=25) as response:
        _validate_public_https_url(response.geturl())
        raw = response.read(max_bytes + 1)
        if not raw or len(raw) > max_bytes:
            raise ValueError("The report exceeds the download size limit.")
        return response.headers.get("Content-Type", "").lower(), raw


def _normalize_open_license(value):
    normalized = " ".join(str(value or "").lower().split())
    for marker, license_name in OPEN_LICENSES.items():
        if marker in normalized:
            return license_name
    return None


def _report_license(reader, reported_license):
    license_name = _normalize_open_license(reported_license)
    if not license_name:
        return None
    pages = list(reader.pages[:3])
    if len(reader.pages) > 3:
        pages.append(reader.pages[-1])
    document_text = " ".join((page.extract_text() or "") for page in pages)
    document_license = _normalize_open_license(document_text)
    return license_name if document_license == license_name else None


def _find_chart_crop(page_image, *, chart_title, provider):
    prompt = f"""Locate the chart titled or described as {untrusted_block("chart_title", chart_title)} on this report page.
Return JSON only with this shape:
{{"found": true, "x0": 0, "y0": 0, "x1": 1000, "y1": 1000}}
Coordinates are normalized from 0 to 1000, origin top-left. Include the whole chart, title, axes, legend, and labels, but exclude surrounding report text where possible. If the requested chart is not clearly present, return {{"found": false}}. Do not guess or invent a chart."""
    response = generate_text(
        prompt,
        provider=provider,
        image=page_image,
        json_mode=True,
        temperature=0,
        max_output_tokens=200,
    )
    result = extract_json_value(response, dict, "Could not locate the chart on the report page.")
    if result.get("found") is not True:
        return None
    try:
        bounds = [int(result[name]) for name in ("x0", "y0", "x1", "y1")]
    except (KeyError, TypeError, ValueError) as error:
        raise ValueError("The chart crop coordinates were invalid.") from error
    x0, y0, x1, y1 = bounds
    if not (0 <= x0 < x1 <= 1000 and 0 <= y0 < y1 <= 1000):
        raise ValueError("The chart crop coordinates were outside the report page.")
    if x1 - x0 < 200 or y1 - y0 < 120:
        raise ValueError("The detected chart crop is too small to read.")
    width, height = page_image.size
    crop = page_image.crop((x0 * width // 1000, y0 * height // 1000, x1 * width // 1000, y1 * height // 1000))
    output = io.BytesIO()
    crop.convert("RGB").save(output, format="JPEG", quality=88, optimize=True)
    image_bytes = output.getvalue()
    if len(image_bytes) > 4_000_000:
        raise ValueError("The report chart image exceeds the 4 MB limit.")
    return base64.b64encode(image_bytes).decode("ascii")


def _extract_report_chart(candidate, provider):
    try:
        pdf_url = candidate["pdf_url"]
        report_url = candidate["report_url"]
        page_number = int(candidate["page_number"])
        if not 1 <= page_number <= MAX_REPORT_PAGES:
            raise ValueError("The report page number is outside the supported range.")
        _validate_public_https_url(report_url)
        content_type, pdf_bytes = _open_public_url(pdf_url, max_bytes=MAX_REPORT_PDF_BYTES)
        if not pdf_bytes.startswith(b"%PDF-") or ("pdf" not in content_type and content_type not in {"", "application/octet-stream"}):
            raise ValueError("The report source did not return a PDF.")
        reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
        if len(reader.pages) > MAX_REPORT_PAGES or page_number > len(reader.pages):
            raise ValueError("The report page number is not present in the PDF.")
        license_name = _report_license(reader, candidate["license"])
        if not license_name:
            raise ValueError("The report does not verify the stated open license in its document text.")
        pdf_document = pdfium.PdfDocument(pdf_bytes)
        page = pdf_document[page_number - 1]
        page_width, page_height = page.get_size()
        render_scale = min(1.6, (25_000_000 / (page_width * page_height)) ** 0.5)
        rendered = page.render(scale=render_scale).to_pil().convert("RGB")
        image_data = _find_chart_crop(
            rendered,
            chart_title=candidate["chart_title"],
            provider=provider,
        )
        if not image_data:
            raise ValueError("The requested chart was not clearly present on that report page.")
        return {
            "kind": "report_chart",
            "data": image_data,
            "chart_title": candidate["chart_title"][:180],
            "report_title": candidate["report_title"][:240],
            "publisher": candidate["publisher"][:120],
            "page_number": page_number,
            "source_url": report_url,
            "license": license_name,
            "license_url": candidate.get("license_url", ""),
        }
    except (KeyError, TypeError, ValueError, OSError, urllib.error.URLError, TimeoutError) as error:
        return {"error": str(error)}


def find_report_charts(topic, slides, provider="gemini"):
    slide_summaries = [
        {"index": index, "title": slide.get("title", ""), "bullets": slide.get("bullets", [])[:4]}
        for index, slide in enumerate(slides)
    ]
    prompt = f"""Use live web search to find up to {MAX_REPORT_CHARTS} actual, relevant charts in reports that support the presentation below.
Topic: {untrusted_block("topic", topic)}
Slides: {untrusted_block("slides", json.dumps(slide_summaries))}

Only select freely accessible PDF reports whose document explicitly states one of these reuse terms: CC BY 4.0, CC BY-SA 4.0, CC0 1.0, or Public domain. Do not use paywalled or all-rights-reserved reports. A chart must directly support a listed slide. Return the exact PDF URL, its report landing-page URL, page number, chart title, report title, publisher, stated license, and license URL when available. Do not invent URLs, page numbers, titles, or licenses. If no clearly reusable chart is found, return an empty charts list.

Return JSON only:
{{"charts":[{{"slide_index":0,"chart_title":"...","report_title":"...","publisher":"...","report_url":"https://...","pdf_url":"https://...pdf","page_number":1,"license":"CC BY 4.0","license_url":"https://creativecommons.org/licenses/by/4.0/"}}]}}"""
    response = generate_text(
        prompt,
        provider=provider,
        web_search=True,
        json_mode=True,
        temperature=0.1,
        max_output_tokens=1400,
    )
    result = extract_json_value(response, dict, "Could not find chart sources for this presentation.")
    candidates = result.get("charts", [])
    if not isinstance(candidates, list):
        raise ValueError("The report chart search returned an invalid result.")
    charts = []
    messages = []
    for candidate in candidates[:MAX_REPORT_CHARTS]:
        if not isinstance(candidate, dict):
            messages.append("A report result was invalid and was skipped.")
            continue
        try:
            index = int(candidate["slide_index"])
            if not 0 <= index < len(slides):
                raise ValueError("The report chart did not map to a presentation slide.")
            for name in ("chart_title", "report_title", "publisher", "report_url", "pdf_url", "license"):
                if not isinstance(candidate.get(name), str) or not candidate[name].strip():
                    raise ValueError(f"The report result is missing {name.replace('_', ' ')}.")
            if candidate.get("license_url"):
                _validate_public_https_url(candidate["license_url"])
            chart = _extract_report_chart(candidate, provider)
            if "error" in chart:
                messages.append(chart["error"])
                continue
            chart["slide_index"] = index
            slides[index]["image"] = chart
            charts.append(chart)
        except (KeyError, TypeError, ValueError) as error:
            messages.append(str(error))
    return charts, messages


def _plain(value):
    value = re.sub(r"<[^>]+>", " ", str(value or ""))
    return " ".join(html.unescape(value).split())


def _get_json(url):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=20, context=ssl.create_default_context(cafile=certifi.where())) as response:
        if response.status != 200:
            raise RuntimeError("Wikimedia Commons image lookup failed.")
        return json.loads(response.read(2_000_000).decode("utf-8"))


def find_commons_image(query):
    query = " ".join(str(query or "").split())[:160]
    if not query:
        return None
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": "6",
        "gsrlimit": "6",
        "prop": "imageinfo",
        "iiprop": "url|mime|extmetadata",
        "iiurlwidth": "1400",
        "iiextmetadatafilter": "Artist|Credit|LicenseShortName|LicenseUrl|UsageTerms",
        "iiextmetadatalanguage": "en",
    }
    try:
        data = _get_json(COMMONS_API + "?" + urllib.parse.urlencode(params))
    except (OSError, ValueError, urllib.error.URLError, TimeoutError):
        return None
    pages = (data.get("query") or {}).get("pages") or {}
    for page in pages.values():
        image_infos = page.get("imageinfo") or []
        if not image_infos:
            continue
        info = image_infos[0]
        if info.get("mime") not in {"image/jpeg", "image/png"}:
            continue
        metadata = info.get("extmetadata") or {}
        license_name = _plain((metadata.get("LicenseShortName") or {}).get("value"))
        if not license_name:
            continue
        title = page.get("title", "")
        file_name = title.removeprefix("File:")
        file_page = "https://commons.wikimedia.org/wiki/" + urllib.parse.quote(title.replace(" ", "_"), safe=":()'!,")
        artist = _plain((metadata.get("Artist") or {}).get("value")) or _plain((metadata.get("Credit") or {}).get("value")) or "Wikimedia Commons contributor"
        return {
            "url": info.get("thumburl") or info.get("url"),
            "file_page_url": file_page,
            "file_name": file_name,
            "artist": artist[:160],
            "license": license_name[:100],
            "license_url": _plain((metadata.get("LicenseUrl") or {}).get("value"))[:300],
        }
    return None
