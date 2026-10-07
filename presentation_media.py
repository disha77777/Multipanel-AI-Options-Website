"""Licensed, attributed image lookup from Wikimedia Commons for presentation slides."""

import html
import json
import re
import ssl
import urllib.error
import urllib.parse
import urllib.request

import certifi


COMMONS_API = "https://commons.wikimedia.org/w/api.php"
USER_AGENT = "MultipanelAI/1.0 (presentation image lookup)"


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
