"""Small Google Translate language-detection helper.

The Translate endpoint used by the current Instant Translate add-on returns
the detected source language alongside the translation. Detection is
deliberately best-effort: description generation must still succeed when this
unofficial endpoint is unavailable.
"""

import json
import re
from dataclasses import dataclass
from urllib.parse import urlencode
from urllib.request import Request, urlopen


_GOOGLE_TRANSLATE_URL = "https://translate-pa.googleapis.com/v1/translate"
_GOOGLE_TRANSLATE_API_KEY = "AIzaSyDLEeFI5OtFBwYBIoK_jj5m32rZK5CkCXA"
_MAX_SAMPLE_CHARS = 1500
_MIN_LETTER_COUNT = 8


@dataclass(frozen=True)
class LanguageDetection:
    language: str
    confidence: float | None = None


def normalize_language_code(language):
    """Return the base ISO language code used for comparisons."""
    normalized = (language or "").strip().lower().replace("_", "-")
    aliases = {"iw": "he", "in": "id", "jw": "jv"}
    normalized = aliases.get(normalized, normalized)
    return normalized.split("-", 1)[0]


def languages_match(detected, expected):
    detected = normalize_language_code(detected)
    expected = normalize_language_code(expected)
    return bool(detected and expected and detected == expected)


def _confidence_from_response(payload):
    if isinstance(payload, list):
        # The language metadata is currently at index 4. Its third item is a
        # list containing the confidence for the detected language.
        try:
            confidence = payload[4][2][0]
        except (IndexError, TypeError):
            return None
        return float(confidence) if isinstance(confidence, (int, float)) else None

    # Retain tolerant parsing for recorded responses from older builds.
    confidence = payload.get("confidence")
    if isinstance(confidence, (int, float)):
        return float(confidence)
    values = payload.get("ld_result", {}).get("srclangs_confidences", [])
    if values and isinstance(values[0], (int, float)):
        return float(values[0])
    return None


def _language_from_response(payload):
    if isinstance(payload, list):
        try:
            language = normalize_language_code(payload[5])
        except (IndexError, TypeError):
            language = ""
        if language:
            return language
        # Be defensive if Google omits the top-level detected-language slot.
        for path in ((4, 0, 0), (4, 3, 0)):
            try:
                language = normalize_language_code(payload[path[0]][path[1]][path[2]])
            except (IndexError, TypeError):
                continue
            if language:
                return language
        return ""

    language = normalize_language_code(payload.get("src"))
    if language:
        return language
    source_languages = payload.get("ld_result", {}).get("srclangs", [])
    return normalize_language_code(source_languages[0]) if source_languages else ""


def detect_language(text, target_language="en", timeout=10, opener=urlopen):
    """Detect *text* through Google Translate, or return ``None`` if unsure.

    The translated text is intentionally ignored.  A bounded sample keeps the
    GET URL small. Callers validate each independent Gemini response before
    combining it with other chunks, so isolated wrong-language responses do
    not get hidden by the dominant language of a full-video result.
    Network and response errors are allowed to propagate so the caller can log
    them and continue without blocking generation.
    """
    sample = " ".join((text or "").split())[:_MAX_SAMPLE_CHARS]
    if len(re.findall(r"[^\W\d_]", sample, flags=re.UNICODE)) < _MIN_LETTER_COUNT:
        return None

    target = normalize_language_code(target_language) or "en"
    query = urlencode([
        ("params.client", "gtx"),
        ("query.source_language", "auto"),
        ("query.target_language", target),
        ("query.display_language", target),
        ("query.text", sample),
        ("key", _GOOGLE_TRANSLATE_API_KEY),
        ("data_types", "TRANSLATION"),
        ("data_types", "SENTENCE_SPLITS"),
    ])
    request = Request(
        f"{_GOOGLE_TRANSLATE_URL}?{query}",
        headers={
            "Content-Type": "application/json+protobuf",
            "User-Agent": "Mozilla/5.0",
        },
    )
    with opener(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))

    language = _language_from_response(payload)
    if not language:
        return None
    return LanguageDetection(language, _confidence_from_response(payload))
