"""Small Google Translate language-detection helper.

The free Translate endpoint is already used by the companion Sonarpad project.
With ``sl=auto`` and ``dj=1`` its JSON response includes the detected source
language in ``src``.  Detection is deliberately best-effort: description
generation must still succeed when this unofficial endpoint is unavailable.
"""

import json
import re
from dataclasses import dataclass
from urllib.parse import urlencode
from urllib.request import Request, urlopen


_GOOGLE_TRANSLATE_URL = "https://translate.googleapis.com/translate_a/single"
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
    confidence = payload.get("confidence")
    if isinstance(confidence, (int, float)):
        return float(confidence)
    values = payload.get("ld_result", {}).get("srclangs_confidences", [])
    if values and isinstance(values[0], (int, float)):
        return float(values[0])
    return None


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

    query = urlencode({
        "client": "gtx",
        "sl": "auto",
        "tl": normalize_language_code(target_language) or "en",
        "dt": "t",
        "dj": "1",
        "q": sample,
    })
    request = Request(
        f"{_GOOGLE_TRANSLATE_URL}?{query}",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with opener(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))

    language = normalize_language_code(payload.get("src"))
    if not language:
        source_languages = payload.get("ld_result", {}).get("srclangs", [])
        if source_languages:
            language = normalize_language_code(source_languages[0])
    if not language:
        return None
    return LanguageDetection(language, _confidence_from_response(payload))
