"""Text normalization and filename slug utilities.

Ensures exported filenames contain no Romanian diacritics, non-ASCII characters,
or unsafe special characters.
"""

import re
import unicodedata

# Transliteration for characters that do not decompose cleanly via NFKD
SPECIAL_CHAR_MAP = {
    "ß": "ss",
    "æ": "ae",
    "Æ": "AE",
    "œ": "oe",
    "Œ": "OE",
    "đ": "d",
    "Đ": "D",
    "ø": "o",
    "Ø": "O",
    "ł": "l",
    "Ł": "L",
}


def remove_diacritics(text: str) -> str:
    """Normalize text by replacing diacritics and accented characters with ASCII equivalents.

    Handles Romanian diacritics:
      - 'ț', 'ţ' -> 't' / 'Ț', 'Ţ' -> 'T'
      - 'ș', 'ş' -> 's' / 'Ș', 'Ş' -> 'S'
      - 'ă' -> 'a' / 'Ă' -> 'A'
      - 'â' -> 'a' / 'Â' -> 'A'
      - 'î' -> 'i' / 'Î' -> 'I'
    as well as other European accents and ligatures.
    """
    if not text:
        return ""
    for char, replacement in SPECIAL_CHAR_MAP.items():
        text = text.replace(char, replacement)
    decomposed = unicodedata.normalize("NFKD", str(text))
    return "".join(c for c in decomposed if unicodedata.category(c) != "Mn")


def filename_slug(value: str) -> str:
    """ASCII filename-safe form of a value without diacritics or special characters.

    1. Normalizes accented/Romanian characters to plain ASCII equivalents (e.g. 'ț' -> 't').
    2. Replaces any non-alphanumeric character (except hyphens) with underscores.
    3. Collapses consecutive underscores and strips leading/trailing underscores.
    """
    text = remove_diacritics(str(value))
    # Replace non-alphanumeric characters (excluding hyphens) with an underscore
    text = re.sub(r"[^a-zA-Z0-9\-]+", "_", text)
    # Collapse multiple consecutive underscores
    text = re.sub(r"_+", "_", text)
    return text.strip("_")

