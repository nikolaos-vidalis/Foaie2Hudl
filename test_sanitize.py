"""Tests for text normalization and filename slug generation."""

import pytest
from sanitize import filename_slug, remove_diacritics


def test_remove_diacritics_romanian():
    assert remove_diacritics("ț") == "t"
    assert remove_diacritics("Ț") == "T"
    assert remove_diacritics("ţ") == "t"
    assert remove_diacritics("Ţ") == "T"
    assert remove_diacritics("ș") == "s"
    assert remove_diacritics("Ș") == "S"
    assert remove_diacritics("ş") == "s"
    assert remove_diacritics("Ş") == "S"
    assert remove_diacritics("ă") == "a"
    assert remove_diacritics("Ă") == "A"
    assert remove_diacritics("â") == "a"
    assert remove_diacritics("Â") == "A"
    assert remove_diacritics("î") == "i"
    assert remove_diacritics("Î") == "I"


def test_remove_diacritics_mixed_names():
    assert remove_diacritics("Bădăluţă Robert Gabriel") == "Badaluta Robert Gabriel"
    assert remove_diacritics("Oanţă Răzvan Gabriel") == "Oanta Razvan Gabriel"
    assert remove_diacritics("Roşu Cosmin Cristian") == "Rosu Cosmin Cristian"
    assert remove_diacritics("Paraschiv Matei Ștefan") == "Paraschiv Matei Stefan"
    assert remove_diacritics("Țurcanu Ionuț") == "Turcanu Ionut"
    assert remove_diacritics("Ţurcanu Ionuţ") == "Turcanu Ionut"
    assert remove_diacritics("FC Csíkszereda") == "FC Csikszereda"


def test_filename_slug_romanian_names():
    assert filename_slug("Oțelul Galați") == "Otelul_Galati"
    assert filename_slug("CSȘ Craiova") == "CSS_Craiova"
    assert filename_slug("Cupa României") == "Cupa_Romaniei"
    assert filename_slug("Șoimii Pâncota") == "Soimii_Pancota"
    assert filename_slug("Bădăluţă & Oanţă") == "Badaluta_Oanta"


def test_filename_slug_special_characters():
    assert filename_slug('SC "Dinamo 1948" SA') == "SC_Dinamo_1948_SA"
    assert filename_slug("Club / Echipa: Gazde *") == "Club_Echipa_Gazde"
    assert filename_slug("Special #!@$%^&*()[]{} chars") == "Special_chars"
    assert filename_slug("Multiple   spaces   and---hyphens") == "Multiple_spaces_and---hyphens"
    assert filename_slug("___leading_and_trailing___") == "leading_and_trailing"


def test_filename_slug_empty_and_edge_cases():
    assert filename_slug("") == ""
    assert filename_slug("   ") == ""
    assert filename_slug("???!!!") == ""
    assert filename_slug(None) == "None"


def test_teamsheet_filename_romanian_and_special():
    from app import teamsheet_filename

    data = {
        "competition": "Cupa României U17",
        "date_iso": "2026-09-21",
        "home": {"name": "FC Oțelul Galați"},
        "away": {"name": "CSȘ Craiova / LPS"},
    }
    expected = "Cupa_Romaniei_U17_2026-09-21_FC_Otelul_Galati_vs_CSS_Craiova_LPS.docx"
    assert teamsheet_filename(data) == expected


def test_teamsheet_filename_edge_cases():
    from app import teamsheet_filename

    # Missing parts
    assert teamsheet_filename({}) == "Teamsheet.docx"
    assert teamsheet_filename({"competition": "", "date_iso": "2026-09-21"}) == "2026-09-21.docx"
