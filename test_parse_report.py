"""Tests for reading and applying the "Schimbare tricou de joc" section.

Synthetic lines and squads only: real reports carry minors' personal data and are not
kept in the repository.
"""

from fill_teamsheet import player_display_name
from parse_report import SHIRT_CHANGE_ROW, _apply_shirt_changes, _shirt_changes

REPORT_LINES = [
    "Observator: -",
    "Delegate: -",
    "Schimbare tricou de joc",
    "CS Dinamo Bucureşti",
    "Min. Jucator Nr. nou Motiv",
    "0' 6 - Petrache Nicolas Cristian 20 Lipsa tricou",
    "0' 60 - Volos  Ivan 19 Lipsa tricou",
    "0' 61 - Tcaci Ioan Marian 16 Lipsa Tricou",
    "0' 16 - Jataruc Rareș Andrei 28 Lipsa Tricou",
    "0' 19 - Cârstea Luca Andrei 77 Lipsa Tricou",
    "Inlocuiri medicale",
    "Titular Rezerva Aditional",
    "Nu au fost inregistrate inlocuiri medicale de jucatori.",
    "Inlocuiri jucatori",
    "CS Dinamo Bucureşti",
    "Min. Jucator iesit Jucator intrat",
    "56' 10 - Bodron Laurențiu Ștefan 47 - Doliș Luca Gabriel",
]


def player(number, name, gk=False, captain=False):
    return {"number": number, "name": name, "gk": gk, "captain": captain}


def squads():
    return {
        "home": {
            "name": "SC Dinamo 1948",
            "starters": [player(6, "Vărzan Tudor Mihai"), player(8, "Dateș Yanis Cristian")],
            "subs": [player(19, "Mititelu Erik Alexandru")],
        },
        "away": {
            "name": "CS Dinamo Bucureşti",
            "starters": [
                player(6, "Damaschin Cătălin Andrei"),
                player(6, "Petrache Nicolas Cristian"),
                player(60, "Volos Ivan"),
                player(61, "Tcaci Ioan Marian"),
                player(17, "Popescu Mihai Alexandru"),
            ],
            "subs": [player(16, "Jataruc Rareş Andrei"), player(19, "Cârstea Luca Andrei")],
        },
    }


def numbers(squad):
    return {p["name"]: p["number"] for p in squad["starters"] + squad["subs"]}


def test_shirt_changes_reads_section_only():
    changes = _shirt_changes(REPORT_LINES)
    assert [(c["old"], c["name"], c["new"]) for c in changes] == [
        (6, "Petrache Nicolas Cristian", 20),
        (60, "Volos Ivan", 19),
        (61, "Tcaci Ioan Marian", 16),
        (16, "Jataruc Rareș Andrei", 28),
        (19, "Cârstea Luca Andrei", 77),
    ]
    assert {c["team"] for c in changes} == {"CS Dinamo Bucureşti"}
    assert {c["minute"] for c in changes} == {"0'"}


def test_shirt_changes_empty_or_missing_section():
    assert _shirt_changes(["Observator: -", "Inlocuiri medicale"]) == []
    assert _shirt_changes(
        ["Schimbare tricou de joc", "Nu exista inregistrari.", "Inlocuiri medicale"]
    ) == []


def test_shirt_changes_across_page_break_and_both_teams():
    lines = [
        "Schimbare tricou de joc",
        "SC Dinamo 1948",
        "Min. Jucator Nr. nou Motiv",
        "0' 6 - Vărzan Tudor Mihai 26 Lipsa tricou",
        "Pagina 1 din10 Raport de arbitru generat cu Football Connect © 11.10.2026 - 16:42",
        "CS Dinamo Bucureşti",
        "Min. Jucator Nr. nou Motiv",
        "La pauza 17 - Popescu Mihai Alexandru 21 Tricou insangerat",
        "Inlocuiri medicale",
    ]
    changes = _shirt_changes(lines)
    assert [(c["team"], c["minute"], c["old"], c["new"]) for c in changes] == [
        ("SC Dinamo 1948", "0'", 6, 26),
        ("CS Dinamo Bucureşti", "La pauza", 17, 21),
    ]


def test_substitution_row_is_not_a_shirt_change():
    assert not SHIRT_CHANGE_ROW.match("56' 10 - Bodron Laurențiu Ștefan 47 - Doliș Luca Gabriel")
    assert not SHIRT_CHANGE_ROW.match("72' 8 - Mustaţă Ovidiu Nicolae 28 - Jataruc Rareș Andrei (16)")


def test_apply_matches_name_and_number_not_number_alone():
    data = squads()
    issues = _apply_shirt_changes(data, _shirt_changes(REPORT_LINES))
    assert issues == []
    away = numbers(data["away"])
    assert away["Petrache Nicolas Cristian"] == 20
    assert away["Damaschin Cătălin Andrei"] == 6
    assert away["Volos Ivan"] == 19          # double space in the report
    assert away["Tcaci Ioan Marian"] == 16
    assert away["Jataruc Rareş Andrei"] == 28  # ș in the report, ş in the squad
    assert away["Cârstea Luca Andrei"] == 77
    assert numbers(data["home"]) == numbers(squads()["home"])


def test_apply_mid_match_change_keeps_number_and_adds_note():
    data = squads()
    changes = [
        {"team": "CS Dinamo Bucureşti", "minute": "60'", "old": 17,
         "name": "Popescu Mihai Alexandru", "new": 21},
        {"team": "CS Dinamo Bucureşti", "minute": "La pauza", "old": 60,
         "name": "Volos Ivan", "new": 19},
    ]
    assert _apply_shirt_changes(data, changes) == []
    popescu = next(p for p in data["away"]["starters"] if p["name"].startswith("Popescu"))
    volos = next(p for p in data["away"]["starters"] if p["name"].startswith("Volos"))
    assert popescu["number"] == 17
    assert popescu["shirt_change"] == {"number": 21, "minute": "60'"}
    assert volos["shirt_change"] == {"number": 19, "minute": "HT"}


def test_apply_unknown_label_searches_both_squads():
    data = squads()
    changes = [{"team": "", "minute": "0'", "old": 8, "name": "Dateș Yanis Cristian", "new": 18}]
    assert _apply_shirt_changes(data, changes) == []
    assert numbers(data["home"])["Dateș Yanis Cristian"] == 18


def test_apply_unmatched_change_is_reported_not_guessed():
    data = squads()
    changes = [{"team": "CS Dinamo Bucureşti", "minute": "0'", "old": 6,
                "name": "Necunoscut Ion", "new": 20}]
    issues = _apply_shirt_changes(data, changes)
    assert issues == [{"team": "CS Dinamo Bucureşti", "name": "Necunoscut Ion", "old": 6, "new": 20}]
    assert numbers(data["away"]) == numbers(squads()["away"])


def test_display_name_with_shirt_change_note():
    changed = {"number": 21, "minute": "60'"}
    assert player_display_name({**player(17, "Popescu Mihai"), "shirt_change": changed}) == (
        "Popescu Mihai (#21 from 60')"
    )
    assert player_display_name({**player(1, "Florea Matei", gk=True), "shirt_change": changed}) == (
        "Florea Matei (GK) (#21 from 60')"
    )
    assert player_display_name(player(5, "Dobre Ioan", captain=True)) == "Dobre Ioan (C)"
