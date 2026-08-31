"""Foaie2Hudl -- upload FRF referee reports, download filled Hudl teamsheets.

Bilingual (Romanian by default), wide three-panel layout: upload on the left, the
branding graphic in the middle, generated teamsheets on the right. Several reports
can be uploaded at once; each one converts into its own teamsheet.

Nothing is written to disk: PDFs are parsed and .docx files built in memory.
"""

import io
import re
import zipfile
from pathlib import Path

import streamlit as st

from fill_teamsheet import fill
from parse_report import parse

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "Wyscout teamsheet template.docx"
IMAGE = ROOT / "assets" / "frf-to-hudl.png"

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_UPLOAD_MB = 5
SOURCE_URL = "https://www.footballconnect.ro"

# Characters Windows forbids in filenames.
ILLEGAL_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

TEXTS = {
    "ro": {
        "tagline": "Din raportul arbitrului direct în foaia de echipă Hudl.",
        "upload_header": "1 · Rapoarte de arbitraj",
        "upload_intro": (
            "Încarcă **raportul complet de arbitraj** (`Raport Arbitru`, `.pdf`) descărcat din "
            f"[www.footballconnect.ro]({SOURCE_URL}) — cel care conține foile de meci ale "
            "**ambelor** echipe, nu foaia unei singure echipe."
        ),
        "uploader_label": "Rapoarte de arbitraj (PDF)",
        "uploader_help": (
            "Poți încărca mai multe rapoarte simultan; fiecare este convertit separat. "
            "Este necesar raportul complet de arbitraj din www.footballconnect.ro, care "
            "include ambele echipe — foaia de meci a unei singure echipe nu este suficientă."
        ),
        "download_header": "2 · Foi de echipă",
        "empty_state": "Foile de echipă apar aici după încărcarea rapoartelor.",
        "download_one": "Descarcă .docx",
        "download_all": "Descarcă toate ({count}) ca .zip",
        "squad_summary": "{starters} titulari · {subs} rezerve",
        "processing": "Se procesează rapoartele…",
        "language_label": "Limbă / Language",
        "error_prefix": "Nu am putut procesa **{name}**",
        "missing_sides": (
            "Raportul nu conține foile de meci ale ambelor echipe (lipsește: {sides}). "
            "Încarcă raportul complet de arbitraj din www.footballconnect.ro, "
            "nu foaia de meci a unei singure echipe."
        ),
        "no_date": (
            "Nu am găsit data meciului în document. "
            "Verifică dacă este raportul de arbitraj din www.footballconnect.ro."
        ),
        "empty_pdf": "PDF-ul este gol.",
        "unreadable_pdf": (
            "Fișierul nu a putut fi citit ca PDF valid. Verifică dacă este raportul de "
            "arbitraj descărcat din www.footballconnect.ro."
        ),
        "privacy_title": "Datele tale nu sunt stocate.",
        "privacy_body": (
            "Rapoartele încărcate sunt procesate temporar, în memorie, doar pe durata "
            "sesiunii curente. Nimic nu este salvat pe server, iar foile de echipă sunt "
            "generate în memorie și trimise direct în browserul tău."
        ),
        "side_home": "gazde",
        "side_away": "oaspeți",
    },
    "en": {
        "tagline": "From the referee report straight to the Hudl team sheet.",
        "upload_header": "1 · Referee reports",
        "upload_intro": (
            "Upload the **complete referee report** (`Raport Arbitru`, `.pdf`) downloaded from "
            f"[www.footballconnect.ro]({SOURCE_URL}) — the one containing the team sheets of "
            "**both** clubs, not a single team's own sheet."
        ),
        "uploader_label": "Referee reports (PDF)",
        "uploader_help": (
            "You can upload several reports at once; each is converted separately. "
            "The complete referee report from www.footballconnect.ro is required, as it "
            "includes both teams -- a single team's own sheet is not enough."
        ),
        "download_header": "2 · Team sheets",
        "empty_state": "Generated team sheets will appear here once you upload reports.",
        "download_one": "Download .docx",
        "download_all": "Download all ({count}) as .zip",
        "squad_summary": "{starters} starters · {subs} substitutes",
        "processing": "Processing reports…",
        "language_label": "Limbă / Language",
        "error_prefix": "Could not process **{name}**",
        "missing_sides": (
            "The report does not contain the team sheets of both clubs (missing: {sides}). "
            "Upload the complete referee report from www.footballconnect.ro, "
            "not a single team's own sheet."
        ),
        "no_date": (
            "Could not find the match date in the document. "
            "Check that this is the referee report from www.footballconnect.ro."
        ),
        "empty_pdf": "The PDF is empty.",
        "unreadable_pdf": (
            "The file could not be read as a valid PDF. Check that it is the referee "
            "report downloaded from www.footballconnect.ro."
        ),
        "privacy_title": "Your data is not stored.",
        "privacy_body": (
            "Uploaded reports are processed temporarily, in memory, for the current "
            "session only. Nothing is saved on the server, and the team sheets are built "
            "in memory and sent straight to your browser."
        ),
        "side_home": "home",
        "side_away": "away",
    },
}

LANGUAGE_NAMES = {"ro": "Română", "en": "English"}


@st.cache_data(show_spinner=False)
def convert(pdf_bytes):
    """Parse one report and build its teamsheet. Cached on the file's bytes.

    Any download click reruns the script, so without this every uploaded PDF would be
    re-parsed on each click.

    This is the boundary between the localized UI and the language-neutral conversion:
    only bytes go in, and the generated document keeps the template's own labels. The
    language setting affects the website only -- never pass `texts` past this point.
    """
    data = parse(io.BytesIO(pdf_bytes))
    return data, fill(data, TEMPLATE)


def slug(value):
    """Filename-safe form of a value, keeping diacritics."""
    cleaned = ILLEGAL_FILENAME_CHARS.sub("", str(value)).strip()
    return re.sub(r"\s+", "_", cleaned).strip("_")


def teamsheet_filename(data):
    """Competition, date and both teams, e.g.

    Liga_Elitelor_U17_2026-08-30_SC_Dinamo_1948_vs_FC_Voluntari.docx
    """
    parts = [
        slug(data["competition"]),
        data["date_iso"],
        f"{slug(data['home']['name'])}_vs_{slug(data['away']['name'])}",
    ]
    return "_".join(part for part in parts if part) + ".docx"


def unique_name(name, taken):
    """Suffix a duplicate zip entry name with _2, _3, ..."""
    if name not in taken:
        return name
    stem, _, suffix = name.rpartition(".")
    index = 2
    while f"{stem}_{index}.{suffix}" in taken:
        index += 1
    return f"{stem}_{index}.{suffix}"


def error_message(error, texts, filename):
    """Localized message for a failed report, falling back to the raw text."""
    code = getattr(error, "code", None)
    if code == "missing_sides":
        sides = ", ".join(texts[f"side_{side}"] for side in getattr(error, "sides", []))
        detail = texts["missing_sides"].format(sides=sides)
    elif code in texts:
        detail = texts[code]
    else:
        # Not a ParseError -- most likely a corrupt file. Keep the library's own text as
        # a parenthetical so the cause is still visible.
        detail = f"{texts['unreadable_pdf']} ({error})"
    return f"{texts['error_prefix'].format(name=filename)} — {detail}"


st.set_page_config(page_title="Foaie2Hudl", page_icon="⚽", layout="wide")

language = st.session_state.get("language", "ro")
texts = TEXTS[language]

left, middle, right = st.columns(3, gap="large")

# --- Left: upload -------------------------------------------------------------------
with left.container(border=True, height="stretch"):
    st.subheader(texts["upload_header"])
    st.caption(texts["upload_intro"])
    uploads = st.file_uploader(
        texts["uploader_label"],
        type="pdf",
        accept_multiple_files=True,
        max_upload_size=MAX_UPLOAD_MB,
        help=texts["uploader_help"],
    )
    st.info(
        f"**{texts['privacy_title']}** {texts['privacy_body']}",
        icon=":material/lock:",
    )

# --- Middle: branding and language --------------------------------------------------
with middle.container(border=True, height="stretch"):
    with st.container(horizontal_alignment="center"):
        st.image(str(IMAGE), width="stretch")
        st.title("Foaie2Hudl")
        st.caption(texts["tagline"])
        st.space("small")
        st.segmented_control(
            texts["language_label"],
            list(LANGUAGE_NAMES),
            default=language,
            required=True,
            format_func=LANGUAGE_NAMES.get,
            key="language",
        )

# --- Right: generated teamsheets ----------------------------------------------------
with right.container(border=True, height="stretch"):
    st.subheader(texts["download_header"])

    if not uploads:
        st.info(texts["empty_state"])
    else:
        results, failures = [], []
        with st.spinner(texts["processing"]):
            for upload in uploads:
                try:
                    data, document = convert(upload.getvalue())
                except Exception as error:  # one bad report must not stop the others
                    failures.append(error_message(error, texts, upload.name))
                else:
                    results.append((data, document))

        if len(results) > 1:
            archive = io.BytesIO()
            names = set()
            with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
                for data, document in results:
                    name = unique_name(teamsheet_filename(data), names)
                    names.add(name)
                    bundle.writestr(name, document)

            dates = {data["date_iso"] for data, _ in results}
            zip_name = f"Teamsheets_{dates.pop()}.zip" if len(dates) == 1 else "Teamsheets.zip"
            st.download_button(
                texts["download_all"].format(count=len(results)),
                archive.getvalue(),
                file_name=zip_name,
                mime="application/zip",
                type="primary",
                width="stretch",
                key="download_all",
            )

        for index, (data, document) in enumerate(results):
            home, away = data["home"], data["away"]
            with st.container(border=True):
                st.markdown(f"**{home['name']} {data['score']} {away['name']}**")
                st.caption(f"{data['competition']} · {data['date']}")
                st.caption(
                    f"{home['name']}: "
                    + texts["squad_summary"].format(
                        starters=len(home["starters"]), subs=len(home["subs"])
                    )
                    + f"  \n{away['name']}: "
                    + texts["squad_summary"].format(
                        starters=len(away["starters"]), subs=len(away["subs"])
                    )
                )
                st.download_button(
                    texts["download_one"],
                    document,
                    file_name=teamsheet_filename(data),
                    mime=DOCX_MIME,
                    width="stretch",
                    key=f"download_{index}",
                )

        for message in failures:
            st.error(message)
