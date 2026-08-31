"""Foaie2Hudl -- upload an FRF referee report, download a filled Hudl teamsheet.

Nothing is written to disk: the PDF is parsed and the .docx built in memory.
"""

from pathlib import Path

import streamlit as st

from fill_teamsheet import fill
from parse_report import parse

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "Wyscout teamsheet template.docx"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

st.set_page_config(page_title="Foaie2Hudl", page_icon="⚽")

st.image(str(ROOT / "assets" / "frf-to-hudl.png"))
st.title("Foaie2Hudl")
st.caption("Încarcă raportul de arbitraj (.pdf) și descarcă foaia de echipă Hudl completată.")

uploaded = st.file_uploader("Raport de arbitraj (PDF)", type="pdf")

if uploaded:
    try:
        data = parse(uploaded)
        document = fill(data, TEMPLATE)
    except Exception as error:
        st.error(f"Nu am putut procesa raportul: {error}")
    else:
        home, away = data["home"], data["away"]
        st.success(
            f"**{home['name']} {data['score']} {away['name']}** — "
            f"{data['competition']}, {data['date']}"
        )
        st.caption(
            f"{home['name']}: {len(home['starters'])} titulari, {len(home['subs'])} rezerve · "
            f"{away['name']}: {len(away['starters'])} titulari, {len(away['subs'])} rezerve"
        )

        slug = f"{home['name']}_vs_{away['name']}".replace(" ", "_")
        st.download_button(
            "Descarcă foaia de echipă (.docx)",
            document,
            file_name=f"Teamsheet_{slug}_{data['date_iso']}.docx",
            mime=DOCX_MIME,
            type="primary",
        )
