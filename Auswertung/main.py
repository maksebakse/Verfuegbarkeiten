# app.py
import streamlit as st
import pandas as pd
import os, base64
from pathlib import Path

from funktionen_app import (
    BUS_TO_HERSTELLER,
    setup_page,
    get_data,
    sidebar_filters,
    filter_and_add_km,
    page_analyse,
    page_statistik,
    page_km_betrachtung,
    page_kategorien,
    page_uebersicht,
    page_monatliche_auswertungen,
    prepare_filtered_summary
)

# ─────────────────────────────────────────────────────────────────────────────
# Pfade / Dateinamen
RAW_SUMMARY    = "Zusammenfassung.xlsx"
DATE_FILE      = "Zulassung-Verkauf.xlsx"
PROCESSED_XLS  = "Zusammenfassung_bearbeitet.xlsx"
PROCESSED_PQ   = "Zusammenfassung_bearbeitet.parquet"
BUS_MAP_FILE   = "bus_hersteller_zuordnung.xlsx"


# ── 1) Einmaliges Laden + Parquet-Dump ───────────────────────────────────────
@st.cache_resource(show_spinner=False)
def load_all_data() -> tuple[dict[int,str], pd.DataFrame]:
    """
    Lädt Bus→Hersteller, erzeugt ggf. die gefilterte XLSX,
    liest sie ein (load_data via get_data) und schreibt einen Parquet-Dump.
    """
    # 1. Bus→Hersteller
    df_map = pd.read_excel(BUS_MAP_FILE, engine="openpyxl")
    df_map.columns = ["BusNr","Hersteller"]
    bus_to_hersteller = df_map.set_index("BusNr")["Hersteller"].to_dict()

    # 2. Gefilterte XLSX einmalig erzeugen
    if not Path(PROCESSED_XLS).exists():
        prepare_filtered_summary(
            summary_path=RAW_SUMMARY,
            date_path=DATE_FILE,
            output_path=PROCESSED_XLS,
            sheet_dates=0
        )

    # 3. Parquet-Cache prüfen
    if Path(PROCESSED_PQ).exists():
        df = pd.read_parquet(PROCESSED_PQ)
    else:
        # hier benutzen wir get_data → ruft intern load_data() auf
        df = get_data(PROCESSED_XLS)
        # Parquet-Dump fürs nächste Mal
        df.to_parquet(PROCESSED_PQ, index=False)
    
    df_dates = pd.read_excel(DATE_FILE, engine="openpyxl", sheet_name=0)
    df_dates = df_dates.rename(columns={
        "KOM-Nr.":       "BusNr",
        "Einsatz":       "ZulassungDatum",
        "Verkauf":       "VerkaufDatum"
    })
    # BusNr und Datumsfelder auf sauberen Typ bringen
    df_dates["BusNr"] = df_dates["BusNr"].astype(str).str.strip()
    df_dates["ZulassungDatum"] = pd.to_datetime(df_dates["ZulassungDatum"], errors="coerce")
    df_dates["VerkaufDatum"]   = pd.to_datetime(df_dates["VerkaufDatum"],   errors="coerce")

    # Merge mit unseren bereits geladenen Daten
    # (BusNr in beiden auf string bringen!)
    df["BusNr"] = df["BusNr"].astype(str).str.strip()
    df = df.merge(
        df_dates[["BusNr", "ZulassungDatum", "VerkaufDatum"]],
        on="BusNr", how="left"
    )

    # Maske: Datum zwischen Zulassung und Verkauf (oder unendlich, falls kein VerkaufDatum)
    mask = (
        (df["Datum"] >= df["ZulassungDatum"]) &
        (
            df["VerkaufDatum"].isna() |
            (df["Datum"] <= df["VerkaufDatum"])
        )
    ).fillna(False)

    # Auf diese Maske einschränken
    df = df.loc[mask].copy()

    # Wir brauchen die beiden Hilfsspalten nicht mehr
    df = df.drop(columns=["ZulassungDatum", "VerkaufDatum"])
    # ────────────────────────────────────────────────────────────────────────

    return bus_to_hersteller, df


# ── 2) Filter & KM-Berechnung cachen ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def filter_and_add_km_cached(df: pd.DataFrame, filt: dict):
    """
    Wrapper um Deine filter_and_add_km-Funktion, damit sie bei
    gleichen Filtern nicht erneut rechnet.
    """
    return filter_and_add_km(df, filt)


# ── 3) DVD-Logo Helfer (optional) ────────────────────────────────────────────
def inject_dvd_css(width: int = 150, duration: int = 12):
    css = f"""
    <style>
      @keyframes dvd {{ 0% {{top:0;left:0;}} 25% {{top:0;left:calc(100vw-{width}px);}}
                       50% {{top:calc(100vh-{width}px);left:calc(100vw-{width}px);}}
                       75% {{top:calc(100vh-{width}px);left:0;}} 100% {{top:0;left:0;}} }}
      .dvd-logo {{ position: fixed; width:{width}px;
                  animation:dvd {duration}s linear infinite; z-index:9999; }}
    </style>"""
    st.markdown(css, unsafe_allow_html=True)

def load_base64_gif(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

def show_dvd_logo(path: str = "dvd.gif", width: int = 150, duration: int = 12):
    inject_dvd_css(width, duration)
    b64 = load_base64_gif(path)
    st.markdown(f'<img src="data:image/gif;base64,{b64}" class="dvd-logo"/>', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Hauptprogramm
def main():
    setup_page(title="Ausfall‐Analyse Busflotte", layout="wide")

    # -------------------------------------------------------------------------
    # Spinner + Progress bei erstem Laden
    # -------------------------------------------------------------------------
    with st.spinner("✨ Initialisiere und lade Daten, einen Moment bitte…"):
        progress = st.sidebar.progress(0)
        bus_to_hersteller, df = load_all_data()
        progress.progress(50)
        # falls Du später noch weitere DataFrames hast → progress.progress(…)
        progress.progress(100)

    # leere Zeilen entfernen
    df = df.dropna(subset=["Ausfall-Typ","Ausfallgrund","Serie","BusNr"], how="all")

    # Sidebar-Filter
    filt = sidebar_filters(df)

    # Filter + KM berechnen (gecached)
    df_filt, df_km = filter_and_add_km_cached(df, filt)

    # Navigation
    st.sidebar.markdown("## 📑 Navigation")
    page = st.sidebar.radio("Seite wählen:",
        ["Analyse","Statistik","KM-Betrachtung","Übersicht","Kategorien","Monatliche Auswertungen"]
    )

    if page == "Analyse":
        page_analyse(df_filt, df_km, filt)
    elif page == "Statistik":
        page_statistik(df_filt, df_km,
                       km_fahren=filt["km_fahren"],
                       kontinuierlich=filt["kontinuierlich"])
    elif page == "KM-Betrachtung":
        page_km_betrachtung(df_filt, df_km, km_fahren=filt["km_fahren"])
    elif page == "Übersicht":
        page_uebersicht(df_filt, filt)
    elif page == "Kategorien":
        page_kategorien(df_filt, diskret=filt["diskret"])
    else:
        # für „Monatliche Auswertungen“ ggf. das rohe Excel statt Parquet
        page_monatliche_auswertungen(df_filt, bus_to_hersteller)

if __name__ == "__main__":
    main()
