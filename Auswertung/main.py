# app.py
import streamlit as st
import pandas as pd
import os, base64
from pathlib import Path
import time 
from itertools import product
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






# Ordner für zwischengespeicherte Daten
CACHE_DIR = Path("data_cache")
CACHE_DIR.mkdir(exist_ok=True)



# Funktion zum Laden aller Daten und Speichern im Cache
# Funktion zum Preloading und Caching
def preload_filter_combinations_with_progress(df: pd.DataFrame, progress_bar, progress_text) -> dict:
    """
    Preloading aller möglichen Filterkombinationen und Speichern im Cache.
    Wenn Cache existiert, lade die Daten aus dem Cache.
    """
    cache_file = CACHE_DIR / "filter_cache.parquet"

    # Prüfen, ob Cache vorhanden ist
    if cache_file.exists():
        progress_text.text("🔄 Lade vorkalkulierte Daten aus dem Cache...")
        cached_data = pd.read_parquet(cache_file).to_dict(orient="index")
        progress_bar.progress(100)
        return {tuple(k.split('|')): v for k, v in cached_data.items()}

    # Alle Filteroptionen ermitteln
    progress_text.text("🛠️ Berechne Filterkombinationen...")
    bus_numbers = sorted(df["BusNr"].unique())
    series = sorted(df["Serie"].unique())
    quartals = sorted(df["Jahr-Quartal"].unique())
    ausfall_typen = sorted(df["Ausfall-Typ"].unique())

    # Preloading mit Fortschrittsanzeige
    total_combinations = len(bus_numbers) * len(series) * len(quartals) * len(ausfall_typen)

    # Cache erstellen
    filter_cache = {}
    for i, (bus_nr, serie, quartal, ausfall_typ) in enumerate(
        product(bus_numbers, series, quartals, ausfall_typen), 1
    ):
        filtered_df = df[
            (df["BusNr"] == bus_nr)
            & (df["Serie"] == serie)
            & (df["Jahr-Quartal"] == quartal)
            & (df["Ausfall-Typ"] == ausfall_typ)
        ]
        filter_cache[(bus_nr, serie, quartal, ausfall_typ)] = filtered_df

        # Fortschritt aktualisieren
        progress_bar.progress(int((i / total_combinations) * 100))

    # Fortschrittsanzeige beenden
    progress_bar.progress(100)
    progress_text.text("✅ Berechnungen abgeschlossen. Speichere Cache...")

    # Cache speichern
    cache_df = pd.DataFrame.from_dict(
        {f"{k[0]}|{k[1]}|{k[2]}|{k[3]}": v for k, v in filter_cache.items()},
        orient="index"
    )
    cache_df.to_parquet(cache_file)

    progress_text.text("✅ Cache gespeichert!")
    return filter_cache



# ── 1) Einmaliges Laden + Parquet-Dump ───────────────────────────────────────
@st.cache_resource(show_spinner=False)
def load_all_data() -> tuple[dict[int, str], pd.DataFrame]:
    """
    Lädt Bus→Hersteller-Mapping und die gefilterte Datenbank (mit Parquet-Cache).
    Alle Daten werden nur einmal geladen.
    """
    RAW_SUMMARY = "Zusammenfassung.xlsx"
    DATE_FILE = "Zulassung-Verkauf.xlsx"
    PROCESSED_XLS = "Zusammenfassung_bearbeitet.xlsx"
    PROCESSED_PQ = "Zusammenfassung_bearbeitet.parquet"
    BUS_MAP_FILE = "bus_hersteller_zuordnung.xlsx"

    # 1. Bus→Hersteller-Mapping laden
    df_map = pd.read_excel(BUS_MAP_FILE, engine="openpyxl")
    df_map.columns = ["BusNr", "Hersteller"]
    bus_to_hersteller = df_map.set_index("BusNr")["Hersteller"].to_dict()

    # 2. Gefilterte XLSX erzeugen, falls nicht vorhanden
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
        # Daten laden und als Parquet speichern
        df = pd.read_excel(PROCESSED_XLS, sheet_name=None, engine="openpyxl")
        combined_df = pd.concat(df.values(), ignore_index=True)
        combined_df.to_parquet(PROCESSED_PQ, index=False)

    # 4. Zulassungs-/Verkaufsdaten einlesen und mit Daten verbinden
    df_dates = pd.read_excel(DATE_FILE, engine="openpyxl", sheet_name=0)
    df_dates = df_dates.rename(columns={
        "KOM-Nr.": "BusNr",
        "Einsatz": "ZulassungDatum",
        "Verkauf": "VerkaufDatum"
    })
    df_dates["BusNr"] = df_dates["BusNr"].astype(str).str.strip()
    df_dates["ZulassungDatum"] = pd.to_datetime(df_dates["ZulassungDatum"], errors="coerce")
    df_dates["VerkaufDatum"] = pd.to_datetime(df_dates["VerkaufDatum"], errors="coerce")
    df["BusNr"] = df["BusNr"].astype(str).str.strip()

    # Nur Daten innerhalb des Zulassungs- und Verkaufszeitraums
    df = df.merge(
        df_dates[["BusNr", "ZulassungDatum", "VerkaufDatum"]],
        on="BusNr", how="left"
    )
    mask = (
        (df["Datum"] >= df["ZulassungDatum"]) &
        (
            df["VerkaufDatum"].isna() |
            (df["Datum"] <= df["VerkaufDatum"])
        )
    ).fillna(False)
    df = df.loc[mask].copy()
    df = df.drop(columns=["ZulassungDatum", "VerkaufDatum"])

    return bus_to_hersteller, df

# Neue Funktion: Preloading aller möglichen Filterkombinationen
@st.cache_resource(show_spinner=False)
def preload_filter_combinations(df: pd.DataFrame) -> dict:
    """
    Lädt alle möglichen Filterkombinationen vorab und speichert die Ergebnisse in einem Cache.
    """
    # Alle möglichen Werte für die Filter
    bus_numbers = sorted(df["BusNr"].unique())
    series = sorted(df["Serie"].unique())
    quartals = sorted(df["Jahr-Quartal"].unique())
    ausfall_typen = sorted(df["Ausfall-Typ"].unique())

    # Alle Kombinationen berechnen
    filter_combinations = list(product(bus_numbers, series, quartals, ausfall_typen))

    # Cache für die vorkalkulierten Ergebnisse
    cache = {}
    for (bus_nr, serie, quartal, ausfall_typ) in filter_combinations:
        filtered_df = df[
            (df["BusNr"] == bus_nr) &
            (df["Serie"] == serie) &
            (df["Jahr-Quartal"] == quartal) &
            (df["Ausfall-Typ"] == ausfall_typ)
        ]
        cache[(bus_nr, serie, quartal, ausfall_typ)] = filtered_df

    return cache

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
def main():
    # Streamlit-Seiteneinstellungen
    st.set_page_config(page_title="Ausfall‐Analyse Busflotte", layout="wide")

    # -------------------------------------------------------------------------
    # Fortschrittsanzeige und Ladeanzeige
    # -------------------------------------------------------------------------
    st.title("🚍 Ausfall‐Analyse Busflotte")
    st.markdown("### Initialisierung der Daten – Bitte warten...")

    # Fortschrittsanzeige
    progress_placeholder = st.empty()  # Platzhalter für Fortschrittsanzeige
    progress_bar = progress_placeholder.progress(0)  # Fortschrittsbalken
    progress_text = st.empty()  # Platzhalter für Fortschrittstext

    # -------------------------------------------------------------------------
    # Schritt 1: Hauptdaten laden
    # -------------------------------------------------------------------------
    progress_text.text("🔄 Lade Hauptdaten...")
    bus_to_hersteller, df = load_all_data()  # Bestehende Funktion
    progress_bar.progress(30)  # Fortschritt aktualisieren

    # -------------------------------------------------------------------------
    # Schritt 2: Preloading aller möglichen Filterkombinationen
    # -------------------------------------------------------------------------
    progress_text.text("🛠️ Berechne Filterkombinationen und lade Cache...")
    filter_cache = preload_filter_combinations_with_progress(df, progress_bar, progress_text)
    progress_bar.progress(100)  # Fortschritt abschließen
    progress_text.text("✅ Daten erfolgreich geladen und vorkalkuliert!")

    # Fortschrittsanzeige entfernen
    progress_placeholder.empty()

    # -------------------------------------------------------------------------
    # Sidebar: Filteroptionen
    # -------------------------------------------------------------------------
    st.sidebar.markdown("## 🔎 Filter")
    bus_nr = st.sidebar.selectbox("Busnummer auswählen", options=sorted(df["BusNr"].unique()))
    serie = st.sidebar.selectbox("Serie auswählen", options=sorted(df["Serie"].unique()))
    quartal = st.sidebar.selectbox("Quartal auswählen", options=sorted(df["Jahr-Quartal"].unique()))
    ausfall_typ = st.sidebar.selectbox("Ausfall-Typ auswählen", options=sorted(df["Ausfall-Typ"].unique()))

    # -------------------------------------------------------------------------
    # Gefilterte Daten aus dem Cache abrufen
    # -------------------------------------------------------------------------
    filtered_data = filter_cache.get((bus_nr, serie, quartal, ausfall_typ), pd.DataFrame())

    # -------------------------------------------------------------------------
    # Navigation: Bestehende Menüstruktur beibehalten
    # -------------------------------------------------------------------------
    st.sidebar.markdown("## 📑 Navigation")
    page = st.sidebar.radio("Seite wählen:",
        ["Analyse", "Statistik", "KM-Betrachtung", "Übersicht", "Kategorien", "Monatliche Auswertungen"]
    )

    # -------------------------------------------------------------------------
    # Bestehende Seitenlogik
    # -------------------------------------------------------------------------
    if page == "Analyse":
        page_analyse(filtered_data, df_km=None, filt=None)  # Anpassbar, falls nötig
    elif page == "Statistik":
        page_statistik(filtered_data, df_km=None, km_fahren=250, kontinuierlich="Viridis")
    elif page == "KM-Betrachtung":
        page_km_betrachtung(filtered_data, df_km=None, km_fahren=250)
    elif page == "Übersicht":
        page_uebersicht(filtered_data, filt=None)
    elif page == "Kategorien":
        page_kategorien(filtered_data, diskret="Plotly")
    else:
        page_monatliche_auswertungen(filtered_data, bus_to_hersteller)
if __name__ == "__main__":
    main()