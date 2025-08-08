# functions.py
import pandas as pd
from io import BytesIO
import plotly.express as px
import numpy as np
import streamlit as st

# Farb-Schemata
DISCRETE_SCHEMAS = {
    "Plotly": px.colors.qualitative.Plotly,
    "Bold":   px.colors.qualitative.Bold,
    "Pastel": px.colors.qualitative.Pastel,
    "D3":     px.colors.qualitative.D3,
    "Set1":   px.colors.qualitative.Set1
}
CONTINUOUS_SCHEMAS = {
    "Viridis": px.colors.sequential.Viridis,
    "Cividis": px.colors.sequential.Cividis,
    "Inferno": px.colors.sequential.Inferno,
    "Magma":   px.colors.sequential.Magma,
    "Plasma":  px.colors.sequential.Plasma,
    "Turbo":   px.colors.sequential.Turbo
}






@st.cache_data
def load_data(
    summary_path: str,
    date_path:    str
) -> pd.DataFrame:
    """
    Lädt die Excelsheets 'Osten' und 'Moosach' aus summary_path,
    wandelt sie ins Long-Format, füllt Nan-Werte, weist Ausfall-Typen
    zu, und filtert anschliessend alle Zeilen, die vor der Zulassung
    oder nach dem Verkauf liegen (gemäss date_path).
    """
    # 1) Zulassungs-/Verkaufs-Daten einlesen und aufbereiten
    df_dates = pd.read_excel(date_path, engine="openpyxl")
    df_dates = df_dates.rename(columns={
        "KOM-Nr.":       "BusNr",
        "Einsatz":      "ZulassungDatum",
        "Verkauf":      "VerkaufDatum"
    })
    # sicherstellen, dass BusNr als String vorliegt
    df_dates["BusNr"] = df_dates["BusNr"].astype(str).str.strip()
    df_dates["ZulassungDatum"] = pd.to_datetime(df_dates["ZulassungDatum"], errors="coerce")
    df_dates["VerkaufDatum"]   = pd.to_datetime(df_dates["VerkaufDatum"],   errors="coerce")
    
    # 2) Deine bereits existierende Logik: Einlesen und Long-Format etc.
    sheets = pd.read_excel(summary_path, sheet_name=["Osten","Moosach"], engine="openpyxl")
    df_list = []
    for bereich, df in sheets.items():
        df = df.copy()
        df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True)
        bus_cols = [c for c in df.columns if c!="Datum"]
        df_long = df.melt(
            id_vars="Datum",
            value_vars=bus_cols,
            var_name="Bus",
            value_name="Ausfallgrund"
        )
        # … hier alle Deine Schritte 4–10 (NaN-Füllen, Typ zuweisen, Serie, Q, …)
        # z.B.:
        df_long["Ausfallgrund"] = df_long["Ausfallgrund"].replace("", np.nan).fillna("Keine Ausfälle")
        df_long["BusNr"]        = df_long["Bus"].str.extract(r"(\d+)", expand=False).astype("Int64")
        # … usw.
        df_long["Bereich"] = bereich
        df_list.append(df_long)
    df_all = pd.concat(df_list, ignore_index=True)
    
    # 3) Jetzt den Merge mit Zulassungs-/Verkaufsdaten und Filtern
    #    erst BusNr in beiden als String
    df_all["BusNr"] = df_all["BusNr"].astype(str).str.strip()
    df_merged = df_all.merge(
        df_dates[["BusNr","ZulassungDatum","VerkaufDatum"]],
        on="BusNr",
        how="left"
    )
    # 4) Maske: Datum zwischen ZulassungDatum und VerkaufDatum (oder kein VerkaufDatum)
    mask = (
        (df_merged["Datum"] >= df_merged["ZulassungDatum"]) &
        (
            df_merged["VerkaufDatum"].isna() |
            (df_merged["Datum"] <= df_merged["VerkaufDatum"])
        )
    ).fillna(False)
    df_filtered = df_merged[mask].copy()
    
    # 5) Falls Du die Zusatsspalten nicht mehr brauchst, droppen:
    df_filtered = df_filtered.drop(columns=["ZulassungDatum","VerkaufDatum"])
    
    return df_filtered


def export_excel_with_charts(
    df_export: pd.DataFrame,
    top_df: pd.DataFrame,
    bus_counts: pd.DataFrame,
    serie_counts: pd.DataFrame,
    df_time: pd.DataFrame,
    pivot: pd.DataFrame,
    quart_tab: pd.DataFrame,
    kpi_dict: dict,
    diskret: str,
    kontinuierlich: str,  # wird hier aktuell nicht verwendet
    top_n: int
) -> bytes:
    """
    Exportiert ein Excel-Workbook mit:
      - Rohdaten
      - Pie-, Bus-, Serie- und Zeitreihen-Daten
      - Pivot- und Quartalstabellen
      - KPI-Übersicht
      - einer vollständigen Datentabelle ('FullDates')
      - einem separaten Chartsheet
    """

    # 1) Vollständiges Datum vorbereiten und 'BusNr' sauber konvertieren
    full_dates = df_export.copy()

    # erst alle nicht-numerischen Einträge auf NaN coeren und dann in den Pandas-nullable-Int konvertieren
    full_dates["BusNr"] = (
        pd.to_numeric(full_dates["BusNr"], errors="coerce")
          .astype("Int64")
    )

    # 2) Start des Excel-Schreibvorgangs
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Rohdaten
        df_export.to_excel(writer, sheet_name="Rohdaten", index=False)

        # Pie-Daten
        top_df.to_excel(writer, sheet_name="PieData", index=False)

        # Bus-Daten
        bus_counts.to_excel(writer, sheet_name="BusData", index=False)

        # Serie-Daten
        serie_counts.to_excel(writer, sheet_name="SerieData", index=False)

        # Zeitreihen-Daten
        df_time.to_excel(writer, sheet_name="TimeData", index=False)

        # Pivot-Tabelle
        pivot.to_excel(writer, sheet_name="Pivot", startrow=1, index=False)

        # Quartalstabelle
        quart_tab.to_excel(writer, sheet_name="Quartal", index=False)

        # KPI-Übersicht
        kpi_df = pd.DataFrame(list(kpi_dict.items()), columns=["Metrik", "Wert"])
        kpi_df.to_excel(writer, sheet_name="KPIs", index=False)

        # FullDates (mit sauberem BusNr-Typ)
        full_dates.to_excel(writer, sheet_name="FullDates", index=False)

        # Chartsheet anlegen
        chartsheet = workbook.add_worksheet("Charts")

        # Pie-Chart
        pie = workbook.add_chart({"type": "pie"})
        pie.add_series({
            "name":       f"Top {top_n} Ausfallgründe",
            "categories": ["PieData", 1, 0, len(top_df), 0],
            "values":     ["PieData", 1, 1, len(top_df), 1],
        })
        pie.set_title({"name": f"Top {top_n} Ausfallgründe"})
        chartsheet.insert_chart("B2", pie, {"x_scale": 1.5, "y_scale": 1.5})

        # Bar-Chart Ausfälle pro Bus
        ch2 = workbook.add_chart({"type": "column"})
        ch2.add_series({
            "name":       "Ausfälle pro Bus",
            "categories": ["BusData", 1, 0, len(bus_counts), 0],
            "values":     ["BusData", 1, 1, len(bus_counts), 1],
            "fill":       {"color": DISCRETE_SCHEMAS[diskret][0]}
        })
        ch2.set_title({"name": "Ausfälle pro Bus"})
        chartsheet.insert_chart("J2", ch2, {"x_scale": 1.3, "y_scale": 1.3})

        # Bar-Chart Ausfälle pro Serie
        ch3 = workbook.add_chart({"type": "column"})
        ch3.add_series({
            "name":       "Ausfälle pro Serie",
            "categories": ["SerieData", 1, 0, len(serie_counts), 0],
            "values":     ["SerieData", 1, 1, len(serie_counts), 1],
            "fill":       {"color": DISCRETE_SCHEMAS[diskret][1]}
        })
        ch3.set_title({"name": "Ausfälle pro Serie"})
        chartsheet.insert_chart("B22", ch3, {"x_scale": 1.3, "y_scale": 1.3})

        # Liniendiagramm Ausfälle über Zeit
        ch4 = workbook.add_chart({"type": "line"})
        ch4.add_series({
            "name":       "Ausfälle über Zeit",
            "categories": ["TimeData", 1, 0, len(df_time), 0],
            "values":     ["TimeData", 1, 1, len(df_time), 1],
            "line":       {"color": DISCRETE_SCHEMAS[diskret][2]}
        })
        ch4.set_title({"name": "Ausfälle über die Zeit"})
        chartsheet.insert_chart("J22", ch4, {"x_scale": 1.3, "y_scale": 1.3})

    # roh als bytes zurückgeben
    return out.getvalue()

def to_excel_raw(df_export: pd.DataFrame) -> bytes:
    """
    Exportiert einfach die Rohdaten in eine Excel-Datei.
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Rohdaten")
    return out.getvalue()