# functions.py
import pandas as pd
from io import BytesIO
import plotly.express as px
from pathlib import Path
import numpy as np
import streamlit as st
from typing import Union
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

RAW_SUMMARY   = "Zusammenfassung.xlsx"
DATE_FILE     = "Zulassung-Verkauf.xlsx"
PROCESSED_XLS = "Zusammenfassung_bearbeitet.xlsx"
PROCESSED_PQ  = "Zusammenfassung_bearbeitet.parquet"
BUS_MAP_FILE  = "bus_hersteller_zuordnung.xlsx"

def assign_series(
    df: pd.DataFrame,
    date_file: str,
    sheet_name: Union[int, str] = 0,
    col_busnr: str = "BusNr",
    col_serie_orig: str = "Serie"
) -> pd.DataFrame:
    """
    Mappt jede BusNr auf die Serie aus date_file und schreibt das in df[col_serie_orig].
    Fehlende BusNr werden mit dem alten Wert oder "Unbekannt" gefüllt.
    """
    df_map = pd.read_excel(
        date_file,
        sheet_name=sheet_name,
        usecols=["KOM-Nr.","Serie"],
        engine="openpyxl"
    )
    df_map.columns = [col_busnr, "Serie_neu"]
    df_map[col_busnr] = df_map[col_busnr].astype(str).str.strip()
    bus_to_serie = df_map.set_index(col_busnr)["Serie_neu"].to_dict()

    df = df.copy()
    df[col_busnr] = df[col_busnr].astype(str).str.strip()
    df["_serie_alt"] = df.get(col_serie_orig, pd.NA)
    df[col_serie_orig] = df[col_busnr].map(bus_to_serie)
    df[col_serie_orig] = (
        df[col_serie_orig]
          .fillna(df["_serie_alt"])
          .fillna("Unbekannt")
    )
    return df.drop(columns=["_serie_alt"])




def load_data(path: str, date_file: str = DATE_FILE, sheet_name: Union[int,str]=0) -> pd.DataFrame:
    import pandas as pd
    import numpy as np

    sheets = pd.read_excel(path, sheet_name=["Osten","Moosach"])
    df_list = []

    for bereich, df in sheets.items():
        df = df.copy()
        df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True)
        bus_cols = [c for c in df.columns if c != "Datum"]
        df_long = df.melt(
            id_vars=["Datum"],
            value_vars=bus_cols,
            var_name="Bus",
            value_name="Ausfallgrund"
        )
        # 1) Strings trimmen
        df_long["Ausfallgrund"] = df_long["Ausfallgrund"].replace("", np.nan).str.strip()

        # 2) Fehlende Grund → "Keine Ausfälle"
        df_long["Ausfallgrund"] = df_long["Ausfallgrund"].fillna("Keine Ausfälle")

        # 3) BusNr extrahieren
        df_long["BusNr"] = df_long["Bus"].str.extract(r"(\d+)")[0].astype(int)

        # 4) Ausfall-Typ
        df_long["Ausfall-Typ"] = "Sonstiges"
        # fahren
        df_long.loc[df_long["Ausfallgrund"] == "Keine Ausfälle", "Ausfall-Typ"] = "Fahren"
        # Standtage
        df_long.loc[df_long["Ausfallgrund"].str.startswith(("St","st"), na=False), "Ausfall-Typ"] = "Standtage"
        # Einrücker
        df_long.loc[df_long["Ausfallgrund"].str.lower().str.startswith("e"),          "Ausfall-Typ"] = "Einrücker"
        # alles übrige bleibt bei "Sonstiges"

        # 5) Bool-Spalte Ausfall?
        df_long["Ausfall"] = df_long["Ausfall-Typ"] != "Fahren"

        # 6) Jetzt die Serie aus dem Date-Mapping ziehen
        #    (anstatt 1–10,11–20…)
        df_long = assign_series(
            df_long,
            date_file=date_file,
            sheet_name=sheet_name,
            col_busnr="BusNr",
            col_serie_orig="Serie"
        )

        # 7) Quartal, Bereich
        df_long["Jahr-Quartal"] = df_long["Datum"].dt.to_period("Q").astype(str)
        df_long["Bereich"]      = bereich

        df_list.append(df_long)

    df_all = pd.concat(df_list, ignore_index=True)

    # finale Sicherstellung, dass es keinen Aberranten gibt
    df_all.loc[df_all["Ausfallgrund"] == "Keine Ausfälle", "Ausfall-Typ"] = "Fahren"
    df_all["Ausfall"] = df_all["Ausfall-Typ"] != "Fahren"

    return df_all













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