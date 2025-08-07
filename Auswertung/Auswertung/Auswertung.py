import streamlit as st
import itertools
import pandas as pd
import plotly.express as px
from io import BytesIO

# ── 1) App‐Konfiguration ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Ausfall‐Analyse Busflotte",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 2) Daten laden & transformieren ────────────────────────────────────────────
@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True)
    df_long = (
        df.melt(id_vars=["Datum"], var_name="Bus", value_name="Ausfallgrund")
          .dropna(subset=["Ausfallgrund"])
    )
    # Busnummer extrahieren
    df_long["BusNr"] = df_long["Bus"].str.extract(r"(\d+)").astype(int)

    # Typ zuweisen (Standtage / Einrücker / Sonstiges)
    df_long["Typ"] = "Sonstiges"
    mask_st = df_long["Ausfallgrund"].str.match(r"^(St|st|ST)\b")
    mask_e = df_long["Ausfallgrund"].str.match(r"^[Ee]\b")
    df_long.loc[mask_st, "Typ"] = "Standtage"
    df_long.loc[mask_e, "Typ"] = "Einrücker"

    # Serie bestimmen (1–10, 11–20, …)
    serie_start = ((df_long["BusNr"] - 1) // 10) * 10 + 1
    df_long["Serie"] = (
        serie_start.astype(str) + "–" + (serie_start + 9).astype(str)
    )

    # Jahr‐Quartal als String
    df_long["Jahr-Quartal"] = df_long["Datum"].dt.to_period("Q").astype(str)

    return df_long


df = load_data("Zusammenfassung.xlsx")

# ── 3) Seitenauswahl ───────────────────────────────────────────────────────────
page = st.sidebar.radio(
    "Seite wählen",
    ["Analyse", "Statistik", "KM-Betrachtung"]
)

# ── 4) Sidebar‐Filter ──────────────────────────────────────────────────────────
st.sidebar.markdown("## 🔎 Filter")
min_d, max_d = df["Datum"].min(), df["Datum"].max()
datum_start, datum_ende = st.sidebar.date_input(
    "Datum von–bis",
    [min_d, max_d],
    min_value=min_d,
    max_value=max_d
)
quartal_selektion = st.sidebar.multiselect(
    "Quartal",
    sorted(df["Jahr-Quartal"].unique()),
    default=sorted(df["Jahr-Quartal"].unique())
)
bus_selektion = st.sidebar.multiselect(
    "Busnummer(n)",
    sorted(df["BusNr"].unique()),
    default=sorted(df["BusNr"].unique())
)
serie_selektion = st.sidebar.multiselect(
    "Busserie(n)",
    sorted(df["Serie"].unique()),
    default=sorted(df["Serie"].unique())
)
gruende_selektion = st.sidebar.multiselect(
    "Ausfallgrund/Gründe",
    sorted(df["Ausfallgrund"].unique()),
    default=sorted(df["Ausfallgrund"].unique())
)
typ_selektion = st.sidebar.multiselect(
    "Ausfall-Typ",
    ["Standtage", "Einrücker", "Sonstiges"],
    default=["Standtage", "Einrücker", "Sonstiges"]
)
top_n = st.sidebar.slider("Top N Ausfallgründe im Pie", 3, 15, 7)
zeit_gruppe = st.sidebar.radio("Zeit gruppieren nach", ["Täglich", "Wöchentlich", "Monatlich"])
ts_diagramm = st.sidebar.selectbox("Typ Zeitreihe", ["Linie", "Fläche", "Balken"])

diskrete_schemata = {
    "Plotly": px.colors.qualitative.Plotly,
    "Bold":   px.colors.qualitative.Bold,
    "Pastel": px.colors.qualitative.Pastel,
    "D3":     px.colors.qualitative.D3,
    "Set1":   px.colors.qualitative.Set1
}
kont_schemata = {
    "Viridis": px.colors.sequential.Viridis,
    "Cividis": px.colors.sequential.Cividis,
    "Inferno": px.colors.sequential.Inferno,
    "Magma":   px.colors.sequential.Magma,
    "Plasma":  px.colors.sequential.Plasma,
    "Turbo":   px.colors.sequential.Turbo
}
diskret = st.sidebar.selectbox("Diskretes Farbschema", list(diskrete_schemata.keys()), index=0)
kontinuierlich = st.sidebar.selectbox("Kontinuierliches Farbschema", list(kont_schemata.keys()), index=0)

# ── 5) Filter anwenden ─────────────────────────────────────────────────────────
mask = (
    (df["Datum"] >= pd.to_datetime(datum_start)) &
    (df["Datum"] <= pd.to_datetime(datum_ende)) &
    (df["Jahr-Quartal"].isin(quartal_selektion)) &
    (df["BusNr"].isin(bus_selektion)) &
    (df["Serie"].isin(serie_selektion)) &
    (df["Ausfallgrund"].isin(gruende_selektion)) &
    (df["Typ"].isin(typ_selektion))
)
df_filt = df[mask].copy()

# ── 5a) Neue Eingaben für Default‐Kilometer ────────────────────────────────────
st.sidebar.markdown("## ⚙️ Standard-Kilometer pro Typ")
einr_km = st.sidebar.number_input("Default km für Einrücker", min_value=0, value=50, step=10)
stand_km = st.sidebar.number_input("Default km für Standtage", min_value=0, value=0, step=10)
sonst_km = st.sidebar.number_input("Default km für Sonstiges", min_value=0, value=250, step=10)
fahren_km = st.sidebar.number_input("Default km für Fahren", min_value=0, value=250, step=10)

# 5b) KM‐Spalte mit dynamischen Defaults nach Typ
defaults = {
    "Einrücker": einr_km,
    "Standtage": stand_km,
    "Sonstiges": sonst_km
}
df_filt["km"] = df_filt["Typ"].map(defaults).fillna(fahren_km)

# 5c) Interaktives Bearbeiten der KM direkt in der App
st.markdown("## 🛣️ KM‐Betrachtung")
df_km = st.data_editor(
    df_filt[["Datum", "BusNr", "Typ", "km"]],
    num_rows="fixed",
    use_container_width=True
)

# 5d) «Fahren» nicht als Ausfall werten
df_filt = df_filt[df_filt["Typ"] != "Fahren"]


# ── Gemeinsame Export‐Funktion für Excel mit Charts ────────────────────────────
def export_excel_with_charts(
    df_export: pd.DataFrame,
    top_df: pd.DataFrame,
    bus_counts: pd.DataFrame,
    serie_counts: pd.DataFrame,
    df_time: pd.DataFrame,
    pivot: pd.DataFrame,
    quart_tab: pd.DataFrame,
    kpi_dict: dict
) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book

        # 1) Rohdaten
        df_export.to_excel(writer, sheet_name="Rohdaten", index=False)
        # 2) Pie‐Daten
        top_df.to_excel(writer, sheet_name="PieData", index=False)
        # 3) Bus‐Daten
        bus_counts.to_excel(writer, sheet_name="BusData", index=False)
        # 4) Serie‐Daten
        serie_counts.to_excel(writer, sheet_name="SerieData", index=False)
        # 5) Zeitreihe‐Daten
        df_time.to_excel(writer, sheet_name="TimeData", index=False)
        # 6) Pivot
        pivot.to_excel(writer, sheet_name="Pivot", startrow=1)
        # 7) Quartalstabelle
        quart_tab.to_excel(writer, sheet_name="Quartal", index=False)
        # 8) KPIs
        kpi_df = pd.DataFrame(list(kpi_dict.items()), columns=["Metrik", "Wert"])
        kpi_df.to_excel(writer, sheet_name="KPIs", index=False)

        # Chartsheet erstellen
        chartsheet = workbook.add_worksheet("Charts")

        # Pie‐Chart
        pie = workbook.add_chart({"type": "pie"})
        pie.add_series({
            "name":       f"Top {top_n} Ausfallgründe",
            "categories": ["PieData", 1, 0, len(top_df), 0],
            "values":     ["PieData", 1, 1, len(top_df), 1],
        })
        pie.set_title({"name": f"Top {top_n} Ausfallgründe"})
        chartsheet.insert_chart("B2", pie, {"x_scale": 1.5, "y_scale": 1.5})

        # Bar‐Chart Bus
        ch2 = workbook.add_chart({"type": "column"})
        ch2.add_series({
            "name":       "Ausfälle pro Bus",
            "categories": ["BusData", 1, 0, len(bus_counts), 0],
            "values":     ["BusData", 1, 1, len(bus_counts), 1],
            "fill":       {"color": diskrete_schemata[diskret][0]}
        })
        ch2.set_title({"name": "Ausfälle pro Bus"})
        chartsheet.insert_chart("J2", ch2, {"x_scale": 1.3, "y_scale": 1.3})

        # Bar‐Chart Serie
        ch3 = workbook.add_chart({"type": "column"})
        ch3.add_series({
            "name":       "Ausfälle pro Serie",
            "categories": ["SerieData", 1, 0, len(serie_counts), 0],
            "values":     ["SerieData", 1, 1, len(serie_counts), 1],
            "fill":       {"color": diskrete_schemata[diskret][1]}
        })
        ch3.set_title({"name": "Ausfälle pro Serie"})
        chartsheet.insert_chart("B22", ch3, {"x_scale": 1.3, "y_scale": 1.3})

        # Zeitreihe‐Chart
        ch4 = workbook.add_chart({"type": "line"})
        ch4.add_series({
            "name":       "Ausfälle über Zeit",
            "categories": ["TimeData", 1, 0, len(df_time), 0],
            "values":     ["TimeData", 1, 1, len(df_time), 1],
            "line":       {"color": diskrete_schemata[diskret][2]}
        })
        ch4.set_title({"name": "Ausfälle über die Zeit"})
        chartsheet.insert_chart("J22", ch4, {"x_scale": 1.3, "y_scale": 1.3})

    return out.getvalue()


# ── ANALYSE‐Seite ───────────────────────────────────────────────────────────────
if page == "Analyse":
    st.title("🚍 Ausfall‐Analyse")

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Zeitraum", f"{datum_start} bis {datum_ende}")
    c2.metric("Quartale", ", ".join(quartal_selektion))
    c3.metric("Ausfälle gesamt", len(df_filt))
    tage = df_filt["Datum"].nunique() or 1
    c4.metric("Ø Ausfälle/Tag", f"{len(df_filt)/tage:.2f}")
    st.markdown("---")

    # 1) Piechart Top N + „Sonstige“
    gr_counts = (
    df_filt["Ausfallgrund"]
      .value_counts()
      .rename_axis("Ausfallgrund")           # macht das Index‐Label sauber
      .reset_index(name="Anzahl")            # benennt die Zähl‐Spalte direkt in 'Anzahl'
)
    top_df = gr_counts.nlargest(top_n, "Anzahl").copy()
    others = gr_counts["Anzahl"].sum() - top_df["Anzahl"].sum()
    if others > 0:
        top_df = pd.concat(
            [top_df, pd.DataFrame({"Ausfallgrund": ["Sonstige"], "Anzahl": [others]})],
            ignore_index=True
        )
    fig1 = px.pie(
        top_df,
        names="Ausfallgrund",
        values="Anzahl",
        title=f"Top {top_n} Ausfallgründe",
        color_discrete_sequence=diskrete_schemata[diskret]
    )
    st.plotly_chart(fig1, use_container_width=True)

    # 2) Ausfälle pro Bus
    bus_counts = (
    df_filt
      .groupby("BusNr")
      .size()
      .reset_index(name="Anzahl")    # erzeugt zwei Spalten: "BusNr" und "Anzahl"
      .sort_values("BusNr")
    )
    fig2 = px.bar(
        bus_counts,
        x="BusNr",
        y="Anzahl",
        title="Ausfälle pro Bus",
        color="BusNr",
        color_discrete_sequence=diskrete_schemata[diskret]
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 3) Ausfälle pro Serie
    serie_counts = (
        df_filt.groupby("Serie").size().reset_index(name="Anzahl").sort_values("Serie")
    )
    fig3 = px.bar(
        serie_counts,
        x="Serie",
        y="Anzahl",
        title="Ausfälle pro Serie",
        color="Serie",
        color_discrete_sequence=diskrete_schemata[diskret]
    )
    st.plotly_chart(fig3, use_container_width=True)

    # 4) Zeitreihe
    if zeit_gruppe == "Wöchentlich":
        df_time = df_filt.set_index("Datum").resample("W").size().reset_index(name="Anzahl")
    elif zeit_gruppe == "Monatlich":
        df_time = df_filt.set_index("Datum").resample("M").size().reset_index(name="Anzahl")
    else:
        df_time = df_filt.groupby("Datum").size().reset_index(name="Anzahl")

    if ts_diagramm == "Linie":
        fig4 = px.line(
            df_time,
            x="Datum",
            y="Anzahl",
            title="Ausfälle über die Zeit",
            markers=True,
            color_discrete_sequence=[diskrete_schemata[diskret][0]]
        )
    elif ts_diagramm == "Fläche":
        fig4 = px.area(
            df_time,
            x="Datum",
            y="Anzahl",
            title="Ausfälle über die Zeit",
            color_discrete_sequence=[diskrete_schemata[diskret][0]]
        )
    else:
        fig4 = px.bar(
            df_time,
            x="Datum",
            y="Anzahl",
            title="Ausfälle über die Zeit",
            color_discrete_sequence=[diskrete_schemata[diskret][0]]
        )
    st.plotly_chart(fig4, use_container_width=True)

    # 5) Ausfälle pro Quartal
    quart_tab = (
        df_filt["Jahr-Quartal"]
        .value_counts()
        .rename_axis("Jahr-Quartal")
        .reset_index(name="Anzahl")
        .sort_values("Jahr-Quartal")
    )
    fig_q = px.bar(
        quart_tab,
        x="Jahr-Quartal",
        y="Anzahl",
        title="Ausfälle pro Quartal",
        color="Anzahl",
        color_continuous_scale=kont_schemata[kontinuierlich]
    )
    st.plotly_chart(fig_q, use_container_width=True)

    # Vorbereitung Export
    pivot = df_filt.pivot_table(
        index="Serie", columns="Ausfallgrund", aggfunc="size", fill_value=0
    )
    kpi_dict = {
        "Zeitraum":        f"{datum_start} bis {datum_ende}",
        "Quartale":        ", ".join(quartal_selektion),
        "Ausfälle gesamt": len(df_filt),
        "Ø Ausfälle/Tag":  f"{len(df_filt)/(df_filt['Datum'].nunique() or 1):.2f}"
    }

    excel_bytes = export_excel_with_charts(
        df_filt, top_df, bus_counts, serie_counts, df_time,
        pivot, quart_tab, kpi_dict
    )
    st.download_button(
        "📥 Komplette Auswertung als Excel herunterladen",
        data=excel_bytes,
        file_name="auswertung_interaktiv.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ── STATISTIK‐Seite ─────────────────────────────────────────────────────────────
elif page == "Statistik":
    st.title("📊 Grundstatistik & KM‐Betrachtung")

    # 1) Häufigkeiten der Ausfallgründe
    gr_tab = (
        df_filt["Ausfallgrund"]
        .value_counts()
        .rename_axis("Ausfallgrund")
        .reset_index(name="Anzahl")
    )
    st.subheader("Ausfallgründe")
    st.dataframe(gr_tab, use_container_width=True)

    # 2) Häufigkeiten der Busse
    bus_tab = (
        df_filt["BusNr"]
        .value_counts()
        .rename_axis("BusNr")
        .reset_index(name="Anzahl")
    )
    st.subheader("Busse")
    st.dataframe(bus_tab, use_container_width=True)

    # 3) Häufigkeiten der Serien
    ser_tab = (
        df_filt["Serie"]
        .value_counts()
        .rename_axis("Serie")
        .reset_index(name="Anzahl")
    )
    st.subheader("Serien")
    st.dataframe(ser_tab, use_container_width=True)

    # 4) Pivot‐Tabelle Serie × Ausfallgrund
    st.markdown("### Pivot‐Tabelle (Serie × Ausfallgrund)")
    pivot_stat = df_filt.pivot_table(
        index="Serie", columns="Ausfallgrund", aggfunc="size", fill_value=0
    )
    st.dataframe(pivot_stat, use_container_width=True)

    # 5) Ausfälle pro Quartal
    st.markdown("### Ausfälle pro Quartal")
    quart_tab_stat = (
        df_filt["Jahr-Quartal"]
        .value_counts()
        .rename_axis("Jahr-Quartal")
        .reset_index(name="Anzahl")
        .sort_values("Jahr-Quartal")
    )
    st.dataframe(quart_tab_stat, use_container_width=True)

    fig_q2 = px.bar(
        quart_tab_stat,
        x="Jahr-Quartal",
        y="Anzahl",
        title="Ausfälle pro Quartal",
        color="Anzahl",
        color_continuous_scale=kont_schemata[kontinuierlich]
    )
    st.plotly_chart(fig_q2, use_container_width=True)

    # 6) KM‐Auswertung pro Bus
    st.markdown("### 🛣️ KM‐Auswertung pro Bus")
    bus_km = (
        df_km
        .groupby("BusNr")
        .agg(
            Tage=("Datum", "nunique"),
            km_ist=("km", "sum")
        )
        .reset_index()
    )
    bus_km["km_soll"] = bus_km["Tage"] * fahren_km
    bus_km["Verf_%"] = (bus_km["km_ist"] / bus_km["km_soll"] * 100).round(1)
    bus_km = bus_km[["BusNr", "Tage", "km_ist", "km_soll", "Verf_%"]]
    st.dataframe(bus_km, use_container_width=True)

    # 7) Rohdaten‐Export
    def to_excel_raw(df_export: pd.DataFrame) -> bytes:
        out = BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Rohdaten")
        return out.getvalue()

    if not df_filt.empty:
        excel_bytes = to_excel_raw(df_filt)
        st.download_button(
            "📥 Rohdaten als Excel herunterladen",
            data=excel_bytes,
            file_name="rohdaten_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Keine Daten zum Export. Bitte Filter anpassen.")

# ── KM-Betrachtung‐Seite ─────────────────────────────────────────────────────────
else:  # page == "KM-Betrachtung"
    st.title("🛣️ KM-Betrachtung")
    st.markdown(
        "In diesem Menüpunkt kannst Du die km-Daten interaktiv anpassen "
        "und siehst dazu passende Diagramme."
    )

    # 1) Interaktiver Editor für Datum, BusNr, Typ, km
    df_km_edit = st.data_editor(
        df_filt[["Datum", "BusNr", "Typ", "km"]],
        num_rows="fixed",
        use_container_width=True,
        key="km_editor"
    )

    # 2) Vollmatrix ergänzen: für alle Datums-Bus-Kombinationen
    all_dates = pd.date_range(datum_start, datum_ende, freq="D")
    all_buses = df_filt["BusNr"].unique()
    full_idx = pd.MultiIndex.from_product(
        [all_dates, all_buses], names=["Datum", "BusNr"]
    )
    df_full = (
        pd.DataFrame(index=full_idx)
        .reset_index()
        .merge(df_km_edit, on=["Datum", "BusNr"], how="left")
    )
    df_full["Typ"] = df_full["Typ"].fillna("Fahren")
    df_full["km"]  = df_full["km"].fillna(fahren_km)

    # Serie beifügen
    serie_map = df_filt[["BusNr", "Serie"]].drop_duplicates()
    df_full = df_full.merge(serie_map, on="BusNr", how="left")

    # 3) Histogramm: Verteilung der km-Werte nach Typ
    st.subheader("Verteilung der gefahrenen km je Typ")
    fig_dist = px.histogram(
        df_full,
        x="km",
        color="Typ",
        nbins=30,
        title="Histogramm der km-Werte pro Typ",
        labels={"km": "gefahrene km"},
        color_discrete_sequence=px.colors.qualitative.Plotly
    )
    st.plotly_chart(fig_dist, use_container_width=True)

    # 4) km-Zusammenfassung pro Bus
    st.subheader("km-Zusammenfassung pro Bus")
    bus_km2 = (
        df_full
        .groupby("BusNr")
        .agg(
            Tage=("Datum", "nunique"),
            km_ist=("km", "sum")
        )
        .reset_index()
    )
    bus_km2["km_soll"] = bus_km2["Tage"] * fahren_km
    bus_km2["Verf_%"]  = (bus_km2["km_ist"] / bus_km2["km_soll"] * 100).round(1)
    st.dataframe(bus_km2, use_container_width=True)

    # 5) Bar-Charts: Ist-Kilometer und Verfügbarkeit
    col1, col2 = st.columns(2)
    with col1:
        fig_ist = px.bar(
            bus_km2,
            x="BusNr",
            y="km_ist",
            title="Ist-Kilometer pro Bus",
            labels={"km_ist": "Ist-Kilometer", "BusNr": "Busnummer"},
            color="BusNr",
            color_continuous_scale="Viridis"
        )
        st.plotly_chart(fig_ist, use_container_width=True)
    with col2:
        fig_verf = px.bar(
            bus_km2,
            x="BusNr",
            y="Verf_%",
            title="Verfügbarkeit in %",
            labels={"Verf_%": "Verfügbarkeit (%)", "BusNr": "Busnummer"},
            color="Verf_%",
            color_continuous_scale="Magma"
        )
        st.plotly_chart(fig_verf, use_container_width=True)

    # 6) km-Zusammenfassung pro Bus-Serie
    st.subheader("km-Zusammenfassung pro Bus-Serie")
    serie_km = (
        bus_km2
        .merge(serie_map, on="BusNr", how="left")
        .groupby("Serie")
        .agg(
            km_ist=("km_ist", "sum"),
            km_soll=("km_soll", "sum")
        )
        .reset_index()
    )
    serie_km["Verf_%"] = (serie_km["km_ist"] / serie_km["km_soll"] * 100).round(1)
    st.dataframe(serie_km, use_container_width=True)

    fig_serie = px.bar(
        serie_km,
        x="Serie",
        y="Verf_%",
        title="Verfügbarkeit pro Serie (%)",
        labels={"Verf_%": "Verfügbarkeit (%)"},
        color="Verf_%",
        color_continuous_scale="Plasma"
    )
    st.plotly_chart(fig_serie, use_container_width=True)