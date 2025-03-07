import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Durchschnittliche Abverkaufsmengen", layout="wide")

st.title("Berechnung der Ø Abverkaufsmengen pro Woche von Werbeartikeln zu Normalpreisen")

# Funktion zur Umwandlung der Originaldatei in das benötigte Format
def convert_original_file(uploaded_file):
    df_original = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    df_original.columns = df_original.iloc[1]  # Zweite Zeile als Spaltenüberschriften setzen
    df_original = df_original[2:]  # Erste zwei Zeilen entfernen
    df_transformed = df_original[[df_original.columns[0], df_original.columns[1], "Woche", "VerkaufsME | Wochentag", "Gesamtergebnis"]].copy()
    df_transformed.columns = ["Artikel", "Name", "Woche", "VerkaufsME", "Menge"]
    df_transformed["Woche"] = pd.to_numeric(df_transformed["Woche"], errors='coerce')
    df_transformed["Menge"] = pd.to_numeric(df_transformed["Menge"], errors='coerce')
    return df_transformed

# Datei-Uploader
uploaded_file = st.file_uploader("Bitte laden Sie Ihre Abverkaufsdatei hoch (Excel)", type=["xlsx"])

if uploaded_file:
    data = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("Wählen Sie das Blatt aus", data.sheet_names)
    df = data.parse(sheet_name)
    
    required_columns = {"Artikel", "Woche", "Menge", "Name"}
    if not required_columns.issubset(df.columns):
        st.warning("Die Datei scheint nicht das richtige Format zu haben. Die App wird sie nun umwandeln.")
        df = convert_original_file(uploaded_file)

    if df.isnull().values.any():
        st.error("Fehler: Die Datei enthält fehlende Werte. Bitte stellen Sie sicher, dass alle Zellen ausgefüllt sind.")
    else:
        st.sidebar.title("Artikel-Filter")
        artikel_filter = st.sidebar.text_input("Nach Artikelnummer filtern (optional)")
        artikel_name_filter = st.sidebar.text_input("Nach Artikelname filtern (optional)")

        if artikel_filter:
            df = df[df['Artikel'].astype(str).str.contains(artikel_filter, case=False, na=False)]

        if artikel_name_filter:
            df = df[df['Name'].str.contains(artikel_name_filter, case=False, na=False)]

        result = df.groupby(['Artikel', 'Name'], sort=False).agg({'Menge': 'mean'}).reset_index()
        result.rename(columns={'Menge': 'Durchschnittliche Menge pro Woche'}, inplace=True)

        round_option = st.sidebar.selectbox("Rundungsoption für alle Artikel:", ['Nicht runden', 'Aufrunden', 'Abrunden', 'Kaufmännisch runden'], index=0)

        if round_option == 'Aufrunden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x + 0.5))
        elif round_option == 'Abrunden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x - 0.5))
        elif round_option == 'Kaufmännisch runden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x))

        st.subheader("Ergebnisse")
        st.dataframe(result)
        st.info("Verarbeitung abgeschlossen. Die Ergebnisse stehen zur Verfügung.")

        output = BytesIO()
        result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button(label="Ergebnisse herunterladen", data=output, file_name="durchschnittliche_abverkaeufe.xlsx")

st.markdown("---")
st.markdown("⚠️ **Hinweis:** Diese Anwendung speichert keine Daten und hat keinen Zugriff auf Ihre Dateien.")
st.markdown("🌟 **Erstellt von Christoph R. Kaiser mit Hilfe von Künstlicher Intelligenz.**")
