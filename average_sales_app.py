import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Durchschnittliche Abverkaufsmengen", layout="wide")

st.title("Berechnung der Ø Abverkaufsmengen pro Woche von Werbeartikeln zu Normalpreisen")

st.markdown("""
### Anleitung zur Nutzung dieser App
1. Bereiten Sie Ihre Abverkaufsdaten vor:
   - Die Datei muss die Spalten **'Artikel', 'Woche', 'Menge' (in Stück) und 'Name'** enthalten.
   - Speichern Sie die Datei im Excel-Format.
2. Laden Sie Ihre Datei hoch:
   - Nutzen Sie die Schaltfläche **„Durchsuchen“**, um Ihre Datei auszuwählen.
3. Überprüfen Sie die berechneten Ergebnisse:
   - Die App zeigt die durchschnittlichen Abverkaufsmengen pro Woche an.
4. Filtern und suchen Sie die Ergebnisse (optional):
   - Nutzen Sie das Filterfeld in der Seitenleiste, um nach bestimmten Artikeln zu suchen.
5. Vergleichen Sie die Ergebnisse (optional):
   - Laden Sie eine zweite Datei hoch, um die Ergebnisse miteinander zu vergleichen.
""")

# Datei-Uploader
uploaded_file = st.file_uploader("Bitte laden Sie Ihre Datei hoch (Excel, .xlsx)", type=["xlsx"])

if uploaded_file:
    # Excel-Datei laden und verarbeiten
    data = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("Wählen Sie das Blatt aus", data.sheet_names)  # Blattauswahl ermöglichen
    df = data.parse(sheet_name)

    # Erweiterte Datenvalidierung
    required_columns = {"Artikel", "Woche", "Menge", "Name"}
    if not required_columns.issubset(df.columns):
        st.error("Fehler: Die Datei muss die Spalten 'Artikel', 'Woche', 'Menge' und 'Name' enthalten.")
    elif df.isnull().values.any():
        st.error("Fehler: Die Datei enthält fehlende Werte. Bitte stellen Sie sicher, dass alle Zellen ausgefüllt sind.")
    else:
        # Filter- und Suchmöglichkeiten
        st.sidebar.title("Filteroptionen")
        artikel_filter = st.sidebar.text_input("Nach Artikelnummer filtern (optional)")
        artikel_name_filter = st.sidebar.text_input("Nach Artikelname filtern (optional)")

        if artikel_filter:
            df = df[df['Artikel'].astype(str).str.contains(artikel_filter, case=False, na=False)]

        if artikel_name_filter:
            df = df[df['Name'].str.contains(artikel_name_filter, case=False, na=False)]

        # Durchschnittliche Abverkaufsmengen berechnen und Originalreihenfolge beibehalten
        result = df.groupby(['Artikel', 'Name'], sort=False).agg({'Menge': 'mean'}).reset_index()
        result.rename(columns={'Menge': 'Durchschnittliche Menge pro Woche'}, inplace=True)

        # Rundungsoptionen in der Sidebar für alle Artikel
        round_option = st.sidebar.selectbox(
            "Rundungsoption für alle Artikel:",
            ['Nicht runden', 'Aufrunden', 'Abrunden'],
            index=0
        )

        if round_option == 'Aufrunden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x + 0.5))
        elif round_option == 'Abrunden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x - 0.5))

        # Ergebnisse anzeigen
        st.subheader("Ergebnisse")
        st.dataframe(result)

        # Fortschrittsanzeige
        st.info("Verarbeitung abgeschlossen. Die Ergebnisse stehen zur Verfügung.")

        # Ergebnisse herunterladen
        output = BytesIO()
        result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button(
            label="Ergebnisse herunterladen",
            data=output,
            file_name="durchschnittliche_abverkaeufe.xlsx"
        )

        # Vergleich von Ergebnissen ermöglichen
        if st.checkbox("Vergleiche mit einer anderen Datei anzeigen"):
            uploaded_file_compare = st.file_uploader("Vergleichsdatei hochladen (Excel, .xlsx)", type=["xlsx"], key="compare")
            if uploaded_file_compare:
                compare_data = pd.read_excel(uploaded_file_compare)
                compare_result = compare_data.groupby(['Artikel', 'Name']).agg({'Menge': 'mean'}).reset_index()
                compare_result.rename(columns={'Menge': 'Durchschnittliche Menge pro Woche'}, inplace=True)

                # Ergebnisse der beiden Dateien nebeneinander anzeigen
                st.subheader("Vergleich der beiden Dateien")
                merged_results = result.merge(compare_result, on='Artikel', suffixes=('_Original', '_Vergleich'))
                st.dataframe(merged_results)

st.markdown("---")
st.markdown("⚠️ **Hinweis:** Diese Anwendung speichert keine Daten und hat keinen Zugriff auf Ihre Dateien.")
st.markdown("🌟 **Erstellt von Christoph R. Kaiser mit Hilfe von Künstlicher Intelligenz.**")
