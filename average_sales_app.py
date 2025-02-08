import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Durchschnittliche Abverkaufsmengen", layout="wide")

st.title("Berechnung der Ø Abverkaufsmengen pro Woche")

uploaded_file = st.file_uploader("Bitte laden Sie Ihre Datei hoch (Excel)", type=["xlsx"])

if uploaded_file:
    data = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("Wählen Sie das Blatt aus", data.sheet_names)
    df = data.parse(sheet_name)
    
    required_columns = {"Artikel", "Woche", "Menge", "Name"}
    if not required_columns.issubset(df.columns):
        st.error("Die Datei muss die Spalten 'Artikel', 'Woche', 'Menge' und 'Name' enthalten.")
    else:
        artikel_filter = st.sidebar.text_input("Nach Artikel filtern (optional)")
        artikel_name_filter = st.sidebar.text_input("Nach Artikelname filtern (optional)")
        
        if artikel_filter:
            df = df[df['Artikel'].astype(str).str.contains(artikel_filter, case=False, na=False)]
        if artikel_name_filter:
            df = df[df['Name'].str.contains(artikel_name_filter, case=False, na=False)]
        
        result = df.groupby(['Artikel', 'Name'], sort=False).agg({'Menge': 'mean'}).reset_index()
        result.rename(columns={'Menge': 'Durchschnittliche Menge pro Woche'}, inplace=True)
        
        round_option = st.sidebar.selectbox("Rundungsoption:", ['Nicht runden', 'Aufrunden', 'Abrunden'], index=0)
        if round_option == 'Aufrunden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x + 0.5))
        elif round_option == 'Abrunden':
            result['Durchschnittliche Menge pro Woche'] = result['Durchschnittliche Menge pro Woche'].apply(lambda x: round(x - 0.5))
        
        st.subheader("Ergebnisse")
        st.dataframe(result)
        
        output = BytesIO()
        result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button(label="Ergebnisse herunterladen", data=output, file_name="durchschnittliche_abverkaeufe.xlsx")
