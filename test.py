import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import requests

# Stammdaten aus GitHub laden
def load_stammdaten():
    url = "https://raw.githubusercontent.com/christoph47112/Stammdaten/main/Stammdaten.xlsx"  # URL der Datei im GitHub-Repository

    try:
        response = requests.get(url)
        response.raise_for_status()  # Fehler auslösen, wenn Download fehlschlägt
        with open("stammdaten.xlsx", "wb") as file:
            file.write(response.content)  # Speichert die Datei lokal
        st.success("Die Stammdaten-Datei wurde erfolgreich integriert. Sie müssen nur die Markt Daten hochladen.")
    except requests.exceptions.RequestException as e:
        st.error("Fehler beim Laden der Stammdaten. Bitte überprüfen Sie den Repository-Link.")
        raise e

    stammdaten_data = pd.read_excel("stammdaten.xlsx")
    return stammdaten_data

def process_files(umsatz_file, stammdaten_data, output_file):
    # Umsatz- und Stammdaten einlesen
    umsatz_data = pd.read_excel(umsatz_file)

    # Sicherstellen, dass Artikelnummern als Strings behandelt werden
    umsatz_data['Artikel'] = umsatz_data['Artikel'].astype(str).str.strip()
    stammdaten_data['Artikel'] = stammdaten_data['Artikel'].astype(str).str.strip()

    # Filter: Artikel in der Stammdaten-Datei, die nicht in der Umsatz-Datei sind
    artikel_diff = stammdaten_data[~stammdaten_data['Artikel'].isin(umsatz_data['Artikel'])]

    # Entfernen aller Displays
    artikel_diff_no_displays = artikel_diff[artikel_diff['Artikeltyp'] != 'Display']

    # Neues Arbeitsbuch erstellen
    wb = Workbook()
    ws_data = wb.active
    ws_data.title = "Daten"

    # Daten in das Arbeitsblatt einfügen
    for r in dataframe_to_rows(artikel_diff_no_displays, index=False, header=True):
        ws_data.append(r)

    # Daten als Tabelle formatieren (erforderlich für Pivot-Funktionen)
    tab = Table(displayName="ArtikelDaten", ref=f"A1:E{len(artikel_diff_no_displays) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    ws_data.add_table(tab)

    # Datei speichern
    wb.save(output_file)

    return artikel_diff_no_displays

# Streamlit App
st.title("Prüfung Kern- und Discount- Sortiment")

st.write("Bitte laden Sie die Markt Daten hoch. Die Stammdaten-Datei ist bereits integriert.")

umsatz_file = st.file_uploader("Markt Daten hochladen (Excel)", type=["xlsx"])

if st.button("Verarbeiten"):
    if umsatz_file is not None:
        try:
            stammdaten_data = load_stammdaten()
            output_file = "Artikel_Differenz_Ergebnis.xlsx"
            artikel_diff_no_displays = process_files(umsatz_file, stammdaten_data, output_file)

            # Vorschau der gefilterten Daten
            st.subheader("Vorschau des Ergebnisses:")
            st.dataframe(artikel_diff_no_displays.head(50))  # Zeigt die ersten 50 Zeilen an
            
            with open(output_file, "rb") as file:
                st.download_button(
                    label="Ergebnis herunterladen",
                    data=file,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except FileNotFoundError as e:
            st.error(str(e))
    else:
        st.error("Bitte laden Sie die Markt Daten-Datei hoch!")

# Hinweise
st.markdown("---")
st.markdown("### Hinweise")
st.markdown("**Datenschutz:** Diese Anwendung speichert **keine Daten**. Alle hochgeladenen Dateien werden ausschließlich lokal verarbeitet.")
st.markdown("**Entwicklung:** Diese Anwendung wurde von Christoph R. Kaiser mithilfe künstlicher Intelligenz erstellt.")
