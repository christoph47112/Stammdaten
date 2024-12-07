import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

def process_files(umsatz_file, stammdaten_file, output_file):
    # Umsatz- und Stammdaten einlesen
    umsatz_data = pd.read_excel(umsatz_file)
    stammdaten_data = pd.read_excel(stammdaten_file)

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

    # Pivot-Tabelle vorbereiten
    ws_pivot = wb.create_sheet(title="Pivot-Tabelle")
    ws_pivot["A1"] = "Hinweis: Gehe zu 'Daten' und erstelle eine Pivot-Tabelle in Excel."

    # Datei speichern
    wb.save(output_file)

# Streamlit App
st.title("Prüfung Kern- und Discount- Sortiment")

st.write("⚠️ Dieses Modul speichert **keine Daten**. Es wurde von Christoph R. Kaiser entwickelt, unterstützt durch modernste künstliche Intelligenz-Technologie. 🚀")

umsatz_file = st.file_uploader("Markt Daten hochladen (Excel)", type=["xlsx"])
stammdaten_file = st.file_uploader("Stammdatendatei hochladen (Excel)", type=["xlsx"])

if st.button("Verarbeiten"):
    if umsatz_file is not None and stammdaten_file is not None:
        output_file = "Artikel_Differenz_Ergebnis.xlsx"
        process_files(umsatz_file, stammdaten_file, output_file)
        
        with open(output_file, "rb") as file:
            st.download_button(
                label="Ergebnis herunterladen",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Bitte laden Sie beide Dateien hoch!")
