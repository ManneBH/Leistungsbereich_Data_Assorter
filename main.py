import streamlit as st
import tempfile
import os
from processing import run_processing

# # Configure Streamlit page
# st.set_page_config(page_title="LV Datenaufbereiter", layout="centered")

# Page title and instructions
st.title("TITEL")
st.write("Bitte lade die Excel-Vorlage und die Rohdaten-Datei hoch.")

# Upload inputs
template_file = st.file_uploader("Excel-Vorlage hochladen", type=["xlsx"])
raw_file = st.file_uploader("Rohdaten-Datei hochladen", type=["xlsx"])

# Run processing once both files are uploaded and button is clicked
if template_file and raw_file:
    if st.button("Verarbeiten"):
        # Save uploaded template to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_template:
            tmp_template.write(template_file.read())
            template_path = tmp_template.name

        # Save uploaded raw data to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_raw:
            tmp_raw.write(raw_file.read())
            raw_path = tmp_raw.name

        # Run the main processing function
        run_processing(template_path, raw_path)

        # Read the processed file for download
        with open(template_path, "rb") as f:
            processed_data = f.read()

        st.write("Erstellt von Manne B. Hansen")

        # Provide download button to user
        st.download_button(
            label="Verarbeitete Datei herunterladen",
            data=processed_data,
            file_name="JJMMTT_XXX_Budgetbildung auf Grundlage Kobe_HP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Clean up temporary files
        os.remove(template_path)
        os.remove(raw_path)
