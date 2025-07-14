import xlwings as xw
import pandas as pd
import warnings
import streamlit as st

## python -m streamlit run Leistungsbereich_Data_Assorter.py ## in Terminal

# Nulstil hele session_state ved første kørsel
if "initialized" not in st.session_state:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.session_state.initialized = True

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.max_colwidth', None)
warnings.filterwarnings("ignore", message="Workbook contains no default style")

# Titel og inputfelter (vises kun hvis ikke færdig)
if not st.session_state.get("completed", False):
    st.title("Vorlage: Finde die Vorlagendatei, rechtsklicken und „Als Pfad kopieren“.")
    st.text("Achtung: Diese Datei WIRD überschrieben – wenn du die Vorlage behalten willst, dupliziere sie vorher!")
    st.text_input("Vorlage Pfad", key="template_path")
    st.text(r'Beispiel: "C:\Users\hansen\Desktop\VORLAGE_XXX_Budgetbildung auf Grundlage Kobe_HP.xlsx"')

    st.title("Data: Finde die Quelldatei, rechtsklicken und „Als Pfad kopieren“.")
    st.text_input("Data Pfad", key="raw_path")

    # Hent og strip inputværdier fra session_state
    template_path = st.session_state.get("template_path", "").strip('"')
    raw_path = st.session_state.get("raw_path", "").strip('"')

    if template_path and raw_path:
        if st.button("Starte Verarbeitung"):
            st.session_state.run_process = True

        if st.session_state.get("run_process", False):
            st.text("Lädt...")  # Tekst vist mens processen kører
    else:
        st.stop()
else:
    st.title("Fertig!")
    st.text("Made by: Manne Bach Hansen")
    st.stop()

# Stop hvis ikke brugeren har trykket knappen
if not st.session_state.get("run_process"):
    st.stop()

# ========== PROCESSEN STARTER HER ==========

try:
    df_template = pd.read_excel(template_path, sheet_name=0)
    df_raw = pd.read_excel(raw_path)
except Exception as e:
    st.error(f"Fehler beim Einlesen der Dateien: {e}")
    st.stop()

df_raw = pd.read_excel(raw_path, header=10)
df_raw = df_raw.dropna(axis=1, how='all').dropna(axis=0, how='all').reset_index(drop=True)
df_raw = df_raw.drop(df_raw.columns[[3, 4, 5, 9, 10, 11]], axis=1)
df_raw['Leistungsbereich'] = pd.NA
df_raw.columns = ['LB', 'KGR', 'Bezeichnung', 'Menge', 'ME', 'EP', 'GB', 'Leistungsbereich']
df_raw = df_raw[~(df_raw['LB'].isna() & df_raw['KGR'].isna())].reset_index(drop=True)

# Fyld Leistungsbereich
col_source = df_raw.columns[0]
col_target = df_raw.columns[-1]
current_value = None
leistungsbereich_liste = []

for idx, row in df_raw.iterrows():
    value = row[col_source]
    if pd.notna(value) and not str(value).startswith("0"):
        continue
    if pd.notna(value):
        current_value = value
        bezeichnung = row[df_raw.columns[2]]
        leistungsbereich_liste.append([current_value, bezeichnung])
    df_raw.at[idx, col_target] = current_value

df_leistungsbereiche = pd.DataFrame(leistungsbereich_liste, columns=['Leistungsbereich', 'Bezeichnung'])

# Fyld KGR
current_value = None
for idx, row in df_raw.iterrows():
    value = row[col_source]
    if pd.notna(value) and not str(value).startswith("3"):
        continue
    if pd.notna(value):
        current_value = value
    df_raw.at[idx, df_raw.columns[1]] = current_value

# Rens inden indsættelse
df_raw = df_raw[df_raw[col_source].apply(lambda x: pd.isna(x))].reset_index(drop=True)
df_raw = df_raw[df_raw[df_raw.columns[5]] != 0].reset_index(drop=True)
df_raw.drop(df_raw.columns[0], axis=1, inplace=True)
df_raw = df_raw.fillna("").infer_objects(copy=False)

# xlwings
app = xw.App(visible=True)
book = app.books.open(template_path)
sheet = book.sheets[0]
format_sheet = book.sheets['Vorlage (DO NOT DELETE)']


def insert_dataframe_with_formatting(start_row, df):
    n_rows, n_cols = df.shape
    sheet.range((start_row, 1), (start_row + n_rows - 1, 1)).api.EntireRow.Insert()
    source_range = format_sheet.range((1, 1), (1, n_cols))
    target_range = sheet.range((start_row, 1), (start_row + n_rows - 1, n_cols))
    source_range.api.Copy()
    target_range.api.PasteSpecial(Paste=-4104)  # xlPasteFormats
    sheet.range((start_row, 1)).value = df.values.tolist()


insert_dataframe_with_formatting(start_row=4, df=df_raw)
sheet_format = book.sheets['Vorlage (DO NOT DELETE)']
sheet_format.range((5, 1)).value = df_leistungsbereiche.values.tolist()

book.save(template_path)
book.close()
app.quit()

# Marker som færdig og vis "Fertig!" direkte
st.session_state.completed = True
st.session_state.run_process = False

st.title("Fertig!")
st.text("Made by: Manne Bach Hansen")

# Tilføj knap for at genstarte
if st.button("Neustarten"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.experimental_rerun()

st.stop()
