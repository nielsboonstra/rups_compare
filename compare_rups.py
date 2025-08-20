import streamlit as st
import pandas as pd
import regex as re

st.set_page_config(page_title="🐛 RUPS vergelijkingstool", layout="wide")

@st.dialog("Uitleg over de RUPS vergelijkingstool", width="large")
def explanation_dialog():
    st.markdown("### Input")
    st.markdown("Om deze tool te kunnen gebruiken, heb je twee Excel-bestanden nodig met RUPS-maatregelen.")
    st.markdown("Zorg ervoor dat de bestanden de juiste structuur hebben, met in ieder geval de volgende elementen:")
    st.markdown("- Een kolom voor Maatregelnr. (maatregelnummer)")
    st.markdown("- Een kolom met de Maatregelnaam")
    st.markdown("- Meerdere kolommen met jaartallen (b.v. 2024, 2025), met een 'X' in de rijen onder deze kolommen")
    st.markdown("ℹ️ _Het maakt niet uit of je Excel één of meerdere werkbladen heeft. Bij > 1 werkbladen zal de webapp je vragen om een werkblad te selecteren._")
    st.divider()
    st.markdown("### Output")
    st.markdown("Na het uploaden van de bestanden en het selecteren van de juiste kolommen, zal de tool één .xlsx-bestand genereren met de volgende 3 tabbladen:")
    st.markdown("- **Vergelijking**: Dit tabblad bevat de vergelijking tussen RUPS-maatregelen die zowel in de oude als de nieuwe planning staan, gebaseerd op maatregelnummer. De tool zal laten zien welke maatregelnamen zijn gewijzigd, en of het jaar waarin de maatregel gepland staat is gewijzigd.")
    st.markdown("- **Verwijderd**: Dit tabblad bevat de RUPS-maatregelen die zijn verwijderd in de nieuwe planning, maar wel aanwezig waren in de oude planning.")
    st.markdown("- **Toegevoegd**: Dit tabblad bevat de RUPS-maatregelen die zijn toegevoegd in de nieuwe planning, en niet aanwezig waren in de oude planning.")

@st.cache_data
def load_excel(file : str, header : int = 0, sheet_name: int = 0) -> pd.DataFrame:
    """
    Load an Excel file into a pandas DataFrame.

    :param file: The path to the Excel file.
    :param header: The row to use as the header.
    :param sheet_name: The name or index of the Excel-sheet to load.
    :return: A pandas DataFrame containing the Excel data.
    """
    try:
        df = pd.read_excel(file, engine='openpyxl', header=header, sheet_name=sheet_name, converters={"Maatregelnr.":str})
        return df
    except Exception as e:
        st.error(f"Fout bij het laden van het Excel-bestand: {e}")
        return None

def find_year_with_x(df: pd.DataFrame) -> pd.DataFrame:
    """
    The year is currently marked with an 'X' in the corresponding column (e.g., an X in column '2025' means the RUPS-maatregel is planned for 2025).
    This function finds the year for each RUPS-maatregel based on the 'X' marking, and transforms it into an integer value in column 'year'.

    :param df: The input DataFrame containing RUPS-maatregelen.
    :return: The modified DataFrame with the 'year' column added.
    """
    # Find the column (year) where the value is 'X' for each row
    df['year'] = df.apply(lambda row: next((int(col) for col in df.columns if row[col] == 'X'), None), axis=1)
    #drop original year int cols
    df = df.drop(columns=[col for col in df.columns if str(col).isdigit()])
    df['year'] = df['year'].fillna(0)
    # Convert the 'year' column to string type
    df['year'] = df['year'].astype(int)
    df['year'] = df['year'].astype(str)
    return df

def find_year_with_col(df : pd.DataFrame, year_column: str) -> pd.DataFrame:
    """
    Rename specified column to 'year'

    :param df: The input DataFrame.
    :param year_column: The name of the column to rename.
    :return: The modified DataFrame with the renamed column.
    """

    df = df.rename(columns={year_column: 'year'})
    return df

def remove_na_vals(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    """
    Remove NA values in the specified columns.

    :param df: The input DataFrame.
    :param cols: The columns to check for NA values.
    :return: The modified DataFrame with NA values removed.
    """
    for col in cols:
        if col in df.columns:
            # If empty or NA, drop
            df = df[df[col].notna()]
    return df

def strip_zeroes(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    """
    Strip leading zeroes from the specified columns.

    :param df: The input DataFrame.
    :param cols: The columns to strip leading zeroes from.
    :return: The modified DataFrame with leading zeroes stripped.
    """
    for col in cols:
        if col in df.columns:
            df[col] = df[col].str.lstrip("0")
    return df

def remove_old_cols(df: pd.DataFrame, cols_to_keep: list) -> pd.DataFrame:
    """
    Remove all columns with suffix _old from the dataframe, except for cols_to_keep. Rename columns with suffix _new to the original name, except for cols_to_keep.

    :param df: The input DataFrame.
    :param cols_to_keep: The columns to remain unchanged (i.e., do not remove or rename).
    :return: The modified DataFrame with _old columns removed and _new columns renamed.
    """
    cols_to_remove = [col for col in df.columns if col.endswith('_old') and col not in cols_to_keep]
    df = df.drop(columns=cols_to_remove)
    df = df.rename(columns={col: col[:-4] for col in df.columns if col.endswith('_new') and col not in cols_to_keep})
    return df

def compare_dataframes(df_old : pd.DataFrame, df_new: pd.DataFrame, cols: dict) -> tuple:
    """
    This is where the comparison between the old and new DataFrames happens after the user has provided all data and triggered the comparison.
    Two DataFrames are compared to identify differences in RUPS-maatregelen. The results are stored in three DataFrames: dropped, added, and kept (maatregelen).

    :param df_old: The older RUPS-maatregelen DataFrame (e.g., 2024).
    :param df_new: The newer RUPS-maatregelen DataFrame (e.g., 2025).
    :param cols: The column mappings that the user has provided as input.

    :return: A tuple containing the comparison DataFrame, dropped DataFrame, and added DataFrame.
    """
    # Fill NaN values with '-1' in the specified columns
    df_old = remove_na_vals(df_old, [cols['maatregelnr_col_old']])
    df_new = remove_na_vals(df_new, [cols['maatregelnr_col_new']])

    # Strip leading zeroes from relevant columns
    df_old = strip_zeroes(df_old, [cols['maatregelnr_col_old']])
    df_new = strip_zeroes(df_new, [cols['maatregelnr_col_new']])

    df_old_maatregelnr = df_old[cols['maatregelnr_col_old']].tolist()
    df_new_maatregelnr = df_new[cols['maatregelnr_col_new']].tolist()

    diff_dropped = list(set(df_old_maatregelnr) - set(df_new_maatregelnr))
    diff_added = list(set(df_new_maatregelnr) - set(df_old_maatregelnr))
    diff_kept = list(set(df_old_maatregelnr) & set(df_new_maatregelnr))

    df_dropped = df_old[df_old[cols['maatregelnr_col_old']].isin(diff_dropped)]
    df_added = df_new[df_new[cols['maatregelnr_col_new']].isin(diff_added)]
    df_kept = df_old[df_old[cols['maatregelnr_col_old']].isin(diff_kept)]

    # Create a comparison DataFrame
    df_compare = df_kept.merge(df_new, left_on=cols['maatregelnr_col_old'], right_on=cols['maatregelnr_col_new'], suffixes=('_old', '_new'), how='left')

    df_compare['year_changed'] = df_compare['year_old'] != df_compare['year_new']
    df_compare['maatregelnaam_changed'] = df_compare[cols['maatregelnaam_col_old'] + '_old'].str.lower() != df_compare[cols['maatregelnaam_col_new'] + '_new'].str.lower()
    df_compare['year_difference'] = df_compare['year_new'].astype(int) - df_compare['year_old'].astype(int)

    df_compare = remove_old_cols(df_compare, [cols['maatregelnaam_col_old'] + '_old', cols['maatregelnaam_col_new'] + '_new', "year_old", "year_new"])

    # Sort cols: Start with metadata (all other columns named hereafter). End with maatregelnr., then maatregelnaam (old and new), then year (old and new), then year_changed, maatregelnaam_changed and year_difference.
    final_order_cols = [cols['maatregelnr_col_old'],
                        cols['maatregelnaam_col_old'] + '_old', cols['maatregelnaam_col_new'] + '_new',
                        "year_old", "year_new", "year_changed", "maatregelnaam_changed", "year_difference"]
    metadata_cols = [col for col in df_compare.columns if col not in final_order_cols]
    df_compare = df_compare[metadata_cols + final_order_cols]

    st.session_state['df_compare'] = df_compare

    return df_compare, df_dropped, df_added

if 'df_old' not in st.session_state:
    st.session_state['df_old'] = None
if 'year_type_old' not in st.session_state:
    st.session_state['year_type_old'] = None
if 'year_type_new' not in st.session_state:
    st.session_state['year_type_new'] = None
if 'df_new' not in st.session_state:
    st.session_state['df_new'] = None
if 'comparison_type' not in st.session_state:
    st.session_state['comparison_type'] = None
if 'df_compare' not in st.session_state:
    st.session_state['df_compare'] = None

st.title("🐛 RUPS vergelijkingstool")

st.write("Deze tool helpt jou met het vergelijken van RUPS-planningen tussen twee verschillende jaren.")
st.write("Begin in de zijbalk met het uploaden van de RUPS-planningsbestanden.")

if st.button("Meer uitleg over deze tool", icon="ℹ️"):
    explanation_dialog()

with st.sidebar:
    st.markdown("**1. Upload een ouder RUPS-bestand (Excel-format) om te beginnen:** 👇")
    uploaded_file_old = st.file_uploader("Kies een bestand", type=["xlsx"], key="old_file_uploader")
    if uploaded_file_old is not None:
        year_type_old = st.pills("Hoe zijn de jaartallen in het Excel-bestand aangegeven?", ["Als 'X' onder een jaartal-kolom", "Jaartal is aangegeven als getal onder een Jaar-kolom"], key="year_old")
        st.session_state['year_type_old'] = year_type_old
        if st.session_state['year_type_old'] is not None:
            with st.expander("Pas de instellingen aan:", expanded=True):
                st.write("Standaard wordt het eerste werkblad uit de Excel ingeladen. Als je een ander werkblad wilt gebruiken, kies het dan hier:")
                sheet_names_old = pd.ExcelFile(uploaded_file_old).sheet_names if uploaded_file_old else []
                sheet_name = st.selectbox("Kies een werkblad:", sheet_names_old, key="sheet_name_old")
                st.write("Kies de rij waar de kolomnamen staan in je Excel-bestand.")
                header_row = st.number_input("Kies de rij waar de kolomnamen staan in je Excel. Inspecteer de data rechts in het hoofdscherm.", min_value=0, key="header_row_old")
                df_old = load_excel(uploaded_file_old, header=header_row, sheet_name=sheet_name)
                if st.session_state['df_old'] is None:
                    st.toast("Oud RUPS-bestand succesvol geüpload! Je kunt het nu inspecteren op de hoofdpagina.", icon = "🎉")
                st.session_state['df_old'] = df_old

    st.markdown("**2. Upload een nieuw RUPS-bestand (Excel-format):** 👇")
    uploaded_file_new = st.file_uploader("Kies een bestand", type=["xlsx"], key="new_file_uploader")
    if uploaded_file_new is not None:
        year_type_new = st.pills("Hoe zijn de jaartallen in het Excel-bestand aangegeven?", ["Als 'X' onder een jaartal-kolom", "Jaartal is aangegeven als getal onder een Jaar-kolom"], key="year_new")
        st.session_state['year_type_new'] = year_type_new
        if st.session_state['year_type_new'] is not None:
            with st.expander("Pas de instellingen aan:", expanded=True):
                st.write("Standaard wordt het eerste werkblad gebruikt uit de Excel ingeladen. Als je een ander werkblad wilt gebruiken, kies het dan hier:")
                sheet_names_new = pd.ExcelFile(uploaded_file_new).sheet_names if uploaded_file_new else []
                sheet_name = st.selectbox("Kies een werkblad:", sheet_names_new, key="sheet_name_new")
                st.write("Kies de rij waar de kolomnamen staan in je Excel-bestand. Inspecteer de data rechts in het hoofdscherm.")
                header_row = st.number_input("Kies de rij waar de kolomnamen staan in je Excel", min_value=0, key="header_row_new")
                df_new = load_excel(uploaded_file_new, header=header_row, sheet_name=sheet_name)
                if st.session_state['df_new'] is None:
                    st.toast("Nieuw RUPS-bestand succesvol geüpload! Je kunt het nu inspecteren op de hoofdpagina.", icon = "🎉")
                st.session_state['df_new'] = df_new

    # Check if both dataframes are loaded
    if st.session_state['df_old'] is not None and st.session_state['df_new'] is not None:
        st.success("Beide bestanden zijn succesvol geüpload! Je kunt nu de stappen rechts volgen om de conversie uit te voeren.")

#On the main page, give user the option to preview both uploaded files that can be hidden/collapsed by clicking on an arrow
if st.session_state['df_old'] is not None:
    with st.expander("📊 Bekijk de oudere RUPS-data", expanded=False):
        st.dataframe(st.session_state['df_old'])

if st.session_state['df_new'] is not None:
    with st.expander("📊 Bekijk de nieuwe RUPS-data", expanded=False):
        st.dataframe(st.session_state['df_new'])

# Ask user to select the columns that contain Maatregel naam and Maatregelnr.
if st.session_state['df_old'] is not None and st.session_state['df_new'] is not None:
    st.markdown("**3. Selecteer de kolommen die de Maatregel naam en Maatregelnr. bevatten:**")
    
    if st.session_state['year_type_old'] == "Als 'X' onder een jaartal-kolom":
        df_old = find_year_with_x(st.session_state['df_old'].copy())
    else:
        # Ask user which year column to use
        year_column_old = st.selectbox("Kies de jaar-kolom voor de oude RUPS-data:", st.session_state['df_old'].columns)
        df_old = find_year_with_col(st.session_state['df_old'].copy(), year_column_old)
    
    if st.session_state['year_type_new'] == "Als 'X' onder een jaartal-kolom":
        df_new = find_year_with_x(st.session_state['df_new'].copy())
    else:
        year_column_new = st.selectbox("Kies de jaar-kolom voor de nieuwe RUPS-data:", st.session_state['df_new'].columns)
        df_new = find_year_with_col(st.session_state['df_new'].copy(), year_column_new)
    cols = {}
    #Divide the page in 2 columns
    col1, col2 = st.columns(2)
    with col1:
        cols['maatregelnaam_col_old'] = st.selectbox("Oud - Maatregel naam:", df_old.columns)
        cols['maatregelnr_col_old'] = st.selectbox("Oud - Maatregelnr.:", df_old.columns)

    with col2:
        cols['maatregelnaam_col_new'] = st.selectbox("Nieuw - Maatregel naam:", df_new.columns)
        cols['maatregelnr_col_new'] = st.selectbox("Nieuw - Maatregelnr.:", df_new.columns)

    st.session_state['cols'] = cols
    st.markdown("**4. Klik op de knop hieronder om de vergelijking uit te voeren:**")
    if st.button("Voer vergelijking uit"):
        df_compare, df_dropped, df_added = compare_dataframes(df_old, df_new, cols)
        with pd.ExcelWriter('rups_vergelijking.xlsx') as writer:
            df_compare.to_excel(writer, sheet_name='Vergelijking', index=False)
            df_dropped.to_excel(writer, sheet_name='Verwijderd', index=False)
            df_added.to_excel(writer, sheet_name='Toegevoegd', index=False)
        st.success("Vergelijking uitgevoerd! Download de resultaten hieronder.")
        st.markdown("**5. Download de resultaten als Excel:**")
        st.download_button(
            label="Download Vergelijking",
            data=open('rups_vergelijking.xlsx', 'rb').read(),
            file_name='rups_vergelijking.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        st.info("Voor het interpreteren van tabblad 'Vergelijking' kan het helpen om conditional formatting toe te passen op de laatste drie kolommen in Excel.", icon="💡")
        st.balloons()