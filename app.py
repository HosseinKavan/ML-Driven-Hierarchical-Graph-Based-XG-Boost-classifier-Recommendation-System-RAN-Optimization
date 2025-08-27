# app.py
import streamlit as st
import pandas as pd
import os
import openpyxl
from io import BytesIO

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Network Sector Analysis Dashboard",
    page_icon="📡",
    layout="wide",
)

# --- HELPER FUNCTIONS ---

def find_table_start(df, header_text):
    """Finds the starting row index of a table based on its header text."""
    for index, row in df.iterrows():
        if header_text in row.to_string():
            return index
    return None

def find_table_end(df, start_search_from):
    """Finds the end of a table by looking for the first all-empty row."""
    for i in range(start_search_from, len(df)):
        if df.iloc[i].isnull().all():
            return i
    return len(df)

def extract_images_from_sheet(excel_path, sheet_name):
    """
    Extracts images from an Excel sheet based on their anchor cell locations.
    """
    images = {}
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook[sheet_name]
        
        for image in sheet._images:
            row = image.anchor._from.row
            col = image.anchor._from.col
            
            # Rule for Plot 1: Found at or after row 20 in Column A (index 0).
            if 'plot1' not in images and row >= 19 and col == 0:
                images['plot1'] = image.ref
            
            # Rule for Plot 2: Found at or after row 20 in Column L (index 11).
            elif 'plot2' not in images and row >= 19 and col == 11:
                images['plot2'] = image.ref

    except Exception as e:
        st.error(f"Failed to extract images from Excel: {e}")
        
    return images

def display_table_from_sheet(sheet_df, table_title):
    """Finds, cleans, and displays a table from a given DataFrame."""
    start_row = find_table_start(sheet_df, table_title)
    if start_row is not None:
        header_row = sheet_df.iloc[start_row + 1]
        data_start_row = start_row + 2
        table_end_row = find_table_end(sheet_df, data_start_row)
        
        table_df = sheet_df.iloc[data_start_row:table_end_row].copy()
        table_df.columns = header_row
        
        # Fix for duplicate columns
        seen = {}
        new_columns = []
        for col in table_df.columns.astype(str):
            if col not in seen:
                seen[col] = 1
                new_columns.append(col)
            else:
                count = seen[col]
                new_columns.append(f"{col}_{count}")
                seen[col] += 1
        table_df.columns = new_columns
        
        table_df = table_df.reset_index(drop=True)
        st.dataframe(table_df.dropna(how='all', axis=0).dropna(how='all', axis=1), use_container_width=True)
    else:
        if "Forbidden" in table_title:
            st.info(f"No '{table_title}' found for this sector.")
        else:
            st.warning(f"Could not find the table '{table_title}'.")

@st.cache_data
def load_data(excel_path):
    """
    Loads all sector sheets into a dictionary and the 'Introduction' sheet 
    into a separate DataFrame for filtering.
    """
    if not os.path.exists(excel_path):
        st.error(f"Error: The file '{excel_path}' was not found.")
        return None, None
    
    xls = pd.ExcelFile(excel_path)
    
    try:
        summary_df = xls.parse("Introduction", header=0) 
    except ValueError:
        st.error("Error: A sheet named 'Introduction' was not found in the Excel file.")
        return None, None

    all_sheet_names = xls.sheet_names
    sector_sheet_names = [name for name in all_sheet_names if name != "Introduction"]
    sector_data = {sheet_name: xls.parse(sheet_name, header=None) for sheet_name in sector_sheet_names}
    
    return sector_data, summary_df


# --- PASSWORD PROTECTION ---
def check_password():
    """Returns `True` if the user entered the correct password."""
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == "mtn123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

if check_password():
    # --- LOAD DATA ---
    EXCEL_FILE_PATH = 'Full_Sector_Recommendation_Report_Part_1_V5.7_Final.xlsx'

    sector_data, summary_df = load_data(EXCEL_FILE_PATH)

    # --- MAIN APP LOGIC ---
    if sector_data is None or summary_df.empty:
        st.stop()

    # --- SIDEBAR FOR FILTERS ---
    st.sidebar.header("Filters")

    sub_region_list = ["All"] + sorted(summary_df["LMBB Sub-Region"].dropna().unique().tolist())
    selected_sub_region = st.sidebar.selectbox("Select a Sub-Region", sub_region_list)

    gis_type_list = ["All"] + sorted(summary_df["GIS Type"].dropna().unique().tolist())
    selected_gis_type = st.sidebar.selectbox("Select a GIS Type", gis_type_list)

    filtered_df = summary_df.copy()

    if selected_sub_region != "All":
        filtered_df = filtered_df[filtered_df["LMBB Sub-Region"] == selected_sub_region]

    if selected_gis_type != "All":
        filtered_df = filtered_df[filtered_df["GIS Type"] == selected_gis_type]

    available_sectors = sorted(filtered_df["Congested Sector ID"].tolist())

    selected_sector = st.sidebar.selectbox(
        "Select a Congested Sector ID", 
        available_sectors, 
        help="This list is updated based on the filters above."
    )

    # --- MAIN PAGE VIEW WITH TABS ---
    st.title("📡 Network Congestion Analysis")
    st.markdown("---")

    tab_summary, tab_details = st.tabs(["📊 Region Summary", "📄 Sector Detail"])

    with tab_summary:
        st.header("Sector Count per Sub-Region")
        
        congestion_summary = summary_df.groupby("LMBB Sub-Region")["Congested Sector ID"].count().reset_index()
        congestion_summary = congestion_summary.rename(columns={"Congested Sector ID": "Number of Sectors"})

        st.bar_chart(congestion_summary, x="LMBB Sub-Region", y="Number of Sectors", color="#1f77b4")
        
        st.info("This chart shows the total number of sectors analyzed in each sub-region based on the master 'Introduction' list.")

    with tab_details:
        if selected_sector:
            st.header(f"Detailed Analysis for: `{selected_sector}`")
            
            if selected_sector in sector_data:
                sector_df = sector_data[selected_sector]
                
                sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📊 Main Report", "📶 Cellular Parameters", "🔄 Handover Parameters"])

                with sub_tab1:
                    st.subheader("Recommendation Report for Sector")
                    display_table_from_sheet(sector_df, "Recommendation Report for Sector")

                    st.subheader("Performance Plots")
                    extracted_images = extract_images_from_sheet(EXCEL_FILE_PATH, selected_sector)
        
                    col1, col2 = st.columns(2)
                    with col1:
                        if 'plot1' in extracted_images:
                            st.image(extracted_images['plot1'], caption="Plot 1", use_container_width=True)
                        else:
                            st.warning("Plot 1 could not be extracted from the Excel sheet.")
                    with col2:
                        if 'plot2' in extracted_images:
                            st.image(extracted_images['plot2'], caption="Plot 2", use_container_width=True)
                        else:
                            st.warning("Plot 2 could not be extracted from the Excel sheet.")

                with sub_tab2:
                    st.subheader("Cellular Parameters for Source and Top Neighbors")
                    display_table_from_sheet(sector_df, "Cellular Parameters for Source and Top Neighbors")

                with sub_tab3:
                    st.subheader("Non-Zero Cell Handover Parameter Details")
                    display_table_from_sheet(sector_df, "Non-Zero Cell Handover Parameter Details")
                    
                    st.subheader("Forbidden Relations Table")
                    display_table_from_sheet(sector_df, "Forbidden relations table")
            else:
                st.error(f"Data sheet for the selected sector '{selected_sector}' could not be found in the Excel file.")
        else:
            st.info("Select a sector from the sidebar to view its detailed report.")