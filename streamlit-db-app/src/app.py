import streamlit as st
import pandas as pd
from db.connection import get_db_connection, fetch_data
from components.filters import create_company_dropdown, create_project_dropdown, create_type_dropdown, create_column_search

def main():
    st.set_page_config(layout="wide")
    st.title("HGAD Database Viewer")

    conn = get_db_connection()
    if conn is None:
        st.error("Failed to connect to the database. Please check the connection details and ensure the server is running.")
        return

    # --- Sidebar for Filters ---
    with st.sidebar:
        st.title("Filters")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_selection = create_type_dropdown()

    # --- Main Content Area ---
    if not company_name or not project_name or not type_selection:
        st.info("Please select a Company, Project, and Data Type from the sidebar to view data.")
        return

    # Fetch data based on initial filters
    df = fetch_data(conn, company_name, project_name, type_selection)

    if df.empty:
        st.warning("No data found for the selected filters.")
    else:
        # Column search filter
        search_column, search_term = create_column_search(df)

        # Apply search term filter if provided
        if search_column and search_term:
            # Ensure we search on string representations of the data
            df = df[df[search_column].astype(str).str.contains(search_term, case=False, na=False)]

        # --- New code to make links clickable ---
        # Create a dictionary to hold column configurations
        column_config = {}
        # Find all columns that end with 'link'
        for col in df.columns:
            if col.lower().endswith('link'):
                # Configure the column as a LinkColumn
                column_config[col] = st.column_config.LinkColumn(
                    label=col,  # Use the original column name as the header
                    display_text="Open Link"  # The text that will be displayed in the cell
                )
        
        # Display the dataframe with the new link configuration
        st.dataframe(df, column_config=column_config, use_container_width=True)

if __name__ == "__main__":
    main()