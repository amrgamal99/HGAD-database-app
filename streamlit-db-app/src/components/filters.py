import streamlit as st
import pandas as pd
from db.connection import fetch_companies, fetch_projects_by_company

def create_company_dropdown(conn):
    """Creates a dropdown for company selection."""
    st.header("1. Select Company")
    companies_df = fetch_companies(conn)
    if companies_df.empty:
        st.warning("No companies found in the database.")
        return None

    company_list = [""] + companies_df['CompanyName'].tolist()

    # The selectbox now directly uses the full company list
    selected_company = st.selectbox("Select Company", company_list)

    return selected_company if selected_company else None

def create_project_dropdown(conn, company_name):
    """Creates a dropdown for project selection based on the company."""
    if not company_name:
        return None

    st.header("2. Select Project")
    projects_df = fetch_projects_by_company(conn, company_name)
    if projects_df.empty:
        st.warning(f"No projects found for {company_name}.")
        return None

    project_list = [""] + projects_df['ProjectTitle'].tolist()
    selected_project = st.selectbox("Select Project", project_list)
    return selected_project if selected_project else None

def create_type_dropdown():
    """Creates a dropdown to select the data type (Contract, Guarantee, etc.)."""
    st.header("3. Select Data Type")
    type_options = ["", "Contract", "Guarantee", "Checks", "Invoice"]
    selected_type = st.selectbox("Select Data Type", type_options)
    return selected_type if selected_type else None

def create_column_search(df):
    """Creates a dropdown to select a column and a text input to search it."""
    if df.empty:
        return None, None

    st.header("4. Search Results")
    columns = [""] + df.columns.tolist()
    search_column = st.selectbox("Select column to search", columns)
    search_term = st.text_input(f"Enter search term for '{search_column}'") if search_column else None

    return search_column, search_term