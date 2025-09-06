import streamlit as st
import pandas as pd
from supabase import create_client, Client

# Define the function to get the Supabase client
@st.cache_resource
def get_db_connection():
    """Initializes and returns a Supabase client instance."""
    try:
        url = st.secrets["supabase_url"]
        key = st.secrets["supabase_key"]
        supabase_client: Client = create_client(url, key)
        print("Supabase client initialized successfully!")
        return supabase_client
    except Exception as e:
        print(f"Supabase client initialization failed: {e}")
        st.error(f"Supabase client initialization failed. Please check your secrets. Error: {e}")
        return None

# Function to fetch companies
def fetch_companies(supabase: Client):
    """Fetches all unique company names."""
    try:
        # Use lowercase table and column names for Supabase
        response = supabase.table('company').select('companyname').execute()
        df = pd.DataFrame(response.data)
        # Rename column for the UI dropdown which expects 'CompanyName'
        df = df.rename(columns={'companyname': 'CompanyName'})
        return df.drop_duplicates()
    except Exception as e:
        print(f"Error fetching companies: {e}")
        return pd.DataFrame()

# Function to fetch projects by company
def fetch_projects_by_company(supabase: Client, company_name):
    """Fetches projects for a given company name by first finding the company's ID."""
    if not company_name:
        return pd.DataFrame(columns=['ProjectTitle'])
    
    try:
        # Use lowercase table and column names
        company_response = supabase.table('company').select('companyid').eq('companyname', company_name).single().execute()
        if not company_response.data:
            return pd.DataFrame()
        company_id = company_response.data['companyid']

        # Use lowercase table and column names
        projects_response = supabase.table('contract').select('projecttitle').eq('companyid', company_id).execute()
        df = pd.DataFrame(projects_response.data)
        # Rename column for the UI dropdown which expects 'ProjectTitle'
        df = df.rename(columns={'projecttitle': 'ProjectTitle'})
        return df
    except Exception as e:
        print(f"Error fetching projects for {company_name}: {e}")
        return pd.DataFrame()

# Function to fetch data with filters
def fetch_data(supabase: Client, company_name=None, project_name=None, type_filter=None):
    """Fetches data from a specified table using company and project names for filtering."""
    if not all([company_name, project_name, type_filter]):
        st.error("Please provide all necessary filters: company name, project name, and type filter.")
        return pd.DataFrame()

    try:
        # Use lowercase table and column names
        company_response = supabase.table('company').select('companyid').eq('companyname', company_name).single().execute()
        if not company_response.data:
            return pd.DataFrame()
        company_id = company_response.data['companyid']

        contract_response = supabase.table('contract').select('contractid').eq('projecttitle', project_name).eq('companyid', company_id).single().execute()
        if not contract_response.data:
            return pd.DataFrame()
        contract_id = contract_response.data['contractid']

        # Convert the type_filter to lowercase to match the actual table name
        target_table = type_filter.lower()

        if target_table == 'contract':
            response = supabase.table('contract').select('*').eq('contractid', contract_id).execute()
        elif target_table in ['guarantee', 'invoice', 'checks']:
            response = supabase.table(target_table).select('*').eq('companyid', company_id).eq('contractid', contract_id).execute()
        else:
            st.error("Invalid type filter.")
            return pd.DataFrame()
        if target_table == 'contract':
            response = supabase.table('contract').select('*').eq('contractid', contract_id).execute()
        elif target_table in ['guarantee', 'invoice', 'checks']:
            response = supabase.table(target_table).select('*').eq('companyid', company_id).eq('contractid', contract_id).execute()
        else:
            st.error("Invalid type filter.")
            return pd.DataFrame()
            
        df = pd.DataFrame(response.data)
        
        # Exclude columns ending with 'id' before returning to the interface
        if not df.empty:
            cols_to_drop = [col for col in df.columns if col.lower().endswith('id')]
            df = df.drop(columns=cols_to_drop, errors='ignore')
            
        return df
    except Exception as e:
        print(f"An error occurred while fetching data: {e}")
        return pd.DataFrame()