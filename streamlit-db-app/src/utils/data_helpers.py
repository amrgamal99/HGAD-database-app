def format_data_for_display(data):
    # Function to format data for display in the Streamlit table
    formatted_data = data.copy()
    # Example formatting: Convert date columns to a more readable format
    if 'date' in formatted_data.columns:
        formatted_data['date'] = formatted_data['date'].dt.strftime('%Y-%m-%d')
    return formatted_data

def filter_data_by_company(data, company_name):
    # Function to filter data by selected company name
    return data[data['company_name'] == company_name]

def filter_data_by_project(data, project_name):
    # Function to filter data by selected project name
    return data[data['project_name'] == project_name]

def filter_data_by_type(data, selected_type):
    # Function to filter data by selected type
    return data[data['type'] == selected_type]

def prepare_data_for_table(data, company_name, project_name, selected_type):
    # Function to prepare data for the table based on selected filters
    if company_name:
        data = filter_data_by_company(data, company_name)
    if project_name:
        data = filter_data_by_project(data, project_name)
    if selected_type:
        data = filter_data_by_type(data, selected_type)
    
    return format_data_for_display(data)