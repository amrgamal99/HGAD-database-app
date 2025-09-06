# Streamlit Database Application

This project is a Streamlit web application that connects to a database, allowing users to filter and display data based on various criteria. The application provides an interactive interface for selecting company names, project names, and types of documents (contracts, guarantees, checks, or invoices).

## Project Structure

```
streamlit-db-app
├── src
│   ├── app.py                # Main entry point of the Streamlit application
│   ├── db
│   │   └── connection.py     # Database connection logic
│   ├── components
│   │   └── filters.py        # Filter components for the Streamlit app
│   └── utils
│       └── data_helpers.py   # Utility functions for data manipulation
├── requirements.txt          # Project dependencies
└── README.md                 # Documentation for the project
```

## Setup Instructions

1. **Clone the repository:**
   ```
   git clone <repository-url>
   cd streamlit-db-app
   ```

2. **Install the required dependencies:**
   ```
   pip install -r requirements.txt
   ```

3. **Run the Streamlit application:**
   ```
   streamlit run src/app.py
   ```

## Usage Guidelines

- Upon launching the application, users will be presented with dropdown menus to select a company name, project name, and type of document.
- After making selections, the application will display a table with relevant data filtered according to the user's choices.

## Features

- Interactive dropdown menus for selecting:
  - Company names
  - Project names from the contract table
  - Document types (contract, guarantee, checks, or invoice)
- Dynamic data filtering and display based on user selections
- User-friendly interface built with Streamlit

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue for any suggestions or improvements.