####################################################################################################
# 2025_sdfile_details_extract.py
# SLA Extraction from SD PDFs
####################################################################################################

# ================================================================================================
# Overview:
# This script automates the extraction of key SLA (Service Level Agreement) details from PDF-based
# Service Descriptions (SDs). It processes multiple PDF files, extracts structured data, and
# consolidates it into an Excel file for reporting and analysis.
# ================================================================================================

# ================================================================================================
# Technical Requirements:
# - Python 3.x
# - Required Libraries: pandas, os, re, pdfplumber, openpyxl
# - Ensure `pdfplumber` is installed for PDF text extraction using:
#
#   pip install pandas pdfplumber openpyxl
# ================================================================================================

# ================================================================================================
# Data Requirements:
# - The SD PDFs must contain structured tables for:
#   - BSN Details
#   - Incident Response & Resolution Time
#   - Service Availability
#   - Disaster Recovery Classes (DRC)
#   - Support Hours
#
# - PDF files must be stored in the "Database" folder:
#   C:\Users\rmya5fe\OneDrive - Allianz\01_Automated Reports\07_Sample_SDs\Database
# ================================================================================================

# ================================================================================================
# Pre-Processing Checks:
# - Verify that all SD PDFs are in the correct folder.
# - Ensure tables in the PDFs follow expected formats.
# - Confirm Excel file (`SLA_extract_from_SD.xlsx`) exists or will be created automatically.
# ================================================================================================

# ================================================================================================
# How to Run:
# 1. Place all SD PDF files in the `Database` folder.
# 2. Execute the script.
#    - python 2025_sdfile_details_extract.py
# 3. It will:
#    - Process each PDF.
#    - Extract SLA details (BSN, Incidents, Service Availability, DRC, Support Hours).
#    - Store results in an Excel file (`SLA_extract_from_SD.xlsx`).
# 3. Check the output Excel file for extracted data.
# ================================================================================================

# ================================================================================================
# Output:
# - Extracted SLA details are saved in `SLA_extract_from_SD.xlsx` under the same directory.
# - The output includes key details such as:
#   - BSN Number
#   - Material & Availability
#   - Incident Response & Resolution Times
#   - Disaster Recovery Classes (DRC)
#   - Support Hour Details
# ================================================================================================

import os
import re
import warnings

import camelot
import pandas as pd
import pdfplumber
from PyPDF2 import PdfReader

warnings.filterwarnings("ignore")


# ====================================
# Extracting BSN and Version details
# ====================================

# Function to extract table from a single PDF
def extract_bsn_table_from_pdf(pdf_path):
    try:
        # print(f"\nExtracting BSN table from '{os.path.basename(pdf_path)}'...")
        tables = camelot.read_pdf(pdf_path, pages="1", flavor="lattice")
        if tables:
            return tables[0].df  # Return the first table as a DataFrame
        else:
            print(f"\nNo tables found in '{pdf_path}'.")
            return None
    except Exception as e:
        print(f"\nError processing '{pdf_path}': {e}")
        return None


# Function to extract "Service Description ID" from a table
def extract_bsn_number_from_table(table_df):
    if table_df is not None:
        table_df.columns = ["Name", "Value"]  # Rename columns
        table_df = table_df.map(lambda x: str(x).replace(" ", "").strip() if pd.notna(x) else x)  # Remove extra spaces
        if "ServiceDescriptionID:" in table_df["Name"].values:
            bsn_value = table_df.loc[table_df["Name"] == "ServiceDescriptionID:", "Value"].values[0]

            # Check if the value already starts with 'BSN'
            if not bsn_value.startswith("BSN"):
                bsn_value = "BSN" + bsn_value

            # print(f"Extracted BSN Number: {bsn_value}")
            return bsn_value
        else:
            print("Service Description ID not found in table.\n")
    return None


# =====================================
# Function to match regex search text
# =====================================

# Function to normalize text for accurate searching
def normalize_text(text):
    return re.sub(r'\s+', ' ', text.strip())


# ==================================
# Extracting the Index page number
# ==================================

# Function to get the index page number
def find_index_page_number(pdf_path):
    # Open the PDF
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    # Initialize variables
    index_page_number = None

    # Regex patterns
    list_of_tables_pattern = re.compile(r'List of Tables|List of Tables and Figures', re.IGNORECASE)

    # Detect and skip the 'List of Tables' page
    for page_num in range(total_pages):
        page_text = reader.pages[page_num].extract_text()
        if list_of_tables_pattern.search(page_text):
            index_page_number = page_num + 1  # Start search after this page
            break

    if index_page_number is not None:
        return index_page_number


# ===========================================
# Extracting list of pages from search text
# ===========================================

# Function to find all the occurrences of service availability text
def find_all_service_availability_and_support_hour_pages(pdf_path, search_text):
    occurrences = []  # List to store all occurrences
    compiled_pattern = re.compile(search_text, re.IGNORECASE)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):  # Pages are 1-indexed
                text = page.extract_text()
                if text:  # Ensure the page contains text
                    normalized_text = normalize_text(text)  # Normalize text

                    if compiled_pattern.search(normalized_text):
                        occurrences.append(page_number)  # Store the page number

        # If no occurrences found, return an empty list
        if not occurrences:
            print(f"\nNo occurrences of '{search_text}' found in '{os.path.basename(pdf_path)}'.")

    except Exception as e:
        print(f"\nError processing '{os.path.basename(pdf_path)}': {e}")

    return occurrences  # Return the full list of occurrences


# Function to find all the occurrences of service availability text
def find_all_run_of_service_pages(pdf_path, search_text):
    occurrences = []  # List to store all occurrences
    compiled_pattern = re.compile(search_text, re.IGNORECASE)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):  # Pages are 1-indexed
                text = page.extract_text()
                if text:  # Ensure the page contains text
                    normalized_text = normalize_text(text).replace(' ', '')  # Normalize text

                    if compiled_pattern.search(normalized_text):
                        occurrences.append(page_number)  # Store the page number

        # If no occurrences found, return an empty list
        if not occurrences:
            print(f"\nNo occurrences of '{search_text}' found in '{os.path.basename(pdf_path)}'.")

    except Exception as e:
        print(f"\nError processing '{os.path.basename(pdf_path)}': {e}")

    return occurrences  # Return the full list of occurrences


# =================================================
# Extracting SA Material and Availability details
# =================================================

# Function to extract data from the material tables
def extract_data_from_material_tables(pdf_path, page_numbers):
    extracted_data = {}  # Dictionary to store key-value pairs
    extracted_dataframes = []

    try:
        for page_number in page_numbers:
            page_number_str = str(page_number)  # Camelot requires page numbers as a string

            # Extract tables using Camelot (lattice method for structured tables)
            tables = camelot.read_pdf(pdf_path, pages=page_number_str, flavor='lattice', line_scale=50)

            if not tables or tables.n == 0:
                return extracted_data

            else:
                for i in range(tables.n):
                    df = tables[i].df  # Convert to DataFrame
                    df = df.replace('\n', ' ', regex=True)  # Clean newlines
                    df = df.applymap(
                        lambda x: x.strip().replace("“", "").replace("”", "").replace('"', ''))  # Normalize Text

                    # Normalize column headers by removing hidden quotes & spaces
                    cleaned_headers = [col.strip().replace("“", "").replace("”", "").replace('"', '') for col in
                                       df.iloc[0].values]
                    cleaned_headers = [
                        re.search(r'\bService Availability\b', col).group(0) if re.search(r'\bService Availability\b',
                                                                                          col) else col for col in
                        cleaned_headers]

                    if "Service Availability" in cleaned_headers:
                        # Make second row as the header
                        df.columns = cleaned_headers  # Assign new headers
                        df = df[1:].reset_index(drop=True)  # Drop the first row
                        extracted_dataframes.append(df)

                        num_columns = df.shape[1]  # Number of columns in the dataframe

                        # Logic for 2 columns table
                        if num_columns == 2:
                            val1 = df.iloc[0, 0]
                            val2 = df.iloc[1, 0]

                            if re.search(r'\b\d{6}\b', val2):
                                key = val1 + ' ' + val2
                            else:
                                key = val1
                            # key = df.iloc[0, 0]
                            value = df.iloc[1, 1]

                            cleaned_value = re.sub(r'\bPI\s*(?=\d|[^\w\s])', 'KPI', value,
                                                   flags=re.IGNORECASE).strip()  # Convert PI to KPI first
                            cleaned_value = re.sub(r'\b[A-Za-z]\b', '',
                                                   cleaned_value).strip()  # Remove single characters
                            cleaned_value = re.sub(r'(?<=\d) (?=\d)', '',
                                                   cleaned_value).strip()  # Remove spaces between numbers

                            # If value is empty, check the next available row dynamically
                            if not cleaned_value:
                                for j in range(1, len(df)):  # Iterate through remaining rows
                                    temp_value = df.iloc[j, 1].strip()

                                    if temp_value:  # If a valid value is found, use it
                                        cleaned_value = temp_value
                                        break

                            cleaned_value = re.sub(r' {2,}', ' ', cleaned_value)
                            key = key.replace("\n", "").replace("  ", " ").replace("   ", " ").strip()

                            if key in extracted_data:
                                # Convert existing value to a list if it's a string
                                if isinstance(extracted_data[key], str):
                                    extracted_data[key] = [extracted_data[key]]  # Convert string to list

                                extracted_data[key].append(cleaned_value)  # Append new value to the list
                            else:
                                extracted_data[key] = cleaned_value  # Store first value as a string


                        # Logic for single-column tables
                        elif num_columns == 1:
                            key = df.iloc[0, 0]

                            # Find the row containing "Service Level Target Value" or similar keywords
                            value = ""
                            for row in df.iloc[:, 0]:  # Iterate over the single column
                                match = re.search(
                                    r"(Service Level Target Value|SL Target Value|Target Value)\s*[:,]?\s*(=\s*\d+[.,]?\d*\s*%)",
                                    row, re.IGNORECASE)
                                if match:
                                    value = match.group(2).strip()  # Extract the percentage value
                                    break  # Stop after finding the first match

                            key = key.replace("\n", "").replace("  ", " ").replace("   ", " ").strip()

                            if key in extracted_data:
                                # Convert existing value to a list if it's a string
                                if isinstance(extracted_data[key], str):
                                    extracted_data[key] = [extracted_data[key]]  # Convert string to list

                                extracted_data[key].append(value)  # Append new value to the list
                            else:
                                extracted_data[key] = value  # Store first value as a string

    except Exception as e:
        print(f"\nError processing the PDF file: {e}")

    return extracted_data


# ========================================================
# Extracting Incident Response & Resolution time details
# ========================================================

# Function to find the first or second occurrence of the text
def find_incident_table_page_number(pdf_path, search_text):
    occurrences = []  # Track pages where the search text is found
    compiled_pattern = re.compile(search_text, re.IGNORECASE)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):  # Pages are 1-indexed
                text = page.extract_text()
                if text:  # Ensure the page contains text
                    normalized_text = normalize_text(text)  # Normalize text

                    match = compiled_pattern.search(normalized_text)
                    if match:
                        occurrences.append(page_number)

                        # Stop when the second occurrence is found
                        if len(occurrences) == 2:
                            return page_number

            # Handle the case where there is only one occurrence
            if len(occurrences) == 1:
                return occurrences[0]

        # print(f"\nNo occurrences of DRC search text found in '{os.path.basename(pdf_path)}'.")

    except Exception as e:
        print(f"\nError processing '{os.path.basename(pdf_path)}': {e}")

    return None


# Function to convert dataframe into list
def convert_df_into_list(extracted_dataframe):
    if extracted_dataframe is not None:
        # headers = [col.replace("\n", "").lower().strip() for col in extracted_dataframe.columns]

        # **Remove P1, P2, P3, P4 Row (Second Row in Most Cases)**
        if extracted_dataframe.shape[0] > 1:
            first_col_values = extracted_dataframe.iloc[:, 0].astype(str).str.lower().str.strip().str.replace("\n", "",
                                                                                                              regex=True)
            if first_col_values.iloc[0] == "nan":  # Classification row detected
                extracted_dataframe = extracted_dataframe.drop(index=0).reset_index(drop=True)

        # Initialize Empty Lists
        response_time_list, resolution_time_list = [], []

        # Identify Available Rows in the First Column
        first_col_values = extracted_dataframe.iloc[:, 0].astype(str).str.lower().str.strip().str.replace("\n", "",
                                                                                                          regex=True)

        # If "Incident Response Time" is present, extract values
        if "incident response time" in first_col_values.values:
            response_index = first_col_values[first_col_values == "incident response time"].index[0]
            response_time_list = [
                re.sub(r' {2,}', ' ', val.replace("\n", "").replace("•", "").replace("\uf0b7", "").strip())
                for val in extracted_dataframe.iloc[response_index, 1:]]

        # If "Incident Resolution Time" is present, extract values
        if "incident resolution time" in first_col_values.values:
            resolution_index = first_col_values[first_col_values == "incident resolution time"].index[0]
            resolution_time_list = [
                re.sub(r' {2,}', ' ', val.replace("\n", "").replace("•", "").replace("\uf0b7", "").strip())
                for val in extracted_dataframe.iloc[resolution_index, 1:]]

    else:
        response_time_list, resolution_time_list = [], []

    return response_time_list, resolution_time_list


# Function to extract all the tables from the target page
def extract_all_tables_from_incident_page(pdf_path, page_number):
    try:

        expected_headers = ["classification", "incident response time", "incident resolution time",
                            "incident classification", "response time", "resolution time"]
        row_keywords = ["incident resolution time", "incident response time", "response time", "resolution time"]

        # Camelot requires page numbers as a string
        page_number_str = str(page_number)

        # Extract tables using Camelot (lattice method for structured tables)
        tables = camelot.read_pdf(pdf_path, pages=page_number_str, flavor='lattice', line_scale=50)

        if not tables or tables.n == 0:
            print("\nNo tables found on the page.")
            return None

        # Convert each table to a DataFrame
        dataframes = []
        for i, table in enumerate(tables):
            try:
                # Convert table to DataFrame
                df = table.df  # Camelot returns tables as pandas DataFrames
                df.columns = [col.strip() for col in df.iloc[0]]  # Set headers
                df = df[1:]  # Remove header row from data
                df = df.reset_index(drop=True)  # Reset index

                # Drop empty columns
                df = df.dropna(how="all", axis=1)

                if not df.empty:
                    dataframes.append(df)
                else:
                    print(f"\nTable {i + 1} is empty after cleaning. Skipping...\n")

            except Exception as e:
                print(f"\nError processing Table {i + 1}: {e}")

        # If no valid table found, return None
        if not dataframes:
            return None

        for df in dataframes:
            # Normalize headers: Convert to lowercase and strip spaces
            headers = [str(col).replace("\n", "").strip().lower() for col in df.columns]
            cleaned_row_keywords = [str(val).replace("\n", "").strip().lower() for val in row_keywords]

            # Step 1: Check if any expected header is in table headers
            if any(keyword in headers for keyword in expected_headers):
                return df  # Return the first valid table

            # Step 2: If no matching header, check the rows for keywords
            for _, row in df.iterrows():
                row_text = " ".join(map(str, row.values)).replace("\n", "").strip().lower()
                # print(row_text)  # Debugging to check the cleaned row text

                if any(keyword in row_text for keyword in cleaned_row_keywords):
                    return df  # Return table if a row contains the keywords

            return None  # Return None if no matching table is found

    except Exception as e:
        print(f"\nError processing the PDF file: {e}")
        return None


# ========================
# Extracting DRC details
# ========================

# Function to find the first or second occurrence of the text
def find_drc_table_page_number(pdf_path, search_text):
    occurrences = []  # Track pages where the search text is found
    compiled_pattern = re.compile(search_text, re.IGNORECASE)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):  # Pages are 1-indexed
                text = page.extract_text()
                if text:  # Ensure the page contains text
                    normalized_text = normalize_text(text)  # Normalize text

                    match = compiled_pattern.search(normalized_text)
                    if match:
                        occurrences.append(page_number)

                        # Stop when the second occurrence is found
                        if len(occurrences) == 2:
                            return page_number

            # Handle the case where there is only one occurrence
            if len(occurrences) == 1:
                return occurrences[0]

        # print(f"\nNo occurrences of DRC search text found in '{os.path.basename(pdf_path)}'.")

    except Exception as e:
        print(f"\nError processing '{os.path.basename(pdf_path)}': {e}")

    return None


# Function to convert dataframe into dictionary
def convert_df_to_dict(dataframes):
    extracted_data = {}  # Final dictionary to store results

    # Define the expected column names for different formats
    required_formats = [
        ["Applicable DRC", "Applicable RPO"],
        ["Applicable DRC", "Applicable RTO/RPO"],
        ["Applicable DRCI", "Applicable RPO"],
        ["Applicable DRCI", "Applicable RTO/RPO"]
    ]

    for df in dataframes:
        # Ensure DataFrame is not empty
        if df is None or df.empty:
            continue  # Skip empty DataFrames

        # Clean column names
        df.columns = [re.sub(r'\s{2,}', ' ', col.replace("\n", " ")).strip() for col in df.columns]

        # Processing Format 1
        if len(df.columns) == 3 and list(df.columns[1:]) in required_formats:
            for _, row in df.iterrows():
                key = row.iloc[0].replace("\n", "").strip()

                drc_raw = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                rpo_raw = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""

                drc_values = [chunk.strip() for chunk in drc_raw.split("\n") if chunk.strip()]
                rpo_values = [chunk.strip() for chunk in rpo_raw.split("\n") if chunk.strip()]

                # Initialize dictionary entry if key is new
                if key not in extracted_data:
                    extracted_data[key] = {"Applicable DRC": [], "Applicable RPO": []}

                extracted_data[key]["Applicable DRC"].extend(drc_values)
                extracted_data[key]["Applicable RPO"].extend(rpo_values)

        # Processing Format 2
        elif (df.columns[0] == "" or
              all(df.iloc[:, 1:].applymap(
                  lambda x: bool(re.search(r'\bYES\b|\bNO\b|\bNA\b|\bN/A\b', str(x), re.IGNORECASE))
              ).all(axis=1))):

            for _, row in df.iterrows():
                key = row.iloc[0].replace("\n", "").strip()
                key = re.sub(r' {2,}', ' ', key)

                # Initialize dictionary entry if key is new
                if key not in extracted_data:
                    extracted_data[key] = {"Applicable DRC": [], "Applicable RPO": []}

                # Iterate over columns to extract values
                for col, val in zip(df.columns[1:], row.iloc[1:]):
                    normalized_col = col.replace("\n", " ").strip()
                    if "yes" in str(val).strip().lower():
                        if "DRC" in normalized_col or "EDR" in normalized_col:
                            extracted_data[key]["Applicable DRC"].append(normalized_col)
                        elif "RPO" in normalized_col:
                            extracted_data[key]["Applicable RPO"].append(normalized_col)

    return extracted_data  # Returns a valid dictionary ({} if no matches found)


# Function to extract all the tables from the target page
def extract_all_tables_from_drc_page(pdf_path, page_number):
    try:
        # Define the expected headers for the relevant table
        possible_last_columns = ["Applicable RPO", "Applicable RTO/RPO"]
        required_first_column = "Applicable DRC"

        # Camelot requires page numbers as a string
        page_number_str = str(page_number)

        # Extract tables using Camelot (lattice method for structured tables)
        tables = camelot.read_pdf(pdf_path, pages=page_number_str, flavor='lattice', line_scale=50)

        if not tables or tables.n == 0:
            print("\nNo tables found on the page.")
            return None

        # Convert each table to a DataFrame
        dataframes = []
        for i, table in enumerate(tables):
            try:
                # Convert table to DataFrame
                df = table.df  # Camelot returns tables as pandas DataFrames
                df.columns = [col.strip() for col in df.iloc[0]]  # Set headers
                df = df[1:]  # Remove header row from data
                df = df.reset_index(drop=True)  # Reset index
                df = df.dropna(how="all", axis=1)  # Drop empty columns

                if not df.empty:
                    dataframes.append(df)
                else:
                    print(f"\nTable {i + 1} is empty after cleaning. Skipping...\n")

            except Exception as e:
                print(f"\nError processing Table {i + 1}: {e}")

        # If no valid table found, return None
        if not dataframes:
            return None

        # Check for split table by inspecting the previous page
        combined_dataframes = []
        for target_df in dataframes:
            # Ensure the last two columns match the required headers
            if target_df.shape[1] >= 2 and target_df.columns[-2] == required_first_column and target_df.columns[
                -1] in possible_last_columns:
                # Check the previous page
                previous_page = page_number - 1
                previous_page_str = str(previous_page)
                previous_tables = camelot.read_pdf(pdf_path, pages=previous_page_str, flavor='lattice', line_scale=50)

                if previous_tables.n > 0:  # Ensure tables exist
                    for prev_table in previous_tables:
                        try:
                            df_prev = prev_table.df
                            if df_prev.shape[0] > 1:  # Ensure at least one row
                                df_prev.columns = [col.strip() for col in df_prev.iloc[0]]  # Set headers
                                df_prev = df_prev[1:].reset_index(drop=True)  # Remove header row
                                df_prev = df_prev.dropna(how="all", axis=1)  # Drop empty columns

                                # Match headers with the target DataFrame
                                if df_prev.shape[1] >= 2 and df_prev.columns[-2] == required_first_column and \
                                        df_prev.columns[-1] in possible_last_columns:
                                    # Combine previous page's table with target page's table
                                    combined_table = pd.concat([df_prev, target_df], ignore_index=True)
                                    combined_dataframes.append(combined_table)
                                    continue  # Move to the next table

                        except Exception as e:
                            print(f"\nError processing table from the previous page: {e}")

                # If no split table was found, store the target_df separately
                combined_dataframes.append(target_df)

        # If combined_dataframes has valid data, return it; otherwise, return all extracted tables
        return combined_dataframes if combined_dataframes else dataframes

    except Exception as e:
        print(f"\nError processing the PDF file: {e}")
        return []


# ===========================================================
# Extracting Support Hour details from Service Timing table
# ===========================================================

# Function to extract the support hour value
def extract_support_hours(df):
    # Dictionary to store extracted values
    support_dict = {
        "1st Level Support": "",
        "2nd Level Support": "",
        "Emergency Support": "",
        "Non-Emergency Support": ""
    }

    if df.empty or df.shape[1] < 2:  # Ensure valid DataFrame with at least 2 columns
        return support_dict  # Return empty structure if input is invalid

    # Normalize the first column (remove spaces, convert to lowercase, but keep special characters)
    df.iloc[:, 0] = df.iloc[:, 0].astype(str).fillna("").str.replace("\n", "").str.lower()

    # Define keyword mappings to corresponding dictionary keys
    keyword_map = {
        rf"(?<![\w-]){re.escape('non-emergency')}(?![\w-])": "Non-Emergency Support",
        rf"(?<![\w-]){re.escape('emergency')}(?![\w-])": "Emergency Support",
        r"\b1st\s*level\b|\b1st\s*\+?\s*level\s*support\b": "1st Level Support",
        r"\b2nd\s*level\b|\b2nd\s*\+?\s*level\s*support\b": "2nd Level Support"
    }

    for idx, row in df.iterrows():
        text = row.iloc[0]  # First column (normalized text)
        value = row.iloc[1]  # Second column (support hours value)

        # print(f"\nKey: {text}")
        # print(f"Value: {value}\n")

        if pd.notna(value):  # Ensure the value is valid
            value = str(value).replace("\n", "").strip()

            for regex, category in keyword_map.items():
                if re.search(regex, text, re.IGNORECASE):  # Match found (case-insensitive)
                    if not support_dict[category]:  # Store only first found value
                        support_dict[category] = value

    return support_dict


# Function to extract all the tables from the support time page
def extract_dataframes_from_support_hour_pages(pdf_path, page_numbers):
    extracted_dataframes = []

    # Define possible valid column headers
    valid_headers = [
        ["the service time describes the hours of coverage for this service.", "service time"],
        ["the service time describes the hours of coverage for this service.",
         "service time (cet if not stated otherwise)"],
        ["the service time describes the hours of coverage for this service.",
         "service time (cet/cest if not stated otherwise)"],
        ["term", "service time"],
        ["term", "service time (cet if not stated otherwise)"],
        ["term", "service time (cet/cest if not stated otherwise)"],
        ['term', 'service time (cet if not stated otherwise)']
    ]

    try:
        for page_number in page_numbers:
            page_number_str = str(page_number)
            tables = camelot.read_pdf(pdf_path, pages=page_number_str, flavor='lattice', line_scale=50)

            if not tables or tables.n == 0:
                continue  # Skip if no tables found

            # Convert each table to a DataFrame
            for i, table in enumerate(tables):
                try:
                    df = table.df  # Convert table to DataFrame
                    # print(df)

                    if df.empty:
                        continue  # Skip empty tables

                    # Normalize headers by removing extra spaces and converting to lowercase
                    headers = [re.sub(r'\s+', ' ', str(col)).strip().lower() for col in df.iloc[0]]
                    headers = [col for col in headers if col]
                    # print(f"\nExtracted Headers for Table {i+1}: {headers}")  # Debugging output

                    # Check if valid headers exist in extracted headers
                    if any(set(valid_set).issubset(set(headers)) for valid_set in valid_headers):
                        expected_columns = next(
                            valid_set for valid_set in valid_headers if set(valid_set).issubset(set(headers)))
                        # Trim df columns to match valid header count
                        df = df.iloc[:, :len(expected_columns)]  # Trim extra columns
                        df.columns = expected_columns  # Assign the expected headers
                        df = df[1:].reset_index(drop=True)  # Drop the first row (headers)
                        # print(df)
                        extracted_dataframes.append(df)

                except Exception as e:
                    print(f"\nError processing Table {i + 1}: {e}")  # Catch block

    except Exception as e:
        print(f"\nError processing the PDF file: {e}")

    return extracted_dataframes  # Always return a list


# ===========================================================
# Extracting Support Level details from Run of Service table
# ===========================================================

# Function to extract the ros details from ros tables
def extract_ros_details(dataframes):
    ros_support_details = {
        "1st Level Support": "",
        "2nd / 3rd Level Support": ""
    }

    # Variations for 2nd / 3rd Level Support
    second_third_variations = [
        "2nd / 3rd Level Support",
        "2nd/3rd Level Support",
        "2nd Level Support",
        "3rd Level Support",
        "1st / 2nd / 3rd Level Support"
    ]

    # Check each dataframe
    for df in dataframes:
        # Convert to string to safely handle mixed data
        df_str = df.astype(str)

        # Iterate over rows
        for row_index in range(len(df_str)):
            row_values = df_str.iloc[row_index].tolist()

            # Iterate over each cell in the row to find keywords
            for col_idx, cell_value in enumerate(row_values):

                # 1) Check for "1st Level Support"
                if ("1st Level Support".lower() in cell_value.lower()) or (
                        "1st / 2nd / 3rd Level Support".lower() in cell_value.lower()):
                    # Gather all columns in this row that contain a tick
                    ticked_columns = []
                    for check_col_idx, cell_content in enumerate(row_values):
                        if check_col_idx != col_idx and (
                                "" in cell_content or "✓" in cell_content or "" in cell_content or "*" in cell_content or "yes" in cell_content or "no" in cell_content):
                            ticked_columns.append(df_str.columns[check_col_idx])

                    # Store comma-separated column headers if any
                    if ticked_columns:
                        ros_support_details["1st Level Support"] = ", ".join(ticked_columns)

                # 2) Check for any variant of "2nd / 3rd Level Support"
                for variant in second_third_variations:
                    if variant.lower() in cell_value.lower():
                        ticked_columns = []
                        for check_col_idx, cell_content in enumerate(row_values):
                            if check_col_idx != col_idx and (
                                    "" in cell_content or "✓" in cell_content or "" in cell_content or "*" in cell_content or "yes" in cell_content or "no" in cell_content):
                                ticked_columns.append(df_str.columns[check_col_idx])

                        if ticked_columns:
                            ros_support_details["2nd / 3rd Level Support"] = ", ".join(ticked_columns)

    return ros_support_details


# Function to check if the extracted headers match the given pattern
def header_matches(headers, pattern):
    # Wildcard-based check (requires exact column count)
    if "*" in pattern:
        if len(headers) != len(pattern):
            return False
        for h, p in zip(headers, pattern):
            if p != "*" and h.lower() != p.lower():
                return False
        return True
    else:
        # For non-wildcard patterns, we check for subset presence (order is not important)
        pattern_lower = set(item.lower() for item in pattern)
        headers_lower = set(item.lower() for item in headers)
        return pattern_lower.issubset(headers_lower)


# Clean spaces from a list
def clean_list_spaces(lst):
    return [item.replace(' ', '').strip() for item in lst]


# Function to extract all the tables from the ros pages
def extract_dataframes_from_ros_pages(pdf_path, page_numbers):
    extracted_dataframes = []

    # Sample list of values to check
    target_values = [
        "Delivery Support Process", "Delivery Support Processes",
        "Support Processes", "Support Process",
        "Service Processes", "Service Process", "Process Category"
    ]

    # Define possible valid column headers
    valid_headers = [
        ["Delivery Support Process", "Yes", "No"],
        ["Delivery Support Process / Activity", "Yes", "No"],
        ["Delivery Support Process/Activity", "Yes", "No"],
        ["Delivery Support Processes", "Yes", "No"],
        ["Delivery Support Processes / Activity", "Yes", "No"],
        ["Delivery Support Processes/Activity", "Yes", "No"],
        ["Support Processes", "Yes", "No"],
        ["Service Processes", "Yes", "No"],
        ["Process Category", "Delivery Support Process", "Yes", "No"],
        ["Process Category", "Delivery Support Processes", "Yes", "No"],
        ["Process Category", "Support Processes", "Yes", "No"],
        ["Process Category", "Service Processes", "Yes", "No"],
        ["Delivery Support Process", "*", "*"],
        ["Delivery Support Process", "*", "*", "*"]
    ]

    try:
        for page_number in page_numbers:
            page_number_str = str(page_number)
            tables = camelot.read_pdf(pdf_path, pages=page_number_str, flavor='lattice', line_scale=50)

            if not tables or tables.n == 0:
                continue  # Skip if no tables found

            # Convert each table to a DataFrame
            for i, table in enumerate(tables):
                try:
                    df = table.df  # Convert table to DataFrame
                    df = df.replace(r'^\s*$', None, regex=True)
                    # print(df)

                    if df.empty:
                        continue  # Skip empty tables

                    # Assume first column is the one with categories
                    col_name = df.columns[0]

                    # Reset index to make slicing easier
                    df = df.reset_index(drop=True)

                    # Count target values and identify where any value occurs twice
                    seen = {}
                    cut_index = None

                    for idx, val in df[col_name].items():
                        if val in target_values:
                            seen[val] = seen.get(val, 0) + 1
                            if seen[val] == 2:
                                cut_index = idx
                                break

                    # Trim the DataFrame if a duplicate was found
                    if cut_index is not None:
                        df = df.iloc[:cut_index]

                    # Final cleaned DataFrame
                    df = df.dropna(axis=1, how='all')

                    # Normalize headers by removing extra spaces and converting to lowercase
                    headers = [re.sub(r'\s+', ' ', str(col)).strip() for col in df.iloc[0]]
                    # print(f"\nExtracted Headers for Table {i+1}: {headers}")
                    headers = [col for col in headers if col]
                    # print(f"\nExtracted Headers for Table {i+1}: {headers}")    # Debug

                    for valid in valid_headers:
                        # Clean first entry in both (remove special chars, lower case)
                        header_main = re.sub(r'[^a-zA-Z]', '', headers[0].lower())
                        valid_main = re.sub(r'[^a-zA-Z]', '', valid[0].lower())

                        # Cleaned sub-headers
                        clean_header_tail = clean_list_spaces(headers[1:])
                        clean_valid_tail = clean_list_spaces(valid[1:])

                        if valid_main in header_main and clean_header_tail == clean_valid_tail:
                            headers = valid

                    # print(f"\nCleared Headers: {headers}")

                    # Check if the table's headers match any of our valid header patterns
                    if any(header_matches(headers, valid_pattern) for valid_pattern in valid_headers):
                        df.columns = headers  # Assign the headers
                        df = df[1:].reset_index(drop=True)  # Drop the header row
                        extracted_dataframes.append(df)

                except:
                    pass
                    # print(f"\nError processing Table {i + 1}: {e}")

    except:
        pass
        # print(f"\nError processing the PDF file: {e}")

    return extracted_dataframes  # Always return a list


# ========================
# Generate list of pages
# ========================

# Function to generate a continuous page list
def create_page_list(pdf_path, search_start_text, search_end_text):
    index_page_number = find_index_page_number(pdf_path)

    start_pages = find_all_service_availability_and_support_hour_pages(pdf_path, search_start_text)
    end_pages = find_all_service_availability_and_support_hour_pages(pdf_path, search_end_text)

    # Remove index_page_number if present in both lists
    if index_page_number in start_pages and index_page_number in end_pages:
        start_pages.remove(index_page_number)
        end_pages.remove(index_page_number)

    # Remove any pages that are less than index_page_number
    start_pages = [page for page in start_pages if page >= index_page_number]
    end_pages = [page for page in end_pages if page >= index_page_number]

    # Ensure material_end_pages do not contain pages lower than the lowest start page
    if start_pages:
        lowest_start_page = min(start_pages)
        end_pages = [page for page in end_pages if page >= lowest_start_page]

    # Get the continuous range of page numbers
    if start_pages and end_pages:
        lowest_page = min(start_pages + end_pages)
        highest_page = max(start_pages + end_pages)
        page_numbers = list(range(lowest_page, highest_page + 1))
    else:
        page_numbers = []  # No valid range if lists are empty

    return start_pages, end_pages, page_numbers


# Function to generate a continuous page list
def create_page_list_run_of_service(pdf_path, search_start_text, search_end_text):
    index_page_number = find_index_page_number(pdf_path)

    start_pages = find_all_run_of_service_pages(pdf_path, search_start_text)
    end_pages = find_all_run_of_service_pages(pdf_path, search_end_text)

    # Remove index_page_number if present in both lists
    if index_page_number in start_pages and index_page_number in end_pages:
        start_pages.remove(index_page_number)
        end_pages.remove(index_page_number)

    # Remove any pages that are less than index_page_number
    start_pages = [page for page in start_pages if page >= index_page_number]
    end_pages = [page for page in end_pages if page >= index_page_number]

    # Ensure material_end_pages do not contain pages lower than the lowest start page
    if start_pages:
        lowest_start_page = min(start_pages)
        end_pages = [page for page in end_pages if page >= lowest_start_page]

    # Get the continuous range of page numbers
    if start_pages and end_pages:
        lowest_page = min(start_pages + end_pages)
        highest_page = max(start_pages + end_pages)
        page_numbers = list(range(lowest_page, highest_page + 1))
    else:
        page_numbers = []  # No valid range if lists are empty

    return start_pages, end_pages, page_numbers


# ==================================
# Save the extracted data to Excel
# ==================================

# Function to extract only numeric values from the Availability data
def extract_numeric_availability(availability):
    if isinstance(availability, str):
        # Replace commas with dots for European-style decimals
        availability = availability.replace(',', '.')
        # Extract only decimal or integer values after '>=', '=', or space, and remove unnecessary text
        numbers = re.findall(r'[>=\s]*?(\d+\.\d+|\d+)\s*%', availability)
        return ", ".join(numbers) if numbers else ""
    return ""


# Function to store the extracted values to the Excel
def insert_data_to_excel(excel_path, bsn_value, response_time_list, resolution_time_list, extracted_material_data,
                         extracted_drc_value, support_hour_dict, ros_support_details):
    # Required columns
    columns = [
        "BSN Number", "Material No/Nos", "Availability",
        "Response Time P1", "Response Time P2", "Response Time P3", "Response Time P4",
        "Resolution Time P1", "Resolution Time P2", "Resolution Time P3", "Resolution Time P4",
        "DRC Service", "Applicable DRC", "Applicable RPO", "1st Level Support", "2nd Level Support",
        "Emergency Support", "Non-Emergency Support", "1st Level Support Provided", "2nd/3rd Level Support Provided"
    ]

    # Load existing Excel file or create a new DataFrame
    if os.path.isfile(excel_path):
        df = pd.read_excel(excel_path, dtype=str)  # Load as strings to avoid type mismatches
        if "BSN Number" in df.columns and bsn_value in df["BSN Number"].values:
            print(f"\nBSN Number {bsn_value} already exists in the Excel file. Skipping insertion.")
            return  # Skip insertion
    else:
        df = pd.DataFrame(columns=columns)  # Create new dataframe if file does not exist

    new_rows = []

    # --- Material No & Availability ---
    material_no_list, availability_values = [], []

    for material, availability in extracted_material_data.items():
        if isinstance(availability, list):
            processed_availability = [extract_numeric_availability(val) for val in availability if val.strip()]
            processed_availability = [val for val in processed_availability if val]
            if not processed_availability:
                processed_availability = [""]
            for avail in processed_availability:
                material_no_list.append(material)
                availability_values.append(avail)
        else:
            material_no_list.append(material)
            availability_values.append(extract_numeric_availability(availability))

    # --- DRC Values ---
    drc_service_list, drc_rto_list, drc_rpo_list = [], [], []

    for drc_service, values in extracted_drc_value.items():
        applicable_drc = values.get("Applicable DRC", []) or [""]
        applicable_rpo = values.get("Applicable RPO", []) or [""]

        max_length = max(len(applicable_drc), len(applicable_rpo))
        for i in range(max_length):
            drc_rto = applicable_drc[i] if i < len(applicable_drc) else ""
            drc_rpo = applicable_rpo[i] if i < len(applicable_rpo) else ""
            drc_service_list.append(drc_service)
            drc_rto_list.append(drc_rto)
            drc_rpo_list.append(drc_rpo)

    # --- Determine max rows needed ---
    max_rows = max(len(drc_service_list), len(material_no_list), 1)

    def expand_list(lst):
        return lst + [lst[-1] if lst else ""] * (max_rows - len(lst))

    material_no_list = expand_list(material_no_list)
    availability_values = expand_list(availability_values)
    drc_service_list = expand_list(drc_service_list)
    drc_rto_list = expand_list(drc_rto_list)
    drc_rpo_list = expand_list(drc_rpo_list)

    # --- Expand new support detail columns ---
    first_level_support_provided = ros_support_details.get("1st Level Support", "")
    second_third_level_support_provided = ros_support_details.get("2nd / 3rd Level Support", "")
    first_level_support_list = [first_level_support_provided] * max_rows
    second_level_support_list = [second_third_level_support_provided] * max_rows

    # --- Row creation ---
    for i in range(max_rows):
        row = {
            "BSN Number": bsn_value,
            "Material No/Nos": material_no_list[i],
            "Availability": availability_values[i],
            "Response Time P1": response_time_list[0] if len(response_time_list) > 0 else "",
            "Response Time P2": response_time_list[1] if len(response_time_list) > 1 else "",
            "Response Time P3": response_time_list[2] if len(response_time_list) > 2 else "",
            "Response Time P4": response_time_list[3] if len(response_time_list) > 3 else "",
            "Resolution Time P1": resolution_time_list[0] if len(resolution_time_list) > 0 else "",
            "Resolution Time P2": resolution_time_list[1] if len(resolution_time_list) > 1 else "",
            "Resolution Time P3": resolution_time_list[2] if len(resolution_time_list) > 2 else "",
            "Resolution Time P4": resolution_time_list[3] if len(resolution_time_list) > 3 else "",
            "DRC Service": drc_service_list[i],
            "Applicable DRC": drc_rto_list[i],
            "Applicable RPO": drc_rpo_list[i],
            "1st Level Support": support_hour_dict.get("1st Level Support", ""),
            "2nd Level Support": support_hour_dict.get("2nd Level Support", ""),
            "Emergency Support": support_hour_dict.get("Emergency Support", ""),
            "Non-Emergency Support": support_hour_dict.get("Non-Emergency Support", ""),
            "1st Level Support Provided": first_level_support_list[i],
            "2nd/3rd Level Support Provided": second_level_support_list[i]
        }
        new_rows.append(row)

    # --- Append and display ---
    updated_df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    # print(f"\n{updated_df}")
    updated_df.to_excel(excel_path, index=False)


# =============
# Main method
# =============

def main(pdf_path, excel_path):
    ### Extract BSN details
    bsn_table_df = extract_bsn_table_from_pdf(pdf_path)
    bsn_value = extract_bsn_number_from_table(bsn_table_df)
    # ----------------------------------------------------------------------------------------------------------------------#

    ### Extract Incident details
    response_time_list, resolution_time_list = [""] * 4, [""] * 4
    incident_search_text = r"Table\s+\d+: Incident (Response and Resolution Time|Response Time|Resolution Time)"
    incident_page_number = find_incident_table_page_number(pdf_path, incident_search_text)

    if incident_page_number:
        extracted_dataframe = extract_all_tables_from_incident_page(pdf_path, incident_page_number)
        response_time_list, resolution_time_list = convert_df_into_list(extracted_dataframe)
    # ----------------------------------------------------------------------------------------------------------------------#

    ### Extract Service Availability details
    extracted_material_data = {}
    material_search_start_text = r"\d+\.\d+(\.\d+)?\sService Availability"
    material_search_end_text = r"\d+\.\d+(\.\d+)?\sService (Performance|Reliability|Times)"
    _, _, material_page_numbers = create_page_list(pdf_path, material_search_start_text, material_search_end_text)

    if material_page_numbers:
        extracted_material_data = extract_data_from_material_tables(pdf_path, material_page_numbers) or {}
    # ----------------------------------------------------------------------------------------------------------------------#

    ### Extract DRC details
    extracted_drc_value = {}
    drc_search_text = r"Table.*?Service(?:\s+\S+){0,6}\s+(Disaster|Recovery|Revocery|DR)(?:\s+\S+){0,6}\s+\b(Class|Classes|classes)\b"
    drc_page_number = find_drc_table_page_number(pdf_path, drc_search_text)

    if drc_page_number:
        extracted_drc_table = extract_all_tables_from_drc_page(pdf_path, drc_page_number)
        extracted_drc_value = convert_df_to_dict(extracted_drc_table) if extracted_drc_table is not None else {}
    # ----------------------------------------------------------------------------------------------------------------------#

    ### Extract Support Hour details
    support_hour_start_text = r"\d+\.\d+(\.\d+)?\sService (Time|Times)"
    support_hour_end_text = r"\d+\.\d+(\.\d+)?\sIncident (Management|Response Time|Resolution Time)"

    start_pages, end_pages, support_hour_pages = create_page_list(
        pdf_path, support_hour_start_text, support_hour_end_text
    )

    support_hour_dict = {
        "1st Level Support": "",
        "2nd Level Support": "",
        "Emergency Support": "",
        "Non-Emergency Support": ""
    }

    if support_hour_pages:
        dataframes = extract_dataframes_from_support_hour_pages(pdf_path, support_hour_pages)

        for dataframe in dataframes:
            temp_support_hour_values = extract_support_hours(dataframe)
            # Update only if value is present
            support_hour_dict.update({
                key: temp_support_hour_values[key]
                for key in support_hour_dict
                if temp_support_hour_values.get(key)
            })
    # ----------------------------------------------------------------------------------------------------------------------#

    ### Extract Run of Service details
    ros_start_text = r"\d+\.\d+(?:\.\d+)?\s(?:\w+\s)?Run of Service"
    ros_end_text = r"\d+\.\d+(?:\.\d+)?\sRetirement of(?:\s\w+)? Service"

    start_pages, end_pages, ros_page_numbers = create_page_list_run_of_service(pdf_path, ros_start_text, ros_end_text)

    if not ros_page_numbers:
        ros_support_details = {
            "1st Level Support": "",
            "2nd / 3rd Level Support": ""
        }
    else:
        extracted_ros_tables = extract_dataframes_from_ros_pages(pdf_path, ros_page_numbers)
        ros_support_details = extract_ros_details(extracted_ros_tables)
    # ----------------------------------------------------------------------------------------------------------------------#

    ### Insert extracted data into Excel
    insert_data_to_excel(excel_path, bsn_value, response_time_list, resolution_time_list,
                         extracted_material_data, extracted_drc_value, support_hour_dict, ros_support_details)


# =============================
# Processing all the SD files
# =============================

folder_path = r"C:\Users\rmya5fe\OneDrive - Allianz\01_Automated Reports\07_Sample_SDs"
database_path = os.path.join(folder_path, "Database")
excel_path = os.path.join(folder_path, "01_SLA_extract_from_SD.xlsx")

# List all PDF files in the database folder
pdf_files = [f for f in os.listdir(database_path) if f.lower().endswith(".pdf")]

print("\n=========>Processing started...<=========")

# Run the function for each PDF file
for file_name in pdf_files:
    pdf_path = os.path.join(database_path, file_name)
    print(f"\nProcessing: {file_name}")
    main(pdf_path, excel_path)

print("\n=========>Processing completed for all PDF files.<=========\n")

# 8010222
