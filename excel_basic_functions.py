import pandas as pd
import logging

# Constants
SPECIAL_CHARACTERS = "!@#$%^&*()_+={}[]|\\:;'\"<>,.?/~`"

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def clean_column_name(column_name, special_chars=SPECIAL_CHARACTERS):
    """
    Removes special characters from column names and handles unnamed columns.
    
    Args:
    column_name (str): The original column name.
    special_chars (str): A string of special characters to be removed from the column name.
    
    Returns:
    str: The cleaned column name.
    """
    if 'Unnamed' in column_name:
        column_name = column_name.split(", ")[-1]  # Extract meaningful part of unnamed columns
    for char in special_chars:
        column_name = column_name.replace(char, "")  # Remove special characters
    return column_name.strip()  # Remove any leading/trailing whitespace

def from_excel_to_list_of_dataframes_multiple_headers(excel_file, reference_header_file, percentage_of_overlap):
    """
    Reads an Excel file and splits it into multiple dataframes based on headers specified in a reference file.
    
    Args:
    excel_file (str): Path to the main Excel file.
    reference_header_file (str): Path to the reference Excel file containing headers.
    percentage_of_overlap (float): Minimum percentage of overlap required to identify headers.
    
    Returns:
    dict: A dictionary where keys are sheet names and values are dictionaries of dataframes split by headers.
    """
    logging.info(f"Loading reference headers from {reference_header_file}")
    reference_df = pd.read_excel(reference_header_file, header=None)
    
    logging.info(f"Loading Excel workbook from {excel_file}")
    excel_workbook = pd.ExcelFile(excel_file)
    
    dataframe_dict = {}

    # Precompute reference header sets for efficient comparison
    logging.info("Precomputing reference header sets")
    reference_header_sets = {}
    for header in reference_df[0].dropna().unique():
        header_rows = reference_df[reference_df[0] == header]
        header_sets = [set(row.dropna()) for _, row in header_rows.drop([0], axis=1).iterrows()]
        reference_header_sets[header] = header_sets

    # Process each sheet in the Excel workbook
    for sheet_name in excel_workbook.sheet_names:
        logging.info(f"Processing sheet: {sheet_name}")
        dataframe_dict[sheet_name] = {}
        
        sheet_df = excel_workbook.parse(sheet_name, header=None)

        # Convert all rows to sets for faster comparison
        row_sets = [set(row.dropna()) for _, row in sheet_df.iterrows()]

        # List to store identified header ranges
        header_ranges = []

        # Identify headers based on reference header sets
        for reference_header, header_sets in reference_header_sets.items():
            potential_header_indices = []  # Tracks potential header row indices
            match_count = 0  # Counter for matching headers
            
            for row_idx, row_set in enumerate(row_sets):
                if match_count < len(header_sets):
                    reference_set = header_sets[match_count]
                    overlap = row_set & reference_set  # Intersection of sets
                    universe = row_set | reference_set  # Union of sets
                    overlap_percentage = (len(overlap) / len(universe)) * 100 if len(universe) else 100  # Calculate overlap percentage

                    if overlap_percentage >= percentage_of_overlap:
                        potential_header_indices.append(row_idx)
                        match_count += 1
                    else:
                        match_count = 0  # Reset match count if overlap condition is not met
                        potential_header_indices = []

                if match_count == len(header_sets):
                    header_ranges.append((potential_header_indices, None))  # Found a complete header set
                    match_count = 0  # Reset for next header search
                    potential_header_indices = []

        # Determine end row for each header set
        for i in range(len(header_ranges) - 1):
            header_ranges[i] = (header_ranges[i][0], header_ranges[i + 1][0][0] - 1)

        # Parse dataframes based on identified headers
        for index, (header_indices, end_row) in enumerate(header_ranges, start=1):
            start_row = header_indices[0]
            nrows = (end_row - start_row + 1) if end_row else None
            excel_dataframe = excel_workbook.parse(sheet_name, header=header_indices, nrows=nrows)
            
            # Clean the column names
            excel_dataframe.columns = [clean_column_name(str(col)) if str(col) else f"Unnamed_{i}" for i, col in enumerate(excel_dataframe.columns)]
            
            # Store the dataframe in the dictionary
            dataframe_dict[sheet_name][f'{sheet_name}_table_{index}'] = excel_dataframe
            logging.info(f"Extracted dataframe {index} from sheet {sheet_name}")

    return dataframe_dict

# Example usage:
# dataframe_dict = from_excel_to_list_of_dataframes_multiple_headers('data.xlsx', 'reference_headers.xlsx', 50)
