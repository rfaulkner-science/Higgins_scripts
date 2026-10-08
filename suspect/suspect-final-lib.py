#####
# Suspect Screening - Final Data Processing for LC-QTOF-MS
# Rewrite of v4.3 in MATLAB, Conrad Pritchard 3/19/2025

# LIBRARY VERSION
# Unlike the MATLAB script and other rewrites, this file has been intentionally redesigned into a library instead.

import os
import sys
import pandas as pd

###############################
## --- Utility functions --- ##
###############################

# random util fn; example for documentation
def read_file(path: str, required_columns: list[str], read_options: dict = {}) -> pd.DataFrame:
    """
    Reads an Excel or CSV file into a pandas DataFrame, validating that the
    required columns are present.
    Note that the function will terminate the program if the file cannot be read
    or if required columns are missing.

    Parameters:
    path (str): The path to the Excel or CSV file.
    required_columns (list[str]): A list of required column names.
    read_options (dict): Additional options for reading the file.

    Returns:
    pd.DataFrame: The DataFrame containing the data from the parsed file.
    """
    
    try:
        if path.endswith('.csv') or path.endswith('.tsv') or path.endswith('.txt'):
            df = pd.read_csv(path, **read_options)
        elif path.endswith('.xlsx') or path.endswith('.xls'):
            df = pd.read_excel(path, **read_options)
        else:
            raise ValueError(f"Unsupported file format for {path}. Please provide a CSV or Excel file.")
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        return df
    except Exception as e:
        print(f"Error reading {path}: {e}")
        sys.exit(1)

# Typically, the injection volume was used to assume the method.
# 100 uL - soil
# rest - aqueous