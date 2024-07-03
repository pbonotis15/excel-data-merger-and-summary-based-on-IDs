# Author: Panos Bonotis
# Date: Jul-2024
# Description: This project is a Python-based tool designed 
# to automate the aggregation and processing of multiple Excel files. 
# It provides a user-friendly interface for selecting input and output directories, 
# reads and concatenates data from multiple Excel sheets, standardizes column names, 
# and removes duplicates based on specified criteria. 
# The processed data is then saved into a new Excel file.

import pandas as pd
import glob
import tkinter as tk
from tkinter import filedialog
import os
import random

# Function to read all Excel files and concatenate them into a dictionary of DataFrames by sheet name
def read_excel_files(file_pattern):
    files = glob.glob(file_pattern)
    all_data = {}
    for file in files:
        xl = pd.ExcelFile(file)
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            df = standardize_column_names(df)
            if sheet_name not in all_data:
                all_data[sheet_name] = df
            else:
                if not df.empty:
                    all_data[sheet_name] = pd.concat([all_data[sheet_name], df], ignore_index=True)
    return all_data

# Function to standardize column names for SR ID and AGE
def standardize_column_names(df):
    rename_dict = {
        'SR': 'SR ID',
        'sr': 'SR ID',
        'SRID': 'SR ID',
        'AGECLASS': 'AGE'
    }
    df.rename(columns=rename_dict, inplace=True)
    return df

# Function to resolve duplicates based on the available columns
def resolve_duplicates(df):
    if 'AGE' in df.columns:
        # Sort by SR ID and AGE to ensure the entry with the highest AGE comes first
        df = df.sort_values(by=['SR ID', 'AGE'], ascending=[True, False])
        # Drop duplicates based on SR ID, keeping the first occurrence (which has the highest AGE due to sorting)
        df = df.drop_duplicates(subset='SR ID', keep='first')
    elif 'status' in df.columns:
        df_grouped = df.groupby('SR ID')
        resolved_df_list = []
        for name, group in df_grouped:
            if group['status'].str.lower().isin(['done']).any():
                # Select randomly among the 'done' entries
                resolved_df_list.append(group[group['status'].str.lower() == 'done'].sample(1))
            else:
                # Select randomly among all entries
                resolved_df_list.append(group.sample(1))
        df = pd.concat(resolved_df_list).reset_index(drop=True)
    else:
        # Randomly select one of the duplicates
        df = df.groupby('SR ID').apply(lambda x: x.sample(1)).reset_index(drop=True)
    return df

# Function to process the DataFrame and handle duplicates
def process_data(data_dict):
    processed_data = {}
    for sheet_name, data in data_dict.items():
        if 'SR ID' in data.columns:
            data = resolve_duplicates(data)
        processed_data[sheet_name] = data
    return processed_data

# Function to save the processed DataFrame to a new Excel file
def save_to_excel(data_dict, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, data in data_dict.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

# Function to select folder using tkinter
def select_folder(title):
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    return folder_path

# Main function to execute the script
def main():
    # Allow the user to select the input folder containing the Excel files
    input_folder = select_folder("Select the folder containing the Excel files")
    # Allow the user to select the output folder where the result will be saved
    output_folder = select_folder("Select the output folder")
    
    # Check if the input and output folders are selected
    if not input_folder or not output_folder:
        print("Input or Output folder not selected. Exiting...")
        return

    # Define the input pattern to match all Excel files in the selected input folder
    input_pattern = os.path.join(input_folder, "*.xlsx")
    # Define the output file path where the processed data will be saved
    output_file = os.path.join(output_folder, "unique_tasks.xlsx")
    
    # Read all Excel files and concatenate them into a dictionary of DataFrames by sheet name
    all_data = read_excel_files(input_pattern)
    # Process the DataFrame to remove duplicates based on SR ID, keeping the entry with the highest AGE
    unique_data = process_data(all_data)
    # Save the processed DataFrame to a new Excel file
    save_to_excel(unique_data, output_file)
    print(f"Processed data saved to {output_file}")

# Execute the main function if the script is run directly
if __name__ == "__main__":
    main()