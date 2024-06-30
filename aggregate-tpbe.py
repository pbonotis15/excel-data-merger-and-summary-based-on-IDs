import pandas as pd
import glob

# Function to read all Excel files and concatenate them into a dictionary of DataFrames by sheet name
def read_excel_files(file_pattern):
    files = glob.glob(file_pattern)
    all_data = {}
    for file in files:
        xl = pd.ExcelFile(file)
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            df = standardize_column_names(df)
            if 'SR ID' in df.columns and 'AGE' in df.columns:
                if sheet_name not in all_data:
                    all_data[sheet_name] = df
                else:
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

# Function to process the DataFrame and remove duplicates based on SR ID, keeping the entry with the highest AGE
def process_data(data_dict):
    processed_data = {}
    for sheet_name, data in data_dict.items():
        if 'SR ID' in data.columns and 'AGE' in data.columns:
            # Sort by SR ID and AGE to ensure the entry with the highest AGE comes first
            data_sorted = data.sort_values(by=['SR ID', 'AGE'], ascending=[True, False])
            # Drop duplicates based on SR ID, keeping the first occurrence (which has the highest AGE due to sorting)
            unique_data = data_sorted.drop_duplicates(subset='SR ID', keep='first')
            processed_data[sheet_name] = unique_data
    return processed_data

# Function to save the processed DataFrame to a new Excel file
def save_to_excel(data_dict, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, data in data_dict.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

# Main function to execute the script
def main(input_pattern, output_file):
    all_data = read_excel_files(input_pattern)
    unique_data = process_data(all_data)
    save_to_excel(unique_data, output_file)

# Execute the main function
if __name__ == "__main__":
    input_pattern = "D:/ownCloud/Personal Folders & Files/εταιρία/nbonotis/Legacy code/ΜΑΙΟΣ 24/*.xlsx"  # Adjust the path to your Excel files
    output_file = "unique_tasks.xlsx"  # Adjust the output file name and path as needed
    main(input_pattern, output_file)