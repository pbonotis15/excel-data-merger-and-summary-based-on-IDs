# Author: Panos Bonotis -> https://www.linkedin.com/in/panagiotis-bonotis-351a7996/
# Date: Jul-2024
# Description: This project is a Python-based tool designed 
# to automate the aggregation and processing of multiple Excel files. 
# It provides a user-friendly interface for selecting input and output directories, 
# reads and concatenates data from multiple Excel sheets, standardizes column names, 
# filters data based on specified criteria, and merges relevant information. 
# Additionally, it removes duplicate entries based on the 'SR ID' and 'Ημερομηνία Δημιουργίας' columns, 
# ensuring that only the most recent entry for each 'SR ID' is retained. 
# The processed data is then saved into a new Excel file with multiple sheets: 
# 'Aggregated Data', 'Summary of Actions', and 'Last Drop'. 

import pandas as pd

def get_sheet_names(file_path):
    return pd.ExcelFile(file_path).sheet_names

def read_excel_file(file_path, sheet_name, usecols):
    return pd.read_excel(file_path, sheet_name=sheet_name, usecols=usecols)

def filter_data(df, filter_df, key):
    return df[df[key].isin(filter_df[key])]

def merge_data(df1, df2, key):
    return pd.merge(df1, df2, on=key)

def create_summary(final_df, df1):
    summary_rows = []
    unique_sr_ids = final_df['SR ID'].unique()
    for sr_id in unique_sr_ids:
        row = {
            'SR ID': sr_id,
            'BUILDING ID': final_df[final_df['SR ID'] == sr_id][['BUILDING ID', 'building Id']].bfill(axis=1).iloc[:, 0].values[0],
            'ADDRESS': final_df[final_df['SR ID'] == sr_id][['ADDRESS', 'full_adr']].bfill(axis=1).iloc[:, 0].values[0],
            'FLOOR': final_df[final_df['SR ID'] == sr_id]['FLOOR'].values[0],
            'A/K': final_df[final_df['SR ID'] == sr_id]['A/K'].values[0],
            'AGE/AGE': final_df[final_df['SR ID'] == sr_id]['AGE'].values[0],
            'CREATED': final_df[final_df['SR ID'] == sr_id]['CREATED'].values[0],
            'Τεχνικός σε Ανάθεση (KAM)': df1[df1['SR ID'] == sr_id]['Τεχνικός σε Ανάθεση (KAM)'].values[0],
            'Κατηγορία Αιτήματος': final_df[final_df['SR ID'] == sr_id]['Κατηγορία Αιτήματος'].values[0],
            'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ': final_df[final_df['SR ID'] == sr_id]['ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ'].values[0],
            'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ': final_df[final_df['SR ID'] == sr_id][['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'customer']].bfill(axis=1).iloc[:, 0].values[0],
            'ΚΙΝΗΤΟ ΠΕΛΑΤΗ': final_df[final_df['SR ID'] == sr_id][['ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'mobile']].bfill(axis=1).iloc[:, 0].values[0],
            'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ': final_df[final_df['SR ID'] == sr_id]['ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ'].values[0],
            'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ': final_df[final_df['SR ID'] == sr_id]['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ'].values[0],
            'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ': final_df[final_df['SR ID'] == sr_id]['ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ'].values[0],
            'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ': final_df[final_df['SR ID'] == sr_id]['ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ'].values[0],
            'BEP/FB CODE': final_df[final_df['SR ID'] == sr_id]['BEP/FB CODE'].values[0],
            'BEP/FB PORT': final_df[final_df['SR ID'] == sr_id]['BEP/FB PORT'].values[0],
            'BEP/FB TYPE': final_df[final_df['SR ID'] == sr_id]['BEP/FB TYPE'].values[0]
        }

        # Autopsia related data
        autopsia_df = final_df[(final_df['SR ID'] == sr_id) & (final_df['Τύπος εργασίας'] == 'ΑΥΤΟΨΙΑ FTTH')]
        row.update({
            'Ημ/νία Αίτησης (ΑΥΤΟΨΙΑ)': autopsia_df['Ημ/νία Αίτησης'].max() if not autopsia_df.empty else None,
            'Τύπος εργασίας (ΑΥΤΟΨΙΑ)': autopsia_df[autopsia_df['Ημ/νία Αίτησης'] == autopsia_df['Ημ/νία Αίτησης'].max()]['Τύπος εργασίας'].values[0] if not autopsia_df.empty else None,
            'Κατάσταση (ΑΥΤΟΨΙΑ)': autopsia_df[autopsia_df['Ημ/νία Αίτησης'] == autopsia_df['Ημ/νία Αίτησης'].max()]['Κατάσταση'].values[0] if not autopsia_df.empty else None
        })

        # Kataskevi related data
        kataskevi_df = final_df[(final_df['SR ID'] == sr_id) & (final_df['Τύπος εργασίας'] == 'ΚΑΤΑΣΚΕΥΗ FTTH')]
        row.update({
            'Ημ/νία Αίτησης (ΚΑΤΑΣΚΕΥΗ)': kataskevi_df['Ημ/νία Αίτησης'].max() if not kataskevi_df.empty else None,
            'Τύπος εργασίας (ΚΑΤΑΣΚΕΥΗ)': kataskevi_df[kataskevi_df['Ημ/νία Αίτησης'] == kataskevi_df['Ημ/νία Αίτησης'].max()]['Τύπος εργασίας'].values[0] if not kataskevi_df.empty else None,
            'Κατάσταση (ΚΑΤΑΣΚΕΥΗ)': kataskevi_df[kataskevi_df['Ημ/νία Αίτησης'] == kataskevi_df['Ημ/νία Αίτησης'].max()]['Κατάσταση'].values[0] if not kataskevi_df.empty else None
        })

        # Default fields
        row.update({
            'Ημερομηνία Εκτέλεσης (Χωματουργικές Εργασίες)': None,
            'Κατάσταση (Χωματουργικές Εργασίες)': None,
            'Ημερομηνία Εκτέλεσης (Δικτυακές Εργασίες)': None,
            'Δικτυακές Εργασίες': None,
            'Κατάσταση (Δικτυακές Εργασίες)': None
        })

        summary_rows.append(row)

    return pd.DataFrame(summary_rows)

def create_last_drop(final_df):
    last_drop_rows = []
    unique_sr_ids = final_df['SR ID'].unique()
    for sr_id in unique_sr_ids:
        row = {
            'SR ID': sr_id,
            'PILOT': final_df[final_df['SR ID'] == sr_id][['PILOT', 'pilot']].bfill(axis=1).iloc[:, 0].values[0],
            'Ημερομηνία Εκτέλεσης (As-built)': final_df[final_df['SR ID'] == sr_id]['sr_created'].values[0],
            'Pilot/Last drop': final_df[final_df['SR ID'] == sr_id]['FIELDTASKTYPE'].values[0],
            'Κατάσταση (As-built)': final_df[final_df['SR ID'] == sr_id]['FIELDTASKSTATUS'].values[0],
            'As-built/Απολογισμός': None
        }
        last_drop_rows.append(row)

    return pd.DataFrame(last_drop_rows)

def main(file_path1, file_path2, sheet_name, output_folder):
    # Load the first Excel file
    df1 = read_excel_file(file_path1, sheet_name, usecols=['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος'])

    # Load and filter data from the second Excel file
    sheet_configs = [
        ('Ανατεθειμένα για κατασκευή', ['SR ID', 'AGE', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']),
        ('Ανατεθειμένες αυτοψίες', ['SR ID', 'AGE', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']),
        ('Εντολές στο ίδιο BID', ['SR ID', 'AGE', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']),
        ('New flow', ['SR ID', 'sr_created', 'FIELDTASKTYPE', 'FIELDTASKSTATUS', 'AGE', 'pilot', 'full_adr', 'customer', 'mobile', 'building Id'])
    ]

    filtered_dfs = []
    for sheet, cols in sheet_configs:
        df2 = read_excel_file(file_path2, sheet, cols)
        filtered_dfs.append(filter_data(df2, df1, 'SR ID'))

    if all(df.empty for df in filtered_dfs):
        print("No matching SR IDs found.")
        pd.DataFrame().to_excel(f'{output_folder}/nothing_found.xlsx', index=False)
    else:
        # Merge the filtered data with the corresponding information from the first file
        merged_dfs = [merge_data(df, df1, 'SR ID') for df in filtered_dfs]

        # Concatenate all merged dataframes
        data_df = pd.concat(merged_dfs, ignore_index=True)

        # Remove duplicate SR ID entries based on the column 'Ημ/νία Αίτησης', keeping the latest date
        data_df['Ημ/νία Αίτησης'] = pd.to_datetime(data_df['Ημ/νία Αίτησης'])
        data_df = data_df.sort_values('Ημ/νία Αίτησης').drop_duplicates(subset=['SR ID'], keep='last')

        final_df = pd.concat(merged_dfs, ignore_index=True)
        
        # Create the summary and last drop DataFrames
        summary_df = create_summary(final_df, df1)
        last_drop_df = create_last_drop(final_df)

        # Save the final DataFrame and the summary DataFrame to the specified output folder
        with pd.ExcelWriter(f'{output_folder}/final_results.xlsx') as writer:
            data_df.to_excel(writer, sheet_name='Aggregated Data', index=False)
            summary_df.to_excel(writer, sheet_name='Summary of Actions', index=False)
            last_drop_df.to_excel(writer, sheet_name='Last Drop', index=False)
        
        print("Data saved to final_results.xlsx")

if __name__ == "__main__":
    main()