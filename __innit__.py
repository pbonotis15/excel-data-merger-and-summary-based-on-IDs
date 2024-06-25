#Author Panos Bonotis -> https://www.linkedin.com/in/panagiotis-bonotis-351a7996/

import pandas as pd
import math

def get_sheet_names(file_path):
    return pd.ExcelFile(file_path).sheet_names

def main(file_path1, file_path2, sheet_name, output_folder):
    # Load the first Excel file
    df1 = pd.read_excel(file_path1, sheet_name=sheet_name, usecols=['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος'])

    # Load the second Excel file
    df2_construction = pd.read_excel(file_path2, sheet_name='Ανατεθειμένα για κατασκευή', usecols=['SR', 'AGE', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE'])
    df2_inspection = pd.read_excel(file_path2, sheet_name='Ανατεθειμένες αυτοψίες', usecols=['SR', 'AGE', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE'])
    df2_bid = pd.read_excel(file_path2, sheet_name='Εντολές στο ίδιο BID', usecols=['SR', 'AGE', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE'])
    df2_new_flow = pd.read_excel(file_path2, sheet_name='New flow', usecols=['SRID', 'sr_created', 'FIELDTASKTYPE', 'FIELDTASKSTATUS', 'AGECLASS', 'pilot', 'full_adr', 'customer', 'mobile', 'building Id'])

    # Filter rows based on SR ID from the first file
    filtered_construction = df2_construction[df2_construction['SR'].isin(df1['SR ID'])]
    filtered_inspection = df2_inspection[df2_inspection['SR'].isin(df1['SR ID'])]
    filtered_bid = df2_bid[df2_bid['SR'].isin(df1['SR ID'])]
    filtered_new_flow = df2_new_flow[df2_new_flow['SRID'].isin(df1['SR ID'])]

    # Check if there are no matching SR IDs
    if filtered_construction.empty and filtered_inspection.empty and filtered_bid.empty and filtered_new_flow.empty:
        print("No matching SR IDs found.")
        df_empty = pd.DataFrame()
        df_empty.to_excel('nothing_found.xlsx', index=False)
    else:
        # Merge the filtered data with the corresponding information from the first file
        merged_construction = pd.merge(filtered_construction, df1, left_on='SR', right_on='SR ID')
        merged_inspection = pd.merge(filtered_inspection, df1, left_on='SR', right_on='SR ID')
        merged_bid = pd.merge(filtered_bid, df1, left_on='SR', right_on='SR ID')
        merged_new_flow = pd.merge(filtered_new_flow, df1, left_on='SRID', right_on='SR ID')


        # Prepare the final DataFrame
        data_df = pd.concat([merged_construction[['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'AGE', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']],
                              merged_inspection[['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'AGE', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']],
                              merged_bid[['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'AGE', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']],
                              merged_new_flow[['SR ID', 'sr_created', 'FIELDTASKTYPE', 'FIELDTASKSTATUS', 'full_adr', 'AGECLASS', 'pilot', 'customer', 'mobile', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'building Id']]], ignore_index=True)
        final_df = pd.concat([merged_construction[['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'AGE', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']],
                              merged_inspection[['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'AGE', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']],
                              merged_bid[['SR ID', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'TYPE', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'ADDRESS', 'FLOOR', 'PILOT', 'A/K', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'AGE', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ', 'E-MAIL ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'E-MAIL ΔΙΑΧΕΙΡΙΣΤΗ', 'CREATED', 'BUILDING ID', 'BEP/FB CODE', 'BEP/FB PORT', 'BEP/FB TYPE']],
                              merged_new_flow[['SR ID', 'sr_created', 'FIELDTASKTYPE', 'FIELDTASKSTATUS', 'full_adr', 'AGECLASS', 'pilot', 'customer', 'mobile', 'Τύπος εργασίας', 'Ημ/νία Αίτησης', 'Τεχνικός σε Ανάθεση (KAM)', 'Κατάσταση', 'Ημερομηνία Ολοκλήρωσης', 'Τ.Τ.Λ.Π.', 'Διεύθυνση πελάτη', 'Αριθμός Οδού', 'Έναρξη Ραντεβού', 'Έγκριση Εργασίας', 'Κατηγορία Αιτήματος', 'building Id']]], ignore_index=True)

      # Create a summary DataFrame
        summary_rows = []
        unique_sr_ids = final_df['SR ID'].unique()
        for sr_id in unique_sr_ids:
            row = {}
            row['SR ID'] = sr_id
            buildin1_df = final_df[final_df['SR ID'] == sr_id]['BUILDING ID']
            buildin2_df = final_df[final_df['SR ID'] == sr_id]['building Id']
            if not buildin1_df.empty and not buildin1_df.isna().values[0]:
                row['BUILDING ID'] = final_df[final_df['SR ID'] == sr_id]['BUILDING ID'].values[0]
            elif not buildin2_df.empty :
                row['BUILDING ID'] = final_df[final_df['SR ID'] == sr_id]['building Id'].values[0]
            else:
                row['BUILDING ID'] = None
            adress1_df = final_df[final_df['SR ID'] == sr_id]['ADDRESS']
            adress2_df = final_df[final_df['SR ID'] == sr_id]['full_adr']
            if not adress1_df.empty and not adress1_df.isna().values[0]:
                row['ADDRESS'] = final_df[final_df['SR ID'] == sr_id]['ADDRESS'].values[0]
            elif not adress2_df.empty:
                row['ADDRESS'] = final_df[final_df['SR ID'] == sr_id]['full_adr'].values[0]
            else:
                row['ADDRESS'] = None                        
            row['FLOOR'] = final_df[final_df['SR ID'] == sr_id]['FLOOR'].values[0]
            row['A/K'] = final_df[final_df['SR ID'] == sr_id]['A/K'].values[0]
            age_df = final_df[final_df['SR ID'] == sr_id]['AGE']
            ageclass_df = final_df[final_df['SR ID'] == sr_id]['AGECLASS']
            if not age_df.empty and not age_df.isna().values[0]:
                row['AGE/AGECLASS'] = final_df[final_df['SR ID'] == sr_id]['AGE'].values[0]
            elif not ageclass_df.empty:
                row['AGE/AGECLASS'] = final_df[final_df['SR ID'] == sr_id]['AGECLASS'].values[0]
            else:
                row['AGE/AGECLASS'] = None            
            row['CREATED'] = final_df[final_df['SR ID'] == sr_id]['CREATED'].values[0]
            row['Τεχνικός σε Ανάθεση (KAM)'] = df1[df1['SR ID'] == sr_id]['Τεχνικός σε Ανάθεση (KAM)'].values[0]
            autopsia_df = final_df[(final_df['SR ID'] == sr_id) & (final_df['Τύπος εργασίας'] == 'ΑΥΤΟΨΙΑ FTTH')]
            if not autopsia_df.empty:
                row['Ημ/νία Αίτησης (ΑΥΤΟΨΙΑ)'] = autopsia_df['Ημ/νία Αίτησης'].max()
                row['Τύπος εργασίας (ΑΥΤΟΨΙΑ)'] = autopsia_df[autopsia_df['Ημ/νία Αίτησης'] == row['Ημ/νία Αίτησης (ΑΥΤΟΨΙΑ)']]['Τύπος εργασίας'].values[0]
                row['Κατάσταση (ΑΥΤΟΨΙΑ)'] = autopsia_df[autopsia_df['Ημ/νία Αίτησης'] == row['Ημ/νία Αίτησης (ΑΥΤΟΨΙΑ)']]['Κατάσταση'].values[0]
            else:
                row['Ημ/νία Αίτησης (ΑΥΤΟΨΙΑ)'] = None
                row['Τύπος εργασίας (ΑΥΤΟΨΙΑ)'] = None
                row['Κατάσταση (ΑΥΤΟΨΙΑ)'] = None
            kataskevi_df = final_df[(final_df['SR ID'] == sr_id) & (final_df['Τύπος εργασίας'] == 'ΚΑΤΑΣΚΕΥΗ FTTH')]
            if not kataskevi_df.empty:
                row['Ημ/νία Αίτησης (ΚΑΤΑΣΚΕΥΗ)'] = kataskevi_df['Ημ/νία Αίτησης'].max()
                row['Τύπος εργασίας (ΚΑΤΑΣΚΕΥΗ)'] = kataskevi_df[kataskevi_df['Ημ/νία Αίτησης'] == row['Ημ/νία Αίτησης (ΚΑΤΑΣΚΕΥΗ)']]['Τύπος εργασίας'].values[0]
                row['Κατάσταση (ΚΑΤΑΣΚΕΥΗ)'] = kataskevi_df[(kataskevi_df['Ημ/νία Αίτησης'] == row['Ημ/νία Αίτησης (ΚΑΤΑΣΚΕΥΗ)'])]['Κατάσταση'].values[0]
            else:
                row['Ημ/νία Αίτησης (ΚΑΤΑΣΚΕΥΗ)'] = None
                row['Τύπος εργασίας (ΚΑΤΑΣΚΕΥΗ)'] = None
                row['Κατάσταση (ΚΑΤΑΣΚΕΥΗ)'] = None
            row['Ημερομηνία Εκτέλεσης (Χωματουργικές Εργασίες)'] = None
            row['Κατάσταση (Χωματουργικές Εργασίες)'] = None
            row['Ημερομηνία Εκτέλεσης (Δικτυακές Εργασίες)'] = None
            row['Δικτυακές Εργασίες'] = None
            row['Κατάσταση (Δικτυακές Εργασίες)'] = None
            pilot1_df = final_df[final_df['SR ID'] == sr_id]['PILOT']
            pilot2_df = final_df[final_df['SR ID'] == sr_id]['pilot']
            if not pilot1_df.empty and not pilot1_df.isna().values[0]:
                row['PILOT'] = final_df[final_df['SR ID'] == sr_id]['PILOT'].values[0]
            elif not pilot2_df.empty:
                row['PILOT'] = final_df[final_df['SR ID'] == sr_id]['pilot'].values[0]
            else:
                row['PILOT'] = None 
            row['Ημερομηνία Εκτέλεσης (As-built)'] = final_df[final_df['SR ID'] == sr_id]['sr_created'].values[0]
            row['Pilot/Last drop'] = final_df[final_df['SR ID'] == sr_id]['FIELDTASKTYPE'].values[0]
            row['Κατάσταση (As-built)'] = final_df[final_df['SR ID'] == sr_id]['FIELDTASKSTATUS'].values[0]
            row['As-built/Απολογισμός'] = None
            row['Κατηγορία Αιτήματος'] = final_df[final_df['SR ID'] == sr_id]['Κατηγορία Αιτήματος'].values[0]
            row['ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ'] = final_df[final_df['SR ID'] == sr_id]['ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ'].values[0]
            customer1_df = final_df[final_df['SR ID'] == sr_id]['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ']
            customer2_df = final_df[final_df['SR ID'] == sr_id]['customer']
            if not customer1_df.empty and not customer1_df.isna().values[0]:
                row['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ'] = final_df[final_df['SR ID'] == sr_id]['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ'].values[0]
            elif not customer2_df.empty:
                row['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ'] = final_df[final_df['SR ID'] == sr_id]['customer'].values[0]
            else:
                row['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ'] = None 
            customertel1_df = final_df[final_df['SR ID'] == sr_id]['ΚΙΝΗΤΟ ΠΕΛΑΤΗ']
            customertel2_df = final_df[final_df['SR ID'] == sr_id]['mobile']
            if not customertel1_df.empty and not customertel1_df.isna().values[0]:
                row['ΚΙΝΗΤΟ ΠΕΛΑΤΗ'] = final_df[final_df['SR ID'] == sr_id]['ΚΙΝΗΤΟ ΠΕΛΑΤΗ'].values[0]
            elif not customertel2_df.empty:
                row['ΚΙΝΗΤΟ ΠΕΛΑΤΗ'] = final_df[final_df['SR ID'] == sr_id]['mobile'].values[0]
            else:
                row['ΚΙΝΗΤΟ ΠΕΛΑΤΗ'] = None
            row['ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ'] = final_df[final_df['SR ID'] == sr_id]['ΣΤΑΘΕΡΟ ΠΕΛΑΤΗ'].values[0]
            row['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ'] = final_df[final_df['SR ID'] == sr_id]['ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ'].values[0]
            row['ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ'] = final_df[final_df['SR ID'] == sr_id]['ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ'].values[0]
            row['ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ'] = final_df[final_df['SR ID'] == sr_id]['ΣΤΑΘΕΡΟ ΔΙΑΧΕΙΡΙΣΤΗ'].values[0]
            row['BEP/FB CODE'] = final_df[final_df['SR ID'] == sr_id]['BEP/FB CODE'].values[0]
            row['BEP/FB PORT'] = final_df[final_df['SR ID'] == sr_id]['BEP/FB PORT'].values[0]
            row['BEP/FB TYPE'] = final_df[final_df['SR ID'] == sr_id]['BEP/FB TYPE'].values[0]

            summary_rows.append(row)

        summary_df = pd.DataFrame(summary_rows)

        # Save the final DataFrame and the summary DataFrame to the specified output folder
        with pd.ExcelWriter(f'{output_folder}/final_results.xlsx') as writer:
            data_df.to_excel(writer, sheet_name='Aggregated Data', index=False)
            summary_df.to_excel(writer, sheet_name='Summary of Actions', index=False)
        
        print("Data saved to final_results.xlsx")

if __name__ == "__main__":
    main()