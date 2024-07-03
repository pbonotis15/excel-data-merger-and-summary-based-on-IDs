
# Excel Data Aggregator and Summarizer

## Overview

This project is a Python-based tool designed to automate the aggregation and processing of multiple Excel files. It provides a user-friendly interface for selecting input and output directories, reads and concatenates data from multiple Excel sheets, standardizes column names, filters data based on specified criteria, and merges relevant information. Additionally, it removes duplicate entries based on the 'SR ID' and 'Ημερομηνία Δημιουργίας' columns, ensuring that only the most recent entry for each 'SR ID' is retained. The processed data is then saved into a new Excel file with multiple sheets: 'Aggregated Data', 'Summary of Actions', and 'Last Drop'.

### File Descriptions

#### `summarize_and_merge.py`

This script contains the core logic for processing the Excel files. It includes functions to read Excel sheets, concatenate data, standardize column names, filter data, merge relevant information, and remove duplicate entries.

**Functions:**
- `get_sheet_names(file_path)`: Returns the sheet names from the specified Excel file.
- `main(file_path1, file_path2, sheet_name, output_folder)`: Main function to process the Excel files and save the aggregated data.

#### `wrapper.py`

This script provides a graphical user interface (GUI) for selecting input files and output directories. It uses the `tkinter` library to prompt the user to select Excel files and specify the output folder.

**Functions:**
- `get_file_path()`: Opens a file dialog to select an Excel file.
- `get_excel_sheet(file_path)`: Prompts the user to select a sheet from the specified Excel file.
- `get_output_folder()`: Opens a dialog to select the output folder.

## Installation

To use this project, clone the repository and navigate to the project directory:

```sh
git clone https://github.com/pbonotis15/excel-data-merger-and-summary-based-on-IDs.git
cd your-repo-name
```

Ensure you have Python installed. It's recommended to use a virtual environment to manage dependencies:

```sh
python3 -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
```

Then, install the dependencies using:

```sh
pip install -r requirements.txt
```

## Usage

To run the tool, use the following command:

```sh
python wrapper.py
```

This will open a GUI to select the input Excel files and the output folder. Follow the prompts to complete the data processing.

## Contributing

If you'd like to contribute to this project, please fork the repository and use a feature branch. Pull requests are warmly welcome.