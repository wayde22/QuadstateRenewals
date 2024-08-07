# Quadstate Renewal Processor

## Overview

The Quadstate Renewal Processor is a Python-based desktop application designed to streamline the processing and management of renewal data in Excel files. It provides an intuitive interface for selecting and processing Excel files, applying data validation, conditional formatting, and exporting the processed data to a new Excel file.

## Features

- **Password-Protected Excel Files**: The application can open and process Excel files that are password-protected.
- **Data Validation**: Adds dropdown lists for specific columns to ensure data consistency.
- **Conditional Formatting**: Applies color coding to rows based on the selected state, including a dark gray color for the "Canceled" state.
- **Logging**: Provides detailed logging with colored output to the console and logs errors to a file.
- **Batch Processing**: Process multiple files in sequence and keep track of how many have been processed.
- **User-Friendly Interface**: Built with `Tkinter` and `ttkbootstrap` for a clean, responsive GUI.

## How It Works

1. **File Selection**: Users can select an Excel file and an output directory using the file and directory dialog.
2. **Password Input**: If the file is password-protected, users can enter the password in the application.
3. **Data Reading**: The application reads the Excel file into a Pandas DataFrame, ensuring the necessary columns are present.
4. **Data Processing**:
   - Adds columns for "State," "Notes Filed," and "Completed By," initialized with empty values.
   - Adds dropdown validation for the new columns.
   - Applies conditional formatting based on the state.
5. **Exporting**: The processed DataFrame is exported to a new Excel file in the selected output directory.
6. **Feedback**: The application updates the user on the number of files processed and logs any errors or warnings.

## Usage Instructions

1. **Launch the Application**: Run the script to open the GUI.
2. **Select the Source File**: Click the "Browse" button to select the Excel file you wish to process.
3. **Select the Destination Folder**: Click the "Browse" button to choose where the processed file should be saved.
4. **Enter Password**: If the Excel file is password-protected, enter the password in the provided field.
5. **Process the File**: Click the "Process" button to begin processing the selected Excel file.
6. **View Processed File**: The processed file will be saved in the destination folder with a timestamped filename.

## Technical Details

- **Libraries Used**:
  - `os` and `shutil`: For file and directory operations.
  - `tkinter` and `ttk`: For the GUI interface.
  - `pandas`: For data manipulation.
  - `win32com.client`: For interacting with Excel files, particularly handling password-protected files.
  - `ttkbootstrap`: For modern styling of the Tkinter interface.
  - `colorlog`: For colored logging output to the console.
  - `xlsxwriter`: For writing the processed DataFrame back to an Excel file.

- **Logging**: Logs are written to both the console and a file (`app.log`) for detailed debugging and error tracking. The console logs use color coding to differentiate log levels.

- **Conditional Formatting**: Applies specific background colors to rows based on the value of the "State" column, making it easier to visually identify the status of each entry.

## System Requirements

- **Python Version**: Requires Python 3.x.
- **Dependencies**: Ensure all required Python libraries are installed. You can install them using pip:
  ```sh
  pip install pandas ttkbootstrap pywin32 colorlog xlsxwriter

