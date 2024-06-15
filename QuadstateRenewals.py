import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import os
import time
import win32com.client as win32
from ttkbootstrap import Style


def open_protected_excel(file_path, temp_file_path, password):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False

    try:
        # Open the workbook with the password (if provided)
        if password:
            wb = excel.Workbooks.Open(file_path, Password=password)
        else:
            wb = excel.Workbooks.Open(file_path)

        if not wb:
            print(f'No workbook!')
            return None

        # Save the temporary file
        wb.SaveAs(temp_file_path, Password='')
        wb.Close(SaveChanges=True)
        return wb
    except Exception as e:
        print("Failed to open the Excel file:", e)
        return False
    finally:
        excel.Quit()


# Function to read Excel file
def read_excel_file(file_path, password=None):
    # Retrieve the path to the temporary unlocked file
    temp_file_path = os.path.join(os.path.dirname(file_path), "~$temp.xlsx")
    success = open_protected_excel(file_path, temp_file_path, password)

    if not success:
        print(f'Unlocking Excel file unsuccessful!')
        return None

    try:
        file_path = temp_file_path
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print("Error: Input file not found.")
        return None
    finally:
        print(f'Removing temp file!')
        # Remove the temporary file if exists
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)


# Function to check required columns
def check_required_columns(df, required_columns):
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print("Error: The following columns are missing in the input file:", missing_columns)
        return False
    return True


# Function to export DataFrame to Excel
def export_to_excel(df, output_file_path, state_dropdown, notes_dropdown):
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Apply data validation to the 'State' column
    state_col_letter = chr(ord('A') + df.columns.get_loc('State'))  # Get the column letter for 'State'
    state_range = f'{state_col_letter}2:{state_col_letter}{len(df) + 1}'  # Range for dropdown
    worksheet.data_validation(state_range, {'validate': 'list', 'source': state_dropdown})

    # Apply data validation to the 'Notes Filed' column
    notes_col_letter = chr(ord('A') + df.columns.get_loc('Notes Filed'))  # Get the column letter for 'Notes Filed'
    notes_range = f'{notes_col_letter}2:{notes_col_letter}{len(df) + 1}'  # Range for dropdown
    worksheet.data_validation(notes_range, {'validate': 'list', 'source': notes_dropdown})

    # Adjust column widths
    for column in df.columns:
        column_width = 16 if column in ['State', 'Notes Filed'] else (
            50 if column == 'Notes' else max(df[column].astype(str).map(len).max(), len(column)))
        col_idx = df.columns.get_loc(column)
        worksheet.set_column(col_idx, col_idx, column_width)

    # Header format - Blue color
    header_format = workbook.add_format({'bold': True, 'bg_color': '#368be9', 'align': 'center', 'valign': 'vcenter'})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Conditional formatting based on 'State' column
    state_format = {
        'Renewal Complete': {'bg_color': '#90EE90'},  # #FF7F7F
        'Nowcerts Complete': {'bg_color': '#36bbe9'},
        'Needs Rewritten': {'bg_color': '#EAE455'},
        'Needs Spoke To': {'bg_color': '#9999FF'},
        'Non Renewing': {'bg_color': '#ff6666'}  # No specific color for 'Non-Renewing'
    }

    for state, format_spec in state_format.items():
        format_ = workbook.add_format(format_spec)
        for row in range(1, len(df) + 1):
            worksheet.conditional_format(f'A{row + 1}:{chr(ord("A") + len(df.columns) - 1)}{row + 1}',
                                         {'type': 'formula',
                                          'criteria': f'INDIRECT("${state_col_letter}${row + 1}")="{state}"',
                                          'format': format_})

    # Center the items in the 'Notes Filed' column
    notes_format = workbook.add_format({'align': 'center'})
    notes_col_index = df.columns.get_loc('Notes Filed')
    for row in range(1, len(df) + 1):
        cell_value = df.iloc[row - 1, notes_col_index]
        worksheet.write(row, notes_col_index, cell_value, notes_format)

    grey_format = workbook.add_format({'bg_color': '#f0f0f0'})
    for row in range(1, len(df) + 1, 2):
        for col in range(len(df.columns)):
            cell_format = grey_format if col != notes_col_index else \
                workbook.add_format({'bg_color': '#f0f0f0', 'align': 'center'})
            worksheet.write(row, col, df.iloc[row - 1, col], cell_format)

    writer.close()  # Use close() instead of save()
    return True


# Function to update count label
def update_count_label(label, count):
    label.config(text=f"Files processed: {count}")


# Function to process the Excel file
def process_excel():
    incorrect_password_label.config(text="")
    root.update_idletasks()

    input_file_path = source_var.get()
    output_folder_path = destination_var.get()
    password = password_var.get()  # Retrieve password from the UI

    try:
        # Read Excel file
        df = read_excel_file(input_file_path, password)  # Pass the password

        if df is None:
            incorrect_password_label.config(text="Incorrect password. Please try again.....")
            root.update_idletasks()
            return

        required_columns = ['Expiration Date', 'Insured First Name', 'Insured Last Name', 'Carrier',
                            'Lines Of Business', 'Status', 'Premium', 'Renewal Premium', 'Percentage Change']

        # Check required columns
        if not check_required_columns(df, required_columns):
            return

        # Rename columns
        df.rename(columns={'Insured First Name': 'First Name', 'Insured Last Name': 'Last Name'}, inplace=True)

        # Selecting only the required columns
        df = df[['Expiration Date', 'First Name', 'Last Name', 'Carrier', 'Lines Of Business', 'Status', 'Premium',
                 'Renewal Premium', 'Percentage Change']]

        # Add 'State' and 'Notes Filed' columns
        df['State'] = ""
        df['Notes Filed'] = ""

        # Define the dropdown options for 'State' column
        state_dropdown = ['Renewal Complete', 'Nowcerts Complete', 'Needs Rewritten', 'Needs Spoke To', 'Non Renewing']

        # Define the dropdown options for 'Notes Filed' column
        notes_dropdown = ['Yes', 'No', 'Left VM', 'Sent Email']

        # Sort the DataFrame by 'Expiration Date' column in ascending order
        df.sort_values(by='Expiration Date', inplace=True)

        # Export to a new Excel file
        output_file_path = os.path.join(output_folder_path, f"Updated_Renewals_{time.strftime('%Y%m%d-%H%M%S')}.xlsx")
        if export_to_excel(df, output_file_path, state_dropdown, notes_dropdown):
            update_count_label(count_label, len(df))
            root.update_idletasks()
            time.sleep(3)
            root.destroy()

    except Exception as e:
        print(f"Error: Incorrect password, {e}")
        incorrect_password_label.config(
            text="Incorrect password. Please, try again.")  # Update label text with error message
        root.update_idletasks()


# Function to select source file
def select_source_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    source_var.set(file_path)


# Function to select destination folder
def select_destination_folder():
    folder_path = filedialog.askdirectory()
    destination_var.set(folder_path)


# Create tkinter window
root = tk.Tk()
root.title("Quadstate Renewal Processor")
root.geometry('450x400')  # Increased height for the password field

# Create ttkbootstrap style
style = Style(theme='flatly')

# Default file paths
if os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Copy of Export_RenewalCenter.xlsx")):
    default_input_file = os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Copy of Export_RenewalCenter.xlsx")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Export_RenewalCenter.xlsx")):
    default_input_file = os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Export_RenewalCenter.xlsx")
else:
    print("No file named 'Copy of Export_RenewalCenter.xlsx' or Export_RenewalCenter.xlsx")

if os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - quadstateinsurance.com", "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - quadstateinsurance.com", "Desktop")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "Desktop")

# Create source file input
source_var = tk.StringVar(value=default_input_file)
source_label = ttk.Label(root, text="Select source file:")
source_label.pack(pady=(10, 0))
source_entry = ttk.Entry(root, textvariable=source_var, width=70)
source_entry.pack()
source_button = ttk.Button(root, text="Browse", command=select_source_file)
source_button.pack(pady=5)

# Create destination folder input
destination_var = tk.StringVar(value=default_output_folder)
destination_label = ttk.Label(root, text="Select destination folder:")
destination_label.pack(pady=(10, 0))
destination_entry = ttk.Entry(root, textvariable=destination_var, width=70)
destination_entry.pack()
destination_button = ttk.Button(root, text="Browse", command=select_destination_folder)
destination_button.pack(pady=5)

# Create password input field
password_var = tk.StringVar()
password_label = ttk.Label(root, text="Enter file password (if any):")
password_label.pack(pady=(10, 0))
password_entry = ttk.Entry(root, textvariable=password_var, show="*")
password_entry.pack(pady=5)

# Create count label
count_label = ttk.Label(root, text="Files processed: 0")
count_label.pack(pady=(10, 0))

# Incorrect password label
incorrect_password_label = ttk.Label(root, text="", style="danger.TLabel")
incorrect_password_label.pack(pady=(5, 0))

style.theme_use('superhero')

# Create process button
process_button = ttk.Button(root, text="Process", command=process_excel)
process_button.pack(pady=10)

# Start the tkinter event loop
root.mainloop()

