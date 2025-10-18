import os
import shutil
import time
import tkinter as tk
from tkinter import ttk, filedialog
from win32com.client import constants
import win32com.client as win32
import pandas as pd
from ttkbootstrap import Style
import logging
import colorlog
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Set up logging with colorlog for console and file logging
handler = logging.StreamHandler()
handler.setFormatter(colorlog.ColoredFormatter(
    "%(log_color)s%(asctime)s - %(levelname)s - %(message)s",
    log_colors={
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'bold_red',
    }
))

file_handler = logging.FileHandler('app.log', mode='w')
file_handler.setFormatter(logging.Formatter(
    "%(asctime)s - %(levelname)s - %(message)s"
))

logging.basicConfig(level=logging.DEBUG, handlers=[handler, file_handler])

def clear_com_cache():
    logging.debug('Clearing COM cache.')
    gen_py_path = os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py')
    if os.path.exists(gen_py_path):
        shutil.rmtree(gen_py_path)
        logging.info('COM cache cleared.')

def open_protected_excel(file_path, temp_file_path, password):
    logging.debug(f'Attempting to open Excel file: {file_path}')
    clear_com_cache()
    excel = None
    try:
        # Try to create Excel application with error handling
        excel = win32.Dispatch('Excel.Application')
        
        # Set properties with error handling
        try:
            excel.DisplayAlerts = False
        except:
            logging.debug('Could not set DisplayAlerts property')
        
        try:
            excel.Visible = False
        except:
            logging.debug('Could not set Visible property')
        
        if password:
            logging.debug('Opening Excel file with password protection.')
            wb = excel.Workbooks.Open(file_path, Password=password)
        else:
            logging.debug('Opening Excel file without password protection.')
            wb = excel.Workbooks.Open(file_path)
        
        wb.SaveAs(temp_file_path, Password='')
        wb.Close(SaveChanges=True)
        logging.info('Excel file opened and saved without password.')
        return wb
    except Exception as e:
        logging.error(f"Failed to open the Excel file: {e}")
        return False
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass

def read_excel_file(file_path, password=None):
    logging.debug(f'Reading Excel file: {file_path}')
    
    # First try to read directly with pandas (for non-password protected files)
    if not password:
        try:
            df = pd.read_excel(file_path)
            logging.info('Excel file read directly into DataFrame.')
            return df
        except Exception as e:
            logging.debug(f'Direct read failed: {e}. Trying COM method.')
    
    # If password protected or direct read failed, use COM method
    temp_file_path = os.path.join(os.environ.get('TEMP'), "~$temp.xlsx")
    logging.debug(f'Temporary file path: {temp_file_path}')
    success = open_protected_excel(file_path, temp_file_path, password)
    if not success:
        logging.warning('Failed to read the Excel file with COM. Trying alternative method.')
        # Try alternative method using msoffcrypto-tool if available
        try:
            import msoffcrypto
            temp_file = msoffcrypto.OfficeFile(open(file_path, 'rb'))
            temp_file.load_key(password=password)
            with open(temp_file_path, 'wb') as f:
                temp_file.decrypt(f)
            df = pd.read_excel(temp_file_path)
            logging.info('Excel file read using msoffcrypto-tool.')
            return df
        except ImportError:
            logging.error('msoffcrypto-tool not available. Please install it with: pip install msoffcrypto-tool')
        except Exception as e:
            logging.error(f'Alternative method also failed: {e}')
        return None
    
    try:
        df = pd.read_excel(temp_file_path)
        logging.info('Excel file read into DataFrame.')
        return df
    except FileNotFoundError:
        logging.error('Error: Input file not found.')
        return None
    finally:
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
            logging.debug('Temporary file removed.')

def check_required_columns(df, required_columns):
    logging.debug('Checking for required columns in DataFrame.')
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        logging.error(f"The following columns are missing in the input file: {missing_columns}")
        return False
    logging.info('All required columns are present.')
    return True

def export_to_excel(df, output_file_path, state_dropdown, contacted_via_dropdown, notes_dropdown, completed_by_dropdown):
    logging.debug(f'Exporting DataFrame to Excel file: {output_file_path}')
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    workbook = writer.book
    workbook.use_zip64()
    workbook.nan_inf_to_errors = True
    df.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']

    logging.debug('Setting up data validation for dropdowns.')
    state_col_letter = chr(ord('A') + df.columns.get_loc('State'))
    state_range = f'{state_col_letter}2:{state_col_letter}{len(df) + 1}'
    worksheet.data_validation(state_range, {'validate': 'list', 'source': state_dropdown})

    contacted_via_col_letter = chr(ord('A') + df.columns.get_loc('Contacted VIA'))
    contacted_via_range = f'{contacted_via_col_letter}2:{contacted_via_col_letter}{len(df) + 1}'
    worksheet.data_validation(contacted_via_range, {'validate': 'list', 'source': contacted_via_dropdown})

    notes_col_letter = chr(ord('A') + df.columns.get_loc('Notes Filed'))
    notes_range = f'{notes_col_letter}2:{notes_col_letter}{len(df) + 1}'
    worksheet.data_validation(notes_range, {'validate': 'list', 'source': notes_dropdown})

    completed_by_col_letter = chr(ord('A') + df.columns.get_loc('Completed By'))
    completed_by_range = f'{completed_by_col_letter}2:{completed_by_col_letter}{len(df) + 1}'
    worksheet.data_validation(completed_by_range, {'validate': 'list', 'source': completed_by_dropdown})

    logging.debug('Setting column widths and formats.')
    for column in df.columns:
        column_width = 16 if column in ['State', 'Contacted VIA', 'Notes Filed', 'Completed By'] else (
            50 if column == 'Notes' else max(df[column].astype(str).map(len).max(), len(column)))
        col_idx = df.columns.get_loc(column)
        worksheet.set_column(col_idx, col_idx, column_width)

    header_format = workbook.add_format({'bold': True, 'bg_color': '#368be9', 'align': 'center', 'valign': 'vcenter'})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    state_format = {
        'Renewal Complete': {'bg_color': '#90EE90'},  # Light green
        'Nowcerts Complete': {'bg_color': '#36bbe9'},  # Light blue
        'Needs Rewritten': {'bg_color': '#EAE455'},  # Light yellow
        'Contact Attempted': {'bg_color': '#9999FF'},  # Light purple
        'Try Bundling': {'bg_color': '#FFB6C1'},  # Light pink
        'Already Rewritten': {'bg_color': '#FFA500'},  # Orange
        'Best Option': {'bg_color': '#DDA0DD'},  # Plum
        'Non Renewing': {'bg_color': '#ff6666'},  # Light red
        'Canceled': {'bg_color': '#A9A9A9'}  # Dark gray
    }

    logging.debug('Applying conditional formatting based on state.')
    for state, format_spec in state_format.items():
        format_ = workbook.add_format(format_spec)
        for row in range(1, len(df) + 1):
            worksheet.conditional_format(f'A{row + 1}:{chr(ord("A") + len(df.columns) - 1)}{row + 1}',
                                         {'type': 'formula',
                                          'criteria': f'INDIRECT("${state_col_letter}${row + 1}")="{state}"',
                                          'format': format_})

    # Format for Notes Filed and Contacted VIA columns (center aligned)
    center_format = workbook.add_format({'align': 'center'})
    notes_col_index = df.columns.get_loc('Notes Filed')
    contacted_via_col_index = df.columns.get_loc('Contacted VIA')
    
    for row in range(1, len(df) + 1):
        # Notes Filed column
        cell_value = df.iloc[row - 1, notes_col_index]
        worksheet.write(row, notes_col_index, cell_value, center_format)
        
        # Contacted VIA column
        cell_value = df.iloc[row - 1, contacted_via_col_index]
        worksheet.write(row, contacted_via_col_index, cell_value, center_format)

    grey_format = workbook.add_format({'bg_color': '#f0f0f0'})
    for row in range(1, len(df) + 1, 2):
        for col in range(len(df.columns)):
            if col == notes_col_index or col == contacted_via_col_index:
                cell_format = workbook.add_format({'bg_color': '#f0f0f0', 'align': 'center'})
            else:
                cell_format = grey_format
            worksheet.write(row, col, df.iloc[row - 1, col], cell_format)

    writer.close()
    logging.info(f'DataFrame exported to {output_file_path}')
    return True

def update_count_label(label, count):
    label.config(text=f"Files processed: {count}")
    logging.info(f'Updated count label to: Files processed: {count}')

def process_excel():
    logging.debug('Starting Excel processing.')
    status_var.set("Processing... Please wait")
    progress_var.set(0)
    root.update_idletasks()

    input_file_path = source_var.get()
    output_folder_path = destination_var.get()
    
    # Get password from environment variable
    password = (
        os.getenv('EXCEL_PASSWORD') or  # Primary environment variable
        os.getenv('QUADSTATE_PASSWORD') or  # Alternative environment variable
        None  # No UI fallback since password field is removed
    )
    
    # Update status bar based on password source
    if os.getenv('EXCEL_PASSWORD'):
        status_var.set("Ready - Using password from EXCEL_PASSWORD environment variable")
        logging.debug('Using password from EXCEL_PASSWORD environment variable')
    elif os.getenv('QUADSTATE_PASSWORD'):
        status_var.set("Ready - Using password from QUADSTATE_PASSWORD environment variable")
        logging.debug('Using password from QUADSTATE_PASSWORD environment variable')
    else:
        status_var.set("Warning - No password found in environment variables")
        logging.warning('No password found in environment variables')
    
    # Update progress
    progress_var.set(10)
    root.update_idletasks()

    logging.debug(f'Input file path: {input_file_path}')
    logging.debug(f'Output folder path: {output_folder_path}')

    try:
        df = read_excel_file(input_file_path, password)
        if df is None:
            status_var.set("Error - Incorrect password or file cannot be read")
            progress_var.set(0)
            root.update_idletasks()
            return
        
        # Update progress
        progress_var.set(30)
        root.update_idletasks()

        required_columns = ['Expiration Date', 'Insured', 'Carrier',
                            'Lines Of Business', 'Status', 'Premium', 'Renewal Premium', 'Percentage Change']
        if not check_required_columns(df, required_columns):
            status_var.set("Error - Missing required columns in Excel file")
            progress_var.set(0)
            root.update_idletasks()
            return

        # Update progress
        progress_var.set(50)
        status_var.set("Processing - Preparing data...")
        root.update_idletasks()
        
        logging.debug('Renaming and selecting required columns.')
        df.rename(columns={'Insured': 'Insured Name'}, inplace=True)
        df = df[['Expiration Date', 'Insured Name', 'Carrier', 'Lines Of Business', 'Status', 'Premium',
                 'Renewal Premium', 'Percentage Change']]
        df['State'] = ""
        df['Contacted VIA'] = ""
        df['Notes Filed'] = ""
        df['Completed By'] = ""

        state_dropdown = ['Renewal Complete', 'Nowcerts Complete', 'Needs Rewritten', 'Contact Attempted', 'Try Bundling', 'Already Rewritten', 'Best Option', 'Non Renewing', 'Canceled']
        contacted_via_dropdown = ['Left VM', 'Sent Text', 'Sent Email', 'Spoke to']
        notes_dropdown = ['Yes', 'Call Filed in AMS']
        completed_by_dropdown = ['Danielle Stevens', 'Amber Miller', 'Teresa Morrisette', 'Lane Ross']

        # Update progress
        progress_var.set(70)
        status_var.set("Processing - Sorting data and creating Excel file...")
        root.update_idletasks()
        
        df.sort_values(by='Expiration Date', inplace=True)
        
        # Update progress
        progress_var.set(85)
        root.update_idletasks()
        
        output_file_path = os.path.join(output_folder_path, f"Updated_Renewals_{time.strftime('%Y%m%d-%H%M%S')}.xlsx")
        if export_to_excel(df, output_file_path, state_dropdown, contacted_via_dropdown, notes_dropdown, completed_by_dropdown):
            # Complete progress bar
            progress_var.set(100)
            status_var.set(f"Success - Processed {len(df)} records. File saved to Desktop.")
            update_count_label(count_label, len(df))
            root.update_idletasks()
            time.sleep(3)
            root.destroy()
    except Exception as e:
        logging.error(f"Error: {e}")
        progress_var.set(0)
        status_var.set(f"Error - {str(e)}")
        root.update_idletasks()

def select_source_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    source_var.set(file_path)
    logging.debug(f'Source file selected: {file_path}')

def select_destination_folder():
    folder_path = filedialog.askdirectory()
    destination_var.set(folder_path)
    logging.debug(f'Destination folder selected: {folder_path}')

root = tk.Tk()
root.title("Quadstate Renewal Processor")
root.geometry('450x400')

style = Style(theme='flatly')

logging.debug('Setting default file paths.')
if os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Copy of Export_RenewalCenter.xlsx")):
    default_input_file = os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Copy of Export_RenewalCenter.xlsx")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Export_RenewalCenter.xlsx")):
    default_input_file = os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Export_RenewalCenter.xlsx")
else:
    default_input_file = ''

if os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - quadstateinsurance.com", "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - quadstateinsurance.com", "Desktop")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "Desktop")
else:
    default_output_folder = ''

source_var = tk.StringVar(value=default_input_file)
source_label = ttk.Label(root, text="Select source file:")
source_label.pack(pady=(10, 0))
source_entry = ttk.Entry(root, textvariable=source_var, width=70)
source_entry.pack()
source_button = ttk.Button(root, text="Browse", command=select_source_file)
source_button.pack(pady=5)

destination_var = tk.StringVar(value=default_output_folder)
destination_label = ttk.Label(root, text="Select destination folder:")
destination_label.pack(pady=(10, 0))
destination_entry = ttk.Entry(root, textvariable=destination_var, width=70)
destination_entry.pack()
destination_button = ttk.Button(root, text="Browse", command=select_destination_folder)
destination_button.pack(pady=5)

# Status bar instead of password field
status_var = tk.StringVar(value="Ready - Password loaded from environment")
status_label = ttk.Label(root, textvariable=status_var, relief="sunken", anchor="w")
# status_label.pack(fill="x", pady=(10, 0), padx=10)

# Progress bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400, mode='determinate')
progress_bar.pack(pady=(10, 10), padx=10, fill="x")

count_label = ttk.Label(root, text="Files processed: 0")
count_label.pack(pady=(10, 0))

style.theme_use('superhero')

process_button = ttk.Button(root, text="Process", command=process_excel)
process_button.pack(pady=10)

root.mainloop()
