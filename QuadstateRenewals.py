import os
import shutil
import sys
import time
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog


def ensure_com_cache_dir():
    local_app_data = os.environ.get('LOCALAPPDATA')
    if not local_app_data:
        return

    gen_py_path = os.path.join(
        local_app_data,
        'Temp',
        'gen_py',
        f'{sys.version_info.major}.{sys.version_info.minor}'
    )
    os.makedirs(gen_py_path, exist_ok=True)


ensure_com_cache_dir()
import win32com.client as win32
import pandas as pd
import logging
import colorlog
from dotenv import load_dotenv

# Load environment variables from .env file
# Check multiple locations for .env file (in order of preference)
env_locations = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'),  # Same folder as executable
    'C:\\QuadstateRenewalsProperties\\.env',  # Dedicated properties folder on C drive
    os.path.join(os.path.expanduser('~'), '.env'),  # User home directory
    os.path.join(os.path.expanduser('~'), 'Documents', '.env'),  # Documents folder
    os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'QuadstateRenewals', '.env'),  # AppData folder
]

env_loaded = False
for env_path in env_locations:
    if os.path.exists(env_path):
        load_dotenv(env_path)
        env_loaded = True
        break

if not env_loaded:
    # Fallback to default behavior
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
    local_app_data = os.environ.get('LOCALAPPDATA')
    if not local_app_data:
        logging.warning('LOCALAPPDATA is not set; skipping COM cache clear.')
        return

    gen_py_path = os.path.join(local_app_data, 'Temp', 'gen_py')
    if os.path.exists(gen_py_path):
        shutil.rmtree(gen_py_path)
        logging.info('COM cache cleared.')
    ensure_com_cache_dir()

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
            try:
                wb = excel.Workbooks.Open(file_path, Password=password)
            except Exception as e:
                logging.debug(f'Failed to open with password: {e}')
                # Try without password as fallback
                try:
                    wb = excel.Workbooks.Open(file_path)
                except Exception as e2:
                    logging.debug(f'Failed to open without password: {e2}')
                    raise e  # Re-raise the original password error
        else:
            logging.debug('Opening Excel file without password protection.')
            wb = excel.Workbooks.Open(file_path)
        
        if wb is None:
            logging.error("Failed to open workbook")
            return False
            
        wb.SaveAs(temp_file_path, Password='')
        wb.Close(SaveChanges=True)
        logging.info('Excel file opened and saved without password.')
        return True
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
    
    # Try direct pandas read first (most reliable for non-password protected files)
    if not password:
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            logging.info('Excel file read directly into DataFrame.')
            return df
        except Exception as e:
            logging.debug(f'Direct read failed: {e}. Trying password-protected method.')
    
    # For password-protected files, use msoffcrypto-tool (more reliable than COM)
    temp_file_path = os.path.join(os.environ.get('TEMP'), f"~$temp_{int(time.time())}.xlsx")
    logging.debug(f'Temporary file path: {temp_file_path}')
    
    try:
        import msoffcrypto
        with open(file_path, 'rb') as f:
            temp_file = msoffcrypto.OfficeFile(f)
            if password:
                temp_file.load_key(password=password)
            else:
                temp_file.load_key()
            with open(temp_file_path, 'wb') as decrypted_file:
                temp_file.decrypt(decrypted_file)
        
        df = pd.read_excel(temp_file_path, engine='openpyxl')
        logging.info('Excel file read using msoffcrypto-tool.')
        return df
        
    except ImportError:
        logging.error('msoffcrypto-tool not available. Please install it with: pip install msoffcrypto-tool')
        return None
    except Exception as e:
        logging.error(f'msoffcrypto-tool method failed: {e}')
        return None
    finally:
        # Clean up temporary file
        if os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
                logging.debug('Temporary file removed.')
            except Exception as e:
                logging.warning(f'Could not remove temporary file: {e}')

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
        'Rewritten': {'bg_color': '#a7754d'},  # Gilmore Girl Brown
        'Contact Attempted': {'bg_color': '#9999FF'},  # Light purple
        'Try Bundling': {'bg_color': '#FFB6C1'},  # Light pink
        'Already Rewritten': {'bg_color': '#FFA500'},  # Orange
        'Best Option': {'bg_color': '#DDA0DD'},  # Plum
        'Non Renewing': {'bg_color': '#ff6666'},  # Light red
        'Canceled': {'bg_color': '#A9A9A9'}  # Dark gray
    }

    logging.debug('Applying conditional formatting based on state.')
    # Apply conditional formatting to the entire data range
    data_range = f'A2:{chr(ord("A") + len(df.columns) - 1)}{len(df) + 1}'
    
    for state, format_spec in state_format.items():
        format_ = workbook.add_format(format_spec)
        worksheet.conditional_format(data_range,
                                     {'type': 'formula',
                                      'criteria': f'${state_col_letter}2="{state}"',
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
    label.configure(text=f"Files processed: {count}")
    logging.info(f'Updated count label to: Files processed: {count}')

def set_progress(percent):
    progress_bar.set(percent / 100)
    root.update_idletasks()

def process_excel():
    logging.debug('Starting Excel processing.')
    status_var.set("Processing... Please wait")
    set_progress(0)

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
    set_progress(10)

    logging.debug(f'Input file path: {input_file_path}')
    logging.debug(f'Output folder path: {output_folder_path}')

    try:
        df = read_excel_file(input_file_path, password)
        if df is None:
            status_var.set("Error - Incorrect password or file cannot be read")
            set_progress(0)
            return
        
        # Update progress
        set_progress(30)

        required_columns = ['Expiration Date', 'Insured', 'Carrier',
                            'Lines Of Business', 'Status', 'Premium', 'Renewal Premium', 'Percentage Change']
        if not check_required_columns(df, required_columns):
            status_var.set("Error - Missing required columns in Excel file")
            set_progress(0)
            return

        # Update progress
        status_var.set("Processing - Preparing data...")
        set_progress(50)
        
        logging.debug('Renaming and selecting required columns.')
        df.rename(columns={'Insured': 'Insured Name'}, inplace=True)
        df = df[['Expiration Date', 'Insured Name', 'Carrier', 'Lines Of Business', 'Status', 'Premium',
                 'Renewal Premium', 'Percentage Change']]
        df['State'] = ""
        df['Contacted VIA'] = ""
        df['Notes Filed'] = ""
        df['Completed By'] = ""

        state_dropdown = ['Renewal Complete', 'Nowcerts Complete', 'Needs Rewritten', 'Rewritten', 'Contact Attempted', 'Try Bundling', 'Already Rewritten', 'Best Option', 'Non Renewing', 'Canceled']
        contacted_via_dropdown = ['Left VM', 'Sent Text', 'Sent Email', 'Spoken with']
        notes_dropdown = ['Yes', 'Call Filed in AMS']
        completed_by_dropdown = ['Danielle Stevens', 'Amber Miller', 'Teresa Morrisette', 'Jillian Stevens']

        # Update progress
        status_var.set("Processing - Sorting data and creating Excel file...")
        set_progress(70)
        
        df.sort_values(by='Expiration Date', inplace=True)
        
        # Update progress
        set_progress(85)
        
        output_file_path = os.path.join(output_folder_path, f"Updated_Renewals_{time.strftime('%Y%m%d-%H%M%S')}.xlsx")
        if export_to_excel(df, output_file_path, state_dropdown, contacted_via_dropdown, notes_dropdown, completed_by_dropdown):
            # Complete progress bar
            status_var.set(f"Success - Processed {len(df)} records. File saved to Desktop.")
            set_progress(100)
            update_count_label(count_label, len(df))
            root.update_idletasks()
            time.sleep(3)
            root.destroy()
    except Exception as e:
        logging.error(f"Error: {e}")
        status_var.set(f"Error - {str(e)}")
        set_progress(0)

def select_source_file():
    current_source = source_var.get()
    # Use the directory of the current source file, or Downloads if empty
    initial_dir = os.path.dirname(current_source) if current_source and os.path.exists(current_source) else os.path.expanduser("~/Downloads")
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("Excel 2007+ files", "*.xlsx"),
            ("Excel 97-2003 files", "*.xls"),
            ("All files", "*.*")
        ],
        initialdir=initial_dir
    )
    if file_path:  # Only update if a file was actually selected
        source_var.set(file_path)
        logging.debug(f'Source file selected: {file_path}')

def select_destination_folder():
    current_destination = destination_var.get()
    # Use the current destination folder as initial directory, or Desktop if empty
    initial_dir = current_destination if current_destination and os.path.exists(current_destination) else os.path.expanduser("~/Desktop")
    
    folder_path = filedialog.askdirectory(
        title="Select Destination Folder",
        initialdir=initial_dir
    )
    if folder_path:  # Only update if a folder was actually selected
        destination_var.set(folder_path)
        logging.debug(f'Destination folder selected: {folder_path}')

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.title("Quadstate Renewal Processor")
root.geometry('640x330')
root.minsize(560, 330)
root.grid_columnconfigure(0, weight=1)

logging.debug('Setting default file paths.')
if os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Copy of Export_RenewalCenter.xlsx")):
    default_input_file = os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Copy of Export_RenewalCenter.xlsx")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Export_RenewalCenter.xlsx")):
    default_input_file = os.path.join("C:\\", "Users", os.getlogin(), "Downloads", "Export_RenewalCenter.xlsx")
else:
    default_input_file = ''

if os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - quadstateinsurance.com", "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - quadstateinsurance.com", "Desktop")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - Quadstate Insurance Agency LLC", "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "OneDrive - Quadstate Insurance Agency LLC", "Desktop")
elif os.path.exists(os.path.join("C:\\", "Users", os.getlogin(), "Desktop")):
    default_output_folder = os.path.join("C:\\", "Users", os.getlogin(), "Desktop")
else:
    default_output_folder = ''

source_var = tk.StringVar(value=default_input_file)
source_label = ctk.CTkLabel(root, text="Select Source File:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w", justify="left")
source_label.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 4))
source_row = ctk.CTkFrame(root, fg_color="transparent")
source_row.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 8))
source_row.grid_columnconfigure(0, weight=1)
source_entry = ctk.CTkEntry(source_row, textvariable=source_var)
source_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
source_button = ctk.CTkButton(source_row, text="Browse Source", width=140, command=select_source_file)
source_button.grid(row=0, column=1)

destination_var = tk.StringVar(value=default_output_folder)
destination_label = ctk.CTkLabel(root, text="Select Destination Folder:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w", justify="left")
destination_label.grid(row=2, column=0, sticky="ew", padx=24, pady=(4, 4))
destination_row = ctk.CTkFrame(root, fg_color="transparent")
destination_row.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 10))
destination_row.grid_columnconfigure(0, weight=1)
destination_entry = ctk.CTkEntry(destination_row, textvariable=destination_var)
destination_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
destination_button = ctk.CTkButton(destination_row, text="Browse Destination", width=160, command=select_destination_folder)
destination_button.grid(row=0, column=1)

# Status bar instead of password field
status_var = tk.StringVar(value="Ready - Password loaded from environment")
status_label = ctk.CTkLabel(root, textvariable=status_var, anchor="w", justify="left")
status_label.grid(row=4, column=0, sticky="ew", padx=24, pady=(4, 0))

# Progress bar
progress_bar = ctk.CTkProgressBar(root, mode='determinate')
progress_bar.set(0)
progress_bar.grid(row=5, column=0, sticky="ew", padx=24, pady=(10, 8))

count_label = ctk.CTkLabel(root, text="Files processed: 0")
count_label.grid(row=6, column=0, pady=(2, 6))

process_button = ctk.CTkButton(root, text="Process", width=180, command=process_excel)
process_button.grid(row=7, column=0, pady=(0, 16))

root.mainloop()
