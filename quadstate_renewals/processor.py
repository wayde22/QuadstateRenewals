import logging
import os
import time
from dataclasses import dataclass
from typing import Callable, Optional

from .config import get_excel_password
from .constants import OUTPUT_COLUMNS, REQUIRED_COLUMNS
from .excel_reader import read_excel_file
from .excel_writer import export_to_excel


StatusCallback = Callable[[str], None]
ProgressCallback = Callable[[int], None]


@dataclass
class ProcessResult:
    success: bool
    record_count: int = 0
    output_file_path: Optional[str] = None
    error: Optional[str] = None


def _set_status(callback, message):
    if callback:
        callback(message)


def _set_progress(callback, percent):
    if callback:
        callback(percent)


def check_required_columns(df, required_columns):
    logging.debug('Checking for required columns in DataFrame.')
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        logging.error(f"The following columns are missing in the input file: {missing_columns}")
        return False
    logging.info('All required columns are present.')
    return True


def prepare_renewals_dataframe(df):
    logging.debug('Renaming and selecting required columns.')
    df = df.rename(columns={'Insured': 'Insured Name'})
    df = df[OUTPUT_COLUMNS].copy()
    df['State'] = ""
    df['Contacted VIA'] = ""
    df['Notes Filed'] = ""
    df['Completed By'] = ""
    return df


def process_renewals(input_file_path, output_folder_path, on_status=None, on_progress=None):
    logging.debug('Starting Excel processing.')
    _set_status(on_status, "Processing... Please wait")
    _set_progress(on_progress, 0)

    password, password_status, password_log, is_password_warning = get_excel_password()
    _set_status(on_status, password_status)

    if is_password_warning:
        logging.warning(password_log)
    else:
        logging.debug(password_log)

    _set_progress(on_progress, 10)

    logging.debug(f'Input file path: {input_file_path}')
    logging.debug(f'Output folder path: {output_folder_path}')

    try:
        df = read_excel_file(input_file_path, password)
        if df is None:
            message = "Error - Incorrect password or file cannot be read"
            _set_status(on_status, message)
            _set_progress(on_progress, 0)
            return ProcessResult(success=False, error=message)

        _set_progress(on_progress, 30)

        if not check_required_columns(df, REQUIRED_COLUMNS):
            message = "Error - Missing required columns in Excel file"
            _set_status(on_status, message)
            _set_progress(on_progress, 0)
            return ProcessResult(success=False, error=message)

        _set_status(on_status, "Processing - Preparing data...")
        _set_progress(on_progress, 50)

        df = prepare_renewals_dataframe(df)

        _set_status(on_status, "Processing - Sorting data and creating Excel file...")
        _set_progress(on_progress, 70)
        df.sort_values(by='Expiration Date', inplace=True)
        _set_progress(on_progress, 85)

        output_file_path = os.path.join(
            output_folder_path,
            f"Updated_Renewals_{time.strftime('%Y%m%d-%H%M%S')}.xlsx"
        )

        if export_to_excel(df, output_file_path):
            message = f"Success - Processed {len(df)} records. File saved to Desktop."
            _set_status(on_status, message)
            _set_progress(on_progress, 100)
            return ProcessResult(
                success=True,
                record_count=len(df),
                output_file_path=output_file_path,
            )

        message = "Error - Excel export failed"
        _set_status(on_status, message)
        _set_progress(on_progress, 0)
        return ProcessResult(success=False, error=message)
    except Exception as e:
        logging.error(f"Error: {e}")
        message = f"Error - {str(e)}"
        _set_status(on_status, message)
        _set_progress(on_progress, 0)
        return ProcessResult(success=False, error=message)
