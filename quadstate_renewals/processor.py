import logging
import os
import time
from dataclasses import dataclass
from typing import Callable, Optional

from .config import get_excel_password
from .constants import OUTPUT_COLUMNS, REQUIRED_COLUMNS
from .dependencies import MissingDependencyError
from .excel_reader import MissingExcelPasswordError, read_excel_file
from .excel_writer import export_to_excel


StatusCallback = Callable[[str], None]
ProgressCallback = Callable[[int], None]
MISSING_COLUMNS_STATUS_LIMIT = 3


@dataclass
class ProcessResult:
    success: bool
    record_count: int = 0
    worksheet_row_count: int = 0
    output_file_path: Optional[str] = None
    error: Optional[str] = None


def _set_status(callback: Optional[StatusCallback], message: str):
    if callback:
        callback(message)


def _set_progress(callback: Optional[ProgressCallback], percent: int):
    if callback:
        callback(percent)


def get_missing_required_columns(df, required_columns):
    return [col for col in required_columns if col not in df.columns]


def format_missing_columns_message(missing_columns):
    visible_columns = missing_columns[:MISSING_COLUMNS_STATUS_LIMIT]
    visible_text = ", ".join(visible_columns)

    if len(missing_columns) <= MISSING_COLUMNS_STATUS_LIMIT:
        return f"Error - Source file format changed. Missing columns: {visible_text}"

    remaining_count = len(missing_columns) - MISSING_COLUMNS_STATUS_LIMIT
    return (
        "Error - Source file format changed. "
        f"Missing columns: {visible_text}, and {remaining_count} more. See app.log."
    )


def format_missing_dependency_message(error):
    return f"Error - Missing dependency: {error.package_name}. See app.log."


def format_missing_password_message():
    return "Error - Password required for protected source file. See app.log."


def prepare_renewals_dataframe(df):
    logging.debug('Renaming and selecting required columns.')
    df = df.rename(columns={'Insured': 'Insured Name'})
    df = df[OUTPUT_COLUMNS].copy()
    df['State'] = ""
    df['Contacted VIA'] = ""
    df['Notes Filed'] = ""
    df['Completed By'] = ""
    return df


def process_renewals(
    input_file_path,
    output_folder_path,
    on_status: Optional[StatusCallback] = None,
    on_progress: Optional[ProgressCallback] = None,
):
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

        logging.debug('Checking for required columns in DataFrame.')
        missing_columns = get_missing_required_columns(df, REQUIRED_COLUMNS)
        if missing_columns:
            logging.error(f"The following columns are missing in the input file: {missing_columns}")
            message = format_missing_columns_message(missing_columns)
            _set_status(on_status, message)
            _set_progress(on_progress, 0)
            return ProcessResult(success=False, error=message)

        logging.info('All required columns are present.')

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
                worksheet_row_count=len(df) + 1,
                output_file_path=output_file_path,
            )

        message = "Error - Excel export failed"
        _set_status(on_status, message)
        _set_progress(on_progress, 0)
        return ProcessResult(success=False, error=message)
    except MissingExcelPasswordError as e:
        logging.exception(
            "Password-protected source file could not be read because no "
            "password was loaded. Set EXCEL_PASSWORD or QUADSTATE_PASSWORD."
        )
        message = format_missing_password_message()
        _set_status(on_status, message)
        _set_progress(on_progress, 0)
        return ProcessResult(success=False, error=message)
    except MissingDependencyError as e:
        logging.exception(
            "Missing dependency while processing renewals: package=%s, "
            "missing_module=%s, install_hint=%s",
            e.package_name,
            e.missing_module,
            e.install_hint,
        )
        message = format_missing_dependency_message(e)
        _set_status(on_status, message)
        _set_progress(on_progress, 0)
        return ProcessResult(success=False, error=message)
    except Exception as e:
        logging.error(f"Error: {e}")
        message = f"Error - {str(e)}"
        _set_status(on_status, message)
        _set_progress(on_progress, 0)
        return ProcessResult(success=False, error=message)
