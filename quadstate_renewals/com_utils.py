import logging
import os
import shutil
import sys


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
        excel = win32.Dispatch('Excel.Application')

        try:
            excel.DisplayAlerts = False
        except Exception:
            logging.debug('Could not set DisplayAlerts property')

        try:
            excel.Visible = False
        except Exception:
            logging.debug('Could not set Visible property')

        if password:
            logging.debug('Opening Excel file with password protection.')
            try:
                wb = excel.Workbooks.Open(file_path, Password=password)
            except Exception as e:
                logging.debug(f'Failed to open with password: {e}')
                try:
                    wb = excel.Workbooks.Open(file_path)
                except Exception as e2:
                    logging.debug(f'Failed to open without password: {e2}')
                    raise e
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
            except Exception:
                pass

