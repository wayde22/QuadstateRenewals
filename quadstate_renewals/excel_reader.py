import logging
import os
import tempfile
import time

import pandas as pd


def read_excel_file(file_path, password=None):
    logging.debug(f'Reading Excel file: {file_path}')

    if not password:
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            logging.info('Excel file read directly into DataFrame.')
            return df
        except Exception as e:
            logging.debug(f'Direct read failed: {e}. Trying password-protected method.')

    temp_dir = os.environ.get('TEMP') or tempfile.gettempdir()
    temp_file_path = os.path.join(temp_dir, f"~$temp_{int(time.time())}.xlsx")
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
        if os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
                logging.debug('Temporary file removed.')
            except Exception as e:
                logging.warning(f'Could not remove temporary file: {e}')

