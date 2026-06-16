import logging
import os
import tempfile
import time

from .dependencies import MissingDependencyError


class MissingExcelPasswordError(RuntimeError):
    pass


def _import_pandas():
    try:
        import pandas as pd
    except ImportError as exc:
        raise MissingDependencyError(
            'pandas',
            'pandas>=2.0.0',
            getattr(exc, 'name', None),
        ) from exc
    return pd


def _require_openpyxl():
    try:
        import openpyxl  # noqa: F401
    except ImportError as exc:
        raise MissingDependencyError(
            'openpyxl',
            'openpyxl>=3.1.0',
            getattr(exc, 'name', None),
        ) from exc


def _import_msoffcrypto():
    try:
        import msoffcrypto
    except ImportError as exc:
        raise MissingDependencyError(
            'msoffcrypto-tool',
            'msoffcrypto-tool>=5.0.0',
            getattr(exc, 'name', None),
        ) from exc
    return msoffcrypto


def read_excel_file(file_path, password=None):
    logging.debug(f'Reading Excel file: {file_path}')
    pd = _import_pandas()
    _require_openpyxl()

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
        msoffcrypto = _import_msoffcrypto()
        with open(file_path, 'rb') as f:
            temp_file = msoffcrypto.OfficeFile(f)
            if not password and temp_file.is_encrypted():
                raise MissingExcelPasswordError(
                    "Password-protected Excel file requires EXCEL_PASSWORD "
                    "or QUADSTATE_PASSWORD."
                )

            if password:
                temp_file.load_key(password=password)
            else:
                temp_file.load_key()
            with open(temp_file_path, 'wb') as decrypted_file:
                temp_file.decrypt(decrypted_file)

        df = pd.read_excel(temp_file_path, engine='openpyxl')
        logging.info('Excel file read using msoffcrypto-tool.')
        return df
    except MissingDependencyError:
        raise
    except MissingExcelPasswordError:
        raise
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

