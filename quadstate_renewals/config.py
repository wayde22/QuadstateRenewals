import os
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None


APP_DATA_DIR_NAME = 'QuadstateRenewals'
ENV_TEMPLATE = """# Add your Excel file password here
EXCEL_PASSWORD=PasswordHere
"""
PASSWORD_PLACEHOLDERS = {'PasswordHere', 'your_password_here'}


def get_application_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(sys.argv[0]).resolve().parent


def get_app_data_dir():
    local_app_data = os.environ.get('LOCALAPPDATA')
    if local_app_data:
        return Path(local_app_data) / APP_DATA_DIR_NAME
    return Path.home() / 'AppData' / 'Local' / APP_DATA_DIR_NAME


def get_log_file_path():
    return get_app_data_dir() / 'app.log'


def get_user_env_file_path():
    return get_app_data_dir() / '.env'


def ensure_user_env_file():
    env_path = get_user_env_file_path()
    if env_path.exists():
        return None

    env_path.parent.mkdir(parents=True, exist_ok=True)
    env_path.write_text(ENV_TEMPLATE, encoding='utf-8')
    return env_path


def get_env_locations():
    home_dir = Path.home()
    return [
        get_application_dir() / '.env',
        Path('C:/QuadstateRenewalsProperties/.env'),
        home_dir / '.env',
        home_dir / 'Documents' / '.env',
        get_user_env_file_path(),
    ]


def load_environment():
    if load_dotenv is None:
        return None

    for env_path in get_env_locations():
        if env_path.exists():
            load_dotenv(env_path)
            return str(env_path)

    load_dotenv()
    return None


def get_excel_password():
    password = os.getenv('EXCEL_PASSWORD')
    if password and password not in PASSWORD_PLACEHOLDERS:
        return (
            password,
            'Ready - Using password from EXCEL_PASSWORD environment variable',
            'Using password from EXCEL_PASSWORD environment variable',
            False,
        )

    password = os.getenv('QUADSTATE_PASSWORD')
    if password and password not in PASSWORD_PLACEHOLDERS:
        return (
            password,
            'Ready - Using password from QUADSTATE_PASSWORD environment variable',
            'Using password from QUADSTATE_PASSWORD environment variable',
            False,
        )

    return (
        None,
        'Warning - No password found in environment variables',
        'No password found in environment variables',
        True,
    )


def get_windows_username():
    try:
        return os.getlogin()
    except OSError:
        return os.environ.get('USERNAME', '')


def get_default_input_file():
    username = get_windows_username()
    if not username:
        return ''

    downloads_dir = Path('C:/Users') / username / 'Downloads'
    candidates = [
        downloads_dir / 'Copy of Export_RenewalCenter.xlsx',
        downloads_dir / 'Export_RenewalCenter.xlsx',
    ]

    for file_path in candidates:
        if file_path.exists():
            return str(file_path)

    return ''


def get_default_output_folder():
    username = get_windows_username()
    if not username:
        return ''

    user_dir = Path('C:/Users') / username
    candidates = [
        user_dir / 'OneDrive - quadstateinsurance.com' / 'Desktop',
        user_dir / 'OneDrive - Quadstate Insurance Agency LLC' / 'Desktop',
        user_dir / 'Desktop',
    ]

    for folder_path in candidates:
        if folder_path.exists():
            return str(folder_path)

    return ''

