import os
import sys
from pathlib import Path

from dotenv import load_dotenv


def get_application_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(sys.argv[0]).resolve().parent


def get_env_locations():
    home_dir = Path.home()
    return [
        get_application_dir() / '.env',
        Path('C:/QuadstateRenewalsProperties/.env'),
        home_dir / '.env',
        home_dir / 'Documents' / '.env',
        home_dir / 'AppData' / 'Local' / 'QuadstateRenewals' / '.env',
    ]


def load_environment():
    for env_path in get_env_locations():
        if env_path.exists():
            load_dotenv(env_path)
            return str(env_path)

    load_dotenv()
    return None


def get_excel_password():
    password = os.getenv('EXCEL_PASSWORD')
    if password:
        return (
            password,
            'Ready - Using password from EXCEL_PASSWORD environment variable',
            'Using password from EXCEL_PASSWORD environment variable',
            False,
        )

    password = os.getenv('QUADSTATE_PASSWORD')
    if password:
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

