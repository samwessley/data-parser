from setuptools import setup

APP = ['excel_parser.py']  # Replace 'main.py' with the name of your script
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['pandas', 'openpyxl'],
    'includes': ['tkinter'],  # Use tkinter for file dialog if needed
    'excludes': ['PyInstaller', 'gi', 'gi.repository'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
