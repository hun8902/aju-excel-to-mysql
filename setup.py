# setup.py
import sys
from cx_Freeze import setup, Executable

# Dependencies
build_exe_options = {
    "packages": [
        "tkinter",
        "pandas",
        "mysql.connector",
        "requests",
        "json",
        "datetime",
        "os",
        "threading"
    ],
    "excludes": [],
    "include_files": []
}

# Base for GUI applications
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="ExcelToMySQLImporter",
    version="1.0",
    description="Excel to MySQL Data Importer",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon=None)]
)