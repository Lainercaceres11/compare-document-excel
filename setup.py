# python setup.py build
# python setup.py bdist_msi - instalador

from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": [
        "pandas",
        "openpyxl",
        "numpy",
        "xlrd",    
    ],
    "includes": [
        "tkinter"
    ],
    "include_files": [
        "icon.ico",
        "ui.py",
        "documents.py"
    ],
}

setup(
    name="documents",
    version="1.0",
    description="Comparador de archivos Excel - Expreso del Pacífico",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "documents.py",
            base="Win32GUI",
            icon="icon.ico"
        )
    ],
)