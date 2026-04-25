
# python setup.py build
# python setup.py bdist_msi - instalador
# C:\Users\Lainer Cáceres\AppData\Local\Programs\documents
from cx_Freeze import setup, Executable

setup(
    name="documents",
    version="4.2",
    description="Compara dos documentos excel y genera otro con las coincidencias.",
    executables=[Executable("documents.py", base="Win32GUI", icon="icon.ico")],
    options={
        "build_exe": {
            "packages": [ "openpyxl", "pandas"],
            "include_files": [],  
        }
    },
)

