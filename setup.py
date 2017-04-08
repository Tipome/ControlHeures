import sys
from cx_Freeze import setup, Executable

#Pour lancer la création de l'exe, taper dans la console "py setup.py build"
#inclus les autres fichiers non python
includefiles=["data.csv","Registre_type.xlsx","noms et trigrammes.xlsm","Forfaits HDC.xlsx"]

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os","openpyxl","pyexcel","pyexcel_xls","pyexcel_xlsx"],
                     "excludes": ["tkinter"],
                     "include_files" : includefiles}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"



setup(  name = "ControlHeure Vigie",
        version = "0.3",
        description = "Crée les registres d'heures mensuels",
        options = {"build_exe": build_exe_options},
        executables = [Executable("ControlHeureVigie.py", icon="stopwatch.ico", base=base)])
