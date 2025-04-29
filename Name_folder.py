import os

directory = (r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\SIGCOM\2025\Febrero")
os.chdir(directory)

os.listdir()
folder_name = os.path.basename(os.getcwd())

import pandas as pd
devengado = pd.ExcelFile("DEVENGADO 2024.xlsx")
nombre_hojas = devengado.sheet_names

import regex as re
# Now im matching the sheet name that contains the folder name doesnt matter if its upper or lower
for hojas in nombre_hojas:
    if re.search(folder_name, hojas, re.IGNORECASE):
        hoja = hojas
        break
