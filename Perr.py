import os

os.chdir(r"C:\Users\Usuario\Desktop\Carpeta_Monitor\2025\Febrero")

file = "BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025 (3).xlsx"

import openpyxl as xl

# now import the file
file = xl.load_workbook(file)
print(file)

sheet