import os
import xlwings as xw

def initial_setup():

    directory = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\Xlwings_Python_Practice"
    os.chdir(directory)
    libro_original = xw.Book("BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025 (1).xlsx")
    libro_copia = xw.Book()
    for hoja in libro_original.sheets:
        hoja.api.Copy(Before=libro_copia.sheets[0].api)
    libro_copia.sheets["Hoja1"].delete()

    return libro_original, libro_copia

libro_original, libro_copia = initial_setup()

sheet = libro_copia.sheets["GASTO GENERAL"]


sheet.cells(1, 1).value = "GASTO GENERAL"
sheet.cells(1, 2).value = "SUMINISTROS"

libro_copia.save("copia_seguridad.xlsx")

# 6. Cerrar ambos archivos
libro_original.close()
libro_copia.close()

import os
import xlwings as xw

file_path = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\SIGCOM\No Borrar\BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025 (1).xlsx"

if os.path.exists(file_path):
    try:
        print(f"Opening file: {file_path}")
        app = xw.App(visible=False)
        wb = xw.Book(file_path)
        print("File opened successfully.")

        # Print sheet names
        sheet_names = [sheet.name for sheet in wb.sheets]
        print(f"Sheet names: {sheet_names}")

        wb.close()
        app.quit()
    except Exception as e:
        print(f"Error accessing file: {e}")
else:
    print(f"File does not exist: {file_path}")