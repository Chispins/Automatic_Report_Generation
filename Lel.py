import pandas as pd
import os

# Directorio de trabajo
WORK_DIRECTORY = r'C:\\Users\\Usuario\\Downloads'
os.chdir(WORK_DIRECTORY)

categorias_codigos = pd.read_excel("Codigos_Clasificador_Compilado.xlsx")
criterios_distribucion = pd.read_excel("criterio_distribución.xlsx")
base_distribucion = pd.read_excel("BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025.xlsx")

"""categorias_codigos = categorias_codigos.columns.strip()
criterios_distribucion = criterios_distribucion.columns.strip()
base_distribucion = base_distribucion.columns.strip()"""
# Esto simplemente elimina todo lo posterior al numero, es decir subasignaciones
categorias_codigos["Subasignaciones"] = categorias_codigos["Cod_SIGFE"].str.contains("y").astype(int)
categorias_codigos.loc[categorias_codigos["Subasignaciones"] == 1, "Cod_SIGFE"] = categorias_codigos[categorias_codigos["Subasignaciones"] == 1]["Cod_SIGFE"].apply(lambda x: x.split()[0] if isinstance(x, str) and ' ' in x else x.replace("y", "") if isinstance(x, str) else x)

def format_sigfe_code(code):
    if not isinstance(code, str):
        return code

    # Dividir por puntos para obtener diferentes secciones
    parts = code.split('.')

    if len(parts) < 3:
        return code

    # Mantener las dos primeras secciones (22.12) y solo los últimos 2 dígitos del resto
    result = f"{parts[0]}.{parts[1]}"

    # Agregar los últimos dos dígitos de cada sección restante
    for i in range(2, len(parts)):
        section = parts[i].strip()
        if len(section) >= 2:
            result += f".{section[-2:]}"
        else:
            result += f".{section}"

    return result

# Si encuentro la palabra "y", entonces "Subasignaciones" = 1
categorias_codigos["Cod_SIGFE"] = categorias_codigos["Cod_SIGFE"].apply(format_sigfe_code)
# Eliminar todos los puntos
categorias_codigos["Cod_SIGFE"] = categorias_codigos["Cod_SIGFE"].str.replace(".", "")

devengado = pd.read_excel("DEVENGADO 2025.xlsx", skiprows=5, header=0)
devengado["item_conv"] = devengado.iloc[:,9][:]
devengado["Cod_SIGFE"] = devengado.iloc[:,9][:]

devengado["item_conv"] = devengado["item_conv"].str.split("-").str[1].str.strip()
devengado["Cod_SIGFE"] = devengado["Cod_SIGFE"].str.split("-").str[1].str.strip()

# Eliminar los dos primeros caracteres
categorias_codigos["Cod_SIGFE"] = categorias_codigos["Cod_SIGFE"].str[2:]

# Fusionar datos y agrupar usando first para evitar duplicaciones
merged = pd.merge(devengado, categorias_codigos, how='outer', on='Cod_SIGFE')
exported_v1 = merged.groupby(by="Cod_SIGFE")[["MONTO ", "Item_en_SIGCOM", "Item SIGFE", "Subasignaciones"]].sum()
exported = merged.groupby(by="Cod_SIGFE").agg({
    "MONTO ": "sum",
    "Item_en_SIGCOM": "first",
    "Item SIGFE": "first",
    "Subasignaciones": "max"
})


col1 = exported.iloc[:,0]
col2 = exported.iloc[:,3]
col3 = exported.index
exported_v1 = pd.DataFrame(col3, col2)
exported_v3 = pd.merge(exported, exported_v1, how='outer', on='Cod_SIGFE')
exported = exported_v3

# Limitar longitud de texto a 65 caracteres
exported["Item_en_SIGCOM"] = exported["Item_en_SIGCOM"].str[:65]
exported["Item SIGFE"] = exported["Item SIGFE"].str[:65]

# Dividir los montos por 2
exported["MONTO "] = exported["MONTO "]/2

# Crear columna para subasignaciones
# Si subasignaciones == 1, sumamos los valores que empiezan con el mismo Cod_SIGFE
exported["MONTO_SUBASIGNACIONES"] = exported["MONTO "].copy()

# Identificar códigos base (códigos con Subasignaciones = 1)
base_codes = exported.index[exported["Subasignaciones"] == 1].tolist()

# Mapear cada código a su código base más específico
exported["MONTO_SUBASIGNACIONES"] = exported["MONTO "].copy()
base_codes = exported.index[exported["Subasignaciones"] == 1].tolist()
code_to_base = {}

for code in exported.index:
    matching_bases = [base for base in base_codes if str(code).startswith(str(base))]
    if matching_bases:
        # Convert all items to strings before comparing lengths
        code_to_base[code] = max(matching_bases, key=lambda x: len(str(x)))
    else:
        code_to_base[code] = None

base_sums = {}
for base in base_codes:
    related_codes = [code for code, mapped_base in code_to_base.items() if mapped_base == base]
    base_sums[base] = exported.loc[related_codes, "MONTO "].sum()

for code in exported.index:
    base = code_to_base.get(code)
    if base is not None:
        exported.loc[code, "MONTO_SUBASIGNACIONES"] = base_sums[base]

exported["MONTO_SUBASIGNACIONES"] = (exported["MONTO_SUBASIGNACIONES"] - exported["MONTO "]) * (
    exported["Subasignaciones"])

exported["Item_SIGCOM_Duplicado"] = exported["Item_en_SIGCOM"]

# Abrir los archivos que comienzan con "BASE DISTRIBUCION GASTO GENERAL"
# Crear un nuevo archivo, exactamente igual a BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025.xlsx pero cambiaremos algunos valores
base_distribucion_generada_automaticamente = pd.read_excel("BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025.xlsx", sheet_name="GASTO GENERAL")
base_distribucion_generada_automaticamente_suministros = pd.read_excel("BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025.xlsx", sheet_name="SUMINISTROS")



columna = exported.iloc[:,0]
exported.index = columna
exported = exported.iloc[:,1:]


# Valores para la hoja GASTO GENERAL

gas = exported.loc["0503", "MONTO "]
mantencion_jardines = exported.loc["0803", "MONTO "]
eq_computo = exported.loc["0607", "MONTO "]
agua = exported.loc["0502", "MONTO "]
servicio_energia = exported.loc["0501", "MONTO "]
servicio_vigilancia = exported.loc["0802", "MONTO "]

# Actualizar valores en la hoja de GASTO GENERAL
base_distribucion_generada_automaticamente.iloc[1, 20] = gas
base_distribucion_generada_automaticamente.iloc[1, 22] = mantencion_jardines
base_distribucion_generada_automaticamente.iloc[1, 24] = eq_computo
base_distribucion_generada_automaticamente.iloc[1, 53] = agua
base_distribucion_generada_automaticamente.iloc[1, 57] = servicio_energia

base_distribucion_generada_automaticamente.to_excel("base_distribucion_generada_automaticamente_cruda.xlsx")

# Valores para la hoja SUMINISTROS
base_distribucion_generada_automaticamente_suministros = pd.read_excel("BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS ENE 2025.xlsx", sheet_name="SUMINISTROS")
calefaccion = exported.loc["0301", "MONTO "]
material_medico_quirurjico = exported.loc["0405", "MONTO_SUBASIGNACIONES"]
material_oficina = exported.loc["0401", "MONTO "]
materiales_informaticos = exported.loc["0409", "MONTO "]
gasto_total_medicamentos = exported.loc["040401", "MONTO_SUBASIGNACIONES"]
productos_quimicos = exported.loc["040302", "MONTO "]

# Actualizar valores en la hoja de SUMINISTROS
base_distribucion_generada_automaticamente_suministros.iloc[1,6] = calefaccion
base_distribucion_generada_automaticamente_suministros.iloc[1,19] = material_medico_quirurjico
base_distribucion_generada_automaticamente_suministros.iloc[1,27] = material_oficina
base_distribucion_generada_automaticamente_suministros.iloc[1,30] = materiales_informaticos
base_distribucion_generada_automaticamente_suministros.iloc[1,38] = gasto_total_medicamentos
base_distribucion_generada_automaticamente_suministros.iloc[1,45] = productos_quimicos