# Generador Automatizado de Reportes Financieros
Este trabajo, busca facilitar la generación de los reportes mensuales de los Devengados, para ello se utiliza el siguiente Flujo.
## Estructura de Directorios

Lo primero que sucede al activar el programa es que se comienza una monitorización de forma permanente de una carpeta especificada, el programa por defecto monitorea Compartido Abastecimiento/Otros/SIGCOM, y todos los años y meses dentro de las subcarpetas.
Esta estructura monitoreada se ve a continuación
```
SIGCOM/
├── 2024/
│   ├── Enero/
│   │   ├── BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx
│   │   └── DEVENGADO.xlsx
│   ├── Febrero/
│   └── ...
├── 2025/
└── ...

NO_BORRAR/
└── Codigos_Clasificador_Compilado.xlsx
```
Donde cada uno de esos elementos corresponde 
Lo que hace el monitoreo, es que monitoreoa la carpeta SIGCOM, y revisa las carpetas dentro.
### 1. **Inicio Monitoreo**
. Esta monitorización va a detectar cualquier cambio o movimiento que se genere dentro de la carpeta y en base a eso generará cambios. Asi por ejemplo, un evento puede ser la creación de un archivo en la carpeta de destino.
### 2. **Verificación de requerimientos**
Cuando se detecta algun cambio lo que sucede es que inmedietamente se comienza a verificar lo siguiente:
- Existe el archivo de Devengado en la carpeta 
- Existe el archivo de Base en la carpeta 
- Existe el Compilado con los códigos SIGFE/SIGCOM en la carpeta NO BORRAR.
- No existe ya el documento de salida.
De cumplirse todos los requerimientos entonces se procede a generar_output1
### 3. **Genera output1**
- Modifico los códigos en Devengado para que sigan el mismo formato que en Codigo Clasificador Compilado.
- Al documento devengado le agrega los codigos SIGFE, SIGCOM y la descripción
- Luego guarda el mismo documento con esos cambios como `Modified_Devengado.xlsx`
### 4. **Genera output_2**
- Luego con todas las ordenes de compra, las agrupa por su respectivo código Sigfe, y agrega una columna con el Monto total dado por la suma de todos los elementos con el mismo código, y otra columna con el Monto Subasignaciones, dado por la suma de los Montos Totales de los SubItems en caso de poseerlos.
### 5. **Genera output_3**
- Luego genera el reporte final, cada uno de los elementos de output_2 es asignado manualmente y designado en una columna en específico dentro del archivo de Base.

# Generador Automatizado de Reportes Financieros  

Este flujo de trabajo automatiza la generación de reportes financieros mediante:  
- Monitoreo de cambios en archivos (`watchdog`).  
- Procesamiento de datos con `pandas`.  
- Llenado de plantillas Excel usando `xlwings`.  

## Descripción del Proceso  

### 1. **Inicio y Monitoreo**  
- El script inicia un observador (`watchdog`) que rastrea cambios en la carpeta designada (`base_path`): creación, modificación o renombrado de archivos.  
- Funciona en segundo plano de forma continua.  

### 2. **Validaciones Previas**  
Antes de procesar, verifica:  
- **Archivo correcto**: El nombre debe comenzar con `BASE DISTRIBUCION GASTO GENERAL` (ignora archivos temporales o ya modificados).  
- **Archivos requeridos**:  
  - `DEVENGADO.....xlsx` (datos brutos de transacciones).  
  - `Codigos_Clasificador_Compilado.xlsx` (mapeo de códigos financieros).  
- **Evitar duplicados**: Confirma que no exista un `Modified_BASE...xlsx` previo.  
- *Si falla alguna validación*, el proceso se detiene y vuelve al modo de monitoreo.  

### 3. **Procesamiento de Datos**  
#### **Salida 1: Devengado + Códigos SIGFE/SIGCOM **  
- Fusiona `DEVENGADO.xlsx` con el compilado de códigos (`Codigos_...`).  
- Limpia/formatea códigos para compatibilidad.  
- Guarda como `Modified_DEVENGADO_...xlsx`.  

#### **Salida 2: Montos Agrupados por Código**  
- Agrupa transacciones por `Cod_SIGFE` y suma los montos (`MONTO`).  
- Maneja casos especiales (ej. `Subasignaciones`).  
- Guarda como `Exported_...xlsx`.  

### 4. **Generación del Reporte Final**  
- Abre la plantilla `BASE...xlsx` con `xlwings`.  
- Llena celdas específicas con los datos consolidados (ej. total de gas → celda `U3` en "GASTO GENERAL").
- Guarda el reporte final como `Modified_BASE...xlsx`.  

### 5. **Finalización**  
- Vuelve al modo de monitoreo, listo para el próximo evento.  

---

### Dependencias Clave
- `watchdog`: Monitoreo del sistema de archivos.  
- `pandas`: Fusión y agregación de datos.  
- `xlwings`: Automatización de plantillas Excel.  

--- 

# Generación Automática de Reportes de Centros de Coste HSJM

## Descripción General

Este sistema es una herramienta automatizada que monitorea y procesa archivos Excel donde están los gastos devengados del Hospital San José de Melipilla, El programa vigila continuamente una estructura de directorios organizada por años y meses, detectando cuando se agregan o modifican archivos específicos. Cuando encuentra archivos que cumplen con ciertos criterios, los procesa automáticamente, extrayendo información relevante y actualizando informes financieros.

## Estructura del Sistema

El sistema está compuesto por los siguientes componentes principales:

1. **Monitor de Archivos**: Utiliza la biblioteca `watchdog` para detectar cambios en la estructura de directorios.
2. **Procesador de Datos**: Lee, manipula y procesa datos de los archivos Excel utilizando pandas y xlwings.
3. **Generador de Informes**: Crea informes modificados basados en los archivos originales.

## Requisitos del Sistema

- Python 3.6+
- Bibliotecas: watchdog, pandas, xlwings, openpyxl, re
- Estructura de carpetas específica
- Archivos específicos:
  - Codigos_Clasificador_Compilado.xlsx (en la carpeta NO_BORRAR)
  - DEVENGADO.xlsx (en cada carpeta mensual)
  - BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx (en cada carpeta mensual)

## Configuración

El sistema requiere la siguiente configuración:

```python
# Rutas principales
base_path = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\SIGCOM"  # Ruta de red en producción
# O alternativamente para desarrollo:
base_path = r"C:\Users\Thinkpad\PycharmProjects\Automatic_Report_Generation\Files\SIGCOM"

# Directorio de trabajo (donde se encuentran archivos auxiliares)
work_directory = r'C:\Users\Thinkpad\PycharmProjects\Automatic_Report_Generation\Files\NO_BORRAR'
```

## Estructura de Directorios

El sistema espera una estructura de directorios específica:

```
SIGCOM/
├── 2024/
│   ├── Enero/
│   │   ├── BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx
│   │   └── DEVENGADO.xlsx
│   ├── Febrero/
│   └── ...
├── 2025/
└── ...

NO_BORRAR/
└── Codigos_Clasificador_Compilado.xlsx
```

## Funcionalidades Principales

### 1. Monitoreo de Archivos

El sistema monitorea continuamente la estructura de directorios en busca de cambios:

- Detecta archivos recién creados, modificados o renombrados.
- Identifica archivos específicos que cumplen con los criterios para procesamiento.
- Inicia el procesamiento automático cuando encuentra un archivo válido.

### 2. Procesamiento de Datos

Para cada archivo válido encontrado, el sistema realiza las siguientes operaciones:

- Lee y procesa los datos del archivo DEVENGADO.xlsx, utiliza la página cuyo nombre coincida con el de la carpeta
- Combina estos datos con la información de clasificación de códigos
- Calcula montos por código SIGFE
- Maneja correctamente las subasignaciones
- Genera archivos modificados y exportados con los resultados

### 3. Actualización de Informes

El sistema actualiza los informes financieros:

- Actualiza hojas específicas (GASTO GENERAL y SUMINISTROS)
- Coloca valores calculados en celdas específicas
- Genera nuevos archivos con prefijo "Modified_" y "Exported_"

## Flujo de Proceso

1. El monitor detecta un archivo nuevo o modificado en la estructura de directorios.
2. Verifica que sea un archivo de interés (BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx).
3. Comprueba si existe el archivo DEVENGADO.xlsx necesario para el procesamiento.
4. Procesa los datos y extrae valores específicos.
5. Abre el archivo original con xlwings y actualiza valores en celdas específicas.
6. Guarda el archivo modificado con el prefijo "Modified_".

## Funciones Principales

### configurar_hoja_activa(root_dir)
Configura la hoja activa en los archivos DEVENGADO.xlsx para que coincida con el nombre de la carpeta donde se encuentra.

### verificar_carpetas(carpeta_modificada=None)
Verifica las carpetas de años/meses y busca archivos Excel para procesar.

### process_data(folder_path)
Procesa datos utilizando archivos del directorio de trabajo y la carpeta especificada.

### update_excel_with_xlwings(file_path)
Actualiza un archivo Excel utilizando xlwings.

### safe_get_value(df, code, column)
Extrae de forma segura un valor del DataFrame, devolviendo 0 si no se encuentra.

### format_sigfe_code(code)
Formatea un código SIGFE para que tenga el formato estándar.

### calculate_subasignaciones_amounts(exported)
Calcula los montos de subasignaciones.

## Ejecutando el Sistema

Para iniciar el sistema, simplemente ejecute el archivo Python principal:

```bash
python "Hora 9200 version.py"
```

El sistema comenzará a monitorear la estructura de directorios y procesará automáticamente los archivos que cumplan con los criterios.

## Solución de Problemas

### Problemas comunes y soluciones:

1. **No se encuentra el archivo Codigos_Clasificador_Compilado.xlsx**:
   - Asegúrese de que el archivo exista en la carpeta NO_BORRAR
   - Verifique las rutas configuradas en el código

2. **No se encuentra el archivo DEVENGADO.xlsx**:
   - Asegúrese de que el archivo esté en la misma carpeta que el archivo BASE DISTRIBUCION

3. **Problemas con la hoja de Excel**:
   - Verifique que el archivo tenga las hojas "GASTO GENERAL" y "SUMINISTROS"
   - Compruebe que el formato de las hojas sea compatible

4. **El prefijo "Modified_" ya existe**:
   - El sistema no sobrescribe archivos modificados para evitar la pérdida de datos
   - Elimine o renombre los archivos existentes si desea volver a procesar

## Notas Adicionales

- El sistema está optimizado para trabajar con una estructura específica de archivos SIGCOM.
- Es posible crear códigos adicionales solamente con adicionar nuevos códigos al final del Compilado.
- Al modificar un devengado, basta con modificar el ITEM SIGFE para que se contabilice de forma correcta por el sistema.
- Hay una nota en el código que indica: "Ojo con subasignaciones, tengo duda de que funcionen bien" - Esta es un área que podría requerir revisión adicional.
