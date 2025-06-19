# Generador Automatizado de Reportes Financieros
Este programa crea automáticamente un reporte mensual de gastos. Reemplaza el trabajo manual de copiar datos entre planillas de Excel, ahorrando tiempo y evitando errores

El programa sigue la siguiente secuencia para lograr generar el reporte mensual
![Image](https://github.com/user-attachments/assets/9744baf1-0f87-4605-9acf-1142fe125670)


## Estructura de Directorios

Lo primero que sucede al activar el programa es que se comienza una monitorización de forma permanente de una carpeta especificada, el programa por defecto monitorea Compartido Abastecimiento/Otros/SIGCOM, y todos los años y meses dentro de las subcarpetas.
Esta estructura monitoreada se ve a continuación

Carpetas principales:
1. **SIGCOM**: Aquí se guardan los archivos mensuales
   - Cada año tiene su carpeta (ej: 2024)
   - Cada mes tiene su subcarpeta (ej: Enero)
      - 📄 `DEVENGADO.xlsx` → **Gastos del mes** (obligatorio)
      - 📄 `BASE...xlsx` → **Plantilla para el reporte** (obligatorio)

2. **NO_BORRAR**: Archivos importantes que NUNCA deben faltar
   - 🔐 `Codigos...xlsx` → Lista de categorías de gastos según SIGFE y SIGCOM
   - 🔐 `BASE...xlsx` → Copia de seguridad de la plantilla


### 1. **Inicio Monitoreo**
El programa revisa cada segundo si hay archivos nuevos o modificados en las carpetas. Cuando detecta los archivos necesarios, genera el reporte automáticamente.
### 2. **Verificación de requerimientos**
Para que el reporte se genere se revisa que se cumplan **todos** los requisitos listados a continuación.
| Requisito | ¿Qué pasa si falta? | ¿Cómo solucionarlo? |
|-----------|---------------------|---------------------|
| **`DEVENGADO.xlsx`** en la carpeta del mes | El reporte **NO se genera** | 1. Consigue el archivo de gastos del mes<br>2. Colócalo en la carpeta del mes<br>3. Asegúrate que se llame el nombre comienza con `DEVENGADO` |
| **`BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx`** en la carpeta del mes | El reporte **NO se genera** | Si quieres una versión específica:<br>1. Copia el archivo desde `NO_BORRAR`<br>2. Pégalo en la carpeta del mes |
| **`Codigos_Clasificador_Compilado.xlsx`** en `NO_BORRAR` | El reporte **NO funciona correctamente** | **No lo muevas ni lo borres**<br>Si falta, repónlo desde una copia de seguridad |
| **NO existe el reporte final** en la carpeta del mes | No se crea nuevo reporte | 1. Elimina el reporte antiguo<br>2. O muévelo a otra carpeta |

### 3. **Creación de Archvio intermedio 1**
- El programa abre el Excel devengado **utilizando la hoja con el mismo nombre de la carpeta** y luego le agrega un par de columnas que contienen los nombres del ITEM SIGFE e ITEM SIGCOM, estas columnas provienen del Compilado de Códigos Presupuestarios.

### 4. **Genera output_2**
- Con el archivo anterior, se toman todas las compras y las agrupa por su respectivo código SIGFE, y agrega una columna con el Monto total dado por la suma de todos los elementos con el mismo código, y otra columna con el Monto Subasignaciones, dado por la suma de los Montos Totales de los SubItems en caso de poseerlos.
### 5. **Genera output_3**
- Luego genera el reporte final, cada uno de los elementos de output_2 es asignado manualmente y designado en una columna en específico dentro del archivo de Base.


### Ejemplo de uso para Noviembre 2026

1. Crear la carpeta:  
   `SIGCOM/2026/Noviembre/`

2. Copiar tus archivos:
   - Pega tu archivo de `DEVENGADO.xlsx` en la carpeta
   - Coloca la plantilla `BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx` en la misma carpeta  
     *(Si no tienes una plantilla, copia la versión de respaldo de la carpeta NO_BORRAR)*

3. ¡Listo! El reporte se generará automáticamente en unos minutos.

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

### Ejemplo de uso para Noviembre 2025
Crear carpeta SIGCOM/2025/Marzo
-Copiar tu archivo DEVENGADO.xlsx
-Pegar aquí la plantilla BASE DISTRIBUCION...xlsx (si no tienes una, usa la de NO_BORRAR)
¡El reporte se creará automáticamente!"



1. **Abre el archivo de gastos del mes** (`DEVENGADO.xlsx`)
   - Busca automáticamente **la hoja que coincide con el nombre del mes** (ej: si estás en la carpeta "Marzo", usará la hoja "Marzo" o "MARZO")
   - ⚠️ Si no encuentra una hoja con ese nombre exacto, el proceso se detiene

2. **Realiza estas mejoras al archivo:**
   - 🧹 Elimina cualquier formato (font, size, color, etc)
   - ➕ Añade nueva información importante:
     - Código oficial del tipo de gasto (ITEM SIGFE)
     - Código alternativo (ITEM SIGCOM)
     - Nombre completo del gasto según ambos sistemas
     - Indicador de sub-items (¿Existen sub asignaciones? → Sí=1 / No=0)

3. **Guarda el resultado mejorado**
   - Nombre del nuevo archivo: `Modified_Devengado.xlsx`
   - Ubicación: **Misma carpeta del mes**
   - 
Una vez se confirmaron que se cumplen las condiciones previas, entonces se procede a abrir los gastos DEVENGADOS mensuales, **si es que el excel posee multiples páginas, entonces abre la página que tenga el mismo nombre que la carpeta en la que se encuentra, es decir, si estamos en la Carpeta "Marzo", al Abrir el Devengado utilizará la hoja de "Marzo" o "MARZO"**. En caso de NO existir la hoja de marzo, entonces el proceso fallará. Y no se generará ningún archivo.
La hoja de excel utilizada es la que posee el mismo nombre de la carpeta, y en base a esa se va generar un primer archivo, que es exactamente igual al original solo que sin ningun formato, y con nuevas columnas agregadas, estas columnas son los datos que están presentes en Código Clasificador compilado, entonces por ejemplo un registro posee entre las muchas columnas, una que especifican el ITEM, que es en realidad un Codigo SIGFE, es ese codigo el cual se hace un "match" con los códigos en Clasificador Compilado, y se adicionan las columnas al Devengado, y las columnas que se agregan son ITEM SIGFE, ITEM SIGCOM, el nombre del Item según SIGFE, y el nombre según SIGCOM, además de una columna Subasignaciones que toma el valor 1 si es que el item posee Subasignaciones o Subitems, y 0 si no. Otra modificación que sucede es que elimina los datos de todas las filas que no representan registros individuales, entonces aquellos que por ejemplo son los Totales de un item son Ignorados, por lo que su preesencia o ausencia no genera ningún efecto en el reporte.
Una vez que se inicia la generación del primer archivo.
- Modifico los códigos en Devengado para que sigan el mismo formato que en Codigo Clasificador Compilado.
- Al documento devengado le agrega los codigos SIGFE, SIGCOM y la descripción
- Luego guarda el mismo documento con esos cambios como `Modified_Devengado.xlsx`

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


## Known Issues
El monto que se designa es imputado manualmente, es decir, Siempre se imputa, el Monto Total o Subasignaciones, pero esto no es condicional, por lo que podría suceder que un item que tipicamente no tiene Subasignaciones un día tenga, y aun así en el reporte se mostrará el monto total
