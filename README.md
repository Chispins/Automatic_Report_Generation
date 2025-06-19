# Generador Automatizado de Reportes Financieros
Este programa crea automáticamente un reporte mensual de gastos. Reemplaza el trabajo manual de copiar datos entre planillas de Excel, ahorrando tiempo y evitando errores

El programa sigue la siguiente secuencia para lograr generar el reporte mensual, que se detallará más adelante
![Image](https://github.com/user-attachments/assets/39e2dbfb-90c9-4c11-a1f6-1a77f37fa7fd)


## Estructura de Directorios

Lo primero que sucede al activar el programa es que se crea un vigilante que estará siempre mirando la carpeta principal y todas las carpetas dentro (meses, y años). Esto es para asegurarse de que cuando se tengan los archivos necesarios se cree el reporte. Este vigilante espera la siguiente estructura de carpetas.

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
| **`BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx`** en la carpeta del mes | El reporte **NO se genera** | Copia el archivo desde `NO_BORRAR`<br>2. Pégalo en la carpeta del mes |
| **`Codigos_Clasificador_Compilado.xlsx`** en `NO_BORRAR` | El reporte **NO funciona correctamente** | **No lo muevas ni lo borres**<br>Si falta, repónlo desde una copia de seguridad |
| **NO existe el reporte final** en la carpeta del mes | Si es que **YA EXISTE UN REPORTE** no se crea un nuevo reporte | 1. Elimina el reporte antiguo<br>2. O muévelo a otra carpeta |

### 3. **Creación de Archvio intermedio 1 (Devengado Modificado)**
- El programa abre el Excel devengado **utilizando la hoja con el mismo nombre de la carpeta** y luego le agrega un par de columnas que contienen los nombres del ITEM SIGFE e ITEM SIGCOM, estas columnas provienen del Compilado de Códigos Presupuestarios.

### 4. **Creación de Archivo Intermedio 2 (Resumen por Item)**
- Utiliza el archivo intermedio 1, toma todas las compras y las agrupa por su respectivo código SIGFE, luego agrega una columna con el Monto total dado por la suma de todos los elementos con el mismo código, y otra columna con el Monto Subasignaciones, dado por la suma de los Montos Totales de los SubItems en caso de poseerlos.
### 5. **Genera Reporte Final**
- Luego genera el reporte final, cada uno de los elementos de archivo intermedio 2 es asignado manualmente y designado en una columna en específico dentro del archivo de Base.


### Ejemplo de uso para Noviembre 2026

1. Crear la carpeta:  
   `SIGCOM/2026/Noviembre/`

2. Copiar tus archivos:
   - Pega tu archivo de `DEVENGADO NOVIEMBRE 2026.xlsx` en la carpeta
   - Coloca la plantilla `BASE DISTRIBUCION GASTO GENERAL Y SUMINISTROS.xlsx` en la misma carpeta  
     *(Si no tienes una plantilla, copia la versión de respaldo de la carpeta NO_BORRAR)*

3. ¡Listo! El reporte se generará automáticamente en unos minutos.
