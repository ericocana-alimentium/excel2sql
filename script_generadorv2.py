# %% [markdown]
# ##### Importing the libraries

# %%
import pandas as pd
import warnings
import datetime
import os
import glob
import re

# %% [markdown]
# Ignoring warnings for cleaner output

# %%
warnings.simplefilter(action='ignore', category=UserWarning) # Ignorar warnings

# %%
# file_path = 'excels\Configuracion-WFL-Inicial-v2.xlsx'
# df = pd.read_excel(file_path, index_col=None)
# df.head(3).to_markdown()


# %%
hora_inicio = datetime.datetime.now()
#print(f"Hora de inicio: {hora_inicio.strftime('%H:%M:%S')}")

# %% [markdown]
# ##### Obtener los excels de la carpeta

# %%
# 3. Obtener la ruta del directorio del script y listar todos los archivos Excel en la carpeta
script_path = "excels"
ruta_carpeta_excel = os.path.join(script_path)
archivos_excel = glob.glob(ruta_carpeta_excel + "/*.xlsx")


# %% [markdown]
# Extraer y almacenar solo los nombres de los archivos con y sin extension para posteriormente usarlos y evitar repetir codigo

# %%
archivos_excel_basename = []
nombres_sin_extension = []


for archivo_excel in archivos_excel:
    # Primero, obtener solo el nombre del archivo con extensión usando os.path.basename
    nombre_archivo_con_extension = os.path.basename(archivo_excel)
    archivos_excel_basename.append(nombre_archivo_con_extension)
    # print(nombre_archivo_con_extension)
    # Luego, quitar la extensión para obtener solo el nombre del archivo
    nombre_archivo_sin_extension = os.path.splitext(nombre_archivo_con_extension)[0]
    nombres_sin_extension.append(nombre_archivo_sin_extension)

# %% [markdown]
# ### Función `prepararCabecera`
# 
# La función `prepararCabecera` tiene como objetivo generar una cabecera personalizada para un archivo SQL basándose en la entrada del usuario. Aquí está el desglose de su funcionalidad:
# 
# #### Proceso
# 
# 1. **Inicialización del Link de la Tarea**: La función comienza con un link base para una tarea (`link_tarea`) predefinido, dirigido a un recurso específico en `ClickUp`.
# 
# 2. **Entrada de Datos del Usuario**:
#    - `input_tarea`: Se solicita al usuario que introduzca un identificador único de tarea (ejemplo: `PROCLI-3948`). Este identificador se anexa al `link_tarea`.
#    - `descripcion`: Se pide al usuario que introduzca la descripción de la tarea.
#    - `autor`: Se solicita el nombre del autor o autores de la tarea.
#    - `fecha_creacion`: Se registra automáticamente la fecha actual en formato `AAAA-MM-DD` utilizando `hora_inicio.strftime('%Y-%m-%d')`.
# 
# 3. **Extracción del Código de Tarea**: Utiliza una expresión regular para extraer el número después del guion en `input_tarea`. Si el formato es incorrecto, se notifica al usuario.
# 
# 4. **Generación de la Cabecera SQL**: Crea una cabecera detallada para el archivo SQL que incluye:
#    - Tipos de cambios (EST, DAT, QRY).
#    - Detalles de la tarea como el link, descripción, autor, fechas de creación y despliegue (desarrollo, pre-producción, producción).
# 
# #### Salida
# 
# - Devuelve una tupla que contiene la `cabecera_sql` generada y el `codigo_tarea` extraído. (Si `codigo_tarea` no es correcto devuelve `None`)
# 
# #### Consideraciones Adicionales
# 
# - La función no valida directamente el formato de entrada del usuario excepto en la extracción del código de tarea.
# - Las fechas de despliegue (desarrollo, pre-producción, producción) se inicializan como cadenas vacías y no se modifican dentro de la función.
# - La cabecera SQL generada está formateada para ser insertada directamente en un archivo SQL, con detalles claros y estructurados para facilitar la comprensión del contexto de las modificaciones realizadas.
#  

# %%
# Definir los detalles de la tarea y pedirlos al usuario
def prepararCabecera():
    link_tarea = "https://app.clickup.com/t/36671967/"
    input_tarea = input("Introduce la parte del link de la tarea: (ej: PROCLI-3948)")
    link_tarea += input_tarea
    codigo_tarea = re.search(r'\d+', link_tarea).group(0)
    descripcion = input("Introduce la descripción: ")
    autor = input("Introduce el nombre del autor/es: ")
    fecha_creacion = hora_inicio.strftime('%Y-%m-%d')  # Formato: AAAA-MM-DD
    fecha_despliegue_desarrollo = ""
    fecha_despliegue_preproduccion = ""
    fecha_despliegue_produccion = ""

    # Extraer el número después del guion
    codigo_tarea = re.search(r'-(\d+)', input_tarea)
    if codigo_tarea:
        codigo_tarea = codigo_tarea.group(1)
    else:
        print("Formato de tarea incorrecto. Asegúrate de incluir un guion '-' seguido por números.")

    # Generar la cabecera del archivo SQL
    cabecera_sql = f"""
    /*
    * (EST): Estructura. 
    * (DAT): Modificación Datos.
    * (QRY): Consultas.
    */
    -------------------------------------------------------------------------------------
    /*
    * LINK TAREA: {link_tarea}
    * DESCRIPCIÓN: {descripcion}
    * 
    *
    * AUTOR: {autor}
    * FECHA CREACIÓN: {fecha_creacion}
    * FECHA DESPLIEGUE DESARROLLO: {fecha_despliegue_desarrollo}
    * FECHA DESPLIEGUE PRE-PRODUCCIÓN: {fecha_despliegue_preproduccion}
    * FECHA DESPLIEGUE PRODUCCIÓN: {fecha_despliegue_produccion}
    */
    -------------------------------------------------------------------------------------
    ---
    -------------------------------------------------
    --- 
    -------------------------------------------------
    BEGIN TRAN
    """
    #print(cabecera_sql)
    return cabecera_sql, codigo_tarea

# %% [markdown]
# ### Proceso de Generación de Scripts SQL desde Archivos Excel
# 
# Este script que parte del script de Nico (**script-generador.ps1**), automatiza la generación de scripts SQL a partir de datos extraídos de archivos Excel. Aquí está el desglose de su funcionalidad:
# 
# #### Inicialización de Contadores y Listas
# 
# - `contador_lineas_totales = 0`: Inicia un contador para el total de líneas de SQL generadas.
# - `contador = 1`: Un contador para llevar la cuenta de los scripts generados.
# 
# #### Procesamiento de Archivos Excel
# 
# - **Recorrido por Archivos Excel**: El script itera sobre cada archivo en la lista `archivos_excel`.
#   - Se abre cada archivo Excel usando `pandas.ExcelFile`.
#   - Se llama a la función `prepararCabecera` para obtener la cabecera SQL y el código de tarea.
# 
# - **Inicialización de Contenido SQL**: Se prepara una variable `contenido_columna` para acumular el contenido SQL.
# 
# - **Recorrido por Hojas de Excel**: 
#   - Se itera sobre cada hoja del archivo Excel.
#   - Se lee cada hoja en un DataFrame de pandas.
#   - Si la hoja tiene más de 18 columnas, se selecciona la columna 19 (columna "S").
# 
# - **Procesamiento de Datos de Columna**:
#   - Se busca por sentencias `DECLARE` y se extraen los nombres de las variables.
#   - Se realizan varias transformaciones en los datos de la columna para adecuarlos a la sintaxis SQL (por ejemplo, reemplazo de ciertos caracteres y ajuste de cadenas).
# 
# - **Inserción de Sentencias `GO` y `DECLARE`**: Cada 45 líneas, se inserta una sentencia `GO` y se redeclaran las variables encontradas.
# 
# - **Finalización del Contenido SQL**: Se añade una sentencia `GO` y `commit` al final del contenido.
# 
# #### Generación de Archivos SQL
# 
# - Se genera un nombre para el archivo SQL usando la fecha actual, el código de tarea y el contador de scripts.
# - Se verifica si el archivo ya existe; si no, se crea.
# - Se escribe el contenido SQL en el archivo.
# 
# #### Finalización
# 
# - Se incrementa el contador para el siguiente archivo.
# - Al finalizar el procesamiento de todos los archivos Excel, se imprime un mensaje de finalización.
# 
# #### Consideraciones Adicionales
# 
# - **¡IMPORTANTE: El script no maneja errores de formato en los archivos Excel, por lo que es importante asegurarse de que los archivos cumplan con los requisitos de formato antes de ejecutar el script!**
# - La funcionalidad está altamente enfocada en la estructura específica de los archivos Excel y los requisitos de formato SQL, lo que lo hace útil en escenarios donde esta estructura y requisitos son consistentes.
# - Muestra por pantalla las hojas que no tienen columna S, para que el usuario pueda verificar si es un error o no.
# - La idea es que poco a poco se vaya mejorando para que sea más eficiente y robusto. 
# 
# 
# Este script es una herramienta valiosa para automatizar la tediosa tarea de convertir datos de Excel en scripts SQL, ahorrando tiempo y reduciendo la posibilidad de errores manuales.
# 

# %%
contador_lineas_totales = 0
contador = 1
nombresSQL = []

# Recorrer cada archivo Excel en la lista
for archivo_excel in archivos_excel:
    print(f"Abriendo el archivo {archivo_excel}...")
    workbook = pd.ExcelFile(os.path.join(archivo_excel))

    print(f"Pidiendo datos del archivo Excel: {os.path.basename(archivo_excel)}")

    datos = prepararCabecera()

    cabecera_sql = datos[0]
    codigo_tarea = datos[1]

    # Inicializar el contenido que se escribirá en el archivo SQL
    contenido_columna = "" 
    
    # Añadir la cabecera SQL al contenido
    contenido_columna += cabecera_sql

    # Recorrer cada hoja en el libro Excel
    for sheet_name in workbook.sheet_names:
        #print(f"Hoja: {sheet_name}")
        # Leer la hoja actual en un DataFrame
        df = pd.read_excel(workbook, sheet_name=sheet_name)

        if df.shape[1] > 20:  # Verificando que hay al menos 19 columnas
        # Seleccionar la columna "U"
            columna_a_copiar = df.iloc[:, 20]
        else:
            print(f"La hoja {sheet_name} no tiene contenido en la columna U. Revisa que sea lo esperado.")
            continue  # Saltar al siguiente ciclo del bucle for si no hay columna U

        columna_a_copiar = df.iloc[:, 20]  # Ajusta esto a la columna que deseas copiar!

        # Array para almacenar las sentencias DECLARE
        declares = []
        
        # Agregar el nombre de la tabla al contenido
        contenido_columna += f"---Tabla: {sheet_name}\n"
        
        for valor in columna_a_copiar:
            if 'DECLARE' in str(valor):
            # Extraer el nombre de la variable y almacenarlo
                nombre_variable = re.search(r'(?<=DECLARE\s@)\w+', valor) # Guardamos el nombre de la variable por si en un futuro queremos hacer algo con ella
                if nombre_variable:
                    declares.append(nombre_variable.group())
            if len(str(valor)) > 70 or 'DECLARE' in str(valor):
                
                valor = str(valor).replace("''", "NULL")
                valor = valor.replace("'NULL'", "NULL")
                valor = valor.replace("%%", "''")
                valor = valor.replace("$$", "\n")
                valor = re.sub(r"(\w)'(\w)", r"\1''\2", valor)

                contenido_columna += valor + "\n"
                contador_lineas_totales += 1

                if contador_lineas_totales % 45 == 0:
                    contenido_columna += "GO\n"
                    for variable in declares:
                        contenido_columna += f"DECLARE @{variable} AS INT\nSET @{variable} = 0\n"
                    declares.clear()  # Limpiar el array para el siguiente lote

    contenido_columna += "GO\ncommit\n"
    print(f"Contador de scripts generados: {contador}")

    nombre_archivo_salida = f"{hora_inicio.strftime('%Y%m%d')}-{codigo_tarea}-00{contador}-DAT-{nombres_sin_extension[contador-1]}.sql"

    nombresSQL.append(nombre_archivo_salida)   


    ruta_archivo_salida = os.path.join(f"sql\{nombre_archivo_salida}")

    if os.path.exists(ruta_archivo_salida):
        print(f"El archivo {nombre_archivo_salida} ya existe. Se sobreescribirá.")
    else:
        os.makedirs(os.path.dirname(ruta_archivo_salida), exist_ok=True)
        
        with open(ruta_archivo_salida, "w") as archivo_salida:
            archivo_salida.write(contenido_columna)
            
        print (f"Archivo {nombre_archivo_salida} generado con éxito")


    
    print(f"Cerramos el archivo {archivo_excel}")
    # print("¡Recuerda borrar los archivos Excel que no necesites!")
    contador += 1

print("FIN")

# %% [markdown]
# ##### Renombrar los excels antiguos
# 
# Se renombran los excels antiguos para que no se vuelvan a procesar en caso de que el usuario no los haya borrado. O quiera tenerlos como historico. 
# 
# > Aún no funciona sale un error de que no se puede renombrar el archivo porque esta siendo usado por otro proceso. Se recomienda borrar los archivos excel ya procesados.

# %%
print("Se han generado los siguientes archivos SQL: ")
for nombre in nombresSQL:
    print(nombre)

# TODO: Renombrar los archivos Excel a archivo_old.xlsx Porque no me deja renombrarlos??
for archivo in archivos_excel:
    # os.remove(archivo) podria interesar mantenerlos para tener un historico
    print(f"{archivo}")
    nuevo_nombre = f"{archivo}"+"_old"
    
    try:
        # Renombrar el archivo
        os.rename(archivo, nuevo_nombre)
        print(f"Archivo renombrado: {archivo} a {nuevo_nombre}")
    except PermissionError as e:
        print(f"No se pudo renombrar el archivo {archivo}. Error: {e}")


