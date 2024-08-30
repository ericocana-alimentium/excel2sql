
---

## 📘 Guía de Uso: Aplicación "Excel 2 SQL"

### Índice

1. [Introducción](#introducción)
2. [Instalación](#instalación)
3. [Uso de la aplicación](#uso-de-la-aplicación)
4. [Detalles del código](#detalles-del-código)
   - [Funciones de utilidad](#funciones-de-utilidad)
   - [Procesamiento de archivos](#procesamiento-de-archivos)
   - [Interfaz gráfica de usuario (GUI)](#interfaz-gráfica-de-usuario-gui)
5. [Notas adicionales](#notas-adicionales)

---

### Introducción

La aplicación **"Excel 2 SQL"** está diseñada para facilitar la conversión de archivos Excel a scripts SQL. A través de una interfaz gráfica amigable, el usuario puede seleccionar archivos o directorios que contienen hojas de cálculo y generar automáticamente los scripts SQL correspondientes.

---

### Instalación

Sigue estos pasos para instalar las dependencias necesarias y ejecutar la aplicación:

1. **Clona el repositorio desde GitHub:**
   ```bash
   git clone https://github.com/ericocana-alimentium/excel2sql.git
   ```

2. **Navega al directorio del proyecto:**
   ```bash
   cd excel2sql
   ```

3. **Instala las dependencias requeridas:**
   ```bash
   pip install -r requirements.txt
   ```
   Si usas Anaconda seguramente no te sea necesario

4. **Ejecuta la aplicación:**
   ```bash
   python excel2sql.py
   ```

---

### Uso de la aplicación

Una vez que hayas instalado la aplicación, sigue estos pasos para usarla:

1. **Inicia la aplicación:** Ejecuta el script principal (`excel2sql.py`) para abrir la interfaz gráfica.
También en caso de querer depurar con celdas abrir con el VS Code el  (`excel2sql_v3.5.ipynb`)

2. **Selecciona el modo de operación:**
   - **Multiple excels (Carpeta):** Selecciona una carpeta que contenga varios archivos Excel.
   - **Único archivo:** Selecciona un único archivo Excel.

3. **Configura las opciones necesarias:**
   - **Directorio/Archivo Excel:** Selecciona la carpeta o archivo que deseas procesar.
   - **Directorio de salida:** Especifica el directorio donde se guardarán los scripts SQL generados. (Por defecto será Descargar)
   - **Identificador de la tarea:** Identificador de la tarea.
   - **Descripción y Autor:** Completa la descripción de la tarea y el nombre del autor.

4. **Genera los archivos SQL:**
   - Haz clic en "Generar SQLs" para comenzar el proceso de conversión.
   - La aplicación mostrará mensajes de que ha finalizado el proceso tanto si es correcto como si ha ocurrido un error.

5. **Visualiza los resultados:**
   - Utiliza los botones "Mostrar Logs" y "Outputs de validación" para revisar los detalles del proceso.
   - `Mostrar Logs`: principalmente nos sirve para saber que hojas no se han procesado, por un error o porque no ha encontrado columnas U y/o W (que no siempre significa que este mal como por ejemplo la tabla 00.DatosInicialesCliente)
   - `Outputs de validación`: Nos sirve para saber de cada hoja cuantas queries se han generado, cuantos inserts y updates para así contrastar con el excel
   - Filtrado: En las dos ventanas podemos filtrar por palabras claves.

---

### Detalles del código

#### Funciones de utilidad

- **`generate_sql_header(task_link, description, author, hora_inicio)`**
  - Genera un encabezado SQL basado en la información proporcionada.
  - Extrae el código de tarea del enlace de la tarea y formatea el encabezado.

- **`clean_sql_value(value, declares)`**
  - Limpia y formatea valores SQL para asegurar que sean compatibles y estén correctamente estructurados.
  - Maneja variables `DECLARE`, reemplazos de `NULL`, y otros ajustes de formato.

#### Procesamiento de archivos

- **Clase `SQLFileProcessor`**
  - Se encarga de procesar archivos Excel y convertirlos en scripts SQL.
  - Métodos clave:
    - `__init__`: Inicializa los parámetros del procesador.
    - `process_files`: Procesa los archivos Excel y genera los scripts SQL. 

  - Aclaración: la lógica está para que si hay un INSERT y un UPDATE a la vez, se cogera el UPDATE para evitar que salte error en el DBUP.

#### Interfaz gráfica de usuario (GUI)

- **Clase `SQLGeneratorApp`**
  - Define la interfaz gráfica utilizando Tkinter.
  - Métodos clave:
    - `__init__`: Inicializa la aplicación y configura los widgets.
    - `create_widgets`: Crea y organiza los widgets en la interfaz.
    - `generate_sql_files`: Llama al procesador de archivos y maneja la generación de scripts SQL.

---

### Notas adicionales

- **¡IMPORTANTE: El script no maneja errores de formato en los archivos Excel, por lo que es importante asegurarse de que los archivos cumplan con los requisitos de formato antes de ejecutar el script!**
- La funcionalidad está altamente enfocada en la estructura específica de los archivos Excel y los requisitos de formato SQL, lo que lo hace útil en escenarios donde esta estructura y requisitos son consistentes.
- La aplicación "manejará" los errores de procesamiento y proporcionará mensajes de registro detallados para ayudar en la depuración.
- Si se procesan grandes cantidades de datos, puede ser necesario ajustar la lógica de control de declaraciones `GO` para optimizar el rendimiento.

---
