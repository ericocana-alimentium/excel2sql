# %%
import os
import glob
import pandas as pd
import datetime
import re
import tkinter as tk
from tkinter import filedialog, messagebox, Tk, scrolledtext, Toplevel
import warnings
warnings.simplefilter(action='ignore', category=UserWarning) #Ignorar warnings

# %% [markdown]
# ### Utils
# Esta sección contiene funciones de utilidad que son utilizadas a lo largo del script para generar el encabezado SQL y limpiar valores SQL.

# %%
def generate_sql_header(task_link, description, author, hora_inicio):
    codigo_tarea = re.search(r'-(\d+)', task_link)
    if codigo_tarea and codigo_tarea != " ":

        codigo_tarea = codigo_tarea.group(1)
    #else:
        #raise ValueError("Formato incorrecto en el identificador de la tarea. Ej. PROCLI-7777")
    
    cabecera_sql = f"""
/*
* (EST): Estructura. 
* (DAT): Modificación Datos.
* (QRY): Consultas.
*/
-------------------------------------------------------------------------------------
/*
* LINK TAREA: https://app.clickup.com/t/36671967/{task_link}
* DESCRIPCIÓN: {description}
* 
*
* AUTOR: {author}
* FECHA CREACIÓN: {hora_inicio.strftime('%Y-%m-%d')}
* FECHA DESPLIEGUE DESARROLLO: 
* FECHA DESPLIEGUE PRE-PRODUCCIÓN: 
* FECHA DESPLIEGUE PRODUCCIÓN: 
*/
-------------------------------------------------------------------------------------
---
-------------------------------------------------
--- 
-------------------------------------------------
BEGIN TRAN
    """
    
    return codigo_tarea, cabecera_sql

def clean_sql_value(value, declares):
    # Detectar y almacenar variables DECLARE
    if 'DECLARE' in str(value):
        nombre_variable = re.search(r'(?<=DECLARE\s@)\w+', value)
        if nombre_variable:
            declares.append(nombre_variable.group())
    
    # Convertir 'NULL' y valores vacíos
    value = str(value).replace("''", "NULL").replace("'NULL'", "NULL").replace("%%", "''").replace("$$", "\n")

    # Reemplazar espacios vacíos entre comas con NULL
    value = re.sub(r"(?<=\()\s*,|,\s*(?=\))|(?<=,)\s*(?=,)|(?<=,)\s*(?=$)", "NULL", value)

    # Manejar casos de comillas internas
    value = re.sub(r"(\w)'(\w)", r"\1''\2", value)
    
    return value

# %% [markdown]
# ### Procesamiento de archivos
# Esta sección define una clase para procesar archivos Excel y generar scripts SQL a partir de ellos.

# %%
class SQLFileProcessor:
 
    def __init__(self, path, task_link, description, author, mode="folder", output_dir=None):
        """
        Inicializa el procesador con los parámetros necesarios.
        Args:
        - path (str): Ruta al archivo Excel o directorio.
        - task_link (str): Enlace de la tarea.
        - description (str): Descripción de la tarea.
        - author (str): Nombre del autor.
        - mode (str): Modo de procesamiento ('folder' para carpeta, 'file' para archivo único).
        - output_dir (str, opcional): Directorio de salida para los archivos generados.
        """
         
        self.path = path
        self.task_link = task_link
        self.description = description
        self.author = author
        self.hora_inicio = datetime.datetime.now()
        self.mode = mode
        self.log_messages = []
        self.validation_data = {}
        
        self.archivos_excel = glob.glob(os.path.join(self.path, "*.xlsx")) if mode == "folder" else [self.path]
        self.output_dir = output_dir if output_dir else os.path.join(os.path.expanduser("~"), "Downloads")

    def process_files(self):

        """
        Procesa los archivos Excel y genera los scripts SQL correspondientes.

        Returns:
        - nombresSQL (list): Lista de nombres de archivos SQL generados.
        - log_messages (list): Lista de mensajes de registro.
        - hojas_no_procesadas (int): Número de hojas no procesadas.
        - validation_data (dict): Datos de validación de queries generadas por hoja.
        """

        contador_lineas_totales = 0
        contador = 0
        nombresSQL = []
        hojas_no_procesadas = 0

        for archivo_excel in self.archivos_excel:
            workbook = pd.ExcelFile(archivo_excel)
            codigo_tarea, cabecera_sql = generate_sql_header(
                self.task_link, self.description, self.author, self.hora_inicio
            )

            contenido_columna = cabecera_sql
            for sheet_name in workbook.sheet_names:
                df = pd.read_excel(workbook, sheet_name=sheet_name)

                # Verificar que el DataFrame no esté vacío y tenga al menos las columnas necesarias (21 para 'INSERT', 23 para 'UPDATE')
                if df.empty or (df.shape[1] < 21 and df.shape[1] < 23):
                    self.log_messages.append(f"Hoja {sheet_name}: No hay columnas suficientes para procesar.")
                    hojas_no_procesadas += 1
                    continue

                declares = []
                contenido_columna += f"---Tabla: {sheet_name}\n"
                queries_count = 0
                insert_count = 0
                update_count = 0

                # Procesar cada fila del DataFrame
                for index, row in df.iterrows():
                    try:
                        # Verificar si la columna 'UPDATE' (índice 22) existe y tiene un valor no nulo
                        update_value = row[22] if df.shape[1] > 22 else None
                        insert_value = row[20] if df.shape[1] > 20 else None

                        if pd.notna(update_value) and isinstance(update_value, str) and update_value.strip().lower().startswith("update"):
                            update_value = clean_sql_value(update_value.strip(), declares)
                            contenido_columna += update_value + "\n"
                            queries_count += 1
                            update_count += 1
                            contador_lineas_totales += 1

                        # Si no hay 'UPDATE', procesar 'INSERT'
                        elif pd.notna(insert_value) and isinstance(insert_value, str) and insert_value.strip().lower().startswith("insert into"):
                            insert_value = clean_sql_value(insert_value.strip(), declares)
                            contenido_columna += insert_value + "\n"
                            queries_count += 1
                            insert_count += 1
                            contador_lineas_totales += 1

                        # Control de declaración 'GO' cada 45 líneas
                        if contador_lineas_totales % 45 == 0:
                            contenido_columna += "GO\n"
                            for variable in declares:
                                contenido_columna += f"DECLARE @{variable} AS INT\nSET @{variable} = 0\n"
                            declares.clear()

                    except Exception as e:
                        self.log_messages.append(f"Error: al procesar la fila {index} en la hoja {sheet_name}: {e}")
                        continue

                # Solo agregar datos de validación si hay consultas generadas
                if queries_count > 0:
                    self.validation_data[sheet_name] = {
                        "total_queries": queries_count,
                        "inserts": insert_count,
                        "updates": update_count
                    }

            contenido_columna += "GO\nROLLBACK\n--COMMIT\n"
            nombre_archivo_salida = f"{self.hora_inicio.strftime('%Y%m%d')}-{codigo_tarea}-00{contador}-DAT-{os.path.basename(archivo_excel).split('.')[0]}.sql"
            nombresSQL.append(nombre_archivo_salida)

            ruta_archivo_salida = os.path.normpath(os.path.join(self.output_dir, nombre_archivo_salida))
            print(f"Saving file to: {ruta_archivo_salida}")

            try:
                os.makedirs(os.path.dirname(ruta_archivo_salida), exist_ok=True)

                with open(ruta_archivo_salida, "w", encoding="utf-8") as archivo_salida:
                    archivo_salida.write(contenido_columna)
                    print(f"Archivo {nombre_archivo_salida} generado con éxito en {ruta_archivo_salida}")
            except Exception as e:
                print(f"Error al guardar el archivo: {e}")

        return nombresSQL, self.log_messages, hojas_no_procesadas, self.validation_data


# %% [markdown]
# ### GUI
# Esta sección define una clase para la interfaz gráfica de usuario (GUI) que permite seleccionar archivos y carpetas, configurar opciones, y generar los scripts SQL mediante interacción con la aplicación

# %%
class SQLGeneratorApp:
    def __init__(self):

        """
        Inicializa la aplicación y configura los elementos de la interfaz gráfica.
        """

        self.root = Tk()
        self.root.title("Excel 2 SQL v3.5")
        
        #icon_path = os.path.join(os.getcwd(), 'assets', 'icon.ico')
        #self.root.iconbitmap(icon_path)

        # Variables
        self.directory = tk.StringVar()
        self.filepath = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Downloads"))
        self.task_link = tk.StringVar()
        self.description = tk.StringVar()
        self.author = tk.StringVar()
        self.mode = tk.StringVar(value="file")
        
        self.log_messages = []  # Lista para capturar mensajes de depuración
        self.validation_data = {}

        # Configuración de columnas para centrado
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

        # Widgets
        self.create_widgets()
        self.root.after(0, self.update_mode)  
        self.root.mainloop()
        
        
    def create_widgets(self):
        
        """
        Crea y configura los widgets de la interfaz gráfica de usuario.
        """

        tk.Label(self.root, text="Seleccionar modo:").grid(row=0, column=0, padx=10, pady=10)
        tk.Radiobutton(self.root, text="Multiple excels (Carpeta)", variable=self.mode, value="folder", command=self.update_mode).grid(row=0, column=1, padx=10, pady=10)
        tk.Radiobutton(self.root, text="Único archivo", variable=self.mode, value="file", command=self.update_mode).grid(row=0, column=2, padx=10, pady=10)

        self.directory_label = tk.Label(self.root, text="Seleccionar directorio de los excels:")
        self.directory_label.grid(row=1, column=0, padx=10, pady=10)
        self.directory_entry = tk.Entry(self.root, textvariable=self.directory, width=50)
        self.directory_entry.grid(row=1, column=1, padx=10, pady=10)
        self.directory_button = tk.Button(self.root, text="Browse", command=self.browse_directory)
        self.directory_button.grid(row=1, column=2, padx=10, pady=10)
        
        self.filepath_label = tk.Label(self.root, text="Seleccionar archivo excel:")
        self.filepath_entry = tk.Entry(self.root, textvariable=self.filepath, width=50)
        self.filepath_button = tk.Button(self.root, text="Browse", command=self.browse_file)


        tk.Label(self.root, text="Directorio de salida:").grid(row=2, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_output_directory).grid(row=2, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Identificador de la tarea:").grid(row=3, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.task_link, width=50).grid(row=3, column=1, padx=10, pady=10)
        
        
        tk.Label(self.root, text="Descripción:").grid(row=4, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.description, width=50).grid(row=4, column=1, padx=10, pady=10)
        
        tk.Label(self.root, text="Autor:").grid(row=5, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.author, width=50).grid(row=5, column=1, padx=10, pady=10)
        
       
        tk.Button(self.root, text="Generar SQLs", command=self.generate_sql_files).grid(row=6, column=0, columnspan=3, padx=10, pady=20)

        # Crear un frame para contener los dos botones
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)

        self.show_log_button = tk.Button(button_frame, text="Mostrar Logs", command=self.show_log)
        self.show_log_button.pack(side=tk.LEFT, padx=5)

        self.show_validation_button = tk.Button(button_frame, text="Outputs de validación", command=self.show_validation)
        self.show_validation_button.pack(side=tk.LEFT, padx=5)
        
        # Ocultar inicialmente
        self.show_log_button.pack_forget()
        self.show_validation_button.pack_forget()


    def update_mode(self):

        """
        Actualiza la interfaz de usuario según el modo seleccionado (archivo único o carpeta).
        """

        if self.mode.get() == "folder":
            self.directory_label.grid(row=1, column=0, padx=10, pady=10)
            self.directory_entry.grid(row=1, column=1, padx=10, pady=10)
            self.directory_button.grid(row=1, column=2, padx=10, pady=10)
            
            self.filepath_label.grid_remove()
            self.filepath_entry.grid_remove()
            self.filepath_button.grid_remove()
            #self.process_button.config(text="Generar archivos SQL")
        else:
            self.filepath_label.grid(row=1, column=0, padx=10, pady=10)
            self.filepath_entry.grid(row=1, column=1, padx=10, pady=10)
            self.filepath_button.grid(row=1, column=2, padx=10, pady=10)
            
            self.directory_label.grid_remove()
            self.directory_entry.grid_remove()
            self.directory_button.grid_remove()
            #self.process_button.config(text="Generar SQL")
    
    def browse_directory(self):
        """
        Muestra un cuadro de diálogo para seleccionar un directorio y guarda la selección.
        """
        directory = filedialog.askdirectory()
        if directory:
            self.directory.set(directory)
    
    def browse_file(self):
        """
        Muestra un cuadro de diálogo para seleccionar un archivo Excel y guarda la selección.
        """
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.filepath.set(file_path)

    def browse_output_directory(self):
        """
        Muestra un cuadro de diálogo para seleccionar un directorio de salida y guarda la selección.
        """
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(directory)
    
    def generate_sql_files(self):

        """
        Genera los archivos SQL llamando a SQLFileProcessor y maneja el resultado mostrando mensajes en la GUI.
        """

        task_link = self.task_link.get()
        description = self.description.get()
        author = self.author.get()
        
        if self.mode.get() == "folder":
            path = self.directory.get()
            if not path: #or not task_link or not description or not author:
                messagebox.showwarning("Input Error", "Selecciona un directorio")
                return
            
            processor = SQLFileProcessor(path, task_link, description, author, mode="folder")
        else:
            path = self.filepath.get()
            if not path: # or not task_link or not description or not author:
                messagebox.showwarning("Input Error", "Selecciona un archivo excel")
                return
            
            processor = SQLFileProcessor(path, task_link, description, author, mode="file")
        
        try:
            generated_files, self.log_messages, hojas_no_procesadas, self.validation_data = processor.process_files()

            if self.log_messages:
                self.show_log_button.pack(side=tk.LEFT, padx=5)  # Mostrar el botón si hay logs
            
            if hojas_no_procesadas > 0:
                tk.Label(self.root, text=f"Total hojas no procesadas: {hojas_no_procesadas}").grid(row=8, column=0, columnspan=3, padx=10, pady=10)
                if self.validation_data:
                    self.show_validation_button.pack(side=tk.LEFT, padx=5) 
                messagebox.showinfo("Success", f"Generado {len(generated_files)} archivos SQL correctamente")
            

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def filter_logs(self, filter_text, log_text_widget):

        """
        Filtra los logs según el texto de búsqueda introducido.

        Args:
        - filter_text (str): Texto por el cual filtrar los logs.
        - log_text_widget (tk.Text): Widget de texto donde se muestran los logs.
        """

        log_text_widget.config(state=tk.NORMAL)
        log_text_widget.delete(1.0, tk.END)
        for message in self.log_messages:
            if filter_text.lower() in message.lower():
                log_text_widget.insert(tk.END, message + "\n")
        log_text_widget.config(state=tk.DISABLED)

    def show_log(self):

        """
        Muestra una nueva ventana con los logs de depuración. De cara a las hojas que se procesan o no, al igual que los errores
        """
        log_window = tk.Toplevel(self.root)
        log_window.title("Log de depuración")

        # Entry para buscar en los logs
        search_frame = tk.Frame(log_window)
        search_frame.pack(padx=10, pady=10)
        tk.Label(search_frame, text="Filtrar por:").pack(side=tk.LEFT)
        search_entry = tk.Entry(search_frame)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<KeyRelease>", lambda event: self.filter_logs(search_entry.get(), log_text))

        log_text = scrolledtext.ScrolledText(log_window, width=100, height=30)
        log_text.pack(padx=10, pady=10)
        
        # log messages
        for message in self.log_messages:
            log_text.insert(tk.END, message + "\n")
        
        log_text.config(state=tk.DISABLED)  # solo lectura


    def filter_validation(self, filter_text, validation_text_widget):

        """
        Filtra los resultados de validación según el texto de búsqueda introducido.

        Args:
        - filter_text (str): Texto por el cual filtrar la validación.
        - validation_text_widget (tk.Text): Widget de texto donde se muestran los resultados de validación.
        """

        validation_text_widget.config(state=tk.NORMAL)
        validation_text_widget.delete(1.0, tk.END)
        for sheet_name, data in self.validation_data.items():
            total = data["total_queries"]
            inserts = data["inserts"]
            updates = data["updates"]
            line = f"Hoja {sheet_name}: {total} queries generadas."
            if total > 0:
                if inserts > 0:
                    line += f" INSERTs: {inserts}"
                if updates > 0:
                    line += f" UPDATEs: {updates}"
                line += "\n"
                
                # Verificar si la línea coincide con el filtro
                if filter_text.lower() in line.lower():
                    validation_text_widget.insert(tk.END, line)
        validation_text_widget.config(state=tk.DISABLED)

    def show_validation(self):

        """
        Muestra una nueva ventana con los resultados de validación de las queries generadas.
        """

        validation_window = Toplevel(self.root)
        validation_window.title("Validación de queries generadas")

        # Entry para buscar en la validación
        search_frame = tk.Frame(validation_window)
        search_frame.pack(padx=10, pady=10)
        tk.Label(search_frame, text="Filtrar por:").pack(side=tk.LEFT)
        search_entry = tk.Entry(search_frame)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<KeyRelease>", lambda event: self.filter_validation(search_entry.get(), validation_text))

        validation_text = scrolledtext.ScrolledText(validation_window, width=100, height=30)
        validation_text.pack(padx=10, pady=10)
        
        for sheet_name, data in self.validation_data.items():
            total = data["total_queries"]
            inserts = data["inserts"]
            updates = data["updates"]
             # Siempre muestra el total de queries generadas
            validation_text.insert(tk.END, f"Hoja {sheet_name}: {total} queries generadas.")
            # Solo muestra INSERTs y UPDATEs si son mayores que 0
            if total > 0:
                details = []
                if inserts > 0:
                    details.append(f"INSERTs: {inserts}")
                if updates > 0:
                    details.append(f"UPDATEs: {updates}")
                if details:
                    validation_text.insert(tk.END, f" ({', '.join(details)})")
            validation_text.insert(tk.END, "\n")    
        
        validation_text.config(state=tk.DISABLED)

# %%
# Esta celda se utiliza para iniciar la aplicación desde el notebook
app = SQLGeneratorApp()



