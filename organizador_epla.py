import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
import shutil
import threading
from pathlib import Path
import requests
from urllib.parse import urlparse
import re
import sys

# --- CONFIGURACIÓN ---
# Nombres exactos de las columnas en el archivo Excel.
COLUMNA_NOMBRE = "Nombre del alumno/a"
COLUMNA_PRIMER_APELLIDO = "Primer apellido del alumno/a"
COLUMNA_SEGUNDO_APELLIDO = "Segundo apellido del alumno/a" # Asegúrate de que esta línea exista

# Rango de columnas que contienen URLs de PDF.
COLUMNAS_PDF_URLS = list(range(9, 17)) # Columnas J (9) a Q (16)

# --------------------

class FileOrganizerApp(tk.Tk):
    """
    Aplicación con interfaz gráfica para organizar archivos de alumnos
    basado en un fichero Excel.
    """
    def __init__(self):
        super().__init__()

        # --- Variables de estado ---
        self.target_folder = tk.StringVar()
        self.excel_file_path = tk.StringVar()
        self.is_processing = False

        # --- Configuración de la ventana principal ---
        self.title("Escuelas profesionales Luis Amigó - SECRETARIA (Operaciones de archivo)")
        self.geometry("600x650")
        self.resizable(False, False)

        # --- Logo ---
        try:
            if hasattr(sys, '_MEIPASS'):
                logo_path = os.path.join(sys._MEIPASS, "logo.png")
            else:
                logo_path = "logo.png"
            logo_image = Image.open(logo_path)
            logo_image = logo_image.resize((150, 150), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_image)
            tk.Label(self, image=self.logo).pack(pady=10)
        except FileNotFoundError:
            # Si no se encuentra el logo, se muestra un texto en su lugar.
            tk.Label(self, text="Logo no encontrado", fg="red").pack(pady=10)
            print("Error: Asegúrate de que 'logo.png' está en la misma carpeta que el script o el ejecutable.")
        except Exception as e:
            tk.Label(self, text=f"Error al cargar logo: {e}", fg="red").pack(pady=10)

        # --- Contenedor principal ---
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill="both", expand=True)

        self.create_widgets(main_frame)
        
    def create_widgets(self, parent):
        """Crea todos los widgets de la interfaz."""
        # --- 1. Selección de Carpeta de Destino ---
        ttk.Label(parent, text="Paso 1: Selecciona la carpeta donde se crearán los directorios.", wraplength=500).pack(fill='x', pady=(0, 5))
        
        folder_frame = ttk.Frame(parent)
        folder_frame.pack(fill='x', pady=5)
        
        entry_folder = ttk.Entry(folder_frame, textvariable=self.target_folder, state='readonly', width=60)
        entry_folder.pack(side='left', fill='x', expand=True)
        
        self.btn_select_folder = ttk.Button(folder_frame, text="Seleccionar...", command=self.select_target_folder)
        self.btn_select_folder.pack(side='right', padx=(5, 0))

        # --- 2. Selección de Archivo Excel ---
        ttk.Label(parent, text="Paso 2: Carga el archivo Excel con los datos de los alumnos.", wraplength=500).pack(fill='x', pady=(20, 5))
        
        excel_frame = ttk.Frame(parent)
        excel_frame.pack(fill='x', pady=5)
        
        entry_excel = ttk.Entry(excel_frame, textvariable=self.excel_file_path, state='readonly', width=60)
        entry_excel.pack(side='left', fill='x', expand=True)

        self.btn_select_excel = ttk.Button(excel_frame, text="Cargar Excel...", command=self.select_excel_file)
        self.btn_select_excel.pack(side='right', padx=(5, 0))

        # --- 3. Iniciar Proceso ---
        self.btn_start = ttk.Button(parent, text="Iniciar Organización de Archivos", command=self.start_processing_thread, style='Accent.TButton')
        self.btn_start.pack(pady=30, ipady=10)
        ttk.Style().configure('Accent.TButton', font=('Helvetica', 10, 'bold'))

        # --- Barra de Progreso y Estado ---
        self.progress_bar = ttk.Progressbar(parent, orient='horizontal', length=100, mode='determinate')
        self.progress_bar.pack(fill='x', pady=10)

        self.status_label = ttk.Label(parent, text="Listo para iniciar.", anchor='center')
        self.status_label.pack(fill='x', pady=5)

    def select_target_folder(self):
        """Abre un diálogo para seleccionar la carpeta de destino."""
        folder_selected = filedialog.askdirectory(title="Selecciona la carpeta de destino")
        if folder_selected:
            self.target_folder.set(folder_selected)
            self.status_label.config(text=f"Carpeta de destino: {folder_selected}")

    def select_excel_file(self):
        """Abre un diálogo para seleccionar el archivo Excel."""
        file_selected = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=(("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*"))
        )
        if file_selected:
            self.excel_file_path.set(file_selected)
            self.status_label.config(text=f"Archivo Excel cargado: {Path(file_selected).name}")

    def start_processing_thread(self):
        """Inicia el proceso de organización en un hilo separado para no bloquear la GUI."""
        if not self.target_folder.get() or not self.excel_file_path.get():
            messagebox.showerror("Error", "Debes seleccionar una carpeta de destino y un archivo Excel antes de iniciar.")
            return
        
        if self.is_processing:
            messagebox.showwarning("Atención", "El proceso ya está en ejecución.")
            return

        self.is_processing = True
        self.toggle_buttons(False) # Deshabilitar botones
        
        # Crear y empezar el hilo
        process_thread = threading.Thread(target=self.process_files)
        process_thread.daemon = True
        process_thread.start()

    def get_google_drive_direct_download_url(self, gd_url):
        """
        Intenta transformar una URL de Google Drive a una URL de descarga directa.
        Funciona para archivos compartidos públicamente.
        """
        # Expresión regular para encontrar el ID del archivo en varias formas de URL de Google Drive
        match = re.search(r'drive\.google\.com/(?:file/d/|open\?id=|uc\?id=)([a-zA-Z0-9_-]+)', gd_url)
        if match:
            file_id = match.group(1)
            # URL de descarga directa para Google Drive
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        return None # No se encontró un ID de archivo válido de Google Drive

    def download_file(self, url, destination_folder, file_name=None):
        """Descarga un archivo desde una URL a una carpeta de destino."""
        if not url or not isinstance(url, str):
            return None # Ignorar si la URL no es válida o no es una cadena

        original_url = url
        # Intentar transformar la URL si es de Google Drive
        if "drive.google.com" in url:
            transformed_url = self.get_google_drive_direct_download_url(url)
            if transformed_url:
                url = transformed_url
                print(f"URL de Google Drive transformada a descarga directa: {original_url} -> {url}")
            else:
                print(f"Advertencia: No se pudo transformar la URL de Google Drive: {original_url}")
                return None # Si no se puede transformar, probablemente no se podrá descargar.

        try:
            if not file_name:
                parsed_url = urlparse(original_url) # Usar la original para el nombre del archivo
                file_name = os.path.basename(parsed_url.path)
                # Si el nombre del archivo es el ID de Google Drive (común después de la transformación)
                # o si no tiene extensión, podemos intentar deducir que es un PDF.
                if not file_name or '.' not in file_name or (file_name == parsed_url.path.split('/')[-1] and "drive.google.com" in original_url):
                    if file_name and file_name != "uc":
                        file_name = f"{file_name}.pdf"
                    else:
                        file_id_match = re.search(r'id=([a-zA-Z0-9_-]+)', url)
                        if file_id_match:
                            file_name = f"documento_{file_id_match.group(1)[:8]}.pdf"
                        else:
                            file_name = "documento_descargado.pdf"
                if not file_name.lower().endswith('.pdf'):
                    if '.' in file_name and file_name.split('.')[-1].lower() == 'pdf':
                        pass
                    else:
                        file_name = f"{file_name}.pdf"
            destination_path = os.path.join(destination_folder, file_name)

            # Evitar descargar si el archivo ya existe en el destino
            if os.path.exists(destination_path):
                print(f"Info: El archivo '{file_name}' ya existe en '{destination_folder}'. Saltando descarga.")
                return destination_path

            # Realizar la solicitud HTTP para descargar el archivo
            response = requests.get(url, stream=True, timeout=15) # Aumentado timeout un poco
            response.raise_for_status() # Lanza un error para códigos de estado HTTP incorrectos (4xx o 5xx)

            # Verificar el tipo de contenido para asegurar que es un PDF
            content_type = response.headers.get('Content-Type', '')
            if 'application/pdf' not in content_type and 'octet-stream' not in content_type:
                print(f"Advertencia: La URL '{original_url}' no parece ser un PDF (Content-Type: {content_type}). Descargando de todos modos como {file_name}.")
                # Opcional: podrías decidir no descargar si no es PDF
                # return None

            with open(destination_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"Descargado: {file_name} a {destination_folder}")
            return destination_path

        except requests.exceptions.RequestException as req_err:
            print(f"Error de red/solicitud al descargar {original_url}: {req_err}")
            return None
        except Exception as e:
            print(f"Error inesperado al descargar {original_url}: {e}")
            return None

    def process_files(self):
        """
        Lógica principal: Lee el Excel, crea carpetas y descarga/mueve los PDFs.
        Esta función se ejecuta en un hilo secundario.
        """
        try:
            target_dir = self.target_folder.get()
            excel_path = self.excel_file_path.get()
            
            # Leer el archivo Excel con pandas
            df = pd.read_excel(excel_path)
            
            # Validar que las columnas necesarias existan
            required_columns = [COLUMNA_NOMBRE, COLUMNA_PRIMER_APELLIDO, COLUMNA_SEGUNDO_APELLIDO]
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"El archivo Excel debe contener la columna '{col}'.")

            total_rows = len(df)
            self.update_progress(0, "Iniciando proceso...")

            # Detectar dinámicamente todas las columnas desde la J en adelante
            pdf_url_column_names = list(df.columns[9:])


            for index, row in df.iterrows():
                # Actualizar progreso y estado en el hilo principal de la GUI
                progress = (index + 1) / total_rows * 100
                status_text = f"Procesando alumno: {index + 1}/{total_rows}"
                self.after(0, self.update_progress, progress, status_text)

                # Obtener nombre y apellidos, limpiando posibles espacios y convirtiendo a string
                nombre = str(row[COLUMNA_NOMBRE]).strip()
                primer_apellido = str(row[COLUMNA_PRIMER_APELLIDO]).strip()
                segundo_apellido = str(row[COLUMNA_SEGUNDO_APELLIDO]).strip()

                # Ignorar filas donde el nombre o algún apellido esté vacío (ej. NaN)
                if not nombre or nombre == "nan" or \
                   not primer_apellido or primer_apellido == "nan" or \
                   not segundo_apellido or segundo_apellido == "nan":
                    print(f"Saltando fila {index+1}: Datos de alumno incompletos (nombre/apellidos).")
                    continue

                # Crear el nombre de la subcarpeta con el formato "Primer Apellido Segundo Apellido Nombre"
                subfolder_name = f"{primer_apellido} {segundo_apellido} {nombre}"
                
                # Sanitizar el nombre de la carpeta para evitar caracteres inválidos
                sanitized_subfolder_name = "".join(c for c in subfolder_name if c.isalnum() or c in (' ', '.', '_', '-')).rstrip()
                sanitized_subfolder_name = " ".join(sanitized_subfolder_name.split()) # Eliminar espacios dobles

                student_folder_path = os.path.join(target_dir, sanitized_subfolder_name)

                # Crear la carpeta del alumno si no existe
                os.makedirs(student_folder_path, exist_ok=True)
                
                # Procesar URLs de PDF de las columnas especificadas
                for col_name in pdf_url_column_names:
                    pdf_url = str(row[col_name]).strip()
                    # Verificar si el valor es una URL válida (al menos que empiece por http/https)
                    if pdf_url.lower().startswith('http://') or pdf_url.lower().startswith('https://'):
                        nombre_archivo = f"{col_name}.pdf"
                        self.after(0, lambda n=nombre, p=primer_apellido, f=nombre_archivo: self.status_label.config(text=f"Descargando para {n} {p}: {f}"))
                        downloaded_path = self.download_file(pdf_url, student_folder_path, nombre_archivo)
                        if downloaded_path:
                            print(f"PDF procesado para {sanitized_subfolder_name}: {Path(downloaded_path).name}")
                        else:
                            print(f"Falló la descarga del PDF de la URL: {pdf_url} para {sanitized_subfolder_name}")
                    elif pdf_url and pdf_url != "nan": # Si hay algo pero no es una URL, imprimir advertencia
                        print(f"Advertencia: El contenido de la celda '{col_name}' no parece una URL válida para {sanitized_subfolder_name}: '{pdf_url}'")
            
            self.after(0, self.process_finished)

        except FileNotFoundError:
            self.after(0, messagebox.showerror, "Error", "No se encontró el archivo Excel. Verifica la ruta.")
            self.after(0, self.reset_ui)
        except ValueError as ve:
            self.after(0, messagebox.showerror, "Error de Columnas", str(ve))
            self.after(0, self.reset_ui)
        except pd.errors.EmptyDataError:
            self.after(0, messagebox.showerror, "Error de Excel", "El archivo Excel está vacío o tiene un formato no válido.")
            self.after(0, self.reset_ui)
        except Exception as e:
            self.after(0, messagebox.showerror, "Error Inesperado", f"Ocurrió un error: {e}")
            self.after(0, self.reset_ui)
            
    def update_progress(self, value, text):
        """Actualiza la barra de progreso y el texto de estado."""
        self.progress_bar['value'] = value
        self.status_label.config(text=text)

    def process_finished(self):
        """Se llama cuando el proceso ha terminado con éxito."""
        self.update_progress(100, "¡Proceso completado con éxito!")
        messagebox.showinfo("Finalizado", "Todos los archivos han sido organizados correctamente.")
        self.reset_ui()

    def reset_ui(self):
        """Reinicia la interfaz a su estado inicial."""
        self.is_processing = False
        self.toggle_buttons(True)
        self.progress_bar['value'] = 0
        self.status_label.config(text="Listo para iniciar de nuevo.")

    def toggle_buttons(self, status):
        """Habilita o deshabilita los botones de la interfaz."""
        state = 'normal' if status else 'disabled'
        self.btn_select_folder.config(state=state)
        self.btn_select_excel.config(state=state)
        self.btn_start.config(state=state)


if __name__ == "__main__":
    app = FileOrganizerApp()
    app.mainloop()