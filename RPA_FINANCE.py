import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import os
import pandas as pd
from datetime import datetime
import shutil
import xlrd
import tempfile
import openpyxl

def num_to_excel_col(n):
        col = ""
        while n >= 0:
            col = chr(n % 26 + ord('A')) + col
            n = n // 26 - 1
        return col
    
def convertir_fechas(fecha):
    # Si la fecha ya es un objeto datetime, la convertimos al formato deseado
    if isinstance(fecha, datetime):
        return fecha.strftime('%d-%m-%Y')

    # Diccionario para los meses en inglés y español
    meses = {
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
        'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12',
        'ENE': '01', 'FEB': '02', 'MAR': '03', 'ABR': '04', 'MAY': '05', 'JUN': '06',
        'JUL': '07', 'AGO': '08', 'SET': '09', 'OCT': '10', 'NOV': '11', 'DIC': '12'
    }
    
    # Verificar si la fecha ya está en el formato correcto
    try:
        return datetime.strptime(fecha, '%d-%m-%Y').strftime('%d-%m-%Y')
    except ValueError:
        pass  # Si falla, intentamos con otros formatos

    # Intentar convertir fechas en formato DD-MMM-YYYY
    for mes in meses:
        if mes in fecha:
            partes = fecha.split('-')
            if len(partes) == 3:
                dia = partes[0]
                mes_num = meses[partes[1]]
                anio = partes[2]
                return f"{dia}-{mes_num}-{anio}"
    
    return fecha  # Retornar la fecha original si no se pudo convertir

class RPAFinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RPA FINANCE")

        # Establecer el icono de la ventana
        self.root.iconbitmap(r'C:\Users\yaco_acuna\Desktop\M6\Python\RIM_PENDIENTE\RPA_FINANCE.ico')
        
        # Cambiar el color de fondo de la ventana principal
        self.root.configure(bg="#fdf3e7")
        
        # Variables de control
        self.proceso_actual = ""
        self.rpa_finance_path = ""
        
        # Configuración de la ventana
        self.root.geometry("600x330")
        
        # Crear carpeta RPA FINANCE en OneDrive si no existe
        onedrive_path = os.path.expanduser("~/OneDrive")
        self.rpa_finance_path = os.path.join(onedrive_path, "RPA FINANCE")
        if not os.path.exists(self.rpa_finance_path):
            os.makedirs(self.rpa_finance_path)
        
        # Crear menú
        self.crear_menu()
            
        # Frame principal para organizar los elementos
        self.main_frame = tk.Frame(root, padx=20, pady=20, bg="#fdf3e7")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame para mensaje inicial
        self.mensaje_inicial_frame = tk.Frame(self.main_frame, bg="#fdf3e7")
        self.mensaje_inicial_frame.pack(fill=tk.BOTH, expand=True)
        
        self.mensaje_inicial = tk.Label(
            self.mensaje_inicial_frame, 
            text="Seleccione un proceso del menú 'Procesos' para comenzar",
            font=("Arial", 12),
            bg="#fdf3e7"  # Color de fondo del label
        )
        self.mensaje_inicial.pack(expand=True)
        
        # Crear frames para ambos procesos (inicialmente ocultos)
        self.crear_frames_proceso1()
        self.crear_frames_proceso2()
        
        self.selected_files = []

    def crear_menu(self):
        # Crear barra de menú
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)

        # Menú Reporte
        reporte_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Reporte", menu=reporte_menu)
        reporte_menu.add_command(label="Nuevo", command=self.nuevo_reporte)
        reporte_menu.add_separator()
        reporte_menu.add_command(label="Salir", command=self.salir_aplicacion)

        # Menú Procesos
        procesos_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Procesos", menu=procesos_menu)
        procesos_menu.add_command(label="Proceso 1", command=lambda: self.cambiar_proceso("Proceso 1"))
        procesos_menu.add_command(label="Proceso 2", command=lambda: self.cambiar_proceso("Proceso 2"))

        # Menú Ayuda
        ayuda_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Ayuda", menu=ayuda_menu)
        ayuda_menu.add_command(label="Acerca de", command=self.infoAdicional)
        ayuda_menu.add_command(label="Licencia", command=self.avisoLicencia)

    def crear_frames_proceso1(self):
        # Frame contenedor para Proceso 1
        self.proceso1_frame = tk.Frame(self.main_frame, bg="#fdf3e7")
        
        # Título del proceso
        self.titulo_proceso = tk.Label(
            self.proceso1_frame,
            text="PROCESO 1",
            font=("Arial", 14, "bold"),
            fg="#FF8C00",  # Color naranja oscuro
            pady=20,
            bg="#fdf3e7"  # Color de fondo del label
        )
        self.titulo_proceso.pack()
        
        # Frame para archivos
        files_frame = tk.Frame(self.proceso1_frame, bg="#fdf3e7")
        files_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Label y botón para seleccionar archivos (usando grid)
        self.label = tk.Label(files_frame, text="Seleccione archivos ZIP para procesar:", anchor="w", bg="#fdf3e7")
        self.label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.select_button = tk.Button(files_frame, text="Seleccionar Archivos", command=self.select_files, width=24, height=2, anchor="center", bg="#f2c19c")
        self.select_button.grid(row=0, column=1, sticky="e")
        
        # Frame para ruta de destino
        dest_frame = tk.Frame(self.proceso1_frame, bg="#fdf3e7")
        dest_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Label y botón para ruta de destino (usando grid)
        self.dest_label = tk.Label(dest_frame, text="Seleccione la carpeta de destino:", anchor="w", bg="#fdf3e7")
        self.dest_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.dest_button = tk.Button(dest_frame, text="Seleccionar Destino", command=self.select_destination, width=24, height=2, anchor="center", bg="#f2c19c")
        self.dest_button.grid(row=0, column=1, sticky="e")
        
        # Configurar el peso de las columnas para alineación
        files_frame.grid_columnconfigure(0, weight=1)
        files_frame.grid_columnconfigure(1, weight=0)
        dest_frame.grid_columnconfigure(0, weight=1)
        dest_frame.grid_columnconfigure(1, weight=0)
        
        # Frame para etiquetas de estado
        status_frame = tk.Frame(self.proceso1_frame, bg="#fdf3e7")
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Etiquetas de estado
        self.files_status_label = tk.Label(status_frame, text="", fg="gray", bg="#fdf3e7")
        self.files_status_label.pack(side=tk.LEFT)
        
        self.dest_status_label = tk.Label(status_frame, text="", fg="gray", bg="#fdf3e7")
        self.dest_status_label.pack(side=tk.RIGHT)
        
        # Crear el botón de procesar con estilo
        self.process_button = tk.Button(
            self.proceso1_frame, 
            text="Procesar Archivos", 
            command=self.process_files,
            state='disabled',  # Inicialmente deshabilitado
            bg='#FFA500',     # Color naranja
            fg='white',       # Texto blanco
            activebackground='#FF8C00',  # Color naranja más oscuro al hacer clic
            activeforeground='white',
            width=24,
            height=3
        )
        self.process_button.pack(pady=10)

    def mostrar_proceso1(self):
        self.mensaje_inicial_frame.pack_forget()
        self.proceso1_frame.pack(fill=tk.BOTH, expand=True)

    def ocultar_proceso1(self):
        self.proceso1_frame.pack_forget()
        self.mensaje_inicial_frame.pack(fill=tk.BOTH, expand=True)

    def nuevo_reporte(self):
        self.resetear_archivos()
        self.resetear_archivos2()
        self.proceso_actual = ""
        
        # Ocultar ambos procesos y mostrar el mensaje inicial
        self.ocultar_proceso1()
        self.ocultar_proceso2()
        self.mensaje_inicial_frame.pack(fill=tk.BOTH, expand=True)
        
        # Limpiar títulos
        self.titulo_proceso.config(text="")
        self.titulo_proceso2.config(text="")
        
        messagebox.showinfo("Nuevo Reporte", "Se ha creado un nuevo reporte.")

    def guardar_reporte(self):
        messagebox.showinfo("Guardar", "Función de guardar (en desarrollo)")

    def salir_aplicacion(self):
        if messagebox.askokcancel("Salir", "¿Estás seguro que deseas salir de la aplicación?"):
            self.root.destroy()

    def infoAdicional(self):
        messagebox.showinfo("RPA FINANCE", "RPA desarrollada para la generación de reportes en formato xlsx a partir del tratamiento de datos obtenidos de la extracción automática de un archivo ZIP.\n\nDesarrolladores: \nYaco David Acuña Espinoza \nDayana Miranda Del Castillo")

    def avisoLicencia(self):
        messagebox.showwarning("Licencia", "Este software es propiedad intelectual de SGS PERÚ. No se permite su uso sin autorización.\nLicencia vitalicia.")

    def resetear_archivos(self):
        self.selected_files = []
        self.files_status_label.config(text="")
        self.dest_status_label.config(text="")
        self.rpa_finance_path = ""
        self.process_button.config(state='disabled')

    def cambiar_proceso(self, proceso):
        self.proceso_actual = proceso
        self.resetear_archivos()
        
        if proceso == "Proceso 1":
            self.mostrar_proceso1()
            self.titulo_proceso.config(text="PROCESO 1")
        else:
            self.ocultar_proceso1()
            self.titulo_proceso.config(text="PROCESO 2")
            
        messagebox.showinfo("Cambio de Proceso", f"Se ha cambiado al {proceso}")
        
    def select_destination(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if folder:
            self.rpa_finance_path = folder
            self.dest_status_label.config(text=f"Destino: {os.path.basename(folder)}")
            self.update_process_button()

    def update_process_button(self):
        # Habilitar el botón solo si hay archivos seleccionados y una ruta de destino
        if self.selected_files and self.rpa_finance_path:
            self.process_button.config(state='normal')
        else:
            self.process_button.config(state='disabled')

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=[("Archivos comprimidos", "*.zip")]
        )
        self.selected_files = files
        if files:
            self.files_status_label.config(text=f"Archivos seleccionados: {len(files)}")
        else:
            self.files_status_label.config(text="")
        self.update_process_button()

    def process_files(self):
        if not self.selected_files:
            messagebox.showwarning("Advertencia", "Por favor seleccione archivos primero")
            return
            
        for file_path in self.selected_files:
            try:
                # Extraer archivos
                if file_path.lower().endswith('.zip'):
                    with zipfile.ZipFile(file_path, 'r') as zip_ref:
                        temp_dir = "temp_extract"
                        zip_ref.extractall(temp_dir)
                
                # Procesar archivos TXT
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file.lower().endswith('.txt'):
                            txt_path = os.path.join(root, file)
                            with open(txt_path, 'r', encoding='utf-8') as f:
                                headers = f.readline().strip().split(';')
                                data = f.readlines()

                            # Procesamos cada línea de datos
                            rows = []
                            for line in data:
                                values = line.strip().split(';')
                                # Convertir fechas en las columnas especificadas
                                for i, header in enumerate(headers):
                                    if header in ["Creation Date", "PO Date", "Received Date"]:
                                        values[i] = convertir_fechas(values[i])
                                rows.append(values)

                            # Procesar archivo TXT
                            df = pd.DataFrame(rows, columns=headers)
                            
                            # Generar nombre del archivo Excel
                            current_time = datetime.now().strftime("%H%M")
                            excel_name = f"{os.path.splitext(file)[0]}_{current_time}.xlsx"
                            excel_path = os.path.join(self.rpa_finance_path, excel_name)
                            
                            # Guardar como Excel con formato tabla
                            writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
                            df.to_excel(writer, sheet_name='Datos', index=False)
                            
                            # Dar formato de tabla
                            workbook = writer.book
                            worksheet = writer.sheets['Datos']
                            tabla = worksheet.add_table(0, 0, len(df), len(df.columns)-1, 
                                                     {'style': 'Table Style Medium 2','columns': [{'header': col} for col in headers]})
                            writer.close()
                            
                            # Copiar archivo TXT a carpeta RPA FINANCE
                            shutil.copy2(txt_path, self.rpa_finance_path)
                
                # Limpiar directorio temporal
                shutil.rmtree(temp_dir, ignore_errors=True)
                
            except Exception as e:
                messagebox.showerror("Error", f"Error procesando {file_path}: {str(e)}")
                continue
        
        messagebox.showinfo("Éxito", "Archivos procesados correctamente")
        self.resetear_archivos()

    def crear_frames_proceso2(self):
        # Frame contenedor para Proceso 2
        self.proceso2_frame = tk.Frame(self.main_frame, bg="#fdf3e7")
        
        # Título del proceso
        self.titulo_proceso2 = tk.Label(
            self.proceso2_frame,
            text="PROCESO 2",
            font=("Arial", 14, "bold"),
            fg="#FF8C00",  # Color naranja oscuro
            pady=20,
            bg="#fdf3e7"  # Color de fondo del label
        )
        self.titulo_proceso2.pack()
        
        # Frame para archivos Excel
        excel_frame = tk.Frame(self.proceso2_frame, bg="#fdf3e7")
        excel_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Label y botón para seleccionar archivos Excel
        self.excel_label = tk.Label(excel_frame, text="Seleccione los dos archivos Excel a procesar:", anchor="w", bg="#fdf3e7")
        self.excel_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.excel_button = tk.Button(excel_frame, text="Seleccionar Archivos Excel", command=self.select_excel_files, width=24, height=2, anchor="center", bg="#f2c19c")
        self.excel_button.grid(row=0, column=1, sticky="e")
        
        # Frame para ruta de destino
        dest_frame = tk.Frame(self.proceso2_frame, bg="#fdf3e7")
        dest_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Label y botón para ruta de destino
        self.dest_label2 = tk.Label(dest_frame, text="Seleccione la carpeta de destino:", anchor="w", bg="#fdf3e7")
        self.dest_label2.grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.dest_button2 = tk.Button(dest_frame, text="Seleccionar Destino", command=self.select_destination2, width=24, height=2, anchor="center", bg="#f2c19c")
        self.dest_button2.grid(row=0, column=1, sticky="e")
        
        # Configurar el peso de las columnas para alineación
        excel_frame.grid_columnconfigure(0, weight=1)
        excel_frame.grid_columnconfigure(1, weight=0)
        dest_frame.grid_columnconfigure(0, weight=1)
        dest_frame.grid_columnconfigure(1, weight=0)
        
        # Frame para etiquetas de estado
        status_frame = tk.Frame(self.proceso2_frame, bg="#fdf3e7")
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Etiquetas de estado
        self.excel_status_label = tk.Label(status_frame, text="", fg="gray", bg="#fdf3e7")
        self.excel_status_label.pack(side=tk.LEFT)
        
        self.dest_status_label2 = tk.Label(status_frame, text="", fg="gray", bg="#fdf3e7")
        self.dest_status_label2.pack(side=tk.RIGHT)
        
        # Crear el botón de procesar con estilo
        self.process_button2 = tk.Button(
            self.proceso2_frame, 
            text="Procesar Archivos", 
            command=self.process_excel_files,
            state='disabled',  # Inicialmente deshabilitado
            bg='#FFA500',     # Color naranja
            fg='white',       # Texto blanco
            activebackground='#FF8C00',  # Color naranja más oscuro al hacer clic
            activeforeground='white',
            width=24,
            height=3
        )
        self.process_button2.pack(pady=10)
        
        # Variables para el Proceso 2
        self.excel_files = []
        self.rpa_finance_path2 = ""

    def mostrar_proceso2(self):
        self.mensaje_inicial_frame.pack_forget()
        self.proceso2_frame.pack(fill=tk.BOTH, expand=True)

    def ocultar_proceso2(self):
        self.proceso2_frame.pack_forget()
        self.mensaje_inicial_frame.pack(fill=tk.BOTH, expand=True)

    def select_excel_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if len(files) > 2:
            messagebox.showwarning("Advertencia", "Solo debe seleccionar dos archivos Excel")
            return
            
        self.excel_files = files
        if files:
            self.excel_status_label.config(text=f"Archivos seleccionados: {len(files)} de 2")
        else:
            self.excel_status_label.config(text="")
        self.update_process_button2()

    def select_destination2(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if folder:
            self.rpa_finance_path2 = folder
            self.dest_status_label2.config(text=f"Destino: {os.path.basename(folder)}")
            self.update_process_button2()

    def update_process_button2(self):
        if len(self.excel_files) == 2 and self.rpa_finance_path2:
            self.process_button2.config(state='normal')
        else:
            self.process_button2.config(state='disabled')

    def process_excel_files(self):
        if len(self.excel_files) != 2:
            messagebox.showwarning("Advertencia", "Debe seleccionar exactamente dos archivos Excel")
            return

        try:
            for file in self.excel_files:
                if not (file.lower().endswith('.xlsx') or file.lower().endswith('.xls')):
                    raise ValueError(f"El archivo {os.path.basename(file)} no es un archivo Excel válido")

            df1 = pd.read_excel(self.excel_files[0])
            df2 = pd.read_excel(self.excel_files[1])

            # Convertir fechas en la columna "Need-By"
            for df in [df1, df2]:
                if 'Need-By' in df.columns:
                    df['Need-By'] = df['Need-By'].apply(convertir_fechas)

            required_columns = ['Number', 'Line', 'Item']
            for col in required_columns:
                if col not in df1.columns or col not in df2.columns:
                    raise ValueError(f"Falta la columna '{col}' en uno de los archivos Excel")

            all_headers = list(set(df1.columns) | set(df2.columns))
            merged_df = pd.DataFrame(columns=all_headers)
            merged_df = pd.merge(df1, df2, on=required_columns, how='outer', suffixes=('_1', '_2'))

            for col in merged_df.columns:
                if col.endswith('_1') or col.endswith('_2'):
                    base_col = col[:-2]
                    if base_col + '_1' in merged_df.columns and base_col + '_2' in merged_df.columns:
                        merged_df[base_col] = merged_df[base_col + '_1'].fillna(merged_df[base_col + '_2'])
                        merged_df = merged_df.drop([base_col + '_1', base_col + '_2'], axis=1)

            if 'Number' in merged_df.columns:
                other_columns = [col for col in merged_df.columns if col != 'Number']
                merged_df = merged_df[['Number'] + other_columns]

            # Dividir 'Charge Account' en columnas A-L en su misma posición
            if 'Charge Account' in merged_df.columns:
                charge_idx = merged_df.columns.get_loc('Charge Account')
                charge_split = merged_df['Charge Account'].astype(str).str.split('.', n=11, expand=True)
                col_names = ['Code', 'Cuenta contable', 'Sector', 'Activity', 'Centro de costo', 'Level', 'Localidad', 'C1', 'C2', 'C3', 'C4', 'C5']
                charge_split.columns = col_names
                merged_df = merged_df.drop(columns=['Charge Account'])
                before = merged_df.iloc[:, :charge_idx]
                after = merged_df.iloc[:, charge_idx:]
                merged_df = pd.concat([before, charge_split, after], axis=1)

            current_time = datetime.now().strftime("%d%m%y%H%M")
            excel_name = f"CONCATENADO_{current_time}.xlsx"
            excel_path = os.path.join(self.rpa_finance_path2, excel_name)

            try:
                merged_df.to_excel(excel_path, sheet_name='Datos', index=False)
                workbook = openpyxl.load_workbook(excel_path)
                worksheet = workbook['Datos']
                tab = openpyxl.worksheet.table.Table(displayName="Tabla1", ref=f"A1:{num_to_excel_col(len(merged_df.columns)-1)}{len(merged_df) + 1}")
                style = openpyxl.worksheet.table.TableStyleInfo(name="TableStyleMedium2", showFirstColumn=True, showLastColumn=True, showRowStripes=True, showColumnStripes=True)
                tab.tableStyleInfo = style
                worksheet.add_table(tab)

                for idx, col in enumerate(merged_df.columns):
                    max_length = max(
                        merged_df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    col_letter = num_to_excel_col(idx)
                    worksheet.column_dimensions[col_letter].width = max_length + 2

                workbook.save(excel_path)
                messagebox.showinfo("Éxito", "Archivos Excel procesados correctamente")
                self.resetear_archivos2()
            except Exception as e:
                raise ValueError(f"Error al guardar el archivo Excel: {str(e)}")

        except ValueError as ve:
            messagebox.showerror("Error de Validación", str(ve))
        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado procesando archivos Excel: {str(e)}")

    def resetear_archivos2(self):
        self.excel_files = []
        self.excel_status_label.config(text="")
        self.dest_status_label2.config(text="")
        self.rpa_finance_path2 = ""
        self.process_button2.config(state='disabled')

    def cambiar_proceso(self, proceso):
        self.proceso_actual = proceso
        self.resetear_archivos()
        self.resetear_archivos2()
        
        if proceso == "Proceso 1":
            self.ocultar_proceso2()
            self.mostrar_proceso1()
            self.titulo_proceso.config(text="PROCESO 1")
        else:
            self.ocultar_proceso1()
            self.mostrar_proceso2()
            self.titulo_proceso2.config(text="PROCESO 2")
            
        messagebox.showinfo("Cambio de Proceso", f"Se ha cambiado al {proceso}")

if __name__ == "__main__":
    root = tk.Tk()

    app = RPAFinanceApp(root)
    root.mainloop()