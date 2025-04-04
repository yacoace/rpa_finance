import pysftp
import os
import zipfile
import pandas as pd
import re
import tkinter as tk
from tkinter import messagebox
import configparser
 
config = configparser.ConfigParser()
config.read('config.ini')
 
try:
    local_path = config.get('Paths', 'local_path')
    cruce_cuenta_path = config.get('Paths', 'cruce_cuenta_path')
    cruce_cecos_path = config.get('Paths', 'cruce_cecos_path')
 
 
    os.makedirs(local_path, exist_ok=True)
except Exception as e:
    print(f"❌ Error al leer el archivo de configuración: {e}")
    messagebox.showerror("Error", f"Error al leer el archivo de configuración: {e}")
 
 
host = "boss-sftp.sgs.net"
port = 22
username = "ifsgs_pe"
password = "Ks*Pri84YK63B_ks"
 
def extract_zip(file_path, extract_to):
    extracted_files = []
    if file_path.endswith(".zip"):
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
            extracted_files = zip_ref.namelist()
        print(f"✅ Archivo ZIP extraído en: {extract_to}")
    return extracted_files
 
def clean_text(value):
    if isinstance(value, str):
        return re.sub(r'[^\w\s.,()-]', '', value)
    return value
 
def convert_txt_to_excel(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            first_line = f.readline()
            delimiter = ',' if ',' in first_line else '\t'
        df = pd.read_csv(file_path, sep=delimiter, engine='python', quoting=3, index_col=False)
        df = df.applymap(clean_text)
        excel_path = file_path.replace(".txt", ".xlsx")
        df.to_excel(excel_path, index=False)
        print(f"✅ Convertido a Excel: {excel_path}")
        return excel_path
    except Exception as e:
        print(f"❌ Error convirtiendo a Excel: {e}")
        return None
 
def add_tipo_column(df):
    df['TIPO'] = df['FRP_ACCOUNT'].astype(str).apply(lambda x: 'INGRESO' if x.startswith('4') else ('GASTO' if x.startswith('5') else 'OTRO'))
    return df
 
def add_monto_column(df):
    try:
        df['Monto'] = df['ACCOUNTED_BALANCE'].astype(float) * -1
        print("✅ Columna 'Monto' generada correctamente.")
    except Exception as e:
        print(f"❌ Error al generar la columna 'Monto': {e}")
    return df
 
def remove_duplicates(df):
    before = len(df)
    df = df.drop_duplicates()
    after = len(df)
    print(f"✅ Filas duplicadas eliminadas: {before - after}")
    return df
 
def run_script():
    try:
        cnopts = pysftp.CnOpts()
        cnopts.hostkeys = None
        with pysftp.Connection(host, username=username, password=password, port=port, cnopts=cnopts) as sftp:
            print("✅ Conectado al SFTP.")
            archivos = sftp.listdir_attr("/EBSPRD/outgoing/")
            zip_files = [file for file in archivos if file.filename.endswith(".zip") and "DRILL" in file.filename]
            if zip_files:
                latest_zip = max(zip_files, key=lambda x: x.st_mtime)
                zip_name = latest_zip.filename
                remote_zip_path = f"/EBSPRD/outgoing/{zip_name}"
                local_zip_path = os.path.join(local_path, zip_name)
                print(f"⬇ Descargando {zip_name}...")
                sftp.get(remote_zip_path, local_zip_path)
                print(f"✅ Archivo guardado en: {local_zip_path}")
                extracted_files = extract_zip(local_zip_path, local_path)
                for file in extracted_files:
                    extracted_file_path = os.path.join(local_path, file)
                    if extracted_file_path.endswith(".txt"):
                        excel_path = convert_txt_to_excel(extracted_file_path)
                        if excel_path:
                            df = pd.read_excel(excel_path)
                            cruce_cuenta = pd.read_excel(cruce_cuenta_path)
                            cruce_cecos = pd.read_excel(cruce_cecos_path)
                            df = df.merge(cruce_cuenta[['ACCOUNT', 'Carga Symphony']], on='ACCOUNT', how='left')
                            df = df.merge(cruce_cecos[['COST CENTER', 'SUBNEGOCIO']], on='COST CENTER', how='left')
                            df = add_tipo_column(df)
                            df = add_monto_column(df)
                            df = remove_duplicates(df)
                            final_path = excel_path.replace(".xlsx", "_final.xlsx")
                            df.to_excel(final_path, index=False)
                            print(f"✅ Archivo final generado: {final_path}")
                messagebox.showinfo("Éxito", "Script ejecutado correctamente.")
            else:
                messagebox.showwarning("Advertencia", "No se encontraron archivos ZIP con 'DRILL'.")
    except Exception as e:
        messagebox.showerror("Error", f"❌ Error: {e}")
 
# Interfaz gráfica (Tkinter)
def run_gui():
    root = tk.Tk()
    root.title("Ejecutar Script SFTP")
    root.geometry("400x200")
    root.configure(bg="#2E4053")
 
    label = tk.Label(root, text="Bienvenido, James.", font=("Arial", 20, "bold"), fg="white", bg="#2E4053")
    label.pack(pady=20)
 
    execute_button = tk.Button(root, text="Ejecutar", font=("Arial", 14), command=run_script, bg="#1ABC9C", fg="white", width=15)
    execute_button.pack(pady=5)
 
    exit_button = tk.Button(root, text="Salir", font=("Arial", 14), command=root.quit, bg="#E74C3C", fg="white", width=15)
    exit_button.pack(pady=5)
 
    root.mainloop()
 
if __name__ == "__main__":
    run_gui()