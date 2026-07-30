import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image
import pytesseract
from fuzzywuzzy import process
from pdf2image import convert_from_path

# Configurar Tesseract (ajusta la ruta según tu sistema)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ComprobanteConciliacionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conciliación de Comprobantes")
        self.geometry("400x300")

        self.folder_path = None
        self.create_widgets()

    def create_widgets(self):
        self.select_folder_button = tk.Button(self, text="Seleccionar Carpeta", command=self.select_folder)
        self.select_folder_button.pack(pady=20)

        self.process_button = tk.Button(self, text="Procesar", command=self.process_files)
        self.process_button.pack(pady=10)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory(title="Selecciona la carpeta del día")
        if self.folder_path:
            messagebox.showinfo("Carpeta Seleccionada", f"Carpeta seleccionada: {self.folder_path}")

    def process_files(self):
        if not self.folder_path:
            messagebox.showwarning("Advertencia", "Debes seleccionar una carpeta primero.")
            return

        bank_file, invoices = self.find_files(self.folder_path)
        
        if not bank_file:
            messagebox.showwarning("Advertencia", "No se encontró ningún archivo de banco (.csv o .xlsx) en la carpeta.")
            return
            
        if not invoices:
            messagebox.showwarning("Advertencia", "No se encontraron comprobantes (.png, .jpg, .pdf) en la carpeta.")
            return

        try:
            df_bank, bank_cols, cred_col_real, ley_col_real = self.read_bank_file_full(bank_file)
            
            matched_records = []
            
            # Forzamos una lista limpia de textos de la columna de conceptos/leyendas del banco
            cobradores_lista = df_bank[ley_col_real].dropna().astype(str).tolist()

            print(f"\n[DEBUG] Columna de búsqueda usada: {ley_col_real}")
            print(f"[DEBUG] Total de registros en el banco: {len(cobradores_lista)}")

            for invoice in invoices:
                text = self.extract_text(invoice)
                print(f"\nProcesando: {os.path.basename(invoice)}")
                print(f"Texto leído: {text[:100]}...")

                # Buscamos la mejor coincidencia
                match = process.extractOne(text, cobradores_lista)
                print(f"Mejor coincidencia: {match}")

                # Si encuentra un match razonable (umbral 30 para texto OCR imperfecto)
                if match and match[1] >= 30:
                    client_legend = match[0]
                    # Obtenemos la fila exacta que coincide con esa leyenda
                    row_data = df_bank[df_bank[ley_col_real] == client_legend]

                    if not row_data.empty:
                        matched_records.append(row_data.iloc[0])

            if not matched_records:
                messagebox.showinfo("Sin coincidencias", "Ningún comprobante hizo match con las leyendas del banco.")
                return

            # Creamos el DataFrame con los registros encontrados
            df_matched = pd.DataFrame(matched_records)
            total_autosuma = df_matched[cred_col_real].sum()

            # Guardamos la planilla nueva con la autosuma
            output_excel_name = f"Planilla_Creada_{os.path.basename(self.folder_path)}.xlsx"
            output_excel_path = os.path.join(self.folder_path, output_excel_name)

            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df_matched.to_excel(writer, sheet_name='Conciliacion', index=False)
                summary_df = pd.DataFrame([[ 'TOTAL AUTOSUMA', total_autosuma ]], columns=[ley_col_real, cred_col_real])
                summary_df.to_excel(writer, sheet_name='Resumen', index=False)

            # Eliminamos del DataFrame original del banco las filas que ya se conciliaron
            for client_legend in df_matched[ley_col_real]:
                mask = (df_bank[ley_col_real] == client_legend)
                df_bank = df_bank.loc[~mask]

            # Restauramos las columnas originales
            df_bank.columns = bank_cols

            # Guardamos el archivo del banco modificado (esto cambia su fecha de modificación)
            if bank_file.endswith('.csv'):
                df_bank.to_csv(bank_file, index=False, encoding='utf-8-sig', sep=';')
            else:
                df_bank.to_excel(bank_file, index=False)

            messagebox.showinfo("Proceso Completado", f"¡Proceso exitoso!\nSe creó la planilla: {output_excel_name}\nSe modificó el archivo del banco original.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n{e}")

    def find_files(self, folder_path):
        bank_file = None
        invoices = []

        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_lower = file.lower()
                file_path = os.path.join(root, file)
                if file_lower.endswith(('.csv', '.xlsx', '.txt')) and not bank_file and not file.startswith('Planilla_Creada'):
                    bank_file = file_path
                elif file_lower.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    invoices.append(file_path)

        return bank_file, invoices

    def read_bank_file_full(self, bank_file):
        if bank_file.endswith('.csv'):
            try:
                df = pd.read_csv(bank_file, encoding='utf-8-sig', sep=';', on_bad_lines='skip')
            except Exception:
                df = pd.read_csv(bank_file, encoding='latin1', sep=';', on_bad_lines='skip')
        elif bank_file.endswith('.xlsx'):
            df = pd.read_excel(bank_file)

        original_cols = df.columns
        df.columns = df.columns.astype(str).str.strip().str.replace('ï»¿', '', regex=True)
        
        cred_col_real = None
        ley_col_real = None
        
        # Buscamos de manera inteligente las columnas correctas basándonos en nombres comunes
        for col in df.columns:
            c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            if any(term in c_low for term in ['credito', 'monto', 'haber', 'importe']):
                cred_col_real = col
            if any(term in c_low for term in ['leyenda', 'descripcion', 'detalle', 'concepto', 'referencia']):
                ley_col_real = col

        # Plan B si no encuentra por palabras clave exactas: agarramos la segunda columna para texto y la de créditos por posición o tipo numérico
        if not ley_col_real and len(df.columns) > 1:
            ley_col_real = df.columns[1] 
        if not cred_col_real and len(df.columns) > 2:
            cred_col_real = df.columns[2]

        if not cred_col_real or not ley_col_real:
            raise ValueError(f"No se pudieron identificar las columnas de montos o leyendas. Columnas encontradas: {list(df.columns)}")

        return df, original_cols, cred_col_real, ley_col_real

    def extract_text(self, file_path):
        extracted_text = ""
        try:
            if file_path.lower().endswith('.pdf'):
                pages = convert_from_path(file_path, poppler_path=r'C:\Program Files\poppler-23.07.0\Library\bin')
                for page in pages:
                    extracted_text += pytesseract.image_to_string(page)
            else:
                image = Image.open(file_path).convert('RGB')
                extracted_text = pytesseract.image_to_string(image)
        except Exception as e:
            print(f"Error al extraer texto de {file_path}: {e}")

        return re.sub(r'\s+', ' ', extracted_text.strip())

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()