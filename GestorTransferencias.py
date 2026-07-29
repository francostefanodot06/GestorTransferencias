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
        if not bank_file or not invoices:
            messagebox.showwarning("Advertencia", "No se encontraron archivos bancarios o comprobantes en la carpeta.")
            return

        try:
            df_bank = self.read_bank_file(bank_file)
            processed_data = self.process_invoices(invoices, df_bank)

            for cobrador, data in processed_data.items():
                self.save_conciliacion(data, cobrador, bank_file)

            messagebox.showinfo("Proceso Completado", "El proceso de conciliación ha finalizado con éxito.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n{e}")

    def find_files(self, folder_path):
        bank_file = None
        invoices = []

        for root, _, files in os.walk(folder_path):
            for file in files:
                file_lower = file.lower()
                if file_lower.endswith(('.csv', '.xlsx', '.txt')) and not bank_file:
                    bank_file = os.path.join(root, file)
                elif file_lower.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    invoices.append(os.path.join(root, file))

        return bank_file, invoices

    def read_bank_file(self, bank_file):
        if bank_file.endswith('.csv'):
            try:
                df = pd.read_csv(bank_file, encoding='utf-8-sig', sep=';', on_bad_lines='skip')
            except Exception:
                df = pd.read_csv(bank_file, encoding='latin1', sep=';', on_bad_lines='skip')
        elif bank_file.endswith('.xlsx'):
            df = pd.read_excel(bank_file)

        df.columns = df.columns.astype(str).str.strip().str.replace('ï»¿', '', regex=True)
        
        col_mapping = {}
        for col in df.columns:
            col_lower = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            if 'fecha' in col_lower:
                col_mapping['Fecha'] = col
            elif 'credito' in col_lower or 'monto' in col_lower or 'haber' in col_lower:
                col_mapping['Creditos'] = col
            elif 'leyenda' in col_lower or 'descripcion' in col_lower or 'detalle' in col_lower or 'concepto' in col_lower:
                col_mapping['Leyenda Adicional1'] = col

        missing = [k for k in ['Fecha', 'Creditos', 'Leyenda Adicional1'] if k not in col_mapping]
        if missing:
            raise ValueError(f"No se pudieron identificar automáticamente las columnas: {missing}. Columnas encontradas: {list(df.columns)}")

        df = df.rename(columns={v: k for k, v in col_mapping.items()})
        relevant_columns = ['Fecha', 'Creditos', 'Leyenda Adicional1']
        return df[relevant_columns]

    def process_invoices(self, invoices, df_bank):
        processed_data = {}

        for invoice in invoices:
            text = self.extract_text(invoice)
            cobrador_match = process.extractOne(text, df_bank['Leyenda Adicional1'].dropna())
            
            if cobrador_match and cobrador_match[1] >= 70:
                cobrador_name = cobrador_match[0]
                
                # Búsqueda segura usando .loc para evitar cualquier problema de índices desalineados
                matching_rows = df_bank.loc[df_bank['Leyenda Adicional1'] == cobrador_name]

                if not matching_rows.empty:
                    credit_amount = matching_rows['Creditos'].values[0]

                    if cobrador_name not in processed_data:
                        processed_data[cobrador_name] = {'invoices': [], 'bank': []}

                    processed_data[cobrador_name]['invoices'].append({'file_path': invoice, 'text': text})
                    processed_data[cobrador_name]['bank'].append(credit_amount)

        return processed_data

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

    def save_conciliacion(self, data, cobrador, bank_file):
        folder_name = f"Rendicion_{os.path.basename(self.folder_path)}_{cobrador}"
        output_folder = os.path.join(self.folder_path, folder_name)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        for invoice in data['invoices']:
            src_file = invoice['file_path']
            dst_file = os.path.join(output_folder, os.path.basename(src_file))
            shutil.copy2(src_file, dst_file)

        if bank_file:
            if bank_file.endswith('.csv'):
                df_bank = pd.read_csv(bank_file, encoding='utf-8-sig', sep=';', on_bad_lines='skip')
            else:
                df_bank = pd.read_excel(bank_file)

            original_cols = df_bank.columns
            df_bank.columns = df_bank.columns.astype(str).str.strip().str.replace('ï»¿', '', regex=True)
            
            cred_col_real = None
            ley_col_real = None
            for col in df_bank.columns:
                c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
                if 'credito' in c_low or 'monto' in c_low or 'haber' in c_low:
                    cred_col_real = col
                if 'leyenda' in c_low or 'descripcion' in c_low or 'detalle' in c_low or 'concepto' in c_low:
                    ley_col_real = col

            if cred_col_real and ley_col_real:
                for credit in data['bank']:
                    # Eliminación limpia por índice booleano directo para evitar cualquier error de alineación
                    mask = (df_bank[cred_col_real] == credit) & (df_bank[ley_col_real] == cobrador)
                    df_bank = df_bank.loc[~mask]

            df_bank.columns = original_cols

            if bank_file.endswith('.csv'):
                df_bank.to_csv(bank_file, index=False, encoding='utf-8-sig', sep=';')
            else:
                df_bank.to_excel(bank_file, index=False)

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()