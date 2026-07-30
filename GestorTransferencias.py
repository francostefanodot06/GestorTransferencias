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
            df_bank_original, leyenda_col, credit_col = self.read_bank_file(bank_file)
            processed_data, no_encontrados = self.process_invoices(invoices, df_bank_original)

            conciliacion_df, updated_bank_df = self.prepare_conciliacion(processed_data, df_bank_original)
            no_encontrados_df = pd.DataFrame(no_encontrados, columns=['Archivo', 'Texto Extraído'])

            output_filename = f"Planilla_Creada_{os.path.basename(self.folder_path)}.xlsx"
            output_file_path = os.path.join(self.folder_path, output_filename)

            with pd.ExcelWriter(output_file_path) as writer:
                conciliacion_df.to_excel(writer, sheet_name='Conciliados', index=False)
                no_encontrados_df.to_excel(writer, sheet_name='No_Encontrados', index=False)

            # Actualizar el archivo del banco original
            updated_bank_df.to_csv(bank_file, encoding='utf-8-sig', sep=';', index=False)

            messagebox.showinfo("Proceso Completado", f"El proceso de conciliación ha finalizado con éxito. Archivo generado: {output_filename}")
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

        # Detectar dinámicamente las columnas 'Leyenda Adicional1' y 'Créditos/Monto'
        leyenda_col = None
        credit_col = None

        for col in df.columns:
            if 'leyenda adicional' in col.lower():
                leyenda_col = col
                break

        if not leyenda_col:
            raise ValueError("No se pudo identificar automáticamente la columna 'Leyenda Adicional1'. Columnas encontradas: {list(df.columns)}")

        for col in df.columns:
            if any(keyword in col.lower() for keyword in ['credito', 'monto', 'haber']):
                credit_col = col
                break

        if not credit_col:
            raise ValueError("No se pudo identificar automáticamente la columna de montos ('crédito', 'monto' o 'haber'). Columnas encontradas: {list(df.columns)}")

        return df, leyenda_col, credit_col

    def process_invoices(self, invoices, df_bank):
        processed_data = {}
        no_encontrados = []
        cobradores_lista = df_bank['Leyenda Adicional1'].dropna().tolist()

        for invoice in invoices:
            text = self.extract_text(invoice)
            
            # Buscar coincidencia por nombre
            cobrador_match = process.extractOne(text, cobradores_lista)
            if not cobrador_match or cobrador_match[1] < 75:
                no_encontrados.append((invoice, text))
                continue

            cobrador_name = cobrador_match[0]
            matching_rows = df_bank.loc[df_bank['Leyenda Adicional1'] == cobrador_name]

            if not matching_rows.empty:
                credit_amounts = matching_rows['Creditos'].values
                credit_text = self.extract_monto(text)

                if any(str(amount) in credit_text for amount in credit_amounts):
                    if cobrador_name not in processed_data:
                        processed_data[cobrador_name] = {'invoices': [], 'bank': []}

                    processed_data[cobrador_name]['invoices'].append({'file_path': invoice, 'text': text})
                    processed_data[cobrador_name]['bank'].extend(credit_amounts)
                else:
                    no_encontrados.append((invoice, text))
            else:
                no_encontrados.append((invoice, text))

        return processed_data, no_encontrados

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

    def extract_monto(self, text):
        # Ajustar según el formato del monto en los comprobantes
        return re.sub(r'\D', '', text)  # Extrae solo números

    def prepare_conciliacion(self, processed_data, df_bank_original):
        conciliacion_entries = []
        for cobrador, data in processed_data.items():
            for invoice in data['invoices']:
                conciliacion_entries.append({
                    'Cobrador': cobrador,
                    'Archivo': invoice['file_path'],
                    'Texto Extraído': invoice['text']
                })

        conciliacion_df = pd.DataFrame(conciliacion_entries)

        # Actualizar el archivo del banco
        updated_bank_df = df_bank_original.copy()
        for cobrador, data in processed_data.items():
            mask = updated_bank_df['Leyenda Adicional1'] == cobrador
            updated_bank_df = updated_bank_df.loc[~mask]

        return conciliacion_df, updated_bank_df

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()