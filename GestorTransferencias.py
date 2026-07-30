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
            df_bank, original_cols, cred_col_real, ley_col_real = self.read_bank_file_full(bank_file)
            processed_data = self.process_invoices(invoices, df_bank)

            # Guardamos la planilla multitab con conciliados y no encontrados
            self.save_conciliacion_multitab(processed_data, bank_file)

            messagebox.showinfo("Proceso Completado", f"¡Proceso exitoso!\nSe creó la planilla con solapas múltiples.\nSe modificó el archivo del banco original.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n{e}")

    def find_files(self, folder_path):
        bank_file = None
        invoices = []

        for root, _, files in os.walk(folder_path):
            for file in files:
                file_lower = file.lower()
                if file_lower.endswith(('.csv', '.xlsx', '.txt')) and not bank_file and not file.startswith('Planilla_Conciliacion'):
                    bank_file = os.path.join(root, file)
                elif file_lower.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    invoices.append(os.path.join(root, file))

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
        
        for col in df.columns:
            c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            if any(term in c_low for term in ['credito', 'monto', 'haber', 'importe']):
                cred_col_real = col

        for col in df.columns:
            if col.strip().lower() == 'leyenda adicional1':
                ley_col_real = col
                break

        if not ley_col_real:
            for col in df.columns:
                if 'leyenda' in col.lower() and 'adicional' in col.lower():
                    ley_col_real = col
                    break

        if not cred_col_real and len(df.columns) > 0:
            cred_col_real = df.columns[-1]

        if not ley_col_real:
            raise ValueError("No se encontró la columna 'Leyenda Adicional1' en el archivo del banco.")

        return df, original_cols, cred_col_real, ley_col_real

    def process_invoices(self, invoices, df_bank):
        processed_data = {'conciliados': [], 'no_encontrados': []}
        cobradores_lista = df_bank['Leyenda Adicional1'].dropna().tolist()

        for invoice in invoices:
            text = self.extract_text(invoice)
            # Umbral subido a 78 para evitar falsos positivos
            cobrador_match = process.extractOne(text, cobradores_lista)
            
            encontrado = False
            if cobrador_match and cobrador_match[1] >= 78:
                cobrador_name = cobrador_match[0]
                matching_rows = df_bank.loc[df_bank['Leyenda Adicional1'] == cobrador_name]

                if not matching_rows.empty:
                    credit_amount = matching_rows['Creditos'].values[0] if 'Creditos' in matching_rows.columns else matching_rows.iloc[0, 1]
                    fecha_banco = matching_rows['Fecha'].values[0] if 'Fecha' in matching_rows.columns else ''

                    processed_data['conciliados'].append({
                        'archivo': os.path.basename(invoice),
                        'cliente': cobrador_name,
                        'monto': credit_amount,
                        'fecha': fecha_banco,
                        'file_path': invoice
                    })
                    encontrado = True

            if not encontrado:
                processed_data['no_encontrados'].append({
                    'archivo': os.path.basename(invoice),
                    'texto_leido': text[:150] + "..." if len(text) > 150 else text,
                    'file_path': invoice
                })

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

    def save_conciliacion_multitab(self, processed_data, bank_file):
        output_excel = os.path.join(self.folder_path, f"Planilla_Conciliacion_{os.path.basename(self.folder_path)}.xlsx")
        
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # 1. Solapa Conciliados con Autosuma
            if processed_data['conciliados']:
                df_conc = pd.DataFrame(processed_data['conciliados'])
                # Quitamos la ruta interna para que la planilla quede limpia
                df_clean = df_conc.drop(columns=['file_path'], errors='ignore')
                
                total_monto = df_clean['monto'].sum() if 'monto' in df_clean.columns else 0
                df_totales = pd.DataFrame([{'archivo': 'TOTAL GENERAL', 'cliente': '', 'monto': total_monto, 'fecha': ''}])
                df_final_conc = pd.concat([df_clean, df_totales], ignore_index=True)
                df_final_conc.to_excel(writer, sheet_name='Conciliados', index=False)
            else:
                pd.DataFrame({'Mensaje': ['No hubo comprobantes conciliados']}).to_excel(writer, sheet_name='Conciliados', index=False)

            # 2. Solapa No Encontrados
            if processed_data['no_encontrados']:
                df_no_enc = pd.DataFrame(processed_data['no_encontrados'])
                df_no_enc_clean = df_no_enc.drop(columns=['file_path'], errors='ignore')
                df_no_enc_clean.to_excel(writer, sheet_name='No_Encontrados', index=False)
            else:
                pd.DataFrame({'Mensaje': ['Todos los comprobantes hicieron match con éxito']}).to_excel(writer, sheet_name='No_Encontrados', index=False)

        # Modificación del archivo del banco original (eliminando los ya conciliados)
        if bank_file and processed_data['conciliados']:
            if bank_file.endswith('.csv'):
                df_bank = pd.read_csv(bank_file, encoding='utf-8-sig', sep=';', on_bad_lines='skip')
            else:
                df_bank = pd.read_excel(bank_file)

            original_cols = df_bank.columns
            df_bank.columns = df_bank.columns.astype(str).str.strip().str.replace('ï»¿', '', regex=True)
            
            cred_col_real, ley_col_real = None, None
            for col in df_bank.columns:
                c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
                if any(term in c_low for term in ['credito', 'monto', 'haber', 'importe']):
                    cred_col_real = col
                if col.strip().lower() == 'leyenda adicional1' or ('leyenda' in c_low and 'adicional' in c_low):
                    ley_col_real = col

            if cred_col_real and ley_col_real:
                for item in processed_data['conciliados']:
                    mask = (df_bank[cred_col_real] == item['monto']) & (df_bank[ley_col_real] == item['cliente'])
                    df_bank = df_bank.loc[~mask]

            df_bank.columns = original_cols

            if bank_file.endswith('.csv'):
                df_bank.to_csv(bank_file, index=False, encoding='utf-8-sig', sep=';')
            else:
                df_bank.to_excel(bank_file, index=False)

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()