import os
import re
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image
import pytesseract
from pdf2image import convert_from_path

warnings.filterwarnings('ignore', category=UserWarning)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ComprobanteConciliacionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conciliación de Comprobantes - Urgente")
        self.geometry("450x280")
        self.configure(bg="#f0f0f0")

        self.folder_path = None
        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(self, text="Gestor de Conciliación Urgente", font=("Arial", 12, "bold"), bg="#f0f0f0")
        title_label.pack(pady=15)
        
        self.select_folder_button = tk.Button(self, text="1. Seleccionar Carpeta del Día", command=self.select_folder, bg="#e0e0e0", font=("Arial", 10), width=30, height=2)
        self.select_folder_button.pack(pady=5)

        self.process_button = tk.Button(self, text="2. Procesar y Generar Planilla", command=self.process_files, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=30, height=2)
        self.process_button.pack(pady=10)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory(title="Selecciona la carpeta con los comprobantes y el banco")
        if self.folder_path:
            messagebox.showinfo("Carpeta Seleccionada", f"Carpeta activa:\n{self.folder_path}")

    def process_files(self):
        if not self.folder_path:
            messagebox.showwarning("Atención", "Primero debés seleccionar una carpeta.")
            return

        bank_file, invoices = self.find_files(self.folder_path)

        if not bank_file:
            messagebox.showwarning("Atención", "No se encontró el archivo del banco (.xlsx o .csv).")
            return
        
        if not invoices:
            messagebox.showwarning("Atención", "No se encontraron comprobantes en la carpeta.")
            return

        try:
            df_bank, cred_col, ley_col = self.read_bank_file(bank_file)
            processed_data = self.process_invoices(invoices, df_bank, cred_col, ley_col)
            self.save_results(processed_data, bank_file, df_bank, cred_col, ley_col)

            messagebox.showinfo("¡Listo!", "Proceso finalizado. Revisá tu planilla generada.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def find_files(self, folder_path):
        bank_file = None
        invoices = []
        for root, _, files in os.walk(folder_path):
            for file in files:
                f_low = file.lower()
                if f_low.endswith(('.csv', '.xlsx', '.txt')) and not bank_file and not file.startswith('Planilla_'):
                    bank_file = os.path.join(root, file)
                elif f_low.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    invoices.append(os.path.join(root, file))
        return bank_file, invoices

    def read_bank_file(self, bank_file):
        if bank_file.endswith('.csv'):
            try:
                df = pd.read_csv(bank_file, encoding='utf-8-sig', sep=';', on_bad_lines='skip')
            except:
                df = pd.read_csv(bank_file, encoding='latin1', sep=';', on_bad_lines='skip')
        else:
            df = pd.read_excel(bank_file)

        df.columns = df.columns.astype(str).str.strip().str.replace('ï»¿', '', regex=True)
        
        cred_col, ley_col = None, None
        for col in df.columns:
            c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            if any(term in c_low for term in ['credito', 'monto', 'haber', 'importe', 'valor', 'saldo']):
                cred_col = col
                break

        for col in df.columns:
            c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            if any(term in c_low for term in ['leyenda', 'adicional', 'concepto', 'detalle', 'descripcion']):
                ley_col = col
                break

        if not cred_col and len(df.columns) > 3:
            cred_col = df.columns[3]
        if not ley_col and len(df.columns) > 11:
            ley_col = df.columns[11]
        elif not ley_col:
            ley_col = df.columns[-1]

        return df, cred_col, ley_col

    def process_invoices(self, invoices, df_bank, cred_col, ley_col):
        conciliados = []
        no_encontrados = []
        df_bank_work = df_bank.copy()

        # Limpiar los montos del banco a float puro para evitar errores de comparación
        df_bank_work['monto_limpio'] = df_bank_work[cred_col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df_bank_work['monto_limpio'] = pd.to_numeric(df_bank_work['monto_limpio'], errors='coerce').fillna(0.0)

        for invoice in invoices:
            invoice_name = os.path.basename(invoice)
            text = self.extract_text(invoice)
            
            # Buscar todos los números que parezcan montos (ej: 2.400,00 o 12000 o 2400)
            # Reemplazamos puntos de miles y comas decimales
            found_numbers = re.findall(r'\b\d{1,3}(?:\.\d{3})*(?:,\d+)?\b|\b\d+(?:,\d+)?\b|\b\d+\b', text)
            
            montos_comprobante = set()
            for num_str in found_numbers:
                clean_str = num_str.replace('.', '').replace(',', '.')
                try:
                    val = float(clean_str)
                    if val > 1.0: # Evitar números chicos sueltos
                        montos_comprobante.add(val)
                except:
                    pass

            print(f"--- Comprobante: {invoice_name} --- Montos detectados OCR: {montos_comprobante}")

            match_encontrado = False
            for idx, row in df_bank_work.iterrows():
                banco_val = row['monto_limpio']
                
                # Comparamos si alguno de los montos del comprobante coincide con el banco (con tolerancia de redondeo)
                if any(abs(m - banco_val) < 1.5 for m in montos_comprobante):
                    banco_nombre = str(row[ley_col]).strip() if pd.notna(row[ley_col]) else "Sin Detalle"
                    conciliados.append({
                        'archivo': invoice_name,
                        'cliente': banco_nombre,
                        'monto': row[cred_col],
                        'fecha': row['Fecha'] if 'Fecha' in df_bank_work.columns else ''
                    })
                    df_bank_work = df_bank_work.drop(idx)
                    match_encontrado = True
                    break

            if not match_encontrado:
                no_encontrados.append({
                    'archivo_faltante': invoice_name,
                    'detalle': f"No matcheó ningún monto de {invoice_name}"
                })

        return {'conciliados': conciliados, 'no_encontrados': no_encontrados, 'banco_restante': df_bank_work}

    def extract_text(self, file_path):
        text = ""
        try:
            if file_path.lower().endswith('.pdf'):
                pages = convert_from_path(file_path, poppler_path=r'C:\Program Files\poppler-23.07.0\Library\bin')
                for p in pages:
                    text += pytesseract.image_to_string(p)
            else:
                img = Image.open(file_path).convert('RGB')
                text = pytesseract.image_to_string(img)
        except Exception as e:
            print(f"Error OCR: {e}")
        return text

    def save_results(self, data, bank_file, df_bank, cred_col, ley_col):
        output_excel = os.path.join(self.folder_path, f"Planilla_Conciliacion_{os.path.basename(self.folder_path)}.xlsx")
        
        df_c = pd.DataFrame(data['conciliados'])
        df_n = pd.DataFrame(data['no_encontrados'])

        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_final = pd.DataFrame()
            if not df_c.empty:
                df_final = pd.concat([df_final, df_c], axis=1)
            else:
                df_final['Mensaje_Conciliados'] = ['Sin conciliados']

            df_final['---'] = ''

            if not df_n.empty:
                df_final = pd.concat([df_final, df_n], axis=1)
            else:
                df_final['Mensaje_No_Encontrados'] = ['Sin faltantes']

            df_final.to_excel(writer, sheet_name='Conciliacion_y_Faltantes', index=False)

        # Actualizar banco original eliminando los conciliados
        df_banco_orig = df_bank.copy()
        if data['conciliados']:
            for item in data['conciliados']:
                mask = (df_banco_orig[cred_col].astype(str) == str(item['monto']))
                matches = df_banco_orig[mask].index
                if not matches.empty:
                    df_banco_orig = df_banco_orig.drop(matches[0])

        if bank_file.endswith('.csv'):
            df_banco_orig.to_csv(bank_file, index=False, encoding='utf-8-sig', sep=';')
        else:
            df_banco_orig.to_excel(bank_file, index=False)

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()