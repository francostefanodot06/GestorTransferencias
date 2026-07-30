import os
import re
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image
import pytesseract
from fuzzywuzzy import fuzz
from pdf2image import convert_from_path

# Ignorar advertencias menores
warnings.filterwarnings('ignore', category=UserWarning)

# Configurar ruta de Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ComprobanteConciliacionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conciliación de Comprobantes - Definitivo")
        self.geometry("450x280")
        self.configure(bg="#f0f0f0")

        self.folder_path = None
        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(self, text="Gestor de Conciliación Bancaria", font=("Arial", 12, "bold"), bg="#f0f0f0")
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
            messagebox.showwarning("Atención", "No se encontró el archivo del banco (.xlsx o .csv) en la carpeta.")
            return
        
        if not invoices:
            messagebox.showwarning("Atención", "No se encontraron comprobantes (PDF, PNG, JPG) en la carpeta.")
            return

        try:
            df_bank, cred_col_real, ley_col_real = self.read_bank_file_full(bank_file)
            processed_data = self.process_invoices_strict(invoices, df_bank, cred_col_real, ley_col_real)
            self.save_conciliacion_and_update_bank(processed_data, bank_file, df_bank, cred_col_real, ley_col_real)

            messagebox.showinfo("¡Éxito total!", "El proceso de conciliación finalizó correctamente.\nSe generó la planilla y se actualizó el banco.")
        except Exception as e:
            messagebox.showerror("Error crítico en el proceso", str(e))

    def find_files(self, folder_path):
        bank_file = None
        invoices = []

        for root, _, files in os.walk(folder_path):
            for file in files:
                file_lower = file.lower()
                if file_lower.endswith(('.csv', '.xlsx', '.txt')) and not bank_file and not file.startswith('Planilla_'):
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

        df.columns = df.columns.astype(str).str.strip().str.replace('ï»¿', '', regex=True)
        
        cred_col_real, ley_col_real = None, None
        
        for col in df.columns:
            c_low = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            if any(term in c_low for term in ['credito', 'monto', 'haber', 'importe']):
                cred_col_real = col

        for col in df.columns:
            if 'leyenda' in col.lower() and 'adicional' in col.lower():
                ley_col_real = col
                break

        if not cred_col_real and len(df.columns) > 0:
            cred_col_real = df.columns[-1]

        if not ley_col_real:
            raise ValueError("No se encontró la columna 'Leyenda Adicional1' en el archivo del banco.")

        return df, cred_col_real, ley_col_real

    def process_invoices_strict(self, invoices, df_bank, cred_col, ley_col):
        conciliados = []
        no_encontrados = []
        
        df_bank_work = df_bank.copy()

        for invoice in invoices:
            invoice_name = os.path.basename(invoice)
            text = self.extract_text(invoice)
            text_upper = text.upper()
            
            # Extraer montos del comprobante (formato decimal y entero)
            montos_en_texto = re.findall(r'\b\d{1,3}(?:\.\d{3})*(?:,\d+)?\b|\b\d+(?:,\d+)?\b', text)
            montos_limpios = []
            for m in montos_en_texto:
                try:
                    val = float(m.replace('.', '').replace(',', '.'))
                    if val > 0:
                        montos_limpios.append(val)
                except ValueError:
                    pass

            match_encontrado = False

            # Recorrer las filas del banco pendientes para buscar coincidencia cruzada
            for idx, row in df_bank_work.iterrows():
                banco_nombre = str(row[ley_col]).strip()
                banco_monto = row[cred_col]
                
                try:
                    banco_monto_float = float(str(banco_monto).replace('.', '').replace(',', '.'))
                except ValueError:
                    banco_monto_float = 0.0

                # Validar si el monto del banco coincide con alguno del comprobante (margen de error menor a 1 unidad)
                monto_coincide = any(abs(m - banco_monto_float) < 1.0 for m in montos_limpios)

                if monto_coincide and banco_nombre:
                    # Limpiar el nombre del banco para buscar palabras clave o fragmentos en el texto del OCR
                    # Dividimos el nombre del cliente en palabras significativas (mayores a 3 letras)
                    palabras_cliente = [p for p in re.findall(r'[A-Z0-9]+', banco_nombre.upper()) if len(p) > 3]
                    
                    coincidencias_palabras = 0
                    for palabra in palabras_cliente:
                        if palabra in text_upper:
                            coincidencias_palabras += 1

                    # Si el nombre aparece parcialmente en el texto o el fuzzy match del nombre da alto (>65)
                    nombre_coincide = False
                    if palabras_cliente and (coincidencias_palabras / len(palabras_cliente) >= 0.5):
                        nombre_coincide = True
                    else:
                        # Fallback por similitud de texto general
                        similitud = fuzz.partial_ratio(banco_nombre.upper(), text_upper)
                        if similitud >= 65:
                            nombre_coincide = True

                    if nombre_coincide:
                        conciliados.append({
                            'archivo': invoice_name,
                            'cliente': banco_nombre,
                            'monto': banco_monto,
                            'fecha': row['Fecha'] if 'Fecha' in df_bank_work.columns else ''
                        })
                        # Eliminar del banco de trabajo para que no se vuelva a usar
                        df_bank_work = df_bank_work.drop(idx)
                        match_encontrado = True
                        break

            if not match_encontrado:
                no_encontrados.append({
                    'archivo_faltante': invoice_name,
                    'detalle': f"El comprobante {invoice_name} no matcheó con ninguna transferencia del banco."
                })

        return {'conciliados': conciliados, 'no_encontrados': no_encontrados, 'banco_restante': df_bank_work}

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
            print(f"Error OCR en {file_path}: {e}")

        return re.sub(r'\s+', ' ', extracted_text.strip())

    def save_conciliacion_and_update_bank(self, processed_data, bank_file, df_bank, cred_col, ley_col):
        output_excel = os.path.join(self.folder_path, f"Planilla_Conciliacion_{os.path.basename(self.folder_path)}.xlsx")
        
        df_conciliados = pd.DataFrame(processed_data['conciliados'])
        df_no_encontrados = pd.DataFrame(processed_data['no_encontrados'])

        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_final = pd.DataFrame()
            
            if not df_conciliados.empty:
                df_final = pd.concat([df_final, df_conciliados], axis=1)
            else:
                df_final['Mensaje_Conciliados'] = ['No hubo comprobantes conciliados']

            df_final['---'] = ''

            if not df_no_encontrados.empty:
                df_final = pd.concat([df_final, df_no_encontrados], axis=1)
            else:
                df_final['Mensaje_No_Encontrados'] = ['Todos los comprobantes fueron conciliados']

            df_final.to_excel(writer, sheet_name='Conciliacion_y_Faltantes', index=False)

        # Actualizar el archivo del banco original removiendo los ítems conciliados
        df_banco_original = df_bank.copy()
        if processed_data['conciliados']:
            for item in processed_data['conciliados']:
                mask = (df_banco_original[cred_col] == item['monto']) & (df_banco_original[ley_col].astype(str) == str(item['cliente']))
                idx_to_drop = df_banco_original[mask].index
                if not idx_to_drop.empty:
                    df_banco_original = df_banco_original.drop(idx_to_drop[0])

        if bank_file.endswith('.csv'):
            df_banco_original.to_csv(bank_file, index=False, encoding='utf-8-sig', sep=';')
        else:
            df_banco_original.to_excel(bank_file, index=False)

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()