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
            df_bank, cred_col_real, ley_col_real = self.read_bank_file_full(bank_file)
            
            # Procesamos con validación estricta (Nombre + Monto cruzado)
            processed_data = self.process_invoices_strict(invoices, df_bank, cred_col_real, ley_col_real)

            self.save_conciliacion_and_update_bank(processed_data, bank_file, df_bank, cred_col_real, ley_col_real)

            messagebox.showinfo("Proceso Completado", "¡Proceso exitoso!\nSe conciliaron los datos y se registraron los faltantes.")
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

        return df, cred_col_real, ley_col_real

    def process_invoices_strict(self, invoices, df_bank, cred_col, ley_col):
        conciliados = []
        
        # Pre-extraemos el texto de todos los comprobantes y buscamos números en ellos
        comprobantes_info = []
        for invoice in invoices:
            text = self.extract_text(invoice)
            # Extraemos todos los posibles montos numéricos del comprobante
            montos_en_texto = re.findall(r'\b\d{1,3}(?:\.\d{3})*(?:,\d+)?\b|\b\d+(?:,\d+)?\b', text)
            # Limpiamos los montos a formato float para comparar bien
            montos_limpios = []
            for m in montos_en_texto:
                m_clean = m.replace('.', '').replace(',', '.')
                try:
                    montos_limpios.append(float(m_clean))
                except ValueError:
                    pass

            comprobantes_info.append({
                'archivo': os.path.basename(invoice),
                'text': text,
                'montos': montos_limpios
            })

        # Recorremos cada registro del banco para ver si encuentra su comprobante exacto
        banco_pendientes = df_bank.copy()
        
        for idx, row in df_bank.iterrows():
            banco_nombre = str(row[ley_col])
            banco_monto = row[cred_col]
            
            try:
                banco_monto_float = float(str(banco_monto).replace('.', '').replace(',', '.'))
            except ValueError:
                banco_monto_float = 0.0

            match_encontrado = False
            
            for comp in comprobantes_info:
                # Verificamos si el nombre del banco aparece en el texto del comprobante con alta coincidencia o está contenido
                # O si el FuzzyWuzzy del texto del comprobante da con el nombre del banco
                match_score = process.extractOne(banco_nombre, [comp['text']])[1]
                
                # Validamos que coincida el nombre (score alto) Y que el monto del banco esté presente en los montos leídos del comprobante
                monto_coincide = any(abs(m - banco_monto_float) < 1.0 for m in comp['montos'])
                
                if match_score >= 75 and monto_coincide:
                    conciliados.append({
                        'archivo': comp['archivo'],
                        'cliente': banco_nombre,
                        'monto': banco_monto,
                        'fecha': row['Fecha'] if 'Fecha' in df_bank.columns else ''
                    })
                    # Quitamos este comprobante ya usado para que no se repita
                    comprobantes_info.remove(comp)
                    match_encontrado = True
                    break

        return {'conciliados': conciliados, 'banco_restante': banco_pendientes, 'comprobantes_sobrantes': comprobantes_info}

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

    def save_conciliacion_and_update_bank(self, processed_data, bank_file, df_bank, cred_col, ley_col):
        output_excel = os.path.join(self.folder_path, f"Planilla_Conciliacion_{os.path.basename(self.folder_path)}.xlsx")
        
        df_conciliados = pd.DataFrame(processed_data['conciliados'])
        
        # Preparamos los faltantes (lo que quedó en el banco sin conciliar o comprobantes sobrantes)
        faltantes_list = []
        # Si hay filas en el banco que no se conciliaron, las listamos al costado como faltantes
        conciliados_clientes = [c['cliente'] for c in processed_data['conciliados']]
        
        for idx, row in df_bank.iterrows():
            cli = str(row[ley_col])
            monto = row[cred_col]
            if cli not in conciliados_clientes:
                faltantes_list.append({
                    'Faltante_Cliente': cli,
                    'Faltante_Monto': monto,
                    'Detalle': f"Faltó conciliar: {monto} de {cli}"
                })

        df_faltantes = pd.DataFrame(faltantes_list)

        # Guardamos en un Excel combinando en columnas al costado (Conciliados en A-D, Faltantes en F en adelante)
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # Creamos una única solapa principal limpia donde se ve todo lado a lado
            df_final = pd.DataFrame()
            
            if not df_conciliados.empty:
                df_final = pd.concat([df_final, df_conciliados], axis=1)
            else:
                df_final['Mensaje_Conciliados'] = ['No hubo coincidencias exactas']

            # Agregamos una columna separadora vacía
            df_final['---'] = ''

            if not df_faltantes.empty:
                df_final = pd.concat([df_final, df_faltantes], axis=1)
            else:
                df_final['Mensaje_Faltantes'] = ['No hay faltantes']

            df_final.to_excel(writer, sheet_name='Conciliacion_y_Faltantes', index=False)

        # Actualizar el archivo del banco original borrando exclusivamente los conciliados
        if processed_data['conciliados']:
            for item in processed_data['conciliados']:
                mask = (df_bank[cred_col] == item['monto']) & (df_bank[ley_col].astype(str) == str(item['cliente']))
                idx_to_drop = df_bank[mask].index
                if not idx_to_drop.empty:
                    df_bank = df_bank.drop(idx_to_drop[0])

        if bank_file.endswith('.csv'):
            df_bank.to_csv(bank_file, index=False, encoding='utf-8-sig', sep=';')
        else:
            df_bank.to_excel(bank_file, index=False)