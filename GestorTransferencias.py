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

# Configurar el directorio de Tesseract (ajusta la ruta según tu PC)
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

        df_bank = self.read_bank_file(bank_file)
        processed_data = self.process_invoices(invoices, df_bank)

        for cobrador, data in processed_data.items():
            self.save_conciliacion(data, cobrador, bank_file)

        messagebox.showinfo("Proceso Completado", "El proceso de conciliación ha finalizado con éxito.")

    def find_files(self, folder_path):
        bank_file = None
        invoices = []

        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(('.csv', '.xlsx')):
                    if not bank_file:
                        bank_file = os.path.join(root, file)
                elif file.lower().endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    invoices.append(os.path.join(root, file))

        return bank_file, invoices

    def read_bank_file(self, bank_file):
        if bank_file.endswith('.csv'):
            try:
                df = pd.read_csv(bank_file, encoding='latin1', on_bad_lines='skip')
            except Exception:
                df = pd.read_csv(bank_file, encoding='latin1', sep=';', on_bad_lines='skip')
        elif bank_file.endswith('.xlsx'):
            df = pd.read_excel(bank_file)

        relevant_columns = ['Fecha', 'Creditos', 'Leyenda Adicional1']
        return df[relevant_columns]

    def process_invoices(self, invoices, df_bank):
        processed_data = {}

        for invoice in invoices:
            text = self.extract_text(invoice)
            cobrador_match = process.extractOne(text, df_bank['Leyenda Adicional1'].dropna())
            
            if cobrador_match and cobrador_match[1] >= 70:  # Umbral de confianza
                cobrador_name = cobrador_match[0]
                matching_rows = df_bank[df_bank['Leyenda Adicional1'] == cobrador_name]

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
                # Ajusta la ruta de poppler según tu sistema si es necesario
                pages = convert_from_path(file_path, poppler_path=r'C:\Program Files\poppler-23.07.0\Library\bin')
                for page in pages:
                    extracted_text += pytesseract.image_to_string(page, lang='spa')
            else:
                image = Image.open(file_path).convert('RGB')
                extracted_text = pytesseract.image_to_string(image, lang='spa')
        except Exception as e:
            print(f"Error al extraer texto de {file_path}: {e}")

        return re.sub(r'\s+', ' ', extracted_text.strip())

    def save_conciliacion(self, data, cobrador, bank_file):
        # Crear nombre de carpeta limpio con la fecha de la carpeta raíz y el nombre del cobrador
        folder_name = f"Rendicion_{os.path.basename(self.folder_path)}_{cobrador}"
        output_folder = os.path.join(self.folder_path, folder_name)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Copiar comprobantes procesados a la nueva carpeta
        for invoice in data['invoices']:
            src_file = invoice['file_path']
            dst_file = os.path.join(output_folder, os.path.basename(src_file))
            shutil.copy2(src_file, dst_file)

        # Actualizar el archivo bancario original (quitando las usadas, dejando las pendientes)
        if bank_file:
            df_bank = pd.read_csv(bank_file, encoding='latin1') if bank_file.endswith('.csv') else pd.read_excel(bank_file)
            
            for credit in data['bank']:
                index = df_bank[(df_bank['Creditos'] == credit) & (df_bank['Leyenda Adicional1'] == cobrador)].index
                if not index.empty:
                    df_bank.drop(index, inplace=True)

            if bank_file.endswith('.csv'):
                df_bank.to_csv(bank_file, index=False, encoding='latin1')
            else:
                df_bank.to_excel(bank_file, index=False)

if __name__ == "__main__":
    app = ComprobanteConciliacionApp()
    app.mainloop()
    