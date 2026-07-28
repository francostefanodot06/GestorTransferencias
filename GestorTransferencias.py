import os
import pandas as pd
from PIL import Image
import pytesseract
import re
from pdf2image import convert_from_path
from tkinter import Tk, Button, filedialog, Text, END
from fuzzywuzzy import process

# Configurar Tesseract (asegúrate de que esté instalado y configurado)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Ajusta la ruta según tu sistema

def extract_data_from_image(image_path):
    try:
        # Leer la imagen
        img = Image.open(image_path)
        
        # Extraer texto usando Tesseract
        text = pytesseract.image_to_string(img, lang='spa')
        
        # Parsear el texto para obtener nombre, fecha y monto

        # Buscar el monto al principio con un signo de pesos
        amount_match = re.search(r'\$\s*\d{1,3}(?:\.\d{3})*(,\d{2})?', text)
        amount = float(amount_match.group(0).replace('$', '').replace('.', '').replace(',', '.')) if amount_match else None

        # Buscar el nombre después de las etiquetas "De" o "Titular"
        name_match = re.search(r'(?:De|Titular):\s*(.*)', text, re.IGNORECASE)
        name = name_match.group(1).strip() if name_match else None

        # Buscar la fecha (formato DD/MM/YYYY)
        date_match = re.search(r'\b\d{2}/\d{2}/\d{4}\b', text)
        date = date_match.group(0) if date_match else None

        return {
            'Nombre': name,
            'Fecha': date,
            'Monto': amount
        }
    except Exception as e:
        print(f"Error al extraer datos de la imagen {image_path}: {e}")
        return None

def extract_data_from_pdf(pdf_path):
    try:
        # Convertir PDF a imágenes
        pages = convert_from_path(pdf_path, poppler_path=r'C:\Program Files\poppler-23.07.0\Library\bin')  # Ajusta la ruta según tu sistema
        
        extracted_text = ""
        for page in pages:
            text = pytesseract.image_to_string(page, lang='spa')
            extracted_text += text

        # Parsear el texto para obtener nombre, fecha y monto

        # Buscar el monto al principio con un signo de pesos
        amount_match = re.search(r'\$\s*\d{1,3}(?:\.\d{3})*(,\d{2})?', extracted_text)
        amount = float(amount_match.group(0).replace('$', '').replace('.', '').replace(',', '.')) if amount_match else None

        # Buscar el nombre después de las etiquetas "De" o "Titular"
        name_match = re.search(r'(?:De|Titular):\s*(.*)', extracted_text, re.IGNORECASE)
        name = name_match.group(1).strip() if name_match else None

        # Buscar la fecha (formato DD/MM/YYYY)
        date_match = re.search(r'\b\d{2}/\d{2}/\d{4}\b', extracted_text)
        date = date_match.group(0) if date_match else None

        return {
            'Nombre': name,
            'Fecha': date,
            'Monto': amount
        }
    except Exception as e:
        print(f"Error al extraer datos del PDF {pdf_path}: {e}")
        return None

def process_invoices(invoice_files, bank_statement_csv, progress_text):
    # Cargar el archivo CSV del extracto bancario
    df_bank = pd.read_csv(bank_statement_csv)
    
    # Crear una copia para trabajar con los datos originales
    df_copy = df_bank.copy()
    
    # Lista para almacenar comprobantes pendientes de revisión manual
    pending_reviews = []
    
    # Procesar cada comprobante
    for invoice_file in invoice_files:
        try:
            if invoice_file.lower().endswith(('.png', '.jpg', '.jpeg')):
                data = extract_data_from_image(invoice_file)
            elif invoice_file.lower().endswith('.pdf'):
                data = extract_data_from_pdf(invoice_file)
            else:
                print(f"Formato de archivo no soportado: {invoice_file}")
                pending_reviews.append(invoice_file)
                continue

            if not all(data.values()):
                raise ValueError("Datos incompletos extraídos del archivo")

            # Buscar coincidencias en el CSV usando Fuzzy Matching
            best_match, score = process.extractOne(data['Nombre'], df_copy['Leyenda Adicional1'].dropna(), scorer=partial_ratio)
            
            if score >= 70:  # Umbral de similitud (ajusta según necesidad)
                mask = (df_copy['Leyenda Adicional1'] == best_match) & \
                       ((df_copy['Créditos'] == data['Monto']) | (df_copy['Débitos'] == -data['Monto']))
                
                matching_rows = df_copy[mask]
                
                if not matching_rows.empty:
                    # Crear o abrir el archivo Excel correspondiente al cobrador
                    cobrador = best_match
                    excel_file = f"Cobrador_{cobrador}.xlsx"
                    
                    if os.path.exists(excel_file):
                        writer = pd.ExcelWriter(excel_file, engine='openpyxl', mode='a')
                        matching_rows[['Fecha', 'Leyenda Adicional1', 'Créditos']].to_excel(writer, index=False, header=False, startrow=writer.sheets[cobrador].max_row)
                    else:
                        matching_rows[['Fecha', 'Leyenda Adicional1', 'Créditos']].to_excel(excel_file, sheet_name=cobrador, index=False, columns=['Fecha', 'Leyenda Adicional1', 'Créditos'])
                    
                    # Eliminar las filas procesadas del CSV original
                    df_copy.drop(matching_rows.index, inplace=True)
                else:
                    raise ValueError("No se encontraron coincidencias en el archivo CSV")
            else:
                raise ValueError(f"Coincidencia de nombre insuficiente: {best_match} (Score: {score})")
        except Exception as e:
            print(f"Error procesando el archivo {invoice_file}: {e}")
            pending_reviews.append(invoice_file)

        # Actualizar la interfaz gráfica con el progreso
        progress_text.insert(END, f"Procesado: {os.path.basename(invoice_file)}\n")
        progress_text.see(END)
    
    # Guardar los cambios en el archivo CSV original al final
    df_copy.to_csv(bank_statement_csv, index=False)
    
    # Guardar la lista de pendientes de revisión manual al final
    if pending_reviews:
        with open('pendientes_revisar.txt', 'w') as f:
            for file in pending_reviews:
                f.write(file + '\n')

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        # Actualizar la interfaz gráfica con el progreso
        progress_text.insert(END, f"Carpeta seleccionada: {folder_path}\n")
        progress_text.see(END)
        
        # Obtener archivos de comprobantes y CSV del extracto bancario
        invoice_files = []
        bank_statement_csv = None
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    invoice_files.append(os.path.join(root, file))
                elif file.lower() == 'extracto_bancario.csv':
                    bank_statement_csv = os.path.join(root, file)
        
        # Validar que se haya encontrado el CSV
        if not bank_statement_csv:
            progress_text.insert(END, "Error: No se encontró el archivo 'extracto_bancario.csv' en la carpeta seleccionada.\n")
            progress_text.see(END)
            return

        # Procesar los comprobantes con progreso visualizado en la interfaz
        process_invoices(invoice_files, bank_statement_csv, progress_text)

def create_gui():
    root = Tk()
    root.title("Procesador de Comprobantes")

    # Botón para seleccionar carpeta
    select_button = Button(root, text="Seleccionar Carpeta", command=select_folder)
    select_button.pack(pady=10)

    # Botón para procesar (desactivado hasta que se seleccione una carpeta)
    process_button = Button(root, text="Procesar", state='disabled', command=lambda: process_invoices(invoice_files, bank_statement_csv, progress_text))
    process_button.pack(pady=10)

    # Caja de texto para mostrar el progreso
    global progress_text
    progress_text = Text(root, height=20, width=80)
    progress_text.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()