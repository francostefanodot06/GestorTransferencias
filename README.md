 # Actualmente en desarrollo
 # 🧾 Gestor de Transferencias y Comprobantes

Aplicación de Escritorio desarrollada en **Python** con interfaz gráfica (**Tkinter**) diseñada para automatizar la conciliación de comprobantes de pago (imágenes y PDFs) contra extractos bancarios (CSV o XLSX), facilitando el control de rendiciones y el filtrado de transferencias pendientes.

---

## 🚀 Características Principales
- **Lectura Flexible:** Soporta extractos bancarios tanto en formato `.csv` como `.xlsx`.
- **Filtrado Estricto:** Ignora columnas innecesarias y se centra exclusivamente en las columnas clave: `Fecha`, `Creditos` y `Leyenda Adicional1`.
- **Conciliación Inteligente:** Utiliza **OCR (Pytesseract)** para extraer texto de comprobantes escaneados o imágenes y **Fuzzywuzzy** para realizar coincidencias de nombres (matching difuso).
- **Gestión de Pendientes:** Actualiza automáticamente el archivo del banco eliminando las transferencias ya utilizadas y dejando únicamente el remanente de las que faltan rendir.

---

## 📦 Librerías Necesarias y su Propósito

Para que el programa funcione correctamente, requiere las siguientes dependencias de Python:

| Librería | ¿Por qué se necesita? |
| :--- | :--- |
| **`pandas`** | Es la librería fundamental para la manipulación y análisis de datos. Se encarga de leer, filtrar y modificar las planillas del banco (CSV y Excel). |
| **`pillow`** | (PIL) Permite abrir, manipular y procesar archivos de imágenes (`.png`, `.jpg`, `.jpeg`) antes de pasarlas por el motor de reconocimiento de texto. |
| **`pytesseract`** | Es el puente en Python para utilizar el motor OCR *Tesseract*, encargado de "leer" y extraer el texto impreso o digital dentro de las imágenes de los comprobantes. |
| **`pdf2image`** | Convierte páginas de documentos PDF en imágenes que luego pueden ser procesadas por el OCR para extraer los datos de las transferencias. |
| **`openpyxl`** | Es el motor que permite a `pandas` leer, escribir y manipular archivos de Excel con extensión `.xlsx`. |
| **`fuzzywuzzy`** | Implementa algoritmos de coincidencia difusa (*fuzzy string matching*). Se usa para comparar los nombres y descripciones de los comprobantes contra el extracto bancario tolerando pequeñas diferencias o abreviaturas. |
| **`python-Levenshtein`** | Funciona como un acelerador en segundo plano para `fuzzywuzzy`, haciendo que las búsquedas y comparaciones de texto sean mucho más rápidas. |

---

## 🛠️ Instalación y Configuración

1. **Clonar el repositorio o descargar el código:**
   ```bash
   git clone https://github.com/francostefanodot06/GestorTransferencias.git
2. **Descargar las librerias de Python:**
   ```bash 
   pip install pandas pillow pytesseract pdf2image openpyxl fuzzywuzzy python-Levenshtein
3. **Ejecutar el Programa:**
   ```bash
    python GestorTransferencias.py
