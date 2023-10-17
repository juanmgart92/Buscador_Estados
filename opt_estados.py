import os
import pdfplumber
import pandas as pd
from datetime import date

# Crea una carpeta para almacenar los resultados
result_folder = "resultados"
os.makedirs(result_folder, exist_ok=True)

# Leer los datos desde el archivo Excel
df = pd.read_excel("./DATA/estados.xlsx")

# Inicializar un DataFrame para los resultados
resultados = pd.DataFrame(columns=["radicado", "se_encontro", "fecha_ejecucion", "archivo_encontrado"])

# Ruta de la carpeta con los archivos PDF
pdf_folder = "./archivos"

# Lista los archivos PDF en la carpeta
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

# Fecha en la que se ejecutó el programa
fecha_ejecucion = date.today()

for _, row in df.iterrows():
    radicado = row['radicado']
    encontrado = None
    archivo_encontrado = None
    
    for pdf_file in pdf_files:
        with pdfplumber.open(os.path.join(pdf_folder, pdf_file)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if radicado in page_text:
                    encontrado = 'Revisar'
                    archivo_encontrado = pdf_file
                    break
    
    resultados = resultados.append({"radicado": radicado, "se_encontro": encontrado, "fecha_ejecucion": fecha_ejecucion, "archivo_encontrado": archivo_encontrado}, ignore_index=True)

# Guardar los resultados en un archivo Excel
resultados_file = os.path.join(result_folder, f"resultados_{fecha_ejecucion}.xlsx")
resultados.to_excel(resultados_file, index=False)

print(f"Resultados guardados en {resultados_file}")
