import os
import pdfplumber
import pandas as pd
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configuración de las credenciales de correo electrónico
email_address = "juanma.0627ga@gmail.com"
email_password = "quwx ybbt qmac igsb"
recipient_email = "juanma.627@hotmail.com"
smtp_server = "smtp.gmail.com"
smtp_port = 587

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

# Crear un mensaje de correo electrónico
msg = MIMEMultipart()
msg['From'] = email_address
msg['To'] = recipient_email
msg['Subject'] = f"Resultados de la búsqueda ({fecha_ejecucion})"

# Guardar los resultados en un archivo Excel
resultados_file = os.path.join(result_folder, f"resultados_{fecha_ejecucion}.xlsx")
resultados.to_excel(resultados_file, index=False)

from email.mime.base import MIMEBase
from email import encoders

# Adjuntar los resultados como un archivo Excel al correo electrónico
with open(resultados_file, "rb") as file:
    attachment = MIMEBase("application", "octet-stream")
    attachment.set_payload(file.read())
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', f'attachment; filename={resultados_file}')
    msg.attach(attachment)

# Configurar el servidor SMTP
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(email_address, email_password)

# Enviar el correo electrónico
server.sendmail(email_address, recipient_email, msg.as_string())

# Cerrar la conexión con el servidor SMTP
server.quit()

print(f"Resultados enviados por correo electrónico a {recipient_email}")
