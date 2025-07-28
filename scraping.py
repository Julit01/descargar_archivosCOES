# enviar_despacho.py

import os
import requests
import smtplib
from email.message import EmailMessage

# Fecha deseada (puedes ponerla fija o tomarla del sistema)
fecha = "2025-01-12"
fecha = fecha.replace("-", "")
dia = fecha[6:8]
mes = fecha[4:6]
año = fecha[:4]

meses = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}
mes_nombre = meses[mes]

url = f"https://www.coes.org.pe/portal/browser/download?url=Operación%2FPrograma%20de%20Operación%2FPrograma%20Diario%2F{año}%2F{mes}_{mes_nombre}%2FDía%20{dia}%2FAnexo1_Despacho_{fecha}.xlsx"

carpeta = "descargas_directas"
os.makedirs(carpeta, exist_ok=True)
nombre_archivo = os.path.join(carpeta, f"Anexo1_Despacho_{fecha}.xlsx")

response = requests.get(url)
with open(nombre_archivo, "wb") as f:
    f.write(response.content)
print(f"✅ Archivo descargado en: {nombre_archivo}")

# Envío por correo
gmail_user = os.environ["GMAIL_USER"]
gmail_pass = os.environ["GMAIL_APP_PASSWORD"]
correo_destino = os.environ["GMAIL_TO"]

msg = EmailMessage()
msg['Subject'] = f"📎 Archivo Despacho {fecha}"
msg['From'] = gmail_user
msg['To'] = correo_destino
msg.set_content(f"Hola,\n\nAdjunto el archivo 'Anexo1_Despacho_{fecha}.xlsx' correspondiente al día {dia} de {mes_nombre} de {año}.\n\nSaludos.")

with open(nombre_archivo, 'rb') as f:
    file_data = f.read()
    file_name = os.path.basename(nombre_archivo)
    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(gmail_user, gmail_pass)
    smtp.send_message(msg)

print("📧 Correo enviado con éxito.")
