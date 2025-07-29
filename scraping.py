import os
import requests
import smtplib
from email.message import EmailMessage
from datetime import datetime

# Obtener credenciales de variables de entorno
GMAIL_USER = os.environ['GMAIL_USER']
GMAIL_PASS = os.environ['GMAIL_APP_PASSWORD']
DESTINATARIO = os.environ['GMAIL_TO']

# Obtener fecha actual (formato YYYY-MM-DD)
fecha_actual = datetime.now().strftime("%Y-%m-%d")
fecha = fecha_actual.replace("-", "")
dia = fecha[6:8]
mes = fecha[4:6]
año = fecha[:4]

meses = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}
mes_nombre = meses[mes]

try:
    # Construir URL y ruta de archivo
    url = f"https://www.coes.org.pe/portal/browser/download?url=Operación%2FPrograma%20de%20Operación%2FPrograma%20Diario%2F{año}%2F{mes}_{mes_nombre}%2FDía%20{dia}%2FAnexo1_Despacho_{fecha}.xlsx"
    
    carpeta = "descargas_directas"
    os.makedirs(carpeta, exist_ok=True)
    nombre_archivo = os.path.join(carpeta, f"Anexo1_Despacho_{fecha}.xlsx")

    print(f"⏳ Descargando archivo de COES...")
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    
    with open(nombre_archivo, "wb") as f:
        f.write(response.content)

    print(f"✅ Archivo descargado: {nombre_archivo}")

    # Configurar email
    msg = EmailMessage()
    msg['Subject'] = f"Anexo1 Despacho COES {fecha_actual}"
    msg['From'] = GMAIL_USER
    msg['To'] = DESTINATARIO
    msg['Content-Type'] = 'text/plain; charset="utf-8"'
    
    msg.set_content(f"""
    Hola,

    Adjunto el archivo Anexo1_Despacho_{fecha}.xlsx 
    correspondiente al {dia} de {mes_nombre} de {año}.

    Descargado automáticamente desde COES.
    """)

    # Adjuntar archivo
    with open(nombre_archivo, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(
            file_data,
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=f"Anexo1_Despacho_{fecha}.xlsx"
        )

    # Enviar correo
    print("⏳ Enviando correo...")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(GMAIL_USER, GMAIL_PASS)
        smtp.send_message(msg)
    
    print("📧 Correo enviado con éxito!")
    exit(0)  # Código de salida 0 = éxito

except requests.exceptions.RequestException as e:
    print(f"❌ Error al descargar archivo: {str(e)}")
    exit(1)
except smtplib.SMTPException as e:
    print(f"❌ Error al enviar correo: {str(e)}")
    exit(2)
except Exception as e:
    print(f"❌ Error inesperado: {str(e)}")
    exit(3)