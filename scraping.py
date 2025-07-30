import os
import requests
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage

# ✅ Leer variables desde el entorno
GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_PASS = os.environ["GMAIL_PASS"]
DESTINATARIO = os.environ["DESTINATARIO"]

# ✅ Obtener la fecha de AYER
hoy = datetime.utcnow() #- timedelta(days=1)  # UTC porque GitHub Actions usa UTC
fecha = hoy.strftime("%Y%m%d")
dia = hoy.strftime("%d")
mes = hoy.strftime("%m")
año = hoy.strftime("%Y")

# ✅ Diccionario de nombres de mes
meses = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}
mes_nombre = meses[mes]

try:
    # ✅ Construir URL
    url = f"https://www.coes.org.pe/portal/browser/download?url=Operación%2FPrograma%20de%20Operación%2FPrograma%20Diario%2F{año}%2F{mes}_{mes_nombre}%2FDía%20{dia}%2FAnexo1_Despacho_{fecha}.xlsx"

    carpeta = "descargas_directas"
    os.makedirs(carpeta, exist_ok=True)
    nombre_archivo = os.path.join(carpeta, f"Anexo1_Despacho_{fecha}.xlsx")

    print(f"⏳ Descargando archivo desde: {url}")
    response = requests.get(url, timeout=10)
    response.raise_for_status()

    with open(nombre_archivo, "wb") as f:
        f.write(response.content)

    print(f"✅ Archivo descargado en: {nombre_archivo}")

    # ✅ Crear el mensaje de correo
    msg = EmailMessage()
    msg['Subject'] = f"Archivo Despacho {fecha}"
    msg['From'] = GMAIL_USER
    msg['To'] = DESTINATARIO

    cuerpo = f"""
    Hola,

    Adjunto el archivo 'Anexo1_Despacho_{fecha}.xlsx' correspondiente al día {dia} de {mes_nombre} de {año}.

    Saludos.
    """
    msg.set_content(cuerpo, charset='utf-8')

    # ✅ Adjuntar el archivo
    with open(nombre_archivo, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=os.path.basename(nombre_archivo)
        )

    # ✅ Enviar el correo
    print("📤 Enviando correo...")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(GMAIL_USER, GMAIL_PASS)
        smtp.send_message(msg)

    print("📧 Correo enviado con éxito!")

except requests.exceptions.RequestException as e:
    print(f"❌ Error al descargar el archivo: {e}")
except smtplib.SMTPException as e:
    print(f"❌ Error al enviar el correo: {e}")
except Exception as e:
    print(f"❌ Error inesperado: {str(e)}")

