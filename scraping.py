import os
import requests
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage

# Configuración - Reemplaza con tus datos reales
GMAIL_USER = "operadorcontrolyunqui@gmail.com"
GMAIL_PASS = "xhtn mjon xxvj ifnh"  # Usa contraseña de aplicación
DESTINATARIO = "rj5281@engie.com"   # Cambiar si es necesario

# Obtener la fecha de mañana (esta en horario UTC (0))
hoy = datetime.now() # + timedelta(days=1)
fecha = hoy.strftime("%Y%m%d")
dia = hoy.strftime("%d")
mes = hoy.strftime("%m")
año = hoy.strftime("%Y")

meses = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}
mes_nombre = meses[mes]

try:
    # Construir URL y descargar archivo
    url = f"https://www.coes.org.pe/portal/browser/download?url=Operación%2FPrograma%20de%20Operación%2FPrograma%20Diario%2F{año}%2F{mes}_{mes_nombre}%2FDía%20{dia}%2FAnexo1_Despacho_{fecha}.xlsx"

    carpeta = "descargas_directas"
    os.makedirs(carpeta, exist_ok=True)
    nombre_archivo = os.path.join(carpeta, f"Anexo1_Despacho_{fecha}.xlsx")

    print(f"⏳ Descargando archivo...")
    response = requests.get(url, timeout=10)
    response.raise_for_status()

    with open(nombre_archivo, "wb") as f:
        f.write(response.content)

    print(f"✅ Archivo descargado en: {nombre_archivo}")

    # Crear correo
    msg = EmailMessage()
    msg['Subject'] = f"Archivo Despacho {fecha}"
    msg['From'] = GMAIL_USER
    msg['To'] = DESTINATARIO
    msg.set_content(
        f"""Hola,

Adjunto el archivo 'Anexo1_Despacho_{fecha}.xlsx' correspondiente al día {dia} de {mes_nombre} de {año}.

Saludos.
""", charset='utf-8')

    # Adjuntar archivo
    with open(nombre_archivo, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=os.path.basename(nombre_archivo)
        )

    # Enviar correo
    print("⏳ Enviando correo...")
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
