import os
import requests

# Crear carpeta de descarga
carpeta = "descargas_directas"
os.makedirs(carpeta, exist_ok=True)

fecha = "2025-01-12"  # Fecha del archivo a descargar (Día 28)


#quiero eliminar los - de fecha
fecha = fecha.replace("-", "")
#quiero que extraiga de la fecha el mes
dia = fecha[6:8]  # Extrae el día de la fecha
mes = fecha[4:6]  # Extrae el mes de la fecha
año = fecha[:4]  # Extrae el año de la fecha
#quiero que ahora me genere una variable que en vez del mes 03 me genere marzo, 02 me genere febrero, etc.
meses = {
    "01": "Enero",
    "02": "Febrero",
    "03": "Marzo",
    "04": "Abril",
    "05": "Mayo",
    "06": "Junio",
    "07": "Julio",
    "08": "Agosto",
    "09": "Septiembre",
    "10": "Octubre",
    "11": "Noviembre",
    "12": "Diciembre"
}
mes_nombre = meses[mes]  # Obtiene el nombre del mes


# URL directa del archivo (Día 28)
url = "https://www.coes.org.pe/portal/browser/download?url=Operación%2FPrograma%20de%20Operación%2FPrograma%20Diario%2F"+año+"%2F"+mes+"_"+mes_nombre+"%2FDía%20"+dia+"%2FAnexo1_Despacho_"+fecha+".xlsx"

# Nombre local del archivo
nombre_archivo = os.path.join(carpeta, "Anexo1_Despacho_"+fecha+".xlsx")

# Descargar
response = requests.get(url)
with open(nombre_archivo, "wb") as f:
    f.write(response.content)

print(f"✅ Archivo descargado en: {nombre_archivo}")
