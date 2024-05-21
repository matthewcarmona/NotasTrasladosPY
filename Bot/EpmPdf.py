from os.path import join
from os import listdir
import datetime as dt
import pdfkit
from xlsx2html import xlsx2html

# Variables
path = "C:\\Users\\correo.ejemplo\\OneDrive\\NotasTrasladosPY"
folder_name = 'epm'

# Función para corregir la codificación de un archivo
def correct_file_encoding(path: str) -> None:
    with open(path, mode='r', encoding='latin1') as fp:
        html = """"""
        for line in fp.readlines():
            html += line
    with open(path, mode='w', encoding='utf8') as fp:
        fp.write(html)

# Declaración de una variable para la fecha de hoy
today = dt.datetime.today().strftime("%d-%m-%Y")

# Configuración global de pdfkit
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')

# Ruta de la carpeta donde están almacenados los reportes a convertir
folder = join(path, "reports", today, folder_name)

# Lista de archivos en la carpeta
files = [file.split('.')[0] for file in listdir(folder)]

# Procesamiento de cada archivo
for file in files:
    filepath = join(folder, f'{file}.xlsx')
    html_path = join(folder, f'{file}.html')
    pdf_path = join(path, "reports", today, f'{file}.pdf')
    
    # Generar archivo HTML
    fp = xlsx2html(filepath, html_path)
    fp.close()
    
    # Corregir codificación HTML
    correct_file_encoding(html_path)
    
    # Generar PDF del archivo
    pdfkit.from_file(html_path, pdf_path, configuration=config)

print("Consolidado EPM converted to PDF")
