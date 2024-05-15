from os.path import join
from os import listdir
import datetime as dt
import pdfkit
from pdfkit import configuration

from xlsx2html import xlsx2html


path = "C:\\Users\\correo.automatizacio\\OneDrive - GANA S.A\\NotasTrasladosPY"
config = configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')


# Nombre de la carpeta donde serán almacenados los reportes generados.
folder_name = 'notas'

global today


def correct_file_encoding(path: str) -> None:
    with open(path, mode='r', encoding='latin1') as fp:
        html = """"""
        for line in fp.readlines():
            html += line
    with open(path, mode='w', encoding='utf8') as fp:
        fp.write(html)


# Declaración de una variable para la fecha de hoy
today = dt.datetime.today().strftime("%d-%m-%Y")

# ruta de la carpeta donde estan almacenados los reportes a convertir.
folder = join(path, "reports", today, folder_name)

files = [file.split('.')[0] for file in listdir(folder)]

for file in files:
    filepath = join(folder, f'{file}.xlsx')
    html_path = join(folder, f'{file}.html')
    pdf_path = join(path, "reports", today, f'{file}.pdf')
    # Generar archivo html
    fp = xlsx2html(filepath, html_path)
    fp.close()
    # Corregir codificación html
    correct_file_encoding(html_path)
    # Generar pdf del archivo
    #pdfkit.from_file(html_path, pdf_path, configuration={'path': r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'})
    pdfkit.from_file(html_path, pdf_path, configuration=config)
print("All notas converted to pdf")
