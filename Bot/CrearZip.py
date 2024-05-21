from os import listdir
from os.path import join, isdir, relpath, exists, isfile
import datetime as dt
import zipfile
import os

# Ruta base donde se encuentran los informes
base_path = "C:\\Users\\correo.ejemplo\\OneDrive - GANA S.A\\NotasTrasladosPY"

# Obtenemos la fecha actual en el formato deseado: dia-mes-año
today = dt.datetime.now().strftime("%d-%m-%Y")

# Ruta donde se guardarán los informes y el archivo ZIP
folder = join(base_path, 'reports', today)

# Obtener la lista de archivos en la carpeta de informes que no sean directorios ni archivos .xlsx
files = [file for file in listdir(folder) if not isdir(join(folder, file)) and not file.endswith('.xlsx')]

# Ruta completa del archivo ZIP a crear
zip_path = join(folder, 'Resultado.zip')

# Verificar si el archivo ZIP ya existe
if not os.path.isfile(zip_path):
    # Si no existe, crear el archivo ZIP y agregar los archivos PDF a comprimir
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zfp:
        # Obtener todos los archivos en el directorio
        files = os.listdir(folder)
        for file in files:
            # Verificar si el archivo es un archivo PDF
            if file.lower().endswith(".pdf"):
                # Agregar el archivo PDF al archivo ZIP
                zfp.write(filename=os.path.join(folder, file), arcname=file)
    print("Zip file created.")
else:
    print("Zip file already exists.")
