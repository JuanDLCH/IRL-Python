from dataclasses import replace
from datetime import datetime
from email import message
import os
from tkinter import messagebox

from pyparsing import withAttribute
import wget
from utils.fecha import Fecha

urls = [
    'https://www.supersolidaria.gov.co/sites/default/files/public/data/desviacion_estandar_',
    'https://www.supersolidaria.gov.co/sites/default/files/public/entidades/desviacion_estandar_',
    ]

meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']


# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')

# Descargar un archivo de la web
def descargarArchivo():
    fechaActual = Fecha(1, datetime.now().month, datetime.now().year)
    found = False
    while not found:
        for url in urls:
            try:
                if os.path.exists(rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx'):
                    found = True
                else:
                    wget.download(url + fechaActual.as_Text().lower().replace(' ', '_') + '.xlsx', rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx')
                    found = True
            except:
                pass
        fechaActual = fechaActual.add_months(-1)

descargarArchivo()
