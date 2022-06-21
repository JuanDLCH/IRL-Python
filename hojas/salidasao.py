import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot
doc = 'SALDOS DIARIOS DE AHORRO'


def salidasao(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Salidas de ahorro ordinario. . .')
   

    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    
    ws = wb.sheets['Salidas De Ahorro ordinario']
    tabla = pd.read_csv(archivo, usecols=['Saldo'], encoding='ANSI', sep=';', skiprows=3)
    ws.range('F4')
  
  
    for i in range(31):
        saldo = tabla.get('Saldo')[i]
        ws.range('B' + str(i + 7)).value = saldo  
        ws.range('A' + str(i + 7)).value = '{}/{}/{}'.format(fecha.add_days(i).dia,fecha.mes,fecha.anio)



