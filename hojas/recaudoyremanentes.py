import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

doc = 'CATALOGO DE CUENTAS'



def recaudoyremanentes(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Recaudo y Remanentes. . .')
   
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)

    ws = wb.sheets['Recaudo y remanentes']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia

    ws.range('A9').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    #ws.range('B6:C' + str(ws.range('B6').end('down').row)).clear()
    
    tabla = pd.read_csv(archivo, usecols=['CUENTA','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    saldo1 = tabla[tabla['CUENTA'] == 246000]['Saldo'].sum()
    saldo2 = tabla[tabla['CUENTA'] == 246500]['Saldo'].sum()

    ws.range('B9').value = saldo1
    ws.range('C9').value = saldo2