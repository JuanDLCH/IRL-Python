import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

doc = 'CATALOGO DE CUENTAS'

def cxp(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Cuentas por Pagar (CxP). . .')
   
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)

    ws = wb.sheets['CxP']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia

    ws.range('B4').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    ultimaFilaSaldo = ws.range('C' + str(ws.api.UsedRange.Rows.Count)).end('up').row

    tabla = pd.read_csv(archivo, usecols=['CUENTA','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    saldo = tabla[tabla['CUENTA'] == 240000]['Saldo'].sum()

    ws.range('C' + str(ultimaFilaSaldo + 1)).value = saldo  
  

