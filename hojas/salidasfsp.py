import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

doc = 'CATALOGO DE CUENTAS'

cuenta = 260000


def salidasfsp(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Salidas fondos sociales pasivos. . .')
   
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)

    ws = wb.sheets['Salidas fondos sociales pasivos']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    
    tabla = pd.read_csv(archivo, usecols=['CUENTA','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    saldo = tabla[tabla['CUENTA'] == cuenta]['Saldo'].sum()

    ws.range('B13').value = saldo 
    ws.range('A13').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    ws.range('B5').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia,fecha.add_months(13-int(mes)).add_days(-1).mes,fecha.add_years(1).anio)
    ws.range('C5').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia,fecha.add_months(2).add_days(-1).mes,fecha.anio)
    ws.range('D5').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia,fecha.add_months(3).add_days(-1).mes,fecha.anio)
