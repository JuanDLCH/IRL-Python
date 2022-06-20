import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

doc = 'CATALOGO DE CUENTAS'

cuenta = 260000

Mes30 = [4, 6, 9, 11]


def salidasfsp(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Salidas fondos sociales pasivos. . .')
   
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)

    ws = wb.sheets['Salidas fondos sociales pasivos']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    
    ultimaFilaFecha = ws.range('A' + str(ws.api.UsedRange.Rows.Count)).end('up').row
    ultimaFilaSaldo = ws.range('B' + str(ws.api.UsedRange.Rows.Count)).end('up').row
 
    tabla = pd.read_csv(archivo, usecols=['CUENTA','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    saldo = tabla[tabla['CUENTA'] == cuenta]['Saldo'].sum()

    ws.range('B7').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
   

    if ws.range('A13').value == None : 
        ws.range('A' + str(ultimaFilaSaldo)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)  
        ws.range('B' + str(ultimaFilaFecha)).value = saldo  
    else : 
        ws.range('A' + str(ultimaFilaSaldo + 1)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)   
        ws.range('B' + str(ultimaFilaFecha + 1)).value = saldo 

    if fecha.mes in Mes30:

        ws.range('B6').value = '{}/{}/{}'.format(fecha.add_days(-1).dia,fecha.add_months(13-int(mes)).add_days(-1).mes,fecha.anio)
        ws.range('C6').value = '{}/{}/{}'.format(fecha.add_days(-1).dia,fecha.add_months(2).add_days(-1).mes,fecha.anio)
        ws.range('D6').value = '{}/{}/{}'.format(fecha.add_days(-1).dia,fecha.add_months(3).add_days(-1).mes,fecha.anio)

    else : 
        ws.range('B6').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia,fecha.add_months(13-int(mes)).add_days(-1).mes,fecha.anio)
        ws.range('C6').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia,fecha.add_months(2).add_days(-1).mes,fecha.anio)
        ws.range('D6').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia,fecha.add_months(3).add_days(-1).mes,fecha.anio)

        

    
    
    
   