import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

carpeta = 'INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES'

def recaudoAportes(fecha: Fecha, primeraVez: bool, wb: xw.Book):
    print('Diligenciando Recaudo de Aportes. . .')
    archivos = os.listdir(rutaRobot + '/Archivos/' + carpeta)
    ws = wb.sheets['Recaudo de Aportes']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    ws.range('B8').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia, mes, fecha.anio)
    if primeraVez:
        fecha = fecha.add_months(-24)
        for i in range(25):
            archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
            archivo = os.path.join(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES/', archivo)
            tabla = pd.read_csv(archivo, skiprows=3, usecols=['Saldo a fecha'], encoding='ANSI', sep=';')
            total = tabla['Saldo a fecha'].sum()

            #Fechas
            fecha = fecha.add_months(1)
            fecha = fecha.add_days(-1)
            mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
            ws.range('A' + str(i + 13)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
            ws.range('B' + str(i + 13)).value = total

            fecha = fecha.add_months(1)
            fecha.setDay(1)
    else:
        archivo = rutaRobot + '/Archivos/INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES/INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES ' + fecha.as_Text() + '.csv'
        tabla = pd.read_csv(archivo, usecols=['Saldo a fecha'], encoding='ANSI', sep=';', skiprows=3)
        total = tabla['Saldo a fecha'].sum()

        # Obtener la siguiente fila en blanco de la columna A
        fila = ws.range('A' + str(ws.api.UsedRange.Rows.Count)).end('down').row

        #Fechas
        fecha = fecha.add_months(1)
        fecha = fecha.add_days(-1)
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        ws.range('A' + str(fila)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
        ws.range('B' + str(fila)).value = total

        fecha = fecha.add_months(1)
        fecha.setDay(1)