import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot
from PyQt5.QtWidgets import *

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
            try:
                tabla = pd.read_csv(archivo, skiprows=3, usecols=['Saldo a fecha'], encoding='ANSI', sep=';')
            except:
                QMessageBox.warning(None, 'Error', 'No se pudo leer el archivo {} {}, por favor pongalo en la carpeta ArchivosNuevos'.format(carpeta, fecha.as_Text()))
                os.remove(archivo)
                os.system('explorer ' + rutaRobot + '\ArchivosNuevos')
                os.system('python ui.py')
                exit()

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
        ultimaFilaFecha = ws.range('A' + str(ws.api.UsedRange.Rows.Count+13)).end('up').row
        ultimaFilaSaldo = ws.range('B' + str(ws.api.UsedRange.Rows.Count+13)).end('up').row
        #Fechas
        fecha = fecha.add_months(1)
        fecha = fecha.add_days(-1)
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        ws.range('A' + str(ultimaFilaFecha + 1)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
        ws.range('B' + str(ultimaFilaSaldo + 1)).value = total
        #ws.range('A38').value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
        #ws.range('B38').value = total

        fecha = fecha.add_months(1)
        fecha.setDay(1)