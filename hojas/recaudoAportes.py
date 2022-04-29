import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

def recaudoAportes(wb: xw.Book, fecha: Fecha, primeraVez: bool):
    print('Diligenciando Recaudo de Aportes. . .')
    archivos = os.listdir(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES')
    ws = wb.sheets['Recaudo de Aportes']
    if primeraVez:
        fecha = fecha.add_months(-24)
        for i in range(25):
            archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
            archivo = os.path.join(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES', archivo)
            tabla = pd.read_csv(archivo, usecols=['Saldo a fecha'])
            total = tabla['Saldo a fecha'].sum()

            #Fechas
            fecha = fecha.add_months(1)
            fecha = fecha.add_days(-1)
            mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
            ws.range('A' + str(i + 13)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
            ws.range('B' + str(i + 13)).value = fecha.as_Text()

            fecha = fecha.add_months(1)
            fecha.setDay(1)
    else:
        archivo = rutaRobot + '/Archivos/INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES/' + fecha.as_Text() + '.csv'
        tabla = pd.read_csv(archivo, usecols=['Saldo a fecha'])
        total = tabla['Saldo a fecha'].sum()

        fila = ws.range('A13:A').end('down').row
        ws.range('A' + str(fila)).value = fecha.as_Text()

        #Fechas
        fecha = fecha.add_months(1)
        fecha = fecha.add_days(-1)
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        ws.range('A' + fila).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
        ws.range('B' + fila).value = fecha.as_Text()

        ws.range("C37:I37").api.AutoFill(ws.range("C37:F{row}".format(row=fila)).api, 0 )

        fecha = fecha.add_months(1)
        fecha.setDay(1)