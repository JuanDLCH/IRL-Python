import os
import pandas as pd
from utils.fecha import Fecha
import xlwings as xw
from xlwings import *
from utils.globals import rutaRobot


def obligacionesFinancieras(fecha: Fecha, primeraVez: bool, wb: xw.Book):
    doc = 'CREDITOS DE BANCOS Y OTRAS OBLIGACIONES FINANCIERAS (NUEVO)'
    hoja = 'Obligaciones Financieras'
    columnas = ['NIT', 'FechaDesIni', 'FechaVencimiento', 'ValorCredito', 'Plazo', 'Amortizacion', 'TasaInteresEfectiva', 'SaldoCapital']

    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia

    ws = wb.sheets[hoja]
    print('Diligenciando {}. . .'.format(hoja))

    ws.range('B4').value = '{}/{}/{}'.format(dia, mes, fecha.anio)

    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)

    tabla = pd.read_csv(archivo, usecols=columnas, skiprows=3, encoding='ANSI', sep=';')
    tabla = tabla[columnas]

    if not primeraVez:
        ultimaFila = ws.range('A9').end('down').row
        ws.range('A9:H' + str(ultimaFila)).clear()

    ws.range('A9').value = tabla.values

