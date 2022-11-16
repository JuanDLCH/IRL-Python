import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

def gastosAdministrativos(fecha: Fecha, primeraVez: bool, wb: xw.Book):
    print('Diligenciando Gastos Administrativos. . .')
    doc = 'CATALOGO DE CUENTAS'
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    ws = wb.sheets['Gastos Administrativos']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    ws.range('B5').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    if primeraVez:
        fecha = fecha.add_months(-12)
        for i in range(13):
            mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
            dia = fecha.add_months(1).add_days(-1).dia
            ws.range('A' + str(i + 14)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)

            archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
            archivo = os.path.join(rutaRobot + '/Archivos/' + doc + '/', archivo)
            tabla = pd.read_csv(archivo, skiprows=3, usecols=['CUENTA', 'Saldo'], encoding='ANSI', sep=';')
            cuentas = [510500, 511000]
            tabla = tabla[tabla['CUENTA'].isin(cuentas)]
            #Obtener el saldo de la cuenta 510500
            gastosBeneficios = tabla[tabla['CUENTA'] == 510500]['Saldo'].sum()
            gastosGenerales = tabla[tabla['CUENTA'] == 511000]['Saldo'].sum()

            ws.range('B' + str(i + 14)).value = gastosBeneficios
            ws.range('C' + str(i + 14)).value = gastosGenerales

            fecha = fecha.add_months(1)
    else:
        ultimafila = ws.range('A14').end('down').row
        ws.range('A' + str(ultimafila + 1)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)
        archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
        archivo = os.path.join(rutaRobot + '/Archivos/' + doc + '/', archivo)
        tabla = pd.read_csv(archivo, skiprows=3, usecols=['CUENTA', 'Saldo'], encoding='ANSI', sep=';')
        cuentas = [510500, 511000]
        tabla = tabla[tabla['CUENTA'].isin(cuentas)]
        #Obtener el saldo de la cuenta 510500
        gastosBeneficios = tabla[tabla['CUENTA'] == 510500]['Saldo'].sum()
        gastosGenerales = tabla[tabla['CUENTA'] == 511000]['Saldo'].sum()

        ws.range('B' + str(ultimafila + 1)).value = gastosBeneficios
        ws.range('C' + str(ultimafila + 1)).value = gastosGenerales




