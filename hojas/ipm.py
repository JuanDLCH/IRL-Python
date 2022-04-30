#Indice promedio de morosidad (IPM)

import xlwings as xw
from xlwings import *
from utils.fecha import *
import pandas as pd
from hojas.cartera import obtenerTabla

# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')



def ipm(fecha: Fecha, primeraVez, desviacionEstandar, wb: xw.Book):
    print('Diligenciando Indice promedio de morosidad. . .')
    ws = wb.sheets['Índice promedio de morosidad ']
    ws.range('I7').value = desviacionEstandar
    if primeraVez:
        fecha = fecha.add_months(-12)
        for i in range(13):
            # Obtener tabla
            # Obtener los archivos de la carpeta
            archivos = os.listdir(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)')
            archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
            archivo = os.path.join(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)', archivo)

            tabla = pd.read_csv(archivo, skiprows=3, usecols=['CodigoContable', 'SaldoCapital'], encoding='ANSI', sep=';')

            saldoTotal = tabla['SaldoCapital'].sum()
            codigosContables = [144110, 144210,141110, 141210, 144115, 144215, 141115, 141215, 144120, 144220,141120, 141220, 144125, 144225, 141125, 14122, 140410, 140415, 140420, 140425, 140510, 140515, 140520, 140525, 144810, 144815, 144820, 144825, 145410, 145415, 145420, 145425, 145510, 145515, 145520, 145525, 146110, 146115, 146120, 146125, 146210, 146215, 146220, 146225]
            #tabla = tabla[tabla['CodigoContable'].isin(codigosContables)]
            tabla = tabla.loc[tabla['CodigoContable'].isin(codigosContables)]


            #Fechas
            fecha = fecha.add_months(1)
            fecha = fecha.add_days(-1)
            mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
            ws.range('B' + str(i + 9)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)

            #SaldoCapital
            saldoMora = tabla['SaldoCapital'].sum()
            ws.range('C' + str(i + 9)).value = saldoMora

            #SaldoTotal
            ws.range('D' + str(i + 9)).value = saldoTotal

            fecha = fecha.add_months(1)
            fecha.setDay(1)          
    else:
        archivo = os.path.join(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)/INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO) ' + fecha.as_Text() + '.csv')
        # Obtener la proxima fila disponiible en la columna B
        fila = ws.range('B8').end('down').row + 1

        fecha = fecha.add_months(1)
        fecha = fecha.add_days(-1)
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        ws.range('B' + str(fila)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
        # Obtener tabla
        tabla = pd.read_csv(archivo, usecols=['CodigoContable', 'SaldoCapital'], encoding='ANSI', sep=';')

        saldoTotal = tabla['SaldoCapital'].sum()
        codigosContables = [144110, 144210,141110, 141210, 144115, 144215, 141115, 141215, 144120, 144220,141120, 141220, 144125, 144225, 141125, 14122, 140410, 140415, 140420, 140425, 140510, 140515, 140520, 140525, 144810, 144815, 144820, 144825, 145410, 145415, 145420, 145425, 145510, 145515, 145520, 145525, 146110, 146115, 146120, 146125, 146210, 146215, 146220, 146225]
        #tabla = tabla[tabla['CodigoContable'].isin(codigosContables)]
        tabla = tabla.loc[tabla['CodigoContable'].isin(codigosContables)]

        #SaldoCapital
        saldoMora = tabla['SaldoCapital'].sum()
        ws.range('C' + str(fila)).value = saldoMora

        #SaldoTotal
        ws.range('D' + str(fila)).value = saldoTotal

        ws.range("E9:F9").api.AutoFill(ws.range("E9:F{row}".format(row=fila)).api, 0 )
        ws.range('B8:F' + str(fila)).api.Borders(11).LineStyle = 1 #Linea Vertical
        ws.range('B8:F' + str(fila + 1)).api.Borders(12).LineStyle = 1 #Linea Horizontal
        # Cambiar fuente
        ws.range('B9:F' + str(fila)).api.Font.Name = 'Verdana'



def ipmpat(fecha: Fecha, primeraVez, wb: xw.Book):
    print('Diligenciando Indice promedio de morosidad pat. . .')
    ws = wb.sheets['Índice promedio de morosida Pat']
    if primeraVez:
        fecha = fecha.add_months(-12)
        for i in range(13):
            # Obtener tabla
            # Obtener los archivos de la carpeta
            archivos = os.listdir(rutaRobot + '/Archivos/INFORME DEUDORES PATRONALES Y EMPRESAS')
            archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
            archivo = os.path.join(rutaRobot + '/Archivos/INFORME DEUDORES PATRONALES Y EMPRESAS', archivo)

            tabla = pd.read_csv(archivo, usecols=['SaldoTotal', 'Número Meses de Incumplimiento'], encoding='ANSI', sep=';', skiprows=3)
            saldoTotal = tabla['SaldoTotal'].sum()
            #tabla = tabla[tabla['Número Meses de Incumplimiento'] > 1]
            tabla = tabla.loc[tabla['Número Meses de Incumplimiento'] > 1]

            morosos = tabla['Número Meses de Incumplimiento'].sum()

            fecha = fecha.add_months(1)
            fecha = fecha.add_days(-1)
            mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
            ws.range('B' + str(i + 9)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
            ws.range('C' + str(i + 9)).value = morosos
            ws.range('D' + str(i + 9)).value = saldoTotal

            fecha = fecha.add_months(1)
            fecha.setDay(1)    

    else:
        archivo = rutaRobot + '/Archivos/INFORME DEUDORES PATRONALES Y EMPRESAS/INFORME DEUDORES PATRONALES Y EMPRESAS ' + fecha.as_Text() + '.csv'
        # Obtener la proxima fila disponiible en la columna B
        fila = ws.range('B8').end('down').row + 1

        fecha = fecha.add_months(1)
        fecha = fecha.add_days(-1)
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        ws.range('B' + str(fila)).value = str(fecha.dia) + '/' + str(mes) + '/' + str(fecha.anio)
        # Obtener tabla
        tabla = pd.read_csv(archivo, usecols=['SaldoTotal', 'Número Meses de Incumplimiento'], encoding='ANSI', sep=';')
        saldoTotal = tabla['SaldoTotal'].sum()
        morosos = tabla.loc[tabla['Número Meses de Incumplimiento'] > 1].sum()

        #SaldoCapital
        ws.range('C' + str(fila)).value = morosos
        #SaldoTotal
        ws.range('D' + str(fila)).value = saldoTotal

        ws.range("E9:F9").api.AutoFill(ws.range("E9:F{row}".format(row=fila)).api, 0 )

        ws.range('B8:F' + str(fila)).api.Borders(11).LineStyle = 1 #Linea Vertical
        ws.range('B8:F' + str(fila + 1)).api.Borders(12).LineStyle = 1 #Linea Horizontal









