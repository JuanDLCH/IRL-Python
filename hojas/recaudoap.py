from utils.globals import *
import xlwings as xw
import pandas as pd
from utils.fecha import Fecha



doc = 'INFORME INDIVIDUAL DE LAS CAPTACIONES (MODIFICADO)'
hoja = 'Recaudo de Ahorro Permanente'
cuentas = [213005, 213010, 213015, 213020]

def recaudoap(fecha: Fecha, primeraVez: bool , wb: xw.Book):
    print ('Diligenciando Recaudo de Ahorro Permanente')
    archivos = os.listdir('{}/Archivos/{}'.format(rutaRobot, doc))
    ws = wb.sheets[hoja]
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    ws.range('B5').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia, mes, fecha.anio)
    if primeraVez:
        fechaAux = fecha.add_months(-24)

        for i in range(25):
            dia = fechaAux.add_months(1).add_days(-1).dia
            mes = '0' + str(fechaAux.mes) if fechaAux.mes < 10 else str(fechaAux.mes)
            ws.range('A' + str(i + 11)).value = '{}/{}/{}'.format(dia, mes, fechaAux.anio)
            archivo = [archivo for archivo in archivos if fechaAux.as_Text() in archivo][0]
            archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot, doc), archivo)

            tabla = pd.read_csv(archivo, usecols=['CodigoContable', 'Saldo'], encoding='ANSI', sep=';', skiprows=3)
            tabla = tabla.loc[tabla['CodigoContable'].isin(cuentas)]
            suma = tabla['Saldo'].sum()

            ws.range('B' + str(i + 11)).value = suma

            fechaAux = fechaAux.add_months(1)
    else:
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        dia = fecha.add_months(1).add_days(-1).dia
        ultimaFila = ws.range('A' + str(ws.api.UsedRange.Rows.Count)).end('up').row
        ws.range('A' + str(ultimaFila + 1)).value = '{}/{}/{}'.format(fecha.dia, mes, fecha.anio)
        archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
        archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot, doc), archivo)

        tabla = pd.read_csv(archivo, usecols=['CodigoContable', 'Saldo'], encoding='ANSI', sep=';', skiprows=3)
        tabla = tabla.loc[tabla['CodigoContable'].isin(cuentas)]
        suma = tabla['Saldo'].sum()

        ws.range('B' + str(ultimaFila + 1)).value = suma

        # Arrastrar formulas de la fila anterior
        ws.range("A{row}:I{row}".format(row = ultimaFila)).api.AutoFill(ws.range("A{row}:F{row}".format(row=ultimaFila + 1)).api, 0 )

        # Formato de celdas
        ws.range('A12:I' + str(ultimaFila + 1)).api.Borders(11).LineStyle = 1 #Linea Vertical
        ws.range('A12:I' + str(ultimaFila + 2)).api.Borders(12).LineStyle = 1 #Linea Horizontal
        
        # Cambiar fuente
        ws.range('A12:I' + str(ultimaFila + 1)).api.Font.Name = 'Verdana'