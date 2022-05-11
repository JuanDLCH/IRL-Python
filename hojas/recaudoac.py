from utils.globals import *
import xlwings as xw
import pandas as pd

doc = 'INFORME INDIVIDUAL DE LAS CAPTACIONES (MODIFICADO)'
hoja = 'Recaudo de Ahorro Contractual'
cuentas = [ 212505,212510,212515, 212520 ]


def recaudoac(fecha: Fecha, primeraVez: bool , wb: xw.Book):
    print ('Diligenciando Recaudo de Ahorro C ontractual')
    ws = wb.sheets[hoja]
    if primeraVez:
        fechaAux = fecha.add_months(-24)
        for i in range(25):
            dia = '0' + str(fechaAux.dia) if fechaAux.dia < 10 else str(fechaAux.dia)
            ws.range('A' + str(i + 12)).value = '{}/{}/{}'.format(dia, fechaAux.mes, fechaAux.anio)
            archivos = os.listdir('{}/Archivos/{}'.format(rutaRobot, doc))
            archivo = [archivo for archivo in archivos if fechaAux.as_Text() in archivo][0]
            archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot, doc), archivo)

            tabla = pd.read_csv(archivo, usecols=['CodigoContable', 'Saldo'], encoding='ANSI', sep=';', skiprows=3)
            tabla = tabla.loc[tabla['CodigoContable'].isin(cuentas)]
            suma = tabla['Saldo'].sum()

            ws.range('B' + str(i + 12)).value = suma

            fechaAux = fechaAux.add_months(1)
    else:
        print('xd')
