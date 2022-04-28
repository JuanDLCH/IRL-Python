import xlwings as xw
from xlwings import *
from utils.fecha import *
import openpyxl
import os
import pandas as pd

hojas = [
    'R. cartera consumo Ventanilla',
    'R. cartera consumo Libranza',
    'R. cartera Comercial',
    'R. cartera Microcrédito',
    'R. cartera vivienda Ventanilla',
    'R. cartera vivienda Libranza'
]

codigosContables = [
    [144205, 141205],
    [144105, 141105],
    [146105, 146205],
    [144805],
    [140505],
    [144405]
]

carpeta = 'INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)'

# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')


def obtenerTabla(fecha: Fecha):
    columnas = ['CodigoContable', 'NroCredito', 'ValorPrestamo', 'SaldoCapital', 'FechaDesembolsoInicial', 'FechaVencimiento',
                'TasaInteresEfectiva', 'AlturaCuota', 'ValorCuotaFija', 'Amortizacion']

    archivos = os.listdir(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)')
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)', archivo)

    tabla = pd.read_csv(archivo, usecols=columnas)

    #tabla = tabla[(tabla['fechaultimopago'] >= fecha.as_datetime()) & (
    #    tabla['fechaultimopago'] < fecha.add_months(1).as_datetime())]
    for col in tabla.columns:
        if col not in columnas:
            tabla.drop(col, axis=1, inplace=True)

    tabla = tabla[columnas]
    return tabla


def filtrarTabla(tabla, fecha: Fecha, hoja):

    _tabla = tabla[tabla['CodigoContable'].isin(
        codigosContables[hojas.index(hoja)])]
    _tabla.drop('CodigoContable', axis=1, inplace=True)

    return _tabla


def diligenciarCarteras(wb: xw.Book, fecha: Fecha):
    tabla = obtenerTabla(fecha)
    fechaAux = fecha
    for hoja in hojas:
        tablaAux = filtrarTabla(tabla, fechaAux, hoja)
        ws = wb.sheets[hoja]
        # Eliminar datos desde A9 hasta la última fila
        ws.range('A9:I' + str(ws.range('J9').end('down').row)).clear()
        # Insertar datos sin encabezado
        ws.range('A9').value = tablaAux.values

        if not tablaAux.empty:
            # Poner bordes a los datos insertados
            last_row = ws.range('A9').end('down').row
            ws.range('A9:P' + str(last_row)).api.Borders(11).LineStyle = 1 #Linea Vertical
            ws.range('A9:O' + str(last_row + 1)).api.Borders(12).LineStyle = 1 #Linea Horizontal
            # Rellenar las formulas
            ws.range("J9:O9").api.AutoFill(ws.range("J9:O{row}".format(row=last_row)).api, 0 )



        

