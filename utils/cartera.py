import xlwings as xw
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
        ['144205','141205'],
        ['144105','141105'],
        ['146105','146205'],
        ['144805','nulo'],
        ['140505','nulo'],
        ['144405','nulo']
]

carpeta = 'INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)'

# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')


def obtenerTabla(fecha: Fecha, i):
    columnas = ['CodigoContable','NroCredito', 'ValorPrestamo', 'SaldoCapital', 'FechaDesembolsoInicial', 'FechaVencimiento',
                'TasaInteresEfectiva', 'AlturaCuota', 'ValorCuotaFija', 'Amortizacion']
    archivo = rutaRobot + '/Archivos/' + carpeta + '/' + carpeta + ' ' + fecha.as_Text() + '.xlsx'
    tabla = pd.read_excel(archivo, sheet_name = 'SIAC', skiprows = 3, usecols = 'C:AA')
    for col in tabla.columns:
        if col not in columnas:
            tabla.drop(col, axis = 1, inplace = True)

    tabla.query("`CodigoContable` in @codigosContables[@i]",inplace=True)
    #print(tabla)
    print(tabla)

    return tabla


def diligenciarCarteras(wb: xw.Book, fecha: Fecha):
    i = 0
    tabla = obtenerTabla(fecha, i)
    fechaAux = fecha
    for hoja in hojas:
        tablaAux = tabla
        ws = wb.sheets[hoja]