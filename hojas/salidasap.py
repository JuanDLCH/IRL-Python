import pandas as pd
import numpy as np
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

def salidasap(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Salidas de ahorro permanente. . .')
    doc = 'INFORME INDIVIDUAL DE LAS CAPTACIONES (MODIFICADO)'

    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    ws = wb.sheets['Salida de Ahorro Permanente']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    tabla = pd.read_csv(archivo, usecols=['NIT','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    NAsociadosMesEstudio = tabla.size
   
    

    

    fechasmespasado = fecha.add_months(-1)
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fechasmespasado.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    tablaP = pd.read_csv(archivo, usecols=['NIT','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    NAsociadosMesAnterior = tablaP.size

    



    TotalAsocioadosRetirados = NAsociadosMesAnterior -NAsociadosMesEstudio



    

