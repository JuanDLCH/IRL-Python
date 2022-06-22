from re import A
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

    no = []

    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    ws = wb.sheets['Salida de Ahorro Permanente']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    tabla = pd.read_csv(archivo, usecols=['NIT','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    tablaAux = tabla.drop(columns='Saldo')
    NAsociadosMesEstudio = tabla.size
    fechasmespasado = fecha.add_months(-1)
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fechasmespasado.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    tablaP = pd.read_csv(archivo, usecols=['NIT','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    tablaAuxP = tablaP.drop(columns='Saldo')

    #hola = tabla.merge(tablaP,indicator=True, how='left').loc[lambda v : v['_merge'] !=  'both' and ]
    #print(hola)

    resultado1 = tablaAuxP[~tablaAuxP.apply(tuple,1).isin(tablaAux.apply(tuple,1))]
    resultado2 = resultado1.drop_duplicates(subset='NIT')

    #saldo = resultado2[tabla['Saldo'].sum()]
    #totalretirados = resultado1.size
    print(resultado2.size)
    print(resultado2)
   # print(saldo)
    #TotalAsocioadosRetirados = NAsociadosMesAnterior -NAsociadosMesEstudio
    NAsociadosMesAnterior = tablaP.size
    TotalAsocioadosRetirados = NAsociadosMesAnterior -NAsociadosMesEstudio