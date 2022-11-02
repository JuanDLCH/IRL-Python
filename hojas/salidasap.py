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

    
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    ws = wb.sheets['Salida de Ahorro Permanente']
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    ws.range('B5').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    tabla = pd.read_csv(archivo, usecols=['NIT','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    tablaAux = tabla.drop(columns='Saldo')
 
    fechasmespasado = fecha.add_months(-1)
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fechasmespasado.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    tablaP = pd.read_csv(archivo, usecols=['NIT','Saldo'], encoding='ANSI', sep=';', skiprows=3)
    tablaAuxP = tablaP.drop(columns='Saldo')


    resultado1 = tablaAuxP[~tablaAuxP.apply(tuple,1).isin(tablaAux.apply(tuple,1))]
    resultado2 = resultado1.drop_duplicates(subset='NIT')

    tablaAuxP = tablaP
    #Filtrar en tablaAuxP los NIT que estan en resultado2
    tablaAuxP = tablaAuxP[tablaAuxP['NIT'].isin(resultado2['NIT'])]
 
    saldo = tablaAuxP.sum()['Saldo']
      
    ultimaFilaFecha = ws.range('A' + str(ws.api.UsedRange.Rows.Count)).end('up').row
    ultimaFilaSaldo = ws.range('C' + str(ws.api.UsedRange.Rows.Count)).end('up').row
    ultimaFilaAso   = ws.range('B' + str(ws.api.UsedRange.Rows.Count)).end('up').row

    ws.range('A' + str(ultimaFilaFecha + 1)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)  
    ws.range('B' + str(ultimaFilaAso + 1)).value = resultado2.size
    ws.range('C' + str(ultimaFilaSaldo + 1)).value = saldo  