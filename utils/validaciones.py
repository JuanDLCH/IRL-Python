from utils.convert2csv import convertiracsv
from utils.fecha import *
import os
from tkinter import *
from tkinter import messagebox
from xlsx2csv import Xlsx2csv
import pandas as pd
from utils.convert2csv import convertiracsv

root = Tk()
root.withdraw()

rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')

carpetas = [
    'CATALOGO DE CUENTAS',
    'CREDITOS DE BANCOS Y OTRAS OBLIGACIONES FINANCIERAS (NUEVO)',
    'INFORME DEUDORES PATRONALES Y EMPRESAS',
    'INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES',
    'INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)',
    'INFORME INDIVIDUAL DE LAS CAPTACIONES (MODIFICADO)',
    'RELACION DE INVERSIONES',
    'SALDOS DIARIOS DE AHORRO'
]

archivosPrimeraVez = [13, 1, 13, 25, 13, 25, 1, 1]

archivosSegundaVez = [1 , 1, 1 , 2 , 1 , 2, 1, 1]

meses = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
         'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}


def validarCarpetas():
    # Si no existe la carpeta del robot, crearla
    if not os.path.exists(rutaRobot):
        os.mkdir(rutaRobot)
        os.mkdir(rutaRobot + '/' + 'Archivos')
        for i in carpetas:
            os.mkdir(rutaRobot + '/' + 'Archivos' + '/' + i)
        os.mkdir(rutaRobot + '/' + 'PlanosDiligenciados')
        os.mkdir(rutaRobot + '/' + 'ArchivosNuevos')

        # Abrir el explorador de archivos en la carpeta del robot
        os.system('explorer ' + rutaRobot + '\ArchivosNuevos')

        # MessageBox
        messagebox.showinfo(
            'RobotIRL', 'Parece que no hay archivos para la ejecución, he abierto la carpeta por ti, deja los archivos aqui y vuelve a ejecutar el robot.')
        exit()
    else:
        clasificarArchivos()


def clasificarArchivos():
    # Obtener archivos de la carpeta ArchivosNuevos
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')

    if len(archivos) != 0:
        convertiracsv()


    # Clasificar los archivos
    for i in carpetas:
        for j in archivos:
            if j.startswith(i):
                for l in meses.keys():
                    if l in j.upper():
                        if os.path.exists(rutaRobot + '/Archivos/' + i + '/' + j):
                            os.remove(rutaRobot + '/Archivos/' + i + '/' + j)
                            
                        os.rename(rutaRobot + '/ArchivosNuevos/' + j,
                                rutaRobot + '/Archivos/' + i + '/' + j)
                        break

    # Si quedaron archivos en la carpeta ArchivosNuevos
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')
    if len(archivos) != 0:
        # MessageBox
        messagebox.showinfo(
            'RobotIRL', 'Ups, algunos archivos no pudieron clasificarse, voy a mostrartelos, verifica sus nombres y vuelve a ejecutar el robot.')
        # Abrir el explorador de archivos en la carpeta ArchivosNuevos
        os.system('explorer ' + rutaRobot + '\ArchivosNuevos')
        exit()


def validarArchivos(primeraVez, fecha: Fecha):
    auxFecha = fecha
    for i in carpetas:
        archivos = [os.path.splitext(filename)[0] for filename in os.listdir(rutaRobot + '/Archivos/' + i)]
        repeticiones = archivosPrimeraVez[carpetas.index(i)] if primeraVez else archivosSegundaVez[carpetas.index(i)]
        repeticiones -= 1
        auxFecha = fecha.add_months(-repeticiones)
        for j in range(repeticiones):
            if primeraVez:
                nombreArchivo = i + ' ' + auxFecha.as_Text()
                if nombreArchivo not in archivos:
                    messagebox.showinfo("RobotIRL", "No pude encontrar  el archivo " + nombreArchivo +
                                        " te abriré la carpeta de archivos para que lo muevas o lo busques y corrijas.")
                    os.system('explorer ' + rutaRobot)
                    exit()
                else:
                    auxFecha = auxFecha.add_months(1)

            else:
                auxFecha = fecha.add_months(
                    archivosSegundaVez[carpetas.index(i)])