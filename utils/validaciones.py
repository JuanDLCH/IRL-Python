from utils.convert2csv import convertiracsv
from utils.fecha import *
import os
from tkinter import *
from tkinter import messagebox
import pandas as pd
from utils.convert2csv import convertiracsv
from utils.globals import *
import wget
import urllib


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

def crearCarpetas():
    os.mkdir(rutaRobot)
    os.mkdir(rutaRobot + '/' + 'Archivos')
    for i in carpetas:
        os.mkdir(rutaRobot + '/' + 'Archivos' + '/' + i)
    os.mkdir(rutaRobot + '/' + 'PlanosDiligenciados')
    os.mkdir(rutaRobot + '/' + 'ArchivosNuevos')
    found = False
    url = 'https://www.supersolidaria.gov.co/sites/default/files/public/data/desviacion_estandar_'
    fechaActual = Fecha(1, datetime.now().month, datetime.now().year)
    while not found:
        try:
            if os.path.exists(rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx'):
                found = True
            else:
                found = descargarArchivo(url + fechaActual.as_Text().lower().replace(' ', '_') + '.xlsx', rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx')
                
        except:
            pass
        fechaActual = fechaActual.add_months(-1)

def validarCarpetas():
    print('Validando carpetas. . .')
    # Si no existe la carpeta del robot, crearla
    if not os.path.exists(rutaRobot):
        print('Creando carpeta y archivos del robot. . .')
        crearCarpetas()

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
    
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')

    print('Clasificando los archivos. . .')
    # Clasificar los archivos
    for carpeta in carpetas:
        for archivo in archivos:
            if archivo.startswith(carpeta):
                for mes in meses:
                    if mes in archivo.upper():
                        if os.path.exists(rutaRobot + '/Archivos/' + carpeta + '/' + archivo):
                            os.remove(rutaRobot + '/Archivos/' + carpeta + '/' + archivo)
                            
                        os.rename(rutaRobot + '/ArchivosNuevos/' + archivo,
                                rutaRobot + '/Archivos/' + carpeta + '/' + archivo)
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
    print('Validando archivos para diligenciamiento de ' + str(fecha.as_Text()) + '. . .')
    for carpeta in carpetas:
        archivos = [os.path.splitext(filename)[0] for filename in os.listdir(rutaRobot + '/Archivos/' + carpeta)]
        repeticiones = archivosPrimeraVez[carpetas.index(carpeta)] if primeraVez else archivosSegundaVez[carpetas.index(carpeta)]
        repeticiones -= 1
        auxFecha = fecha.add_months(-repeticiones)
        for j in range(repeticiones):
            if primeraVez:
                nombreArchivo = carpeta + ' ' + auxFecha.as_Text()
                if nombreArchivo not in archivos:
                    messagebox.showinfo("RobotIRL", "No pude encontrar  el archivo " + nombreArchivo +
                                        " te abriré la carpeta de archivos para que lo muevas o lo busques y corrijas.")
                    os.system('explorer ' + rutaRobot)
                    exit()
                else:
                    auxFecha = auxFecha.add_months(1)

            else:
                auxFecha = fecha.add_months(
                    archivosSegundaVez[carpetas.index(carpeta)])