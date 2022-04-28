from unittest import case
import pandas as pd
import os

# Obtener ruta de la carpeta documentos
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

def convertiracsv():
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')
    for archivo in archivos:
        if 'xlsx' or 'xls' or 'XLS' or 'XLSX' in archivo:
            df = pd.read_excel(rutaRobot + '/ArchivosNuevos/' + archivo, sheet_name='SIAC', skiprows=3, engine='xlwt')
            os.remove(rutaRobot + '/ArchivosNuevos/' + archivo)
            nombre = archivo.split('.')[0]
            df.to_csv(rutaRobot + '/ArchivosNuevos/' + nombre + '.csv', index=False)



