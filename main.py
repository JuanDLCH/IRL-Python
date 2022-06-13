import os
import datetime
from datetime import datetime
from pandas import ExcelFile
from hojas.ipm import ipm, ipmpat
from hojas.obligacionesFinancieras import obligacionesFinancieras
from hojas.recaudoAportes import recaudoAportes
from hojas.recaudoac import recaudoac
from hojas.recaudoap import recaudoap
from hojas.salidasap import salidasap
from utils.desviacionEstandar import desviacionEstandar
from utils.fecha import *
import xlwings as xw
from xlwings import *
from xlwings import App
from utils.validaciones import *
from hojas.cartera import diligenciarCarteras
from hojas.ipm import ipm
from hojas.activosLiquidos import activosLiquidos
from hojas.cxc import cxc
from hojas.salidas import salidaCdatyAC
from hojas.creditosAprobados import creditosAprobados
from hojas.gastosAdministrativos import gastosAdministrativos
from hojas.salidasfsp import salidasfsp

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

# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')
    


def main(mes, anio, primeraVez):
    # Iniciar un cronometro
    start_time = datetime.now()  

    path = resource_path('planoirl.xlsm')

    if primeraVez:
        plano = ExcelFile(path)
    else:
        mesPlano = meses.index(mes.upper())
        plano = ExcelFile(buscarPlano(mesPlano, anio))

    fecha = Fecha(1, meses.index(mes.upper()) + 1, anio)
    print(fecha.as_String())
    validarCarpetas()
    clasificarArchivos()
    validarArchivos(primeraVez, fecha)
    desviacion = desviacionEstandar(fecha.as_datetime())

    # Abrir el plano
    print('Abriendo plano. . .')

    try:
        wb = xw.Book(plano)
    except Exception as e:
        print(e)
        print('No se pudo abrir el plano, es posible que ya haya uno con el mismo nombre abierto, cierrelo e intentelo de nuevo')
        # Ejecutar ui.py
        os.system('python ui.py')
        exit()
        
    print('Iniciando sesion. . .')
    planoInvisible = wb.macro('Visibility.makeInvisible') 
    planoVisible = wb.macro('Visibility.makeVisible')

    #planoInvisible()

    #To do: Diligenciar carteras
    diligenciarCarteras(wb, fecha)
    ipm(fecha, primeraVez, desviacion, wb)
    ipmpat(fecha, primeraVez, wb)
    activosLiquidos(fecha, primeraVez, wb)
    recaudoAportes(fecha, primeraVez, wb)
    recaudoac(fecha, primeraVez, wb)
    recaudoap(fecha,primeraVez,wb)
    cxc(fecha, primeraVez, wb)
    salidaCdatyAC(fecha, primeraVez, wb)
    obligacionesFinancieras(fecha, primeraVez, wb)
    creditosAprobados(fecha, wb)
    gastosAdministrativos(fecha, primeraVez, wb)
    salidasfsp(fecha,wb)
    salidasap(fecha,wb)
    wb.api.RefreshAll()
    print('Guardando plano. . .')
    try:
        wb.save(rutaRobot + '/PlanosDiligenciados/PLANOIRL {} {}.xlsm'.format(mes, anio))
    except:
        print('Error al guardar el plano, es posible que ya haya un plano con este nombre abierto, por favor cierrelo.')
    # Terminar el cronometro
    end_time = datetime.now()
    print('Tiempo de ejecucion: {}'.format(end_time - start_time))

    wb = xw.Book(rutaRobot + '/PlanosDiligenciados/PLANOIRL {} {}.xlsm'.format(mes, anio))
    planoVisible = wb.macro('Visibility.makeVisible')

    planoVisible()