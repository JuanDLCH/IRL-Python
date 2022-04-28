import os
from tkinter import *
from tkinter import messagebox
import datetime
from datetime import datetime
from hojas.ipm import ipm
from utils.desviacionEstandar import desviacionEstandar
from utils.fecha import *
import xlwings as xw
from xlwings import *
from xlwings import App
from utils.validaciones import *
from hojas.cartera import diligenciarCarteras
from hojas.ipm import ipm

root = Tk()
root.withdraw()

meses = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
         'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}

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

# Ejemplo para escribir en el excel
def escribirEnElPlano():
    app = xw.app(visible=False)
    wb = app.books.open(rutaRobot + '/planoirl.xlsm')
    #wb = xw.Book(rutaRobot + '/planoirl.xlsm')
    ws = wb.sheets['Recaudo de Aportes']
    ws.range('J8:J15').value = 'ESCRIBIENDO CON PYTHON'
    wb.save(rutaRobot + '/PlanosDiligenciados/planoirl.xlsm')
    messagebox.showinfo("RobotIRL", "Se ha escrito en el planoirl.xlsm")
    


def main():
    # Iniciar un cronometro
    start_time = datetime.now()    

    mes = 'Junio'
    anio = 2021
    primeraVez = True

    fecha = Fecha(1, meses[mes.upper()], anio)
    print(fecha.as_String())
    validarCarpetas()
    clasificarArchivos()
    validarArchivos(primeraVez, fecha)
    desviacion = desviacionEstandar(fecha.as_datetime())

    # Abrir el plano
    wb = xw.Book(rutaRobot + '/planoirl.xlsm')
    planoInvisible = wb.macro('Visibility.makeInvisible') 
    planoVisible = wb.macro('Visibility.makeVisible')

    planoInvisible()
    #escribirEnElPlano()

    #To do: Diligenciar carteras
    diligenciarCarteras(wb, fecha)
    ipm(fecha, primeraVez, desviacion, wb)

    wb.save(rutaRobot + '/PlanosDiligenciados/planoirl.xlsm')

    # Terminar el cronometro
    end_time = datetime.now()
    print('Tiempo de ejecucion: {}'.format(end_time - start_time))

    messagebox.showinfo("RobotIRL", "Termine")

    wb.close()


main()
