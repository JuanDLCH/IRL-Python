from ctypes import wstring_at
import os
from tkinter import *
from tkinter import messagebox
import datetime
from datetime import datetime
from dateutil.relativedelta import relativedelta
from itsdangerous import want_bytes
from utils.desviacionEstandar import descargarArchivo
from utils.fecha import *
import xlwings as xw
from xlwings import *
from utils.validaciones import *
from hojas.cartera import diligenciarCarteras

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
    wb = xw.Book(rutaRobot + '/planoirl.xlsm')
    ws = wb.sheets['Recaudo de Aportes']
    ws.range('J8:J15').value = 'ESCRIBIENDO CON PYTHON'
    wb.save(rutaRobot + '/PlanosDiligenciados/planoirl.xlsm')
    messagebox.showinfo("RobotIRL", "Se ha escrito en el planoirl.xlsm")
    wb.close()
    


def main():
    descargarArchivo()
    messagebox.showerror("RobotIRL", "Se ha descargado el archivo")
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

    # Abrir el plano
    wb = xw.Book(rutaRobot + '/planoirl.xlsm')
    #escribirEnElPlano()

    #To do: Diligenciar carteras
    diligenciarCarteras(wb, fecha)

    wb.save(rutaRobot + '/PlanosDiligenciados/planoirl.xlsm')

    # Terminar el cronometro
    end_time = datetime.now()
    print('Tiempo de ejecucion: {}'.format(end_time - start_time))

    messagebox.showinfo("RobotIRL", "Termine")


main()
