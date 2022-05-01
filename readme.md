# Automatizando el diligenciamiento del Plano IRL

Todas las cooperativas en Medellín tienen que presentar un informe de índice de riesgo de liquidez, existe un documento con macros en excel para calcular este índice y enviarlo a las entidades pertinentes automáticamente, pero su llenado es engorroso y poco eficiente, tomandole mucho tiempo a los encargados de esto su trabajo.

Realizamos un algoritmo en Python que automatiza todo este proceso, haciendo que pase de horas e incluso días a tan sólo unos cuántos segundos.

Inicialmente intentamos realizar esto en UiPath, pero el proceso se hace tan largo y grande que su codificación se volvió tediosa, y a falta de la mitad del trabajo no lográbamos ubicarnos en las diferentes carpetas que teníamos creadas, y al movernos entre ellas teníamos problemas de rendimiento del equipo ya que la interfaz de UiPath es bastante pesada y lenta.

Dándonos cuenta de que no había interacción con la UI y era sólo trabajo de Lectura y escritura, procedimos a hacer esto de otra manera.

# Proceso de automatización:

## Librerías:
Nos apoyamos de algunas liberías disponibles para Python con el fin de procesar los datos de la manera más eficiente posible. Algunas de estas son:

- ### Pandas:
  Inicialmente utilizamos Pandas para leer los archivos de excel y convertirlos a csv (Para una lectura más eficiente y rápida), luego, tratando los archivos como dataframes, fue más sencillo realizar trabajos de cálculo y filtrado de los datos que suelen tener un volúmen considerable.

  ```python
    pip install pandas
  ```
  - ### XLWings y Openpyxl
  Apoyándonos en estas 2 librerías trabajamos lo que fue la edición del Plano, XLWings se encargó de la mayoría del trabajo, Openpyxl a pesar de tener funciones muy similares, no es tan amigable con archivos que tienen macros, como el plano, por lo que únicamente la usamos como motor para la lectura de algunos archivos viejos (XLS) que nos suelen estar empaquetados.

   ```python
    pip install xlwings
    pip install openpyxl
  ```

- ### PyQt5
  Usamos PyQt5 para integrar la interfaz de usuario diseñada previamente con Qt Designer.
  
  ```python
  pip install pyqt5
  ```
  
- ### Os
  Esta libería nos permitió listar rutas y archivos del sistema, borrar y renombrar archivos y carpetas, etc. Facilitando el trabajo con archivos para clasificación y validación
- ### Sys
  Usamos esta libería principalmente para acceder a rutas temporales del sistema donde nuestro ejecutable desempaqueta el plano para trabajarlo, ya que la intención es que el usuario no tenga que lidiar dándole este archivo al robot.

# Entradas y Salidas
El robot va a recibir los archivos necesarios para poder trabajar, estos son conocidos por cada cooperativa aunque también serán indicados dentro del manual de usuario. Estos deberán tener un formato de nombre que será 'ARCHIVO MES AÑO' (Ejemplo: 'CATALOGO DE CUENTAS JUNIO 2021.csv')

Los archivos deberán estar en formato csv (También son permitidos archivos de excel como xls o xlsx, pero con estos el proceso será más tardado dado que el robot igual los convierte a csv antes de iniciar el trabajo). 

El numero de archivos variará dependiendo de si es primera vez que se diligencia o no. Como salida tendremos el plano irl completamente diligenciado y listo para su envío.

# Interfaz de usuario
El robot únicamente necesita por entrada que el usuario le indique el mes, año y si es primera vez que va a diligenciar, por tanto la interfaz no deberá ser muy compleja para ser manipulada, esta la hicimos utilizando Qt Designer y la implementamos mediante a librería PyQt5, recibirá Mes, Año, Primera vez y tendrá 2 opciones extra, una para abrir la carpeta contenedora de los archivos y otra para abrir el manual de usuario del robot, y será algo como esto.

![Interfaz](https://imgur.com/R4yxNFo)



# Hojas del plano a diligenciar:
  - Indice promedio de morosidad &check;
  - Indice promedio de morosidad pat &check;
  - Activos Liquidos &check;


  - R. cartera &check;
    - Consumo ventanilla &check;
    - Consumo libranza &check;
    - Comercial &check;
    - Microcredito &check;
    - Vivienda ventanilla &check;
    - Vivienda libranza &check;
  <br><br>
  - Recaudo &cross;
    - De aportes &check;
    - De ahorro contractual &cross;
    - De ahorro permanente &cross;
    - CxC &cross;
  <br><br>
  - Salidas &cross;
    - De CDAT &cross;
    - De Ahorro contractual &cross;
    - Salidas de aportes &cross;
    - Salidas de ahorro permanente &cross;
    - Salidas fondos sociales pasivos &cross;
  <br><br>
  - Oblicaciones financieras &cross;
  - Creditos aprobados &cross;
  - Gastos administrativos &cross;
  - Recaudo y remanentes &cross;
  - CxP &cross;
  - Saldos de ahorro ordinario &cross;
