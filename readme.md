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

- ### Os
  Esta libería nos permitió listar rutas y archivos del sistema, borrar y renombrar archivos y carpetas, etc. Facilitando el trabajo con archivos para clasificación y validación
- ### Sys
  Usamos esta libería principalmente para acceder a rutas temporales del sistema donde nuestro ejecutable desempaqueta el plano para trabajarlo, ya que la intención es que el usuario no tenga que lidiar dándole este archivo al robot.

# Hojas del plano a diligenciar:
  - [x] Indice promedio de morosidad 
  - [x] Indice promedio de morosidad pat
  - [x] Activos Liquidos


  - [x] R. cartera
    - [x] Consumo ventanilla
    - [x] Consumo libranza
    - [x] Comercial
    - [x] Microcredito
    - [x] Vivienda ventanilla
    - [x] Vivienda libranza
  <br><br>
  - [ ] Recaudo
    - [x] De aportes
    - [ ] De ahorro contractual
    - [ ] De ahorro permanente
    - [ ] CxC
  <br><br>
  - [ ] Salidas
    - [ ] De CDAT
    - [ ] De Ahorro contractual
    - [ ] Salidas de aportes
    - [ ] Salidas de ahorro permanente
    - [ ] Salidas fondos sociales pasivos
  <br><br>
  - [ ] Oblicaciones financieras
  - [ ] Creditos aprobados
  - [ ] Gastos administrativos
  - [ ] Recaudo y remanentes
  - [ ] CxP
  - [ ] Saldos de ahorro ordinario