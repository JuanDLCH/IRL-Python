from datetime import datetime
from dateutil.relativedelta import relativedelta

date_format = '%d/%m/%Y'

meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']

#crear clase
class Fecha:
    def __init__(self, dia, mes, anio):
        self.dia = dia
        self.mes = mes
        self.anio = anio

    def as_String(self):
        return f'{self.dia}/{self.mes}/{self.anio}'

    def as_Text(self):
        return str(meses[self.mes - 1]) + ' ' + str(self.anio)

    def as_datetime(self):
        return datetime(self.anio, self.mes, self.dia)

    def add_days(self, days):
        fecha = self.as_datetime() + relativedelta(days=days)
        return Fecha(fecha.day, fecha.month, fecha.year)

    def add_months(self, months):
        fecha = self.as_datetime() + relativedelta(months=months)
        return Fecha(fecha.day, fecha.month, fecha.year)

    def add_years(self, years):
        fecha = self.as_datetime() + relativedelta(years=years)
        return Fecha(fecha.day, fecha.month, fecha.year)
