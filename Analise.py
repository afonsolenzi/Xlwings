import xlwings as xw
from scipy.stats import zscore
import pandas as pd
from scipy.interpolate import interp1d
import matplotlib.pyplot as plt
import matplotlib
import numpy

#conectar com a workbook
wb = xw.Book.caller()

#usar o expand para atender N linhas
data = xw.Range('B2').expand('down').value

#converter para int
data = list(map(int, data))

#calcular o zscore para a amostra
sheet = wb.sheets['Zscore']
df = pd.DataFrame(zscore(data))
wb.sheets['Zscore'].range("C1").value = df
wb.sheets['Zscore'].range("D1").value = "Zscore"

#media e desvpad
mean = numpy.mean(data, axis=0)
sd = numpy.std(data, axis=0)

#criar uma lista com valores abaixo do limite minimo

a=[]
for x in data:
    if (x < mean - (2 *sd)):
        a.append(x)
wb.sheets['Zscore'].range("H2").value = a

#criar uma lista com valores acima do limite maximo
b=[]
for x in data:
    if (x > mean + (2 *sd)):
        b.append(x)
wb.sheets['Zscore'].range("H3").value = b

#grafico com os zscores para detectar anomalias
#entre menor que -3 e maior que +3 anomalias
chart = sheet.charts.add()
chart.set_source_data(sheet.range('D2').expand())
chart.chart_type = 'line'
chart.top = sheet.range('G10').top
chart.chart_type = 'area'





