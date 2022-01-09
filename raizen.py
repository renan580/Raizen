import xml.etree.ElementTree as ET
from datetime import datetime as dt
import pandas as pd
import win32com.client as win32
from pathlib import Path
import zipfile

#Convert xls file to xlsx
file = str(Path("vendas-combustiveis-m3.xls").resolve())
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(file)
wb.SaveAs(file+'x', FileFormat = 51)
wb.Close()                               
excel.Application.Quit()

#Unzip xlsx file
file = str(Path("vendas-combustiveis-m3.xlsx").resolve())
ext = zipfile.ZipFile(file) 
ext.extractall()

#Parsing xml files
meses = []
combustivel = []
combustivel2 = []
ano = []
ano2 = []
regiao = []
estado = []
data = {}
data2 = {}
cami = str(Path().resolve())
file = cami +'/xl/pivotCache/pivotCacheDefinition1.xml'
tree = ET.parse(file)
root = tree.getroot()
for i in range(5, 17):
    meses.append(i-4)
for i in range(0,8):
    combustivel.append(root[1][0][0][i].get('v'))
for i in range(0,21):
    ano.append(root[1][1][0][i].get('v'))
for i in range(0,5):
    regiao.append(root[1][2][0][i].get('v'))
for i in range(0,27):
    estado.append(root[1][3][0][i].get('v'))
file = cami +'/xl/pivotCache/pivotCacheDefinition2.xml'
tree = ET.parse(file)
root = tree.getroot()
for i in range(0, 8):
    ano2.append(root[1][1][0][i].get('v'))
for i in range(0,5):
    combustivel2.append(root[1][0][0][i].get('v'))
file = cami +'/xl/pivotCache/pivotCacheRecords1.xml'
tree = ET.parse(file)
root = tree.getroot()
count = 0
vector = []
for child in root:
    for gchild in child:
        vector.append(gchild.get('v'))
    vector[len(vector)-1] = (str(dt.now()))
    vector[0] = combustivel[int(vector[0])]
    vector[1] = ano[int(vector[1])]
    vector[2] = regiao[int(vector[2])]
    vector[3] = estado[int(vector[3])]
    data[count] = []
    data[count].append(vector)
    vector = []
    count = count + 1
file = cami +'/xl/pivotCache/pivotCacheRecords2.xml'
tree = ET.parse(file)
root = tree.getroot()
count = 0
vector = []
for child in root:
    for gchild in child:
        vector.append(gchild.get('v'))
    vector[len(vector)-1] = (str(dt.now()))
    vector[0] = combustivel2[int(vector[0])]
    vector[1] = ano2[int(vector[1])]
    vector[2] = regiao[int(vector[2])]
    vector[3] = estado[int(vector[3])]
    data2[count] = []
    data2[count].append(vector)
    vector = []
    count = count + 1

#Data Structuring
def datastructuring (x):
    year_month = []
    uf = []
    product = []
    unit = []
    volume = []
    created_at = []
    count = 0
    for mes in meses:
        while count < len(x):
            d = dt(int(x[count][0][1]),mes, 1)
            #d = d.strftime('%b/%Y')
            year_month.append(d)
            uf.append(str(x[count][0][3]))
            product.append(str(x[count][0][0]))
            unit.append(str(x[count][0][4]))
            try:
                volume.append(float(x[count][0][mes+4]))
            except:
                volume.append(0)
            created_at.append(x[count][0][17])
            count = count + 1
        count = 0
    return(year_month, uf, product, unit, volume, created_at)

year_month = datastructuring(data)[0]
uf = datastructuring(data)[1]
product = datastructuring(data)[2]
unit = datastructuring(data)[3]
volume = datastructuring(data)[4]
created_at = datastructuring(data)[5]

year_month2 = datastructuring(data2)[0]
uf2 = datastructuring(data2)[1]
product2 = datastructuring(data2)[2]
unit2 = datastructuring(data2)[3]
volume2 = datastructuring(data2)[4]
created_at2 = datastructuring(data2)[5]

#Data Storing
with pd.option_context('display.float_format', '{:0.13f}'.format, 'display.max_columns', 6):
    df = pd.DataFrame({'year_month':year_month, 'uf':uf, 'product':product, 'unit':unit, 'volume':volume, 'created_at':created_at})
    df.to_csv('Colected Data.csv', sep=';', index=False, encoding='utf-8', float_format='%.13f', date_format='%b/%Y')
    df2 = pd.DataFrame({'year_month':year_month2, 'uf':uf2, 'product':product2, 'unit':unit2, 'volume':volume2, 'created_at':created_at2})
    df2.to_csv('Colected Data 2.csv', sep=';', index=False, encoding='utf-8', float_format='%.13f', date_format='%b/%Y')
    print(df.head(10))
    print(df.dtypes)
    print(df2.head(10))
    print(df2.dtypes)