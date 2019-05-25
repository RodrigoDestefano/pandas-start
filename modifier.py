import pandas as pd
import numpy as np
from openpyxl import Workbook
from itertools import repeat

def getVariables(v,s):
    return v.append(s.split('|')[0])

def getPeriods(s):
    #Less the value in ASCII
    return ord(s[6])-48

def getFColumns(l):
    aux = [*l]
    aux.insert(0,"Year")
    aux.insert(0,"Period")
    aux.insert(0,"ID")
    aux.insert(0,"Company")
    return aux 

def getFCompanys(p,y,c):
    return [x for item in c for x in repeat(item, p*len(y))]

def getFIDs(p,y,c):
    aux = []
    cont = 1
    aux2 = np.empty(p*len(y))

    for i in range(len(c)):
        aux2.fill(cont)
        aux.extend(aux2)
        cont += 1
    
    return aux 

def getFPeriods(p,y,c):
    aux = []
    
    for i in range(len(c)*len(y)):    
        aux.extend(range(1,p+1))

    return aux

def getFYears(p,y,c):
    aux = []
    y = [x for item in y for x in repeat(item, p)]
    
    for i in range(len(c)):
        aux.extend(y)
        i += i
    
    return aux

def getFinal(n,d,c,p,y,v,w,t):
    aux = []
    for i in range(len(c)):
        for j in range(len(p*y)):    
            aux.append(d.ix[j+t,c[i]])
            j += 1
        i += 1
    
    n[v[w]] = aux
    return 

#Get the .xlsx file from ECONOMATICA
df = pd.read_excel("./source/data.xlsx", "Sheet1")

#Extracting the companys
companys = []
companys.extend(df.columns)
companys.pop(0)
companys.pop(0)
companys.pop(0)

#Extracting the variables
variables = []

mixVariables = list(dict.fromkeys(df["Variáveis"].values))
#Deleting the last elemente in the list (nan)
del mixVariables[-1]

for i in mixVariables:
    getVariables(variables, i)

variables = list(dict.fromkeys(variables))

#Extracting the periods
periods = getPeriods(df.columns[0])

#Extracting the years
years = []
years = list(dict.fromkeys(df.fillna(0)[df.columns[0]].values))
years.remove(0)

#Creating the new .xlsx file 
FColumns = getFColumns(variables)
new = pd.DataFrame(columns = FColumns)

new[FColumns[0]] = getFCompanys(periods,years,companys)
new[FColumns[1]] = getFIDs(periods,years,companys)
new[FColumns[2]] = getFPeriods(periods,years,companys)
new[FColumns[3]] = getFYears(periods,years,companys)

init = 0
whatVar = 0
for i in range(len(variables)):
    getFinal(new,df,companys,periods,years,variables,whatVar,init)
    init += periods*len(years)
    whatVar += 1

new.to_excel("./source/new.xlsx", sheet_name="Final", index=False)
