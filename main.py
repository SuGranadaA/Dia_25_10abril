#Importamos la librería para abrir el archivo excel
import openpyxl

#Importamos la librería para pandas
import pandas as pd

#Leemos el archivo en la Hoja1
num1 = pd.read_excel("panda1.xlsx", sheet_name=0)

#Leemos laas filas específicas del documento a partir de la fila 2
array1 = pd.Series(num1.iloc[0, :])
array2 = pd.Series(num1.iloc[1, :])

#Imprimimos los datos obtenidos
print("\nLa lectura del archivo es: \n", num1, "\n")
print("\nLa fila 2 es: \n", array1, "\n")
print("\nLa fila 3 es: \n", array2, "\n")

#Creamos el documento para guardar los resultados y asignamos la hoja
resultado1 = openpyxl.Workbook()
hoja = resultado1.active

#Guardamos el archivo
resultado1.save('res1.xlsx')

#Establecemos la función para identificar el tamaño de la fila
sizearray1 = array1.size
print("\nEl tamaño de la fila es: \n", sizearray1, "\n")

#Establecemos la función para identificar la desviación típica de los datos de la serie numérica
stdarray2 = array2.std()
print("\nLa desviación típica de los datos de la serie numérica: \n",
      stdarray2, "\n")

#Establecemos la función para identificar el dato menor de una fila
minarray1 = array1.min()
print("\nEl dato con menor valor de la fila 2 es: \n", minarray1, "\n")

#Establecemos la función para identificar el dato mayor de una fila
maxarray2 = array2.max()
print("\nEl dato con mayor valor de la fila 3 es: \n", maxarray2, "\n")

#Escribimos los resultados en el nuevo documento
hoja['B2'] = ("Tamaño de la segunda fila: ")
hoja['C2'] = sizearray1
hoja['B3'] = ("Desviación típica de datos de la serie numérica: ")
hoja['C3'] = stdarray2
hoja['B4'] = ("El minimo de los datos: ")
hoja['C4'] = minarray1
hoja['B5'] = ("El maximo de los datos: ")
hoja['C5'] = maxarray2

#Guardamos el documento
resultado1.save('res1.xlsx')

#Leemos el archivo en la Hoja1
num2 = pd.read_excel("panda2.xlsx", sheet_name=0)

#Leemos laas filas específicas del documento a partir de la fila 2
array3 = pd.Series(num2.iloc[0, :])
array4 = pd.Series(num2.iloc[1, :])

#Imprimimos los datos obtenidos
print("\nLa lectra del archivo es: \n", num2, "\n")
print("\nLa fila 2 es: \n", array3, "\n")
print("\nLa fila 3 es: \n", array4, "\n")

#Creamos el documento para guardar los resultados y asignamos la hoja
resultado2 = openpyxl.Workbook()
hojita = resultado2.active

#Guardamos el archivo
resultado2.save('res2.xlsx')

#Establecemos la función para identificar la media de datos numéricos de la serie
meanarray3 = array3.mean()
print("\nLa media de la fila 2 es: \n", meanarray3, "\n")

#Establecemos la función para concatenar strings y sumar números en una fila
sumaarray3 = array3.sum()
print("\nLa suma de los datos de la fila 2 es: \n", sumaarray3, "\n")
sumaarray4 = array4.sum()
print("\nLa suma de los datos de la fila 4 es: \n", sumaarray4, "\n")

#Establecemos la función para identificar la cantidad de elementos no nulos
countarray4 = array4.count()
print("\nLa cantidad de elementos no nulos es: \n", countarray4, "\n")

#Escribimos los resultados en el nuevo documento
hojita['B1'] = ("Media de datos de fila 2: ")
hojita['C1'] = meanarray3
hojita['B2'] = ("Suma de datos fila 2: ")
hojita['C2'] = sumaarray3
hojita['B3'] = ("Suma o concatenación de datos de fila 3: ")
hojita['C3'] = str(sumaarray4)
hojita['B4'] = ("Cantidad de elementos no nulos: ")
hojita['C4'] = (countarray4)

#Guardamos el documento
resultado2.save('res2.xlsx')
