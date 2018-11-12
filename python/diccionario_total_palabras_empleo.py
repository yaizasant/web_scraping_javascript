#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from nltk.corpus import stopwords

# obtenemos las stopwords del español
stopwords = stopwords.words('spanish')

# paso a mayúsculas los stopwords
for i in range (len(stopwords)):
    stopwords[i] = stopwords[i].upper()

# obtengo en un dataframe las experiencias introducidas en octubre de 2018
data = pd.read_excel('C:\Users\ysantana\Desktop\experiencias_octubre_2018.xls', sheetname='Exportar Hoja de Trabajo')

# obtengo en un dataframe las formaciones introducidas en octubre de 2018
data_f = pd.read_excel('C:\Users\ysantana\Desktop\estudios_octubre_2018.xls', sheetname='Exportar Hoja de Trabajo')

# obtengo en un dataframe los conocimientos introducidas en octubre de 2018
data_c = pd.read_excel('C:\Users\ysantana\Desktop\conocimientos_octubre_2018.xls', sheetname='Exportar Hoja de Trabajo')

# obtengo en un dataframe los códigos personales
codigos_personales = pd.read_excel('C:\Users\ysantana\Desktop\codigos_personales.xls', sheetname='Exportar Hoja de Trabajo')

# saber el número de filas del excel experiencias
tamanyo_dicc=len(data)
print("El fichero experiencias tiene:")
print(tamanyo_dicc)

# saber el número de filas del excel formacion
tamanyo_dicc_f=len(data_f)
print("El fichero formacion tiene:")
print(tamanyo_dicc_f)

# saber el número de filas del excel conocimientos
tamanyo_dicc_c=len(data_c)
print("El fichero conocimientos tiene:")
print(tamanyo_dicc_c)

# vemos el tipo de objeto que es data (dataframe)
print(type(data))

# vemos el tipo de objeto que es data_f (dataframe)
print(type(data_f))

# vemos el tipo de objeto que es data_c (dataframe)
print(type(data_c))

# vemos las columnas que existen, para comprobar que se leyó correctamente
print("Nombre de las columnas del fichero experiencias:")
print(data.columns)
print("Nombre de las columnas del fichero formación:")
print(data_f.columns)
print("Nombre de las columnas del fichero conocimientos:")
print(data_c.columns)


# Declaración de un diccionario
diccionario = dict()

# declaramos la cadena de palabras para contabilizar la frecuencia de palabras
cadenaPalabras=''

#recorremos cada fila de experiencia
for i in range(tamanyo_dicc):
   if (data['TIPO_TRABAJO'][i].encode("utf-8")) != '':

      # pasamos a mayúsculas y separamos por palabras cada fila
      listaPalabras = data['TIPO_TRABAJO'][i].encode("utf-8").upper().split()
      # print('Sin filtrar:')
      # print(listaPalabras)

      # quitamos las stopwords de cada fila
      filtered_words = [word for word in listaPalabras if word not in stopwords]
      # print('Filtrado:')
      # print (filtered_words)

      print('Palabras en la frase filtrada:')
      palabras_en_frase=len(filtered_words)

      for j in range(palabras_en_frase):
         print(str(filtered_words[j]))

         # Insertamos un elemento en el diccionario con su clave:valor
         diccionario[filtered_words[j]] = 'experiencia'



#recorremos cada fila de formación
for k in range(tamanyo_dicc_f):
   if (data_f['CURSO'][k].encode("utf-8")) != '':

      # pasamos a mayúsculas y separamos por palabras cada fila
      listaPalabras_f = data_f['CURSO'][k].encode("utf-8").upper().split()
      # print('Sin filtrar:')
      # print(listaPalabras_f)

      # quitamos las stopwords de cada fila
      filtered_words_f = [word for word in listaPalabras_f if word not in stopwords]
      # print('Filtrado:')
      # print (filtered_words_f)


      print('Palabras en la frase filtrada:')
      palabras_en_frase=len(filtered_words_f)

      for j in range(palabras_en_frase):
         print(str(filtered_words_f[j]))

         # Insertamos un elemento en el diccionario con su clave:valor
         diccionario[filtered_words_f[j]] = 'formacion'


#recorremos cada fila de conocimientos
for l in range(tamanyo_dicc_c):
   if (data_c['DENOMINACION'][l].encode("utf-8")) != '':

      # pasamos a mayúsculas y separamos por palabras cada fila
      listaPalabras_c = data_c['DENOMINACION'][l].encode("utf-8").upper().split()
      # print('Sin filtrar:')
      # print(listaPalabras_c)

      # quitamos las stopwords de cada fila
      filtered_words_c = [word for word in listaPalabras_c if word not in stopwords]
      # print('Filtrado:')
      # print (filtered_words_c)


      print('Palabras en la frase filtrada:')
      palabras_en_frase=len(filtered_words_c)

      for j in range(palabras_en_frase):
         print(str(filtered_words_c[j]))

         # Insertamos un elemento en el diccionario con su clave:valor
         diccionario[filtered_words_c[j]] = 'conocimiento'



# Devuelve el numero de elementos que tiene el diccionario
print('Hay',len(diccionario), 'claves en el diccionario')

# creamos el dataframe para crear el excel con las palabras del diccionario
columns = ['Index','Palabra','Num. CV con la palabra']
palabras_tabla = pd.DataFrame(columns=columns)

# creo vector con ceros
zeros=np.zeros(len(diccionario))

# recorro el diccionario e inserto cada palabra en el excel
palabras_tabla['Palabra']= diccionario.keys()

# inserto los contadores para cada palabra
palabras_tabla['Num. CV con la palabra']= zeros

# inserto los contadores para cada palabra
total=len(diccionario)
palabras_tabla['Index']= np.arange(total)
print(palabras_tabla)

# contamos qué palabras aparecen por cada usuario
for a in range(len(codigos_personales)):

   #declaro el diccionario del usuario
   diccionario_usuario={}

   # averiguo las experiencias, formaciones y conocimientos del usuario
   experiencias_usuario=data[data['COD_PERSONAL'] == codigos_personales['COD_PERSONAL'][a]]
   formacion_usuario=data_f[data_f['COD_PERSONAL'] == codigos_personales['COD_PERSONAL'][a]]
   conocimientos_usuario=data_c[data_c['COD_PERSONAL'] == codigos_personales['COD_PERSONAL'][a]]

   # inicializamos la variable a nula para cada usuario
   cadenaPalabrasUsuario=''

   # saber el número de filas de experiencias, formación y experiencias del usuario
   tamanyo_exp = len(experiencias_usuario)
   tamanyo_form = len(formacion_usuario)
   tamanyo_conoc = len(conocimientos_usuario)


   if (tamanyo_exp > 0):
      for indice_fila, fila in experiencias_usuario.iterrows():

         # unimos todas las experiencias en un texto
         cadenaPalabrasUsuario += ' ' + fila['TIPO_TRABAJO'].encode("utf-8").upper()

   if (tamanyo_form > 0):
      for indice_fila, fila in formacion_usuario.iterrows():

         # unimos todas las formaciones en un texto
         cadenaPalabrasUsuario += ' ' + fila['CURSO'].encode("utf-8").upper()

   if (tamanyo_conoc > 0):
      for indice_fila, fila in conocimientos_usuario.iterrows():

         # unimos todas los conocimientos en un texto
         cadenaPalabrasUsuario += ' ' + fila['DENOMINACION'].encode("utf-8").upper()


   # print(cadenaPalabrasUsuario)

   if (len(cadenaPalabrasUsuario)>0):

      # separamos por palabras cada fila
      listaPalabras_u = cadenaPalabrasUsuario.split()
      # print('Sin filtrar:')
      # print(listaPalabras_u)

      # quitamos las stopwords de cada fila
      filtered_words_u = [word for word in listaPalabras_u if word not in stopwords]
      # print('Filtrado:')
      # print (type(filtered_words_u))

   # creamos el diccionario del usuario
   for r in range(len(filtered_words_u)):

      # Insertamos un elemento en el diccionario del usuario con su clave:valor
      diccionario_usuario[filtered_words_u[r]] = codigos_personales['COD_PERSONAL'][a]

   # print(diccionario_usuario.keys())

   for pal in diccionario_usuario:
      print(pal)
      indice = palabras_tabla['Index'][palabras_tabla['Palabra'] == pal]
      valor = palabras_tabla['Num. CV con la palabra'][palabras_tabla['Palabra'] == pal] + 1
      print(palabras_tabla['Num. CV con la palabra'][indice])
      palabras_tabla['Num. CV con la palabra'][indice] = valor
      print(palabras_tabla['Num. CV con la palabra'][indice])


writer = pd.ExcelWriter('C:\Users\ysantana\Desktop\piden_terminos_diccionario_en_cv.xlsx')
palabras_tabla.to_excel(writer,'Sheet1')
writer.save()