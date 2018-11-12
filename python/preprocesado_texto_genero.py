#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from nltk.corpus import stopwords
from string import maketrans

# obtenemos las stopwords del español
stopwords = stopwords.words('spanish')

# paso a mayúsculas los stopwords
for i in range (len(stopwords)):
    stopwords[i] = stopwords[i].upper()

# obtengo en un dataframe los términos introducidos en octubre de 2018
data = pd.read_excel(r'C:\Users\ysantana\Desktop\piden_terminos_diccionario_en_cv.xlsx', sheetname='Sheet1')


# saber el número de filas del excel
tamanyo_dicc=len(data)
print("El fichero tiene:")
print(tamanyo_dicc)

# vemos el tipo de objeto que es data (dataframe)
print(type(data))


# vemos las columnas que existen, para comprobar que se leyó correctamente
print("Nombre de las columnas del fichero:")
print(data.columns)

# creamos lista para almacenar texto tratado
lista=[]

# creamos el diccionario
diccionario_usuario = dict()

# Dataframe para cada palabra del excel
data_tratado=pd.DataFrame()

#recorremos cada fila del excel
for i in range(tamanyo_dicc):
   if (data['Palabra'][i].encode("utf-8")) != '':


      # quitar caracteres a la izquierda y la derecha (se ha omitido los símbolos + y @)
      tratar=data['Palabra'][i].encode("utf-8").strip("!#$%^&*()[]{};:,./<>?\|`~-'=_")
      tratar = tratar.strip('"')

      # separar palabras con caracteres especiales en medio (se ha omitido los símbolos + @ ?)
      intab = '!#$%^&*()[]{};:,./<>\|`~-"=_'
      outtab = "                            "
      trantab = maketrans(intab, outtab)
      tratar_2=tratar.translate(trantab)

      # tokenizar la palabra tratada
      tokens = tratar_2.split()

      # elimino las stopwords
      filtered_words_u = [word for word in tokens if word not in stopwords]

      # creamos el diccionario del usuario
      for r in range(len(filtered_words_u)):

         # quitamos el símbolo español ¡¿ºª y corregimos los acentos y letra ñ
         sin_simbolo = filtered_words_u[r].replace("¡", "")
         sin_simbolo = sin_simbolo.replace("»", "")
         sin_simbolo = sin_simbolo.replace("«", "")
         sin_simbolo = sin_simbolo.replace("¿", "")
         sin_simbolo = sin_simbolo.replace("°", "")
         sin_simbolo = sin_simbolo.replace("º", "")
         sin_simbolo = sin_simbolo.replace("ª", "")
         sin_simbolo = sin_simbolo.replace("á", "A")
         sin_simbolo = sin_simbolo.replace("Á", "A")
         sin_simbolo = sin_simbolo.replace("é", "E")
         sin_simbolo = sin_simbolo.replace("É", "E")
         sin_simbolo = sin_simbolo.replace("í", "I")
         sin_simbolo = sin_simbolo.replace("Í", "I")
         sin_simbolo = sin_simbolo.replace("ó", "O")
         sin_simbolo = sin_simbolo.replace("Ó", "O")
         sin_simbolo = sin_simbolo.replace("ú", "U")
         sin_simbolo = sin_simbolo.replace("Ú", "U")
         sin_simbolo = sin_simbolo.replace("ñ", "Ñ")

         # corrección de tildes
         sin_simbolo = sin_simbolo.replace("?", "")


         # comprobamos que no sea un número
         es_numero = sin_simbolo.isdigit()

         # guardamos la palabra si no es número y no vacío
         if ((es_numero == False ) and (sin_simbolo!='')):

            # Insertamos un elemento en el diccionario del usuario con su clave:valor
            diccionario_usuario[sin_simbolo] = 'elemento'


# creo la lista con las palabras del diccionario
# data_tratado['Palabra tratada']= lista

# recorro el diccionario e inserto cada palabra en el excel
lista = diccionario_usuario.keys()
data_tratado['Palabra tratada'] = sorted(lista)
data_tratado['Singular'] = sorted(lista)


# averiguamos si en el excel existe el singular y plural de la misma palabra. Se propone en la columna singular
# el singular del plural
for s in range(len(data_tratado)):
   var = data_tratado['Palabra tratada'][s]

   if ((var[len(var) - 1: len(var)]) == 'S'):
      sing = var[0: len(var) - 1]
      fila=data_tratado[data_tratado['Palabra tratada'] == sing]

      if (fila.index>0):
         data_tratado['Singular'][s] = sing

      else:
         sing2 = var[0: len(var) - 2]
         fila = data_tratado[data_tratado['Palabra tratada'] == sing2]

         if (fila.index > 0):
            data_tratado['Singular'][s] = sing2

# Dataframe para cada palabra del excel
data_final=pd.DataFrame()
diccionario_singular=dict()


for j in range(len(data_tratado)):

   # Insertamos un elemento en el diccionario con su clave:valor
   diccionario_singular[data_tratado['Singular'][j]] = 'palabra'


print(len(diccionario_singular))

# recorro el diccionario e inserto cada palabra en el excel
data_final['Palabra']= diccionario_singular.keys()
data_final['Palabra']=sorted(data_final['Palabra'])

# Creamos la nueva columna para el género
data_final['Genero'] = data_final['Palabra']

# averiguamos si en la columna singular existe femenino y masculino. Se propone en la columna genero
# el masculino para no duplicar el término
for g in range(len(data_final)):
   var = data_final['Palabra'][g]

   if ((var[len(var) - 1: len(var)]) == 'A'):
      gen = var[0: len(var) - 1] + 'O'
      print(gen)
      fila2=data_final[data_final['Palabra'] == gen]

      if (fila2.index>0):
         data_final['Genero'][g] = gen

writer = pd.ExcelWriter('C:\Users\ysantana\Desktop\preprocesado_texto.xlsx')
data_tratado.to_excel(writer,'Sheet1')
writer.save()

writer = pd.ExcelWriter('C:\Users\ysantana\Desktop\preprocesado_texto_singular.xlsx')
data_final.to_excel(writer,'Sheet1')
writer.save()