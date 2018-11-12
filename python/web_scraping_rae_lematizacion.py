#!/usr/bin/env python
# -*- coding: utf-8 -*-

# importamos librerías necesarias
import time
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

# --------------------------- CAPTURA DE DATOS ---------------------------

# obtengo en un dataframe los términos introducidos en octubre de 2018
data = pd.read_excel(r'C:\Users\Bomboncito\Documents\Master Ciencia de Datos\tipologia y ciclo de vida de los datos\Práctica 1\diccionario_reducido_octubre_2.xlsx', sheetname='Hoja1')

# saber el número de filas del excel
tamanyo_dicc=len(data)
print("El fichero tiene:")
print(tamanyo_dicc)

# vemos el tipo de objeto que es data (dataframe)
print(type(data))


# vemos las columnas que existen, para comprobar que se leyó correctamente
print("Nombre de las columnas del fichero:")
print(data.columns)

# creamos lista para almacenar lematización de cada palabra
lista_lemas=[]
lista_sugeridas=[]

# creamos lista para almacenar tipo de cada palabra
lista_tipos=[]

# ----------------------- JAVASCRIPT SCRAPING ------------------------------

#recorremos cada fila del excel
for i in range(tamanyo_dicc):
    print('Iteración término')

    # obtengo la url de búsqueda
    parametro=data['Palabra'].iloc[i]
    print(str(parametro))
    url='http://dle.rae.es/?w=' + str(parametro)
    print(url)

    # obtenemos el driver de Chrome, que fue previamente descargado, guardado en C:// y añadido el PATH como variable de entorno
    driver = webdriver.Chrome(r'C:\Users\Bomboncito\chromedriver.exe')

    # Añadimos un retardo para simular que es una persona quien lo hace
    time.sleep(5) # Let the user actually see something!

    try:

        # Acudimos a la página con el parámetro de búsqueda por la URL
        driver.get(url)

        # Buscamos la etiqueta deseada (header) con la clase deseada (f)
        etiqueta = driver.find_element_by_xpath("//header[@class='f']")

        print('Lematización:')
        print(etiqueta.text)

        # añadimos a la lista de lemas
        lista_lemas.append(etiqueta.text)
        lista_sugeridas.append('')

        # Buscamos la etiqueta deseada (abbr) con la clase deseada (d). Me quedo con la primera abbr
        # porque se supone la más usada
        tipo = driver.find_element_by_xpath("//abbr[@class='d']")

        print('Tipo:')
        print(tipo.text)

        # añadimos a la lista de tipos
        lista_tipos.append(tipo.text)

        # Simulamos que el usuario lee el resultado
        time.sleep(5)  # Let the user actually see something!

        # Cerramos el navegador
        driver.quit()

    except NoSuchElementException:

        # creamos diccionario para etiquetas de cada palabra
        diccionario_termino=dict()

        # creamos la lista de enlaces que no son basura
        lista_enlaces=list()

        print('No hay elemento concreto')

        # buscamos si hay varios enlaces para la palabra buscada
        aTagsInLi = driver.find_elements_by_xpath("/html/body/div/div/section/section/div/section/div/div/ul/li/a")

        # creamos la lista de direcciones a visitar
        for b in aTagsInLi:

            enlace = b.get_attribute('href')

            # evitar hacer javascript scraping de enlaces "basura"
            if 'http://dle.rae.es/?id' in enlace:
                lista_enlaces.append(enlace)

        # visualizamos enlaces a visitas
        print('Hay que visitar',len(lista_enlaces),'enlaces')

        # Simulamos que el usuario lee el resultado
        time.sleep(5)  # Let the user actually see something!

        # Cerramos el navegador
        driver.quit()

        try:

            if (len(lista_enlaces) < 3):
                for a in lista_enlaces:
                    try:
                        # obtenemos el driver de Chrome, que fue previamente descargado, guardado en C:// y añadido el PATH como variable de entorno
                        driver = webdriver.Chrome(r'C:\Users\Bomboncito\chromedriver.exe')

                        # Añadimos un retardo para simular que es una persona quien lo hace
                        time.sleep(5)  # Let the user actually see something!

                        # Acudimos al enlace
                        driver.get(a)

                        # Buscamos la etiqueta deseada (header) con la clase deseada (f)
                        etiqueta_e = driver.find_element_by_xpath("//header[@class='f']")

                        print('Lematización:')
                        print(etiqueta_e.text)

                        # Buscamos la etiqueta deseada (abbr) con la clase deseada (d)
                        tipo_e = driver.find_element_by_xpath("//abbr[@class='d']")

                        print('Tipo:')
                        print(tipo_e.text)

                        # añadimos al diccionario
                        diccionario_termino[etiqueta_e.text] = tipo_e.text

                        # Simulamos que el usuario lee el resultado
                        time.sleep(5)  # Let the user actually see something!

                        # Cerramos el navegador
                        driver.quit()

                    except NoSuchElementException:
                        # añadimos nulos a la lista de lemas y tipos
                        lista_lemas.append('')
                        lista_tipos.append('')

                        # Simulamos que el usuario lee el resultado
                        time.sleep(5)  # Let the user actually see something!

                        # Cerramos el navegador
                        driver.quit()


                # Visualizamos el diccionario
                for termino in diccionario_termino:
                    print (diccionario_termino)

                # añadimos a la lista de etiquetas y tipos
                var_e=''
                var_t=''

                for termino in diccionario_termino:


                    if (diccionario_termino[termino]=='m.'):
                        var_e = termino
                        var_t = 'm.'
                    elif (diccionario_termino[termino]=='f.'):
                        if (var_t!='m.'):
                            var_e = termino
                            var_t = 'f.'
                    elif (diccionario_termino[termino]=='adj.'):
                        if ((var_t!='m.') and (var_t!='f.')):
                            var_e = termino
                            var_t = 'adj.'
                    elif (diccionario_termino[termino]=='tr.'):
                        if ((var_t!='m.') and (var_t!='f.') and (var_t!='adj.')):
                            var_e = termino
                            var_t = 'tr.'
                    elif (diccionario_termino[termino]=='intr.'):
                        if ((var_t!='m.') and (var_t!='f.') and (var_t!='adj.') and (var_t!='tr.')):
                            var_e = termino
                            var_t = 'intr.'

                # añadimos la etiqueta y tipo a la lista
                lista_sugeridas.append(var_e)
                lista_lemas.append('')
                lista_tipos.append(var_t)

            else:
                # añadimos nulos a la lista de lemas y tipos
                lista_sugeridas.append('')
                lista_lemas.append('')
                lista_tipos.append('')

        except NoSuchElementException:
            # añadimos nulos a la lista de lemas y tipos
            lista_sugeridas.append('')
            lista_lemas.append('')
            lista_tipos.append('')



    # Simulamos que el usuario lee el resultado
    time.sleep(5) # Let the user actually see something!

    # Cerramos el navegador
    driver.quit()

# Tamaños listas
print(len(lista_lemas))
print(len(lista_tipos))
print(len(lista_sugeridas))
print(lista_lemas)
print(lista_tipos)
print(lista_sugeridas)


# creamos las nuevas columnas del dataframe con el lema y tipo de cada palabra
data['Lema']= lista_lemas
data['Sugeridas']= lista_sugeridas
data['Tipo']= lista_tipos


# visualizamos el dataframe
print(data)


writer = pd.ExcelWriter(r'C:\Users\Bomboncito\Documents\Master Ciencia de Datos\tipologia y ciclo de vida de los datos\Práctica 1\lematizacion_rae.xlsx')
data.to_excel(writer,'Sheet1')
writer.save()


