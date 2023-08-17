# -*- coding: utf-8 -*-
"""
Created on Sun May  7 19:58:51 2023

@author: @AquinoRuilova
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Maneja archivo de participantes 

# Lee archivo excel e indexa celdas

archivo_excel = pd.ExcelFile('graduados.xlsx')

graduado = pd.read_excel(archivo_excel,'participante')
datos = pd.read_excel(archivo_excel,'datos') 

cedula = graduado["cedula"] 
nombre = graduado["nombre"] 
archivo = graduado["archivo"]
rol1 = graduado["rol1"]
rol2 = graduado["rol2"]
rol3 = graduado["rol3"]

titulo1 = datos["titulo1"] 
titulo2 = datos["titulo2"] 
titulo3 = datos["titulo3"] 
fecha = datos["fecha"] 

cantidad_estudiantes =  graduado.shape[0]

print()      
print ("Cantidad de graduandos = ", cantidad_estudiantes)

numero_estudiante = 1

# Prepara certificado

certificado = Image.open("certificado_una.jpg")

titulo1_certificado = ImageDraw.Draw(certificado)
titulo2_certificado = ImageDraw.Draw(certificado)
titulo3_certificado = ImageDraw.Draw(certificado)
fecha_certificado = ImageDraw.Draw(certificado)

coordenadas_titulo1 = (600,200)
coordenadas_titulo2 = (540,280)
coordenadas_titulo3 = (640,330)
coordenadas_fecha = (1300,1050)
 
color_texto = (0,0,0) # código color negro en formato (R,G,B)

letra_titulo1 = ImageFont.truetype("letras.ttf", 48)
letra_titulo2 = ImageFont.truetype("letras.ttf", 36)
letra_titulo3 = ImageFont.truetype("letras.ttf", 36)
letra_fecha = ImageFont.truetype("letras.ttf", 36)
 
titulo1_certificado.text(coordenadas_titulo1, titulo1[0], 
                         fill=color_texto, font=letra_titulo1)
titulo2_certificado.text(coordenadas_titulo2, titulo2[0], 
                         fill=color_texto, font=letra_titulo2)
titulo3_certificado.text(coordenadas_titulo3, titulo3[0], 
                         fill=color_texto, font=letra_titulo3)
fecha_certificado.text(coordenadas_fecha, fecha[0], 
                         fill=color_texto, font=letra_fecha)

certificado.save("certificado.jpg")

# Maneja certificados

while numero_estudiante <= cantidad_estudiantes:
    
    
    i = numero_estudiante - 1
    
    certificado = Image.open("certificado.jpg") 
    
    nombre_certificado = ImageDraw.Draw(certificado)
    cedula_certificado = ImageDraw.Draw(certificado)
    
    rol1_certificado = ImageDraw.Draw(certificado)
    rol2_certificado = ImageDraw.Draw(certificado)
    rol3_certificado = ImageDraw.Draw(certificado)
    
    coordenadas_nombre = (740,650)
    coordenadas_cedula = (800,800)
    
    coordenadas_rol1 = (100,850)
    coordenadas_rol2 = (100,920)
    coordenadas_rol3 = (100,970)
    
    color_texto = (0,0,0) # código color negro en formato (R,G,B) 
    
    
    letra_nombre = ImageFont.truetype("OLDENGL.ttf", 72) 
    letra_cedula = ImageFont.truetype("letras.ttf", 36)
    
    letra_rol1 = ImageFont.truetype("letras.ttf", 36)
    letra_rol2 = ImageFont.truetype("letras.ttf", 24)
    letra_rol3 = ImageFont.truetype("letras.ttf", 24)

    nombre_certificado.text(coordenadas_nombre, nombre[i], 
                            fill=color_texto, font=letra_nombre)
    cedula_certificado.text(coordenadas_cedula, cedula[i], 
                            fill=color_texto, font=letra_cedula)
    
    rol1_certificado.text(coordenadas_rol1, rol1[i], 
                            fill=color_texto, font=letra_rol1)
    rol2_certificado.text(coordenadas_rol2, rol2[i], 
                            fill=color_texto, font=letra_rol2)
    rol3_certificado.text(coordenadas_rol3, rol3[i], 
                            fill=color_texto, font=letra_rol3)
    
    n = str(i+1)
    
    archivog = archivo[i]+n
    
    certificado.save("C://Users//aquin//Proyectos_SEA//diploma//certificados//"
                     +archivo[i]+n+".pdf")
    
    numero_estudiante = numero_estudiante + 1

certificado = Image.open("certificado_una.jpg")
certificado.save("certificado.jpg")
        
        
                 
   
        

