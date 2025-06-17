# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 16:12:55 2022

@author: FRCASTRO
******************************************************************************
                        INTENCIONES ASSET MANAGEMENT
******************************************************************************  
"""

import pandas as pd
import numpy as np
import glob
import win32api
import win32con
import win32event
import win32com.client as win32
from datetime import date, datetime,timedelta
import time
import sys
import os
import pythoncom
import definiciones
import getpass
import tkinter as tk
from tkinter import messagebox


def cargarDatos():
    
    #Aqui se importan los insumos que son útiles para el funcionamiento de la herramienta
    #Tener presente su ubicación
    
    #Desde el archivo de administración que está en el archivo de txt llamado ruta.txt se obtiene varias hojas con información.
    #Parámetros
    definiciones.parametros = pd.read_excel(definiciones.archivoParametros,sheet_name="Parametros",index_col=(0),header=0)
    print("1 de 12")

    #Si no es posible acceder al servidor SBBOGSCL0 por su nombre, lo hacemos por medio de la IP, eso es útil cuando se trabaja con VPN
    if not os.path.exists(definiciones.parametros["Valor"]["rutaAccesoAplicacionExcel"]):
        definiciones.parametros["Valor"] = definiciones.parametros["Valor"].apply(lambda x: x.replace("SBBOGSCL0",definiciones.parametros["Valor"]["ip servidor"]) if "SBBOGSCL0" in str(x) else x)
    
    
    #Relación entre códigos de OyD y código Murex de los portafolios de Valores
    definiciones.nombreYcodigosPortValores = pd.read_excel(definiciones.archivoParametros,sheet_name="Port. Valores",index_col=(0),header=0)
    print("2 de 12")
    
    #Listado de portafolios en CRM
    portafoliosCRM = pd.read_excel(definiciones.parametros["Valor"]["rutaInventarioPortafolios"],sheet_name="Listado PD",skiprows=definiciones.parametros["Valor"]["filasOmitirPortafolios"])
    portafoliosCRM.loc[portafoliosCRM["ADMINISTRADOR"] == "VALORES","CÓD MUREX"] = portafoliosCRM.loc[portafoliosCRM["ADMINISTRADOR"] == "VALORES","CÓD. CONT. | OYD | PERSHING"].apply(lambda x :  definiciones.nombreYcodigosPortValores.loc[x,"Código Murex"])
    definiciones.portafoliosCRM = portafoliosCRM
    print("3 de 12")

    #Títulos de Fiduciaria
    nemoTitulosFiduciaria = obtenerUltimoInventarioTitulos(definiciones.parametros["Valor"]["rutaInventarioTitulosFiduciaria"],definiciones.parametros["Valor"]["filasOmitirTitulosFiduciaria"])
    nemoTitulosFiduciaria = limpiarDatos(nemoTitulosFiduciaria,["Macro Activo","ISIN","Emisor / Contraparte","Especie/Generador","Nemotécnico","Emisor Unificado"],["Macro Activo","ISIN","Emisor / Contraparte","Especie/Generador","Nemotécnico","Emisor Unificado"],["SALDO Macro Activo","SALDO ABA","Nominal Remanente"])
    nemoTitulosFiduciaria["Origen Informacion"] = "INVENTARIO TITULOS"
    print("4 de 12")
    
    #Titulos de Valores
    nemoTitulosValores = obtenerUltimoInventarioTitulos(definiciones.parametros["Valor"]["rutaInventarioTitulosValores"],definiciones.parametros["Valor"]["filasOmitirTitulosValores"]) 
    nemoTitulosValores["Portafolio"] = nemoTitulosValores["Código OyD"].apply(lambda x: definiciones.nombreYcodigosPortValores.loc[x,"Código Murex"] if x in definiciones.nombreYcodigosPortValores.index.tolist() else "-" )
    nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Portafolio"].isin(definiciones.portafoliosCRM.loc[definiciones.portafoliosCRM["ADMINISTRADOR"] == "VALORES","CÓD MUREX"].tolist())]
    nemoTitulosValores = limpiarDatos(nemoTitulosValores,["Macro Activo","Isin","Nemoténico","Emisor Unificado","Nombre Emisor"],["Macro Activo","Isin","Nemoténico","Emisor Unificado","Nombre Emisor"],["SALDO Macro Activo","SALDO ABA","Valor Nominal Actual"])
    nemoTitulosValores["Origen Informacion"] = "INVENTARIO TITULOS"
    nemoTitulosValores["POSICIÓN"] = ""
    print("5 de 12")
    
    #Usuarios
    usuarios = pd.read_excel(definiciones.archivoParametros,sheet_name="Usuarios",index_col=(0),header=0)
    print("6 de 12")
    
    #Especies
    especies = pd.read_excel(definiciones.archivoParametros,sheet_name="Especies",header=0)
    print("7 de 12")
    
    #Cupos de Fiduciaria
    cuposFiduciaria = pd.read_excel(definiciones.rutaCupos.replace("usuario_red",getpass.getuser()) + definiciones.parametros["Valor"]["archivoCuposFiduciaria"],sheet_name="BDCupos",skiprows=2)
    cuposFiduciaria = cuposFiduciaria[cuposFiduciaria["Entidad"] == "Fiduciaria"]
    cuposFiduciaria = limpiarDatos(cuposFiduciaria,colTrim=["Entidad","MUREX","Nombre.1","Cupo","Nemo","ISIN 1"],colUpper=["Entidad","MUREX","Nombre.1","Cupo","Nemo","ISIN 1"],colFloat=["Ocupación Máxima"])
    print("8 de 12")
    
    #CUpos de Valores
    cuposValores = pd.read_excel(definiciones.rutaCupos.replace("usuario_red",getpass.getuser()) + definiciones.parametros["Valor"]["archivoCuposValores"],sheet_name="BDCupos",skiprows=2)
    cuposValores = cuposValores[cuposValores["Entidad"] == "Valores"]
    cuposValores = limpiarDatos(cuposValores,colTrim =["Entidad","Nombre.1","Cupo","Nemo","ISIN 1"],colUpper=["Entidad","Nombre.1","Cupo","Nemo","ISIN 1"],colFloat=["Ocupación Máxima"])
    cuposValores["MUREX"] = cuposValores["OyD"].apply(lambda x: definiciones.nombreYcodigosPortValores.loc[x,"Código Murex"] if x in definiciones.nombreYcodigosPortValores.index.tolist() else "-" )
    cuposValores = cuposValores[cuposValores["MUREX"].isin(definiciones.portafoliosCRM.loc[definiciones.portafoliosCRM["ADMINISTRADOR"] == "VALORES","CÓD MUREX"].tolist())]
    print("9 de 12")
    
    #Operaciones por cumplir en Fiduciaria: 
    #Estas operaciones aqui obtenidas se deben cargar al archivo de titulos.
    #Se deben agregar solo las operaciones cuya fecha es igual a hoy.
    #La especie coincide con la especie de inventario de titulos
    operacionesPorCumplirFidu = cargarOperacionesPorCumplirFiduciaria(definiciones.parametros["Valor"]["rutaOperacionesPorCumplirFidu"],nemoTitulosFiduciaria)
    nemoTitulosFiduciaria = agregarOpercionesPendientesPorCumplirFidu(nemoTitulosFiduciaria,operacionesPorCumplirFidu)
    print("10 de 12")
    
    #Operaciones por cumplir Valores: 
    #Estas operaciones aqui obtenidas se deben cargar al archivo de titulos.
    #Se deben agregar solo las operaciones cuya fecha es igual a hoy.
    #El Nemo coincide con el Nemo del inventario de titulos
    #Las compras entran como positivas y las ventas entran con nomial negativo
    operacionesPorCumplirValores = cargarOperacionesPorCumplirValores(definiciones.parametros["Valor"]["rutaOperacionesPorCumplirValores"],nemoTitulosValores)
    nemoTitulosValores = agregarOpercionesPendientesPorCumplirValores(nemoTitulosValores,operacionesPorCumplirValores)
    nemoTitulosValores["Portafolio"] = nemoTitulosValores["Código OyD"].apply(lambda x: definiciones.nombreYcodigosPortValores.loc[x,"Código Murex"] if x in definiciones.nombreYcodigosPortValores.index.tolist() else "-" )
    nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Portafolio"].isin(definiciones.portafoliosCRM.loc[definiciones.portafoliosCRM["ADMINISTRADOR"] == "VALORES","CÓD MUREX"].tolist())]
    print("11 de 12")
    
    #Sobrepasos 
    ruta_seguimiento_limites = guardarArchivosDesdeCorreo(definiciones.parametros["Valor"]["rutaArchivoSobrepasosLocal"],definiciones.parametros["Valor"]["asuntoSeguimientoLimites"])
    ruta_sobrepasos_limites = guardarArchivosDesdeCorreo(definiciones.parametros["Valor"]["rutaArchivoSobrepasosLocal"],definiciones.parametros["Valor"]["asuntoSobrepasosLimites"])
    tablaSobrepasos = cargarSobrepasos(ruta_seguimiento_limites,ruta_sobrepasos_limites) #Agregar un control para leer el de la fecha correcta
    
    print("12 de 12")
        
    definiciones.porcentajeProtejidoFondos =  pd.read_excel(definiciones.archivoParametros,sheet_name="Por. Retiro Fondos",header=0)
    definiciones.nemoTitulosFiduciaria = nemoTitulosFiduciaria
    definiciones.nemoTitulosValores = nemoTitulosValores
    definiciones.especies = limpiarDatos(especies,["Emisor inventario","Emisor cupos","Nit emisor","Especie","Nemotecnico","Nemo intenciones","Isin","Macro Activo","Macro Activo inventario","Indicador","Moneda"],["Emisor inventario","Emisor cupos","Nit emisor","Especie","Nemotecnico","Nemo intenciones","Isin","Macro Activo","Macro Activo inventario","Indicador","Moneda"],[])
    definiciones.especiesOriginal = definiciones.especies.copy()
    definiciones.especies.loc[especies["Macro Activo"] == "RF INTERNACIONAL","Macro Activo"] = "RV INTERNACIONAL"
    definiciones.cuposFiduciaria = cuposFiduciaria
    definiciones.cuposValores = cuposValores
    definiciones.valorPortafolioFiduciaria = definiciones.nemoTitulosFiduciaria[["Portafolio","SALDO ABA"]].groupby(["Portafolio"]).sum()
    definiciones.valorPortafolioValores = definiciones.nemoTitulosValores[["Portafolio","SALDO ABA"]].groupby(["Portafolio"]).sum()
    definiciones.usuarios = usuarios
    definiciones.precioAcciones = calcularPrecioAcciones(definiciones.nemoTitulosFiduciaria,definiciones.nemoTitulosValores)
    definiciones.UVR = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Moneda Security"] == "UVR","Tasa Formada"].iloc[0]
    definiciones.TRM = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Moneda Security"] == "USD","Tasa Formada"].iloc[0]
    definiciones.excel = win32.DispatchEx("Excel.Application") 
    definiciones.keepOpen = True
    definiciones.tablaSobrepasos = tablaSobrepasos
    definiciones.fechaInvenFidu = nemoTitulosFiduciaria["Fecha reporte"].iloc[0]
    definiciones.fechaInvenVal = nemoTitulosValores["Fecha"].iloc[0]    
    definiciones.fechaOpPendientesValores = operacionesPorCumplirValores["Fecha"].iloc[0]
    definiciones.fechaSeguimientoLimites = ruta_seguimiento_limites.split("/")[-1]
    definiciones.fechaSobrepasoLimites = ruta_sobrepasos_limites.split("/")[-1]
    actualizarEstadoIntencionesVencidas(definiciones.usuarios[definiciones.usuarios["Rol"].isin(["Administrador","PM"])].index.tolist())

    print("HAY UNA VENTANA ABIERTA, REVISE LA INFORMACIÓN")
    root = tk.Tk()
    root.withdraw()  # Hide the main 
    root.attributes('-topmost',True)
    messagebox.showinfo("Carga completada exitosamente","TRM: "+ str(definiciones.TRM) +"\nUVR: "+str(definiciones.UVR) +"\nInventario títulos Fiduciaria: " +str(definiciones.fechaInvenFidu.date()) + "\nInventario títulos Valores: " + str(definiciones.fechaInvenVal.date()) + "\nOperaciones pendientes Fiduciaria: "+ str(definiciones.fechaOpPendientesFidu) + "\nOperaciones pendientes Valores: " + str(definiciones.fechaOpPendientesValores.date())+ "\n" + str(definiciones.fechaSeguimientoLimites) + "\n" + str(definiciones.fechaSobrepasoLimites))
  
def limpiarDatos(tabla,colTrim,colUpper,colFloat):
    """
    tabla: df
    colTrim: list
    colUpper: list
    colFloat: list
    return df
    """
    tabla[colTrim] = tabla[colTrim].apply(lambda x: x.astype("str").str.strip())
    tabla[colUpper] = tabla[colUpper].apply(lambda x: x.astype("str").str.upper())
    tabla[colFloat] = tabla[colFloat].apply(lambda x: x.astype("float"))
    return tabla
   
def guardarArchivosDesdeCorreo(path,subject):
    '''
    Esta funcion guarda el ultimo archivo de excel correspondiente a un asunto predeterminado que llega al correo
    Autor: Daniel Steven Lopez Daza <dldaza@bancolombia.com.co>
	Fecha: 2023-11-24
	:param path: Ruta en donde se va a almacenar el archivo
	:param subject: Palabra inicial del asunto del correo
	'''
    ruta_archivo_adjunto = ""
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # Obtener la carpeta de la bandeja de entrada
        inbox = outlook.GetDefaultFolder(6)  # 6 representa la constante para la carpeta de la bandeja de entrada

        # Obtener todos los correos electrónicos en la bandeja de entrada
        messages = inbox.Items
        mensajes_sobrepasos = list(filter(lambda x: subject in x.Subject , messages))
        mensaje = mensajes_sobrepasos[-1]
        adjuntos = mensaje.Attachments
        for i in range(1,len(adjuntos)+1):
            if str(adjuntos.Item(i)).split('.')[1] == 'xlsx':
                adjunto = adjuntos.Item(i)
                adjunto.SaveAsFile(path + str(adjunto))
                ruta_archivo_adjunto = path + str(adjunto)
                
                
    except Exception as e:
       mostrarMensajeAdvertencia('No se pudo descargar los archivos de sobrepasos desde el correo: ' +str(e))
   
    return ruta_archivo_adjunto 
        
    
   
def cargarSobrepasos(ruta_seguimiento_limites,ruta_sobrepasos_limites):
    
    insumo_sobrepasos = pd.DataFrame([],columns=["PORTAFOLIO","EMISOR","NIT EMISOR","CALIFICACION","FECHA SOBREPASO","TOTAL DIAS TRANSCURRIDOS","SOBREPASO PESOS","SOBREPASO PORCENTAJE","TIPO SOBREPASO","MENSAJE","CAUSA","PLAN DE ACCION"])
    
    
    if ruta_seguimiento_limites == "" or ruta_sobrepasos_limites == "":
        return insumo_sobrepasos
    
    try:
        archivo_sobrepasos_limites = pd.ExcelFile(ruta_sobrepasos_limites)
        nombre_hojas_sobrepasos_limites = archivo_sobrepasos_limites.sheet_names 
        archivo_sobrepasos_limites.close()
        archivo_seguimiento_limites = pd.ExcelFile(ruta_seguimiento_limites)
        nombre_hojas_seguimiento_limites = archivo_seguimiento_limites.sheet_names
        archivo_seguimiento_limites.close()

        
        
        
        if "Sobrepasos" in nombre_hojas_seguimiento_limites:  #Pensionales
            pensionales = pd.read_excel(ruta_seguimiento_limites,sheet_name = "Sobrepasos",skiprows = 6)
            pensionales["TIPO SOBREPASO"] = "PENSIONALES"
            pensionales = pensionales[~pensionales["SOBREPASO"].isna()]
            pensionales["FECHA SOBREPASO"] = pensionales["SOBREPASO"].apply(lambda sobrepaso:sobrepaso[:10])
            pensionales["SOBREPASO PORCENTAJE"] = pensionales["Consumo%"] - pensionales["Max"]
            pensionales = pensionales[["Código","Descripción","FECHA SOBREPASO","Disponible","SOBREPASO PORCENTAJE","TIPO SOBREPASO"]]
            pensionales = pensionales.rename(columns={"Código":"PORTAFOLIO","Descripción":"EMISOR","Disponible":"SOBREPASO PESOS"})
            pensionales["MENSAJE"] = pensionales.apply(lambda dato: " EL portafolio pensional " + str(dato["PORTAFOLIO"]) + " tiene un sobre paso de $" +format(abs(float(dato["SOBREPASO PESOS"])),",.0f") +" que corresponde a un " + format(dato["SOBREPASO PORCENTAJE"],".2%") + " con la siguiente descripción: " + str(dato["EMISOR"]).strip(),axis =1)
            insumo_sobrepasos = pd.concat([insumo_sobrepasos,pensionales],axis=0,ignore_index=True)
        
        if "Emisor Sobrepasos" in nombre_hojas_sobrepasos_limites:
            emisor_sobrepasos = pd.read_excel(ruta_sobrepasos_limites, sheet_name="Emisor Sobrepasos", skiprows=6)
            emisor_sobrepasos.columns = emisor_sobrepasos.columns.str.strip()
            emisor_sobrepasos = emisor_sobrepasos.rename(columns={"Portafolio\n MLC":"PortafolioMLC"})
            emisor_sobrepasos["TIPO SOBREPASO"] = "EMISOR SOBREPASOS"
            emisor_sobrepasos = emisor_sobrepasos[~emisor_sobrepasos["Observaciones"].isna() & ~emisor_sobrepasos["PortafolioMLC"].isna()]
            emisor_sobrepasos["FECHA SOBREPASO"] = emisor_sobrepasos["Observaciones"].apply(lambda observacion: observacion[1:11])
            emisor_sobrepasos = emisor_sobrepasos[["PortafolioMLC","Nombre Emisor","Emisor","FECHA SOBREPASO","Sobrepaso Actual","Sobrepaso en %","TIPO SOBREPASO","Causa","Plan de Acción"]]
            emisor_sobrepasos = emisor_sobrepasos.rename(columns={"PortafolioMLC":"PORTAFOLIO","Nombre Emisor":"EMISOR","Emisor":"NIT EMISOR","Sobrepaso Actual":"SOBREPASO PESOS","Sobrepaso en %":"SOBREPASO PORCENTAJE","Causa":"CAUSA","Plan de Acción":"PLAN DE ACCION"})
            emisor_sobrepasos["MENSAJE"] = emisor_sobrepasos.apply(lambda dato: "El portafolio " + str(dato["PORTAFOLIO"]).strip() + " tiene un sobrepaso en el emisor " + str(dato["EMISOR"]).strip() + " de $" +format(abs(float(dato["SOBREPASO PESOS"])),",.0f") +" que corresponde a un " + format(dato["SOBREPASO PORCENTAJE"],".2%"),axis =1)
            insumo_sobrepasos = pd.concat([insumo_sobrepasos,emisor_sobrepasos],axis=0,ignore_index=True)
            insumo_sobrepasos["EMISOR"] =  insumo_sobrepasos["EMISOR"].astype('str').str.upper()
            
        
        if "Emisor Sobrepasos - Traslados" in nombre_hojas_sobrepasos_limites:
            emisor_sobrepasos_traslados = pd.read_excel(ruta_sobrepasos_limites,sheet_name="Emisor Sobrepasos - Traslados",skiprows=6)
            emisor_sobrepasos_traslados.columns = emisor_sobrepasos_traslados.columns.str.strip()
            emisor_sobrepasos_traslados = emisor_sobrepasos_traslados.rename(columns={"Portafolio\nMLC":"PortafolioMLC"})
            emisor_sobrepasos_traslados["TIPO SOBREPASO"] = "EMISOR SOBREPASOS TRASLADOS"
            emisor_sobrepasos_traslados = emisor_sobrepasos_traslados[~emisor_sobrepasos_traslados["Observaciones"].isna() & ~emisor_sobrepasos_traslados["PortafolioMLC"].isna()]
            emisor_sobrepasos_traslados["FECHA SOBREPASO"] = emisor_sobrepasos_traslados["Observaciones"].apply(lambda observacion: str(observacion)[1:11])
            emisor_sobrepasos_traslados = emisor_sobrepasos_traslados[["PortafolioMLC","Nombre Emisor","Emisor","FECHA SOBREPASO","Sobrepaso Actual","Sobrepaso \nen %","TIPO SOBREPASO","Causa","Plan de Acción"]]
            emisor_sobrepasos_traslados = emisor_sobrepasos_traslados.rename(columns={"PortafolioMLC":"PORTAFOLIO","Nombre Emisor":"EMISOR","Emisor":"NIT EMISOR","Sobrepaso Actual":"SOBREPASO PESOS","Sobrepaso \nen %":"SOBREPASO PORCENTAJE","Causa":"CAUSA","Plan de Acción":"PLAN DE ACCION"})
            emisor_sobrepasos_traslados["MENSAJE"] = emisor_sobrepasos_traslados.apply(lambda dato: "El portafolio " + str(dato["PORTAFOLIO"]).strip() + " tiene un sobrepaso en el emisor " + str(dato["EMISOR"]).strip() + " de $" + format(abs(float(dato["SOBREPASO PESOS"])),",.0f") + " que corresponde a un " + format(dato["SOBREPASO PORCENTAJE"],".2%"),axis =1 )
            
            insumo_sobrepasos = pd.concat([insumo_sobrepasos,emisor_sobrepasos_traslados],axis=0,ignore_index=True)
            insumo_sobrepasos["EMISOR"] =  insumo_sobrepasos["EMISOR"].astype('str').str.upper()
            
        
        if "Grupo Económico Sobrepasos" in nombre_hojas_sobrepasos_limites:
            
            grupo_economico_sobrepasos = pd.read_excel(ruta_sobrepasos_limites, sheet_name="Grupo Económico Sobrepasos", skiprows=6)
            grupo_economico_sobrepasos.columns = grupo_economico_sobrepasos.columns.str.strip()
            grupo_economico_sobrepasos["TIPO SOBREPASO"] = "GRUPO ECONOMICO"
            grupo_economico_sobrepasos = grupo_economico_sobrepasos[~grupo_economico_sobrepasos["Portafolio"].isna()]
            grupo_economico_sobrepasos["FECHA SOBREPASO"] = grupo_economico_sobrepasos["Validación Sobrepasos"].apply(lambda fecha: fecha[:10])
            grupo_economico_sobrepasos = grupo_economico_sobrepasos[["Portafolio","Grupo Económico","FECHA SOBREPASO","Sobrepaso Actual","Sobrepaso \n%","TIPO SOBREPASO","Causa","Plan de Acción"]]
            grupo_economico_sobrepasos = grupo_economico_sobrepasos.rename(columns={"Portafolio":"PORTAFOLIO","Grupo Económico":"EMISOR","Sobrepaso Actual":"SOBREPASO PESOS","Sobrepaso \n%":"SOBREPASO PORCENTAJE","Causa":"CAUSA","Plan de Acción":"PLAN DE ACCION"})
            grupo_economico_sobrepasos["MENSAJE"] = grupo_economico_sobrepasos.apply(lambda dato: "El portafolio " + str(dato["PORTAFOLIO"]).strip() + " tiene un sobrepaso en el grupo económico " + str(dato["EMISOR"]).strip() + " de $" + format(abs(float(dato["SOBREPASO PESOS"])),",.0f") + " que corresponde a un " + format(dato["SOBREPASO PORCENTAJE"],".2%") ,axis =1 )

            insumo_sobrepasos = pd.concat([insumo_sobrepasos,grupo_economico_sobrepasos],axis=0,ignore_index=True)
            insumo_sobrepasos["EMISOR"] =  insumo_sobrepasos["EMISOR"].astype('str').str.upper()


        if "Calificación Sobrepasos" in nombre_hojas_sobrepasos_limites:

            calificacion_sobrepasos = pd.read_excel(ruta_sobrepasos_limites, sheet_name="Calificación Sobrepasos", skiprows=8)   
            calificacion_sobrepasos.columns = calificacion_sobrepasos.columns.str.strip()
            calificacion_sobrepasos["TIPO SOBREPASO"] = "CALIFICACION SOBREPASO"
            calificacion_sobrepasos = calificacion_sobrepasos[~calificacion_sobrepasos["Observaciones"].isna() & ~calificacion_sobrepasos["Portafolio"].isna()]
            calificacion_sobrepasos["FECHA SOBREPASO"] = calificacion_sobrepasos["Observaciones"].apply(lambda  observacion: str(observacion)[:10])
            calificacion_sobrepasos = calificacion_sobrepasos[["Portafolio","Instrumento","Calificación (Fitch)del título","FECHA SOBREPASO","TIPO SOBREPASO","Causa","Plan de Acción"]]
            calificacion_sobrepasos = calificacion_sobrepasos.rename(columns={"Portafolio":"PORTAFOLIO","Instrumento":"EMISOR","Calificación (Fitch)del título":"CALIFICACION","Causa":"CAUSA","Plan de Acción":"PLAN DE ACCION"})
            calificacion_sobrepasos["MENSAJE"] = calificacion_sobrepasos.apply(lambda dato: "El portafolio " + str(dato["PORTAFOLIO"]).strip() + " tiene el siguiente título " + str(dato["EMISOR"]).strip() + " el cual posee la siguiente calificación: " + dato["CALIFICACION"],axis =1 )
            insumo_sobrepasos = pd.concat([insumo_sobrepasos,calificacion_sobrepasos],axis=0,ignore_index=True)
            insumo_sobrepasos["EMISOR"] =  insumo_sobrepasos["EMISOR"].astype('str').str.upper()
            
            
        if "Duración Sobrepasos" in nombre_hojas_sobrepasos_limites:
            
            duracion_sobrepasos = pd.read_excel(ruta_sobrepasos_limites, sheet_name="Duración Sobrepasos", skiprows=7)
            duracion_sobrepasos.columns = duracion_sobrepasos.columns.str.strip()
            duracion_sobrepasos["TIPO SOBREPASO"] = "DURACION SOBREPASOS"
            duracion_sobrepasos = duracion_sobrepasos[~duracion_sobrepasos["Portafolio Mx"].isna()]
            duracion_sobrepasos["FECHA SOBREPASO"] = duracion_sobrepasos["Observaciones"].apply(lambda observacion: observacion[1:11])
            duracion_sobrepasos = duracion_sobrepasos[["Portafolio Mx","Sobrepaso","FECHA SOBREPASO","TIPO SOBREPASO","Causa","Plan de Acción"]]
            duracion_sobrepasos = duracion_sobrepasos.rename(columns={"Portafolio Mx":"PORTAFOLIO","Sobrepaso":"EMISOR","Causa":"CAUSA","Plan de Acción":"PLAN DE ACCION"})
            duracion_sobrepasos["MENSAJE"] = duracion_sobrepasos.apply(lambda dato: "El portafolio " + str(dato["PORTAFOLIO"]).strip() + " tiene un sobrepaso en duración de   " + format(float(dato["EMISOR"]),",.2f") + " dias",axis =1)
            insumo_sobrepasos = pd.concat([insumo_sobrepasos,duracion_sobrepasos],axis=0,ignore_index=True)
            
        
        if "Depósitos Sobrepasos" in nombre_hojas_sobrepasos_limites:
            
            depositos_sobrepasos = pd.read_excel(ruta_sobrepasos_limites, sheet_name="Depósitos Sobrepasos", skiprows=6)
            depositos_sobrepasos.columns = depositos_sobrepasos.columns.str.strip()
            depositos_sobrepasos["TIPO SOBREPASO"] = "DEPOSITOS SOBREPASOS"
            depositos_sobrepasos = depositos_sobrepasos[~depositos_sobrepasos["Portafolio MLC"].isna()]
            depositos_sobrepasos["FECHA SOBREPASO"] = depositos_sobrepasos["Observaciones"].apply(lambda observaciones: observaciones[:10])
            depositos_sobrepasos = depositos_sobrepasos[["Portafolio MLC","Emisor","NIT Emisor","Posición","Sobrepaso Actual","FECHA SOBREPASO","TIPO SOBREPASO","Causa","Plan de Acción"]]
            depositos_sobrepasos = depositos_sobrepasos.rename(columns = {"Portafolio MLC":"PORTAFOLIO","Emisor":"EMISOR","NIT Emisor":"NIT EMISOR","Posición":"SOBREPASO PESOS","Sobrepaso Actual":"SOBREPASO PORCENTAJE","Causa":"CAUSA","Plan de Acción":"PLAN DE ACCION"})
            
            depositos_sobrepasos["MENSAJE"] = depositos_sobrepasos.apply(lambda dato: "El portafolio " + str(dato["PORTAFOLIO"]).strip() + " tiene un sobrepaso de depósito en " + str(dato["EMISOR"]).strip() + " por un valor de $" + format(abs(float(dato["SOBREPASO PESOS"])),",.0f") + " que corresponde a un " + format(dato["SOBREPASO PORCENTAJE"],".2%"),axis =1 )
            insumo_sobrepasos = pd.concat([insumo_sobrepasos,depositos_sobrepasos],axis=0,ignore_index=True)
            insumo_sobrepasos["EMISOR"] =  insumo_sobrepasos["EMISOR"].astype('str').str.upper()
                
        insumo_sobrepasos["TOTAL DIAS TRANSCURRIDOS"] = insumo_sobrepasos["FECHA SOBREPASO"].apply(lambda fecha_sobrepaso: ( datetime.now().date() - datetime.strptime(fecha_sobrepaso,"%d/%m/%Y").date()).days)
        insumo_sobrepasos["PORTAFOLIO"] = insumo_sobrepasos["PORTAFOLIO"].apply(lambda portafolio: portafolio[:3])
        return insumo_sobrepasos
    
    except Exception as e:
        
        mostrarMensajeAdvertencia('No se pudo cargar los archivos de sobrepasos a la herramienta: ' +str(e))
        return insumo_sobrepasos

def cargarOperacionesPorCumplirFiduciaria(ruta_insumo,nemoTitulosFiduciaria):
    
    archivo_operaciones_por_cumplir_fiduciaria = obtenerUltimoArchivoCarpeta(ruta_insumo)        
    fechaOpPendientesFidu = archivo_operaciones_por_cumplir_fiduciaria.split(".")
    definiciones.fechaOpPendientesFidu = fechaOpPendientesFidu[0][-8:]
    operaciones_por_cumplir_fiduciaria = pd.read_excel(archivo_operaciones_por_cumplir_fiduciaria,skiprows=1) 
    
    
    if len(operaciones_por_cumplir_fiduciaria) > 0:
        operaciones_por_cumplir_fiduciaria = operaciones_por_cumplir_fiduciaria[operaciones_por_cumplir_fiduciaria["Mod"]!= "---"]
        operaciones_por_cumplir_fiduciaria.columns =  operaciones_por_cumplir_fiduciaria.columns.str.strip().str.upper()
        operaciones_por_cumplir_fiduciaria =  limpiarDatos(operaciones_por_cumplir_fiduciaria,colTrim=["TRANSACCIO","POR","MON","ESPECIE"],colUpper=["TRANSACCIO","POR","MON","ESPECIE"],colFloat=["VR NOMINAL ACTUAL","VR TRANSACCION"])
        operaciones_por_cumplir_fiduciaria["PORTAFOLIO"] = operaciones_por_cumplir_fiduciaria["POR"].apply(lambda x: x[0:3])
        
        #Este cruce se realiza para conocer el saldo en pesos al mismo precio que se encuentra en el inventario
        operaciones_por_cumplir_fiduciaria = operaciones_por_cumplir_fiduciaria.join(nemoTitulosFiduciaria[["Especie/Generador","Vr Mercado Hoy Moneda Empresa","Nominal Remanente"]].drop_duplicates("Especie/Generador").set_index("Especie/Generador",drop=True),on="ESPECIE",how="left")
        operaciones_por_cumplir_fiduciaria["PRECIO"] = operaciones_por_cumplir_fiduciaria["Vr Mercado Hoy Moneda Empresa"]/operaciones_por_cumplir_fiduciaria["Nominal Remanente"]
        operaciones_por_cumplir_fiduciaria["VALOR MERCADO"] = operaciones_por_cumplir_fiduciaria["PRECIO"] * operaciones_por_cumplir_fiduciaria["VR NOMINAL ACTUAL"]
        operaciones_por_cumplir_fiduciaria = operaciones_por_cumplir_fiduciaria[["FECHA","ESPECIE","TRANSACCIO","PORTAFOLIO","MON","VALOR MERCADO","VR NOMINAL ACTUAL"]]
    else:
        operaciones_por_cumplir_fiduciaria = pd.DataFrame([],columns=["FECHA","ESPECIE","TRANSACCIO","PORTAFOLIO","MON","VALOR MERCADO","VR NOMINAL ACTUAL"]) 
    
    resultado = validarInformacionOperacionesPorCumplir(archivo_operaciones_por_cumplir_fiduciaria)
    if resultado == True:
        return operaciones_por_cumplir_fiduciaria
    else:
        return pd.DataFrame([],columns= ["FECHA","ESPECIE","TRANSACCIO","PORTAFOLIO","MON","VALOR MERCADO","VR NOMINAL ACTUAL"]) 

def cargarOperacionesPorCumplirValores(ruta_insumo, nemoTitulosValores):
    
    archivoOperacionesPorCumpirValores = obtenerUltimoArchivoCarpeta(ruta_insumo)
    operacionesPorCumplirValores = pd.read_excel(archivoOperacionesPorCumpirValores)
    operacionesPorCumplirValores.columns = operacionesPorCumplirValores.columns.str.strip() 
    operacionesPorCumplirValores = limpiarDatos(operacionesPorCumplirValores,colTrim=["Nemoténico","Moneda Especie","Isin","Nombre Emisor","Clase Tasa","Estado Título"],colUpper=["Nemoténico","Moneda Especie","Isin","Nombre Emisor","Clase Tasa","Estado Título"],colFloat=["Valor Nominal Actual"])
    
    operacionesPorCumplirValores = operacionesPorCumplirValores[operacionesPorCumplirValores["Clase Tasa"] == "PRECIO ACCIÓN"]
    operacionesPorCumplirValores = operacionesPorCumplirValores[operacionesPorCumplirValores["Estado Título"].isin(["VENTA POR CUMPLIR","COMPRA POR CUMPLIR"])]
    
    if len(operacionesPorCumplirValores) > 0:
        operacionesPorCumplirValores = operacionesPorCumplirValores.join(nemoTitulosValores[["Nemoténico","Valor Nominal Actual","Valor VPN Actual"]].drop_duplicates(["Nemoténico"]).set_index("Nemoténico",drop=True),on="Nemoténico",how='left',rsuffix= " Inventario")
        operacionesPorCumplirValores["PRECIO"] = operacionesPorCumplirValores["Valor VPN Actual Inventario"]/ operacionesPorCumplirValores["Valor Nominal Actual Inventario"]
        operacionesPorCumplirValores["VALOR MERCADO"] = operacionesPorCumplirValores["PRECIO"] * operacionesPorCumplirValores["Valor Nominal Actual"]
        operacionesPorCumplirValores = operacionesPorCumplirValores[["Fecha","Fecha Cumplimiento","Nemoténico","Isin","Nombre Emisor","Estado Título","Codigo OyD","Moneda Especie","VALOR MERCADO","Valor Nominal Actual"]]
    else:
        operacionesPorCumplirValores = pd.DataFrame([],columns=["Fecha","Fecha Cumplimiento","Nemoténico","Isin","Nombre Emisor","Estado Título","Codigo OyD","Moneda Especie","VALOR MERCADO","Valor Nominal Actual"])
    
    return operacionesPorCumplirValores
       

def validarInformacionOperacionesPorCumplir(nombre_archivo):
    #obtener fecha de creación
    fecha_creacion = datetime.fromtimestamp(os.path.getctime(nombre_archivo))
    fecha_creacion = fecha_creacion.date()
    #obtener fecha del nombre del insumo
    fecha_nombre = nombre_archivo.split(".xlsx")[0][-8:]
    fecha_nombre = datetime.strptime(fecha_nombre,"%Y%m%d").date()
    
    hoy = date.today()
    fecha_ayer = date.today() - timedelta(days=1)
    
    if fecha_creacion == hoy and fecha_ayer == fecha_nombre:
        return True
    else:
        return False
    
def agregarOpercionesPendientesPorCumplirFidu(nemoTitulosFiduciaria,operacionesPorCumplirFidu):
    
    #Se dejan las operaciones pendientes que se sumplen hoy
    #Se agregan las operaciones como si fueran un título adicional, si es una venta el valor nominal se ingresa negativo, si es una compra el valor nominal se ingresa positivo
    #Asi ya quedan reflejadas las operaciones pendientes por cumplir en el inventario de titulos
    hoy = date.today()
    #operacionesPorCumplirFidu = operacionesPorCumplirFidu[operacionesPorCumplirFidu["FECHA"].apply(lambda fecha: True if fecha.date() == hoy else False )]  
    if len(operacionesPorCumplirFidu) == 0:
        return nemoTitulosFiduciaria
    
    operacionesPorCumplirFidu.loc[operacionesPorCumplirFidu["TRANSACCIO"] == "VENTA","VR NOMINAL ACTUAL"] =  -operacionesPorCumplirFidu.loc[operacionesPorCumplirFidu["TRANSACCIO"] == "VENTA","VR NOMINAL ACTUAL"]
    operacionesPorCumplirFidu["Macro Activo"] = operacionesPorCumplirFidu["MON"].apply(lambda moneda: "RV LOCAL" if moneda == "COP" else "RV INTERNACIONAL")
    operacionesPorCumplirFidu["Nemotécnico"] = 'NAN'
    operacionesPorCumplirFidu = operacionesPorCumplirFidu.join(nemoTitulosFiduciaria[["Especie/Generador","ISIN","Emisor / Contraparte","Emisor Unificado"]].drop_duplicates(["Especie/Generador"]).set_index("Especie/Generador",drop=True),on=["ESPECIE"],how='left')
    operacionesPorCumplirFidu["SALDO Macro Activo"] = operacionesPorCumplirFidu["VALOR MERCADO"]
    operacionesPorCumplirFidu["SALDO ABA"] = operacionesPorCumplirFidu["VALOR MERCADO"]
    operacionesPorCumplirFidu.loc[operacionesPorCumplirFidu["TRANSACCIO"] == "VENTA","SALDO Macro Activo"] = -operacionesPorCumplirFidu.loc[operacionesPorCumplirFidu["TRANSACCIO"] == "VENTA","SALDO Macro Activo"]
    operacionesPorCumplirFidu.loc[operacionesPorCumplirFidu["TRANSACCIO"] == "VENTA","SALDO ABA"] = -operacionesPorCumplirFidu.loc[operacionesPorCumplirFidu["TRANSACCIO"] == "VENTA","SALDO ABA"]
    operacionesPorCumplirFidu.rename(columns={"PORTAFOLIO":"Portafolio","ESPECIE":"Especie/Generador","VR NOMINAL ACTUAL":"Nominal Remanente"},inplace=True)
    operacionesPorCumplirFidu["Origen Informacion"] = "OPERACIONES PENDIENTES"
    nemoTitulosFiduciaria = pd.concat([nemoTitulosFiduciaria,operacionesPorCumplirFidu[["Portafolio","Especie/Generador","Emisor / Contraparte","Emisor Unificado","Nemotécnico","ISIN","Macro Activo","SALDO Macro Activo","SALDO ABA","Nominal Remanente","Origen Informacion"]]],ignore_index=True)
   
    return nemoTitulosFiduciaria
    
def agregarOpercionesPendientesPorCumplirValores(nemoTitulosValores, operacionesPorCumplirValores):
    
    hoy = date.today()
    #operacionesPorCumplirValores = operacionesPorCumplirValores[operacionesPorCumplirValores["Fecha Cumplimiento"].apply(lambda fecha: True if fecha.date() == hoy else False)] 
    
    if len(operacionesPorCumplirValores) == 0:
        return nemoTitulosValores
    
    operacionesPorCumplirValores["Macro Activo"] = operacionesPorCumplirValores["Moneda Especie"].apply(lambda moneda: "RV LOCAL" if moneda == "COP" else "RV INTERNACIONAL")
    operacionesPorCumplirValores = operacionesPorCumplirValores.join(nemoTitulosValores[["Nombre Emisor","Emisor Unificado"]].drop_duplicates(["Nombre Emisor"]).set_index("Nombre Emisor",drop=True),on =["Nombre Emisor"],how ='left')
    operacionesPorCumplirValores["SALDO Macro Activo"] = operacionesPorCumplirValores["VALOR MERCADO"]
    operacionesPorCumplirValores["SALDO ABA"] = operacionesPorCumplirValores["VALOR MERCADO"]
    #operacionesPorCumplirValores.loc[operacionesPorCumplirValores["Estado Título"] =="VENTA POR CUMPLIR","Saldo Macro Activo"] = - operacionesPorCumplirValores.loc[operacionesPorCumplirValores["Estado Título"] =="VENTA POR CUMPLIR","Saldo Macro Activo"]
    #operacionesPorCumplirValores.loc[operacionesPorCumplirValores["Estado Título"] =="VENTA POR CUMPLIR","Saldo ABA"] = - operacionesPorCumplirValores.loc[operacionesPorCumplirValores["Estado Título"] =="VENTA POR CUMPLIR","Saldo ABA"]
    operacionesPorCumplirValores = operacionesPorCumplirValores.rename(columns = {"Codigo OyD":"Código OyD"})
    operacionesPorCumplirValores["Origen Informacion"] = "OPERACIONES PENDIENTES"
    nemoTitulosValores = pd.concat([nemoTitulosValores,operacionesPorCumplirValores[["Código OyD","Nemoténico","Isin","Nombre Emisor","Emisor Unificado","Macro Activo","SALDO Macro Activo","SALDO ABA","Valor Nominal Actual","Origen Informacion"]]],ignore_index=True)
    
    
    return nemoTitulosValores

def cargarOperacionesVigentes():
    
    #Se cargan primero todas las intenciones   
    usuarios = definiciones.usuarios 
    gerentesAConsultar = usuarios[usuarios["Rol"].isin(["Administrador","PM"])].index.tolist()      
    ruta = definiciones.parametros["Valor"]["rutaIntenciones"]
    listaCSVs = list(map(lambda x:  x+ ".csv",gerentesAConsultar))
    intenciones = obtenerIntenciones(ruta,listaCSVs)
    
    #Se dejan solo las intenciones que están en los siguientes estados      
    intenciones = intenciones[intenciones["Estado"].isin(["Nueva","Modificada","Renovada","En proceso","Ejecutada/Parcial"])]
    
    #Solo nos interesa tener en cuenta las ventas y los retiros solamente
    intenciones = intenciones[intenciones["TipoOperacion"].isin(["VENTA"])]
    
    #Solo nos interesan 3 mercados: "Renta Variable","Deuda Privada","Deuda Pública"
    intenciones = intenciones[intenciones["TipoActivo"].isin(["Renta Variable","Deuda Pública","Deuda Privada"])] 
    return intenciones

def agregarintencionesVigentesFidu(nemoTitulosFiduciaria, intencionesVigentes):
    
   
    intencionesVigentes = intencionesVigentes.copy()
    especies = definiciones.especies
    intencionesVigentes = intencionesVigentes.join(especies[["Especie","Nemotecnico","Nemo intenciones","Isin","Macro Activo inventario"]].drop_duplicates("Nemo intenciones").set_index("Nemo intenciones",drop= False),on="Nemotecnico",how='left',rsuffix=" especies")
    
    #Para la deuda privada la espicie es la misma que se colocó en el nombre del títtulo
    intencionesVigentes.loc[intencionesVigentes["Especie"].isna(),"Especie"] = intencionesVigentes[intencionesVigentes["Especie"].isna()].apply(lambda intencion: intencion["Nemotecnico"] if intencion["TipoActivo"] == "Deuda Privada"  else intencion["Especie"],axis=1)
    
    #Necesitamos asignar el macro activo
    intencionesVigentes["Macro Activo"] = ""
    intencionesVigentes.loc[intencionesVigentes["TipoActivo"] == "Deuda Privada","Macro Activo"]  = "DEUDA PRIVADA"
    intencionesVigentes.loc[intencionesVigentes["TipoActivo"] == "Deuda Pública","Macro Activo"]  = "DEUDA PÚBLICA"
    intencionesVigentes.loc[(intencionesVigentes["TipoActivo"] == "Renta Variable") &(intencionesVigentes["Mercado"]=="LOCAL"),"Macro Activo"]  = "RV LOCAL"
    intencionesVigentes.loc[(intencionesVigentes["TipoActivo"] == "Renta Variable") &(intencionesVigentes["Mercado"]=="INTERNACIONAL"),"Macro Activo"]  = "RV INTERNACIONAL"
    #Debemos cruzar por medio del Nemo con la base de especies para obtener el Nemo oficial y la especie 
    intencionesVigentes = intencionesVigentes.rename(columns={"Portafolio":"Nombre Portafolio","CodPortafolio":"Portafolio","Especie":"Especie/Generador","Nemotecnico especies":"Nemotécnico","Isin":"ISIN","TipoOperacion":"POSICIÓN"})
    intencionesVigentes["Origen Informacion"] = "INTENCIONES VIGENTES"
    intencionesVigentes["CantidadTotal"] = intencionesVigentes["CantidadTotal"].astype("float")
    intencionesVigentes["Nominal Remanente"] = intencionesVigentes[["Denominacion","CantidadTotal"]].apply(lambda intencion: intencion["CantidadTotal"]*1000000 if "MM" in intencion["Denominacion"] else intencion["CantidadTotal"],axis=1)
    intencionesVigentes.loc[intencionesVigentes["POSICIÓN"] == "VENTA","Nominal Remanente"] =  -intencionesVigentes.loc[intencionesVigentes["POSICIÓN"] == "VENTA","Nominal Remanente"]
    #intencionesVigentes.to_excel("C:/Users/frcastro/downloads/ope_intradia_fidu.xlsx")
    nemoTitulosFiduciaria = pd.concat([nemoTitulosFiduciaria,intencionesVigentes[["Portafolio","Especie/Generador","Nemotécnico","ISIN","Macro Activo","Nominal Remanente","POSICIÓN","Origen Informacion"]]],ignore_index=True)
    return nemoTitulosFiduciaria    
    
def agregarintencionesVigentesValores(nemoTitulosValores, intencionesVigentes):
    
    intencionesVigentes = intencionesVigentes.copy()
    
    especies = definiciones.especies
    intencionesVigentes = intencionesVigentes.join(especies[["Especie","Nemotecnico","Nemo intenciones","Isin","Macro Activo inventario"]].drop_duplicates("Nemo intenciones").set_index("Nemo intenciones",drop= True),on="Nemotecnico",how='left',rsuffix=" especies")
    
    #Asignación de macroactivo
    intencionesVigentes.loc[intencionesVigentes["TipoActivo"] == "Deuda Privada","Macro Activo"]  = "DEUDA PRIVADA"
    intencionesVigentes.loc[intencionesVigentes["TipoActivo"] == "Deuda Pública","Macro Activo"]  = "DEUDA PÚBLICA"
    intencionesVigentes.loc[(intencionesVigentes["TipoActivo"] == "Renta Variable") &(intencionesVigentes["Mercado"]=="LOCAL"),"Macro Activo"]  = "RV LOCAL"
    intencionesVigentes.loc[(intencionesVigentes["TipoActivo"] == "Renta Variable") &(intencionesVigentes["Mercado"]=="INTERNACIONAL"),"Macro Activo"]  = "RV INTERNACIONAL"
    
    #Debemos cruzar por medio del Nemo con la base de especies para obtener el Nemo oficial y la especie 
    intencionesVigentes = intencionesVigentes.rename(columns={"Portafolio":"Nombre Portafolio","CodPortafolio":"Portafolio","Nemotecnico especies":"Nemoténico","TipoOperacion":"POSICIÓN"})
    intencionesVigentes["Origen Informacion"] = "INTENCIONES VIGENTES"
    intencionesVigentes["CantidadTotal"] = intencionesVigentes["CantidadTotal"].astype("float")
    intencionesVigentes["Valor Nominal Actual"] = intencionesVigentes[["Denominacion","CantidadTotal"]].apply(lambda intencion: intencion["CantidadTotal"]*1000000 if "MM" in intencion["Denominacion"] else intencion["CantidadTotal"],axis=1)
    intencionesVigentes.loc[intencionesVigentes["POSICIÓN"] == "VENTA","Valor Nominal Actual"] =  -intencionesVigentes.loc[intencionesVigentes["POSICIÓN"] == "VENTA","Valor Nominal Actual"]
    
    nemoTitulosValores = pd.concat([nemoTitulosValores,intencionesVigentes[["Portafolio","Nemoténico","Isin","Macro Activo","Valor Nominal Actual","POSICIÓN","Origen Informacion"]]],ignore_index=True)
    
    return nemoTitulosValores

def obtenerUltimoArchivoCarpeta(nombre_generico):
    
    archivos = glob.glob(nombre_generico)
    fecha_creacion = list(map(lambda x:os.path.getctime(x),archivos))
    ultimo_archivo = archivos[fecha_creacion.index(max(fecha_creacion))]
    return ultimo_archivo
    
def obtenerUltimoInventarioTitulos(ruta,rowsSkip):
    archivos = glob.glob(ruta)
    fechaCreacion = list(map(lambda x:os.path.getctime(x),archivos))
    ultimoArchivo = archivos[fechaCreacion.index(max(fechaCreacion))]
     
    try:
        data = pd.read_excel(ultimoArchivo,skiprows=rowsSkip )
    except:
        mostrarMensajeAdvertencia("No se encontro el archivo :" + ruta)
        sys.exit()
    return data

def actualizarEstadoIntencionesVencidas(usuariosPortafolios):
    
    ruta = definiciones.parametros["Valor"]["rutaIntenciones"]
    for usuarioPortafolio in usuariosPortafolios:
        intenciones = obtenerIntenciones(ruta, [usuarioPortafolio+".csv"])
        if definiciones.parametros["Valor"]["camposArchivoIntenciones"].split("-") == list(intenciones.columns):
            if len(intenciones) > 0:
                intencionesVencidas = intenciones[~intenciones["Estado"].isin(["Cancelada","Vencida","Ejecutada/Total"])]
                intencionesVencidas = intencionesVencidas[intencionesVencidas.loc[:,["VigenteHasta"]].apply(lambda x:True if datetime.strptime(x["VigenteHasta"], '%d/%m/%Y').date() < date.today() else False,axis=1)]
                intencionesVencidas["Estado"] = "Vencida"
                intencionesVencidas["UltimaModificacion"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
                if len(intencionesVencidas) > 0:
                    intenciones.loc[:,["Estado"]] = intenciones.apply(lambda x: "Vencida" if  x["Estado"] not in ["Cancelada","Vencida","Ejecutada/Total"] and datetime.strptime(x["VigenteHasta"], '%d/%m/%Y').date() < date.today() else x["Estado"],axis =1 )
                    intenciones.to_csv(ruta+"/"+usuarioPortafolio+".csv",index=False)
                    
                    #Guardamos trazabilidad de intenciones vencidas
                    guardarIntenciones(intencionesVencidas,definiciones.parametros["Valor"]["rutaIntencionesTrazabilidad"],usuarioPortafolio)
                    
        else:
            
            print("Existe un error con el siguiente archivo: " + usuarioPortafolio+".csv")
            print("***********************")
            
def actualizarLogIntenciones(idIntencion,quienModifica,accion):
    
    logIntenciones = definiciones.parametros["Valor"]["logIntenciones"]
    df = pd.DataFrame([])
    df.loc[0,"Tiempo"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    df.loc[0,"Id"] = str(idIntencion)
    df.loc[0,"QuienModifica"] = quienModifica
    df.loc[0,"Accion"] = accion
    
    df.to_csv(logIntenciones, header=None, index=None, sep=' ', mode='a')
  
def calcularPrecioAcciones(fidu,valores):

    fidu = fidu[["Especie/Generador","ISIN","Nominal Remanente","Vr Mercado Hoy Moneda Empresa"]].copy()
    valores = valores[["Nemoténico","Isin","Valor Nominal Actual","SALDO Macro Activo"]].copy()
    fidu.rename(columns = {"Especie/Generador":"Titulo","ISIN":"Isin","Nominal Remanente":"Nominal","Vr Mercado Hoy Moneda Empresa":"Valor"},inplace  =True)
    valores.rename(columns = {"Nemoténico":"Titulo","Valor Nominal Actual":"Nominal","SALDO Macro Activo":"Valor"},inplace = True)
    titulos = pd.concat([fidu,valores],axis  =0)
    titulos.drop_duplicates(["Titulo"],inplace = True)
    titulos["Nominal"] = pd.to_numeric(titulos["Nominal"],errors='coerce')
    titulos["Valor"] = pd.to_numeric(titulos["Valor"],errors='coerce')
    titulos = titulos.loc[(titulos["Valor"]!= "NAN") & (titulos["Nominal"]!= "NAN"),]
    titulos = titulos[titulos["Nominal"]  != 0]
    titulos["PrecioAccion"] = titulos["Valor"]/titulos["Nominal"]
    titulos = titulos[titulos["Isin"]!="NAN"]
    return titulos
 
def crearValidadorConFiltro(hojaEspecies,datos,rango):

    hojaEspecies.Range("A:A").ClearContents()
    hojaEspecies.Range("A2:A" + str(len(datos)+1)).Value = pd.DataFrame(datos).to_records(index=False)
    hojaEspecies.Range("B1").Formula2R1C1 = "=UNIQUE(FILTER(OFFSET(RC[-1],1,0,COUNTA(C[-1]),1),ISNUMBER(SEARCH('Formulario Ordenes'!R[" +str(rango.Row-1)+"]C["+ str(rango.Column -2)+"],OFFSET(RC[-1],1,0,COUNTA(C[-1]),1)))))" #Frank       
    numDatos = hojaEspecies.Cells(hojaEspecies.Rows.Count, "B").End(definiciones.xlUp).Row
    Formula1 = "='Especies'!B1:B" + str(numDatos)
    return Formula1         
    

def actualizarVistaGerentes():
   
    intencionesAM= win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        
    hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
    hojaOrdenes.Unprotect()
    gerenteAConsultar = hojaOrdenes.OLEObjects('listaGerentes').Object.Value
    fechaDesde = hojaOrdenes.OLEObjects('fechaDesdeIntenciones').Object.Value
    if gerenteAConsultar == "":
        mostrarMensajeAdvertencia("Por favor seleccione un nombre de gerente")
        return
    if fechaDesde == "":
        mostrarMensajeAdvertencia("Por favor seleccione una fecha.")
        return
    #validaciones:
    #la fecha desde debe ser menor o igual a hoy        
    fechaDesde = ValidarFechaDesdeParaMostrar(fechaDesde)
    
    #Si se solicita mostrar las intenciones de todos los gerentes, entonces debo obtener la lista de todos los gerentes
    if gerenteAConsultar == "(Todos)":
        usuarios = definiciones.usuarios 
        gerentesAConsultar = usuarios[usuarios["Rol"].isin(["Administrador", "PM"])].index.tolist()
    else:
        gerentesAConsultar = [gerenteAConsultar]
    
    
    data = obtenerIntencionesGerentesParaVisualizar(fechaDesde,gerentesAConsultar,definiciones.parametros["Valor"]["rutaIntenciones"],definiciones.parametros["Valor"]["nombreCamposVerGerentes"]) #Esta función se encarga de hacer las validaciones y llamar los archivos csv
    data = data.where(~data.isna(), other="")
    primeraFilaTabla = 8
    limpiarHojaIntencionesPM(hojaOrdenes, primeraFilaTabla)
    if len(data) == 0:
        print("No hay intenciones para mostrar")
        return

    hojaOrdenes.Range(hojaOrdenes.Cells(8,2),hojaOrdenes.Cells(len(data)+7,1 + data.shape[1])).Value = data.to_records(index=False)         
    hojaOrdenes.UsedRange.NumberFormat ="@"
    columnaCantidad = 10
    columnaPorcentaje = 4
    columnaFecha = 21
    hojaOrdenes.Columns(columnaFecha).NumberFormat = "mm/dd/yyyy"
    hojaOrdenes.Columns(columnaPorcentaje).NumberFormat = "0.00"
    hojaOrdenes.Columns(columnaCantidad).NumberFormat = "#,##0.00" #$10.00
    hojaOrdenes.Range("B:B").NumberFormat = "0"
    hojaOrdenes.Range("E4:E5").NumberFormat = "mm/dd/yyyy"
    hojaOrdenes.Columns.AutoFit()
    
     

def actualizarVistaTraders():
    usuarios = definiciones.usuarios
    intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
    hojaMonitor = intencionesAM.Worksheets("Monitor")
    mercadoAConsultar = hojaMonitor.OLEObjects("lstMacroActivosTrader").Object.Value
    periodoAConsultarDesde = hojaMonitor.OLEObjects("txtFechaDesdeMonitor").Object.Value
    
    if not esUnaFecha(periodoAConsultarDesde):
        mostrarMensajeAdvertencia("Por favor ingrese una fecha desde correcta")
        return
       
    gerentesAConsultar = usuarios[usuarios["Rol"].isin(["Administrador","PM"])].index.tolist()
    #Que pasa si se solicita todos los mercados? Pues no se aplica ningún filtro
    #Validaciones
    #fecha a visualizar
    periodoAConsultarDesde = ValidarFechaDesdeParaMostrar(periodoAConsultarDesde)
    
    if mercadoAConsultar =="RV y Fondos":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoRV"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersRV"].split("-")        
    if mercadoAConsultar =="Deuda Privada":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoDPR"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersDPR"].split("-")
    if mercadoAConsultar =="Deuda Pública":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoDPU"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersDPU"].split("-")
    if mercadoAConsultar == "Fondos":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoFondos"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersFondos"].split("-")
    if mercadoAConsultar =="Forex":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoForex"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersForex"].split("-")
    if mercadoAConsultar =="Swaps":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoSwaps"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersSwaps"].split("-")
    if mercadoAConsultar =="Todos":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersTodosArchivo"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersTodos"].split("-")  
    if mercadoAConsultar == "RF Internacional":
        camposMostrar = definiciones.parametros["Valor"]["ColumnasMostrarTradersArchivoRF"].split("-")
        encabezado = definiciones.parametros["Valor"]["ColumnasMostrarTradersRF"].split("-") 
    columnaCantidad = encabezado.index("Cantidad") +2 
    columnaEjecutado = encabezado.index("Ejecutado") +2
    columnaPendiente = encabezado.index("Pendiente") + 2 
    columnaPorcentaje = encabezado.index("% Ejec") + 2
    columnaFechaVigencia = encabezado.index("Vigente hasta") +2
    
    for campo in encabezado:
        hojaMonitor.OLEObjects("cmboxCamposVista").Object.AddItem(campo)
    datosFiltro = hojaMonitor.OLEObjects("DatosFiltro").Object.Value
    campoFiltro = hojaMonitor.OLEObjects("cmboxCamposVista").Object.Value
    datosFiltro = datosFiltro.split(";")
    
    data = obtenerIntencionesTradersParaVisualizar(gerentesAConsultar,mercadoAConsultar,periodoAConsultarDesde,definiciones.parametros["Valor"]["rutaIntenciones"],camposMostrar)        
    
    #Aplicar el filtro que se muestra en al Hoja
    if campoFiltro != '':
        if datosFiltro != ['']:
            data = data[data[camposMostrar[encabezado.index(campoFiltro)]].isin(datosFiltro)]
               
    if len(data) == 0:
        mostrarMensajeAdvertencia("No hay intenciones para mostrar")
        return
    
    filaDesde = 7
    limpiarHojaIntencionesPM(hojaMonitor, filaDesde)
    
    celdasEncabezado = hojaMonitor.Range(hojaMonitor.Cells(7,2),hojaMonitor.Cells(7,len(encabezado)+1))
    celdasEncabezado.Value = encabezado
    
    hojaMonitor.Range(hojaMonitor.Cells(8,2),hojaMonitor.Cells(len(data)+7,1 + data.shape[1])).Value = data.to_records(index=False) 
   
    #Dar formato a la nueva tabla
    celdasEncabezado.Font.Name ="CIBFont Sans"
    celdasEncabezado.Font.Size = 11
    celdasEncabezado.Font.ThemeColor = 1
    celdasEncabezado.Font.Bold = True
    celdasEncabezado.Interior.Color = 0x2C2A29
    celdasEncabezado.ColumnWidth = 20
    celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
    hojaMonitor.Range("B:B").NumberFormat = "0"
    hojaMonitor.Range("C:AZ").NumberFormat = "@"
    hojaMonitor.Columns(columnaFechaVigencia).NumberFormat = "mm/dd/yyyy"
    hojaMonitor.Columns(columnaCantidad).NumberFormat =  "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    hojaMonitor.Columns(columnaEjecutado).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"#$10.00
    hojaMonitor.Columns(columnaPendiente).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"#$10.00
    hojaMonitor.Columns(columnaPorcentaje).NumberFormat = "0.00" # dos decimales
    hojaMonitor.Columns.AutoFit() 
    hojaMonitor.Range(hojaMonitor.Cells(8,data.shape[1]+2),hojaMonitor.Cells(8,data.shape[1]+30)).EntireColumn.Hidden = True
  
        

def mostrarMensajeAdvertencia(mensaje):
    
    root = tk.Tk()
    root.withdraw()  # Hide the main 
    root.attributes('-topmost',True)
    messagebox.showinfo("Notificación",mensaje)


def mostrarMensajeAdvertenciaSiNo(mensaje):
    
    respuesta = win32api.MessageBox(None, mensaje, "Pregunta", win32con.MB_YESNO | win32con.MB_ICONQUESTION |  win32con.MB_SYSTEMMODAL)
    return respuesta == win32con.IDYES
    
def mostrarMensajeAdvertenciaSiNoReconfirmacion(mensaje):
    
    respuesta = win32api.MessageBox(None, mensaje, "Advertencia", win32con.MB_YESNO | win32con.MB_ICONWARNING |  win32con.MB_SYSTEMMODAL)
    
    return respuesta == win32con.IDYES
    
 

def ValidarFechaDesdeParaMostrar(fechaDesde):
    hoy = date.today()
    if esUnaFecha(fechaDesde):
        fechaDesde =  datetime.strptime(fechaDesde, '%d/%m/%Y').date()
    else:
        mostrarMensajeAdvertencia("La fecha desde no fué reconocida.")
        fechaDesde = hoy
    
    difFechas = hoy - fechaDesde
    if difFechas.days < 0:
        mostrarMensajeAdvertencia("La fecha desde no puede ser superior al dia de hoy.")
        fechaDesde = hoy
    return fechaDesde

def validarFechaHastaParaMostrar(fechaHasta,fechaDesde):
    hoy = date.today()
    if esUnaFecha(fechaHasta):
        fechaHasta =  datetime.strptime(fechaHasta, '%d/%m/%Y').date()
    else:
        mostrarMensajeAdvertencia("La fecha hasta no fué reconocida.")
        fechaHasta = hoy
        
    difFechas = hoy - fechaHasta
    if difFechas.days < 0:
        mostrarMensajeAdvertencia("La fecha hasta no puede ser superior al dia de hoy.")
        fechaHasta = hoy
    
    difFechas = fechaHasta -fechaDesde
    if difFechas.days < 0:
        mostrarMensajeAdvertencia("La fecha hasta no puede ser inferior a la fecha desde.")
        fechaHasta = hoy
    
    return fechaHasta

def validarFechaDesdeHastaNueva(fechas):
    #Valida que la fecha ingresada sea igual osuperior a hoy, si es asi devuelve True, de lo contrario devulve False
    
    fechaDesde = fechas[0]
    fechaHasta = fechas[1]    
    if esUnaFecha(fechaDesde) and esUnaFecha(fechaHasta):
        fechaDesde = datetime.strptime(fechaDesde, '%d/%m/%Y').date()
        fechaHasta = datetime.strptime(fechaHasta, '%d/%m/%Y').date()
    elif esUnaFecha(fechaDesde) and not esUnaFecha(fechaDesde):
        return [True, False]
    elif not esUnaFecha(fechaDesde) and esUnaFecha(fechaDesde):
        return [False, True]
    else:
        return [False, False]
    
    hoy = date.today()
    difFechas =  fechaDesde - hoy
    difFechasVigencia = fechaHasta - fechaDesde
    if difFechas.days >= 0 and difFechasVigencia.days >= 0: 
        return [True, True]
    elif difFechas.days < 0 and difFechasVigencia.days >= 0: 
        return [False, True]
    elif difFechas.days >= 0 and difFechasVigencia.days < 0: 
        return [True, False]
    else:
        return [False, False]
      

    

def esUnaFecha(fecha):
    try:
        fechaDesde =  datetime.strptime(fecha, '%d/%m/%Y').date()
        return True
    except:
        return False
    
def esUnNumero(x):
    
    try:
        float(x)
        return True
    except:
        return False

def esUnaLista(x):
    
    try:
        len(x)
        return True
    except:
        return False
    
    
def obtenerIntencionesParaDescargar(hojaDescargas):
    #1.Validar datos de entrada
    #2.Obtener dataframe de datos
    #3.Mostrar info tabla en la hoja de descargas
    
    #1.
    fechaDesde = hojaDescargas.OLEObjects('txtFechaDesdeDescargas').Object.Value 
    fechaHasta = hojaDescargas.OLEObjects("txtFechaHastaDescargas").Object.Value
    mercadoAConsultar = hojaDescargas.OLEObjects("lstMacroActivoDescargas").Object.Value
    gerentesAConsultar = hojaDescargas.OLEObjects("lstGerentesDescarga").Object.Value
    
    if not esUnaFecha(fechaDesde):
        mostrarMensajeAdvertencia("Por favor ingrese una fecha desde correcta")
        return pd.DataFrame([])   
    if not esUnaFecha(fechaHasta):
        mostrarMensajeAdvertencia("Por favor ingrese una fecha hasta correcta")
        return pd.DataFrame([])
    
    if gerentesAConsultar == "(Todos)":
        usuarios = definiciones.usuarios 
        gerentesAConsultar = usuarios[usuarios["Rol"].isin(["Administrador","PM"])].index.tolist()
    else:
        gerentesAConsultar = [gerentesAConsultar]
    fechaDesde = ValidarFechaDesdeParaMostrar(fechaDesde)
    fechaHasta = validarFechaHastaParaMostrar(fechaHasta,fechaDesde)  
        
    camposMostrar = definiciones.parametros["Valor"]["camposArchivoIntenciones"].split("-")
    
    ruta = definiciones.parametros["Valor"]["rutaIntenciones"]
    listaCSVs = list(map(lambda x:  x+ ".csv",gerentesAConsultar))
    intencionesGerentes = obtenerIntenciones(ruta,listaCSVs)
    if len(intencionesGerentes) == 0:
        return pd.DataFrame([])
       
    intencionesGerentes.reset_index(drop=True, inplace=True)    
    fechaUltimaModificacion = pd.to_datetime(intencionesGerentes["UltimaModificacion"],format="%d/%m/%Y-%H:%M:%S")   
    intencionesGerentes = intencionesGerentes.loc[fechaUltimaModificacion.apply(lambda x: True if (x.date() >=fechaDesde) and (x.date() <=fechaHasta) else False),]
    ntencionesGerentes = intencionesGerentes[intencionesGerentes["TipoActivo"] == mercadoAConsultar]
    intencionesGerentes = intencionesGerentes.fillna("")
    intencionesGerentes = intencionesGerentes.sort_values(by=['Id'],ascending=False)
    
    
    return intencionesGerentes
    
        
        
def obtenerIntencionesGerentesParaVisualizar(fechaDesde,gerentesAConsultar,ruta,camposMostrar):
    
    listaCSVs = list(map(lambda x:  x+ ".csv",gerentesAConsultar))
    intencionesGerentes = obtenerIntenciones(ruta,listaCSVs)
    if len(intencionesGerentes) == 0:
        return pd.DataFrame([])
    #Ahora se deben aplicar los filtros correspondientes antes le mostrar las intenciones en el excel
    #Filtrar desde la fecha especificada
    
    ultimaModificacion = pd.to_datetime(intencionesGerentes["UltimaModificacion"],format="%d/%m/%Y-%H:%M:%S")
    vigenteHasta = pd.to_datetime(intencionesGerentes["VigenteHasta"],format="%d/%m/%Y")
    intencionesGerentes = intencionesGerentes[(ultimaModificacion.apply(lambda x: True if x.date() >=fechaDesde else False)) |( (~intencionesGerentes["Estado"].isin(["Ejecutada/Total", "Cancelada"])) & (vigenteHasta.apply(lambda x: True if x.date() >= fechaDesde else False))) | (intencionesGerentes["Estado"].isin(["Vencida"]))]
    
    #devolver solo los campos que me interesan mostrar
    camposMostrar = camposMostrar.split("-")
    #Mostrar las intenciones cuya fecha de vigencia desde ya esté activa
    vigenciaDesde = pd.to_datetime(intencionesGerentes["VigenciaDesde"],format="%d/%m/%Y")
    intencionesGerentes = intencionesGerentes.fillna("")
    intencionesGerentes = intencionesGerentes.sort_values(by=['Id'],ascending=False)
    
    return intencionesGerentes[camposMostrar]
  
def obtenerIntencionesTradersParaVisualizar(gerentesAConsultar,mercadoAConsultar,periodoAConsultarDesde,ruta,camposMostrar) :
 
    listaCSVs = list(map(lambda x:  x+ ".csv",gerentesAConsultar))
    intencionesGerentes = obtenerIntenciones(ruta,listaCSVs)
    if len(intencionesGerentes) == 0:
        return pd.DataFrame([])
    #Ahora se deben aplicar los filtros correspondientes antes le mostrar las intenciones en el excel    
    intencionesGerentes.reset_index(drop=True, inplace=True)
    ## Se filtran las intenciones que la vigencia hasta aún esté activa en cierto rango de fechas
    ## VigenteHasta debe ser superior o igual a periodoAConsultarDesde
    
    ### Vamos a traer al monitor todas las inenciones que:
    # han sido modificadas en el periodo de consulta
    # y tambien se traeran todas las intenciones vigentes en la fecha de consulta y que no estan canceladas ni ejecutadas al 100%
    vigenteHasta = pd.to_datetime(intencionesGerentes["VigenteHasta"],format="%d/%m/%Y")   
    ultimaModificacion = pd.to_datetime(intencionesGerentes["UltimaModificacion"],format="%d/%m/%Y-%H:%M:%S")
    intencionesGerentes = intencionesGerentes[(ultimaModificacion.apply(lambda x: True if x.date() >=periodoAConsultarDesde else False)) |( (~intencionesGerentes["Estado"].isin(["Ejecutada/Total", "Cancelada"])) & (vigenteHasta.apply(lambda x: True if x.date() >= periodoAConsultarDesde else False))) ]
    ##
    if mercadoAConsultar == "RV y Fondos" or mercadoAConsultar == "Renta Variable":
        intencionesGerentes = intencionesGerentes.loc[(intencionesGerentes["TipoActivo"].isin(["Renta Variable","Fondos"])) &(intencionesGerentes["Mercado"].isin(["LOCAL","INTERNACIONAL","FONDO MUTUO"])),]
    if mercadoAConsultar == "Deuda Privada":
        intencionesGerentes = intencionesGerentes.loc[(intencionesGerentes["TipoActivo"] == "Deuda Privada") & (intencionesGerentes["Mercado"] == "LOCAL"),]
    if mercadoAConsultar == "Deuda Pública":
        intencionesGerentes = intencionesGerentes.loc[(intencionesGerentes["TipoActivo"] == "Deuda Pública")& (intencionesGerentes["Mercado"] == "LOCAL") ,]
    if mercadoAConsultar == "Fondos":
        intencionesGerentes = intencionesGerentes.loc[(intencionesGerentes["TipoActivo"] == "Fondos") & (intencionesGerentes["Mercado"] == "FIC") ,]
    if mercadoAConsultar == "Liquidez":
        intencionesGerentes = intencionesGerentes.loc[intencionesGerentes["TipoActivo"] == "Liquidez",]
    if mercadoAConsultar == "Forex":
        intencionesGerentes = intencionesGerentes.loc[intencionesGerentes["TipoActivo"] == "Forex",]
    if mercadoAConsultar == "Swaps":
        intencionesGerentes = intencionesGerentes.loc[intencionesGerentes["TipoActivo"] == "Swaps",]
    if mercadoAConsultar == "RF Internacional":
        intencionesGerentes = intencionesGerentes.loc[(intencionesGerentes["TipoActivo"].isin(["Deuda Privada","Deuda Pública"])) & (intencionesGerentes["Mercado"] == "INTERNACIONAL"),]
    if "Todos" in  mercadoAConsultar:
        intencionesGerentes = intencionesGerentes
        
    vigenciaDesde = pd.to_datetime(intencionesGerentes["VigenciaDesde"],format="%d/%m/%Y")
    intencionesGerentes = intencionesGerentes.loc[vigenciaDesde.apply(lambda x: True if x.date() <=date.today() else False),]
    
    
    intencionesGerentes = intencionesGerentes.fillna("")
    intencionesGerentes = intencionesGerentes.sort_values(by=['Id'],ascending=False)
    return intencionesGerentes[camposMostrar]
    
def obtenerIntenciones(ruta, listaArchivos):
    return cargueArchivosCSV(ruta, listaArchivos)    

def cargueArchivosCSV(ruta, archivos_csv):
    
    df = {}
    for file in archivos_csv:
        try:
            df[file] = pd.read_csv(ruta + file,sep = ",")
        except:
            None
    
    if len(df) >0:
        return pd.concat(list(df.values()))
    else:
        return pd.DataFrame([])
    
def celdaConValidador(rango):
    
    try:
        rango.Validation.Type
        return True
    except:
        return False

def limpiarHojaIntencionesPM(hoja,filaDesde):    
    
    ultimaFila = hoja.Range(hoja.Cells(hoja.Rows.Count, 2), hoja.Cells(hoja.Rows.Count, 2)).End(definiciones.xlUp).Row
    rango = str(filaDesde) + ":" + str(ultimaFila +1)
    if hoja.ProtectContents == True:
        hoja.Unprotect()
    hoja.Range(rango).Delete()
    
    
def obtenerMacroActivos():
    
    macroActivos = definiciones.parametros["Valor"]["macroActivos"]
    macroActivos = macroActivos.split("-")
    return macroActivos
    
def obtenerMacroActivosTrader():
    macroActivos = definiciones.parametros["Valor"]["macroActivosTraders"]
    macroActivos = macroActivos.split("-")
    return macroActivos
    
    
def traducirMacroActivo(macroActivoSeleccionado,mercado):
    
    if macroActivoSeleccionado == "":
        return ""
    macroActivosInventarios = definiciones.parametros["Valor"]["MacroActivosNombreInventario"]
    macroActivosInventarios = macroActivosInventarios.split("-")
    macroActivos = definiciones.parametros["Valor"]["macroActivos"]
    macroActivos = macroActivos.split("-")
    traductor = dict(zip(macroActivos,macroActivosInventarios))
    traducido = traductor[macroActivoSeleccionado]
    if traducido == "RV" :
        traducido = traducido + " " + mercado
    return traducido

def graficarEstadisticasIntenciones(intenciones,hojaEstadisticas,hojaDescargas,intencionesAM):
    #1. Crear tabla de etiquetas y datos a mostrar
    #2. limpiar hoja
    #3. hacer gráfico
    fechaHasta = hojaDescargas.OLEObjects("txtFechaHastaDescargas").Object.Value
    fechaHasta = datetime.strptime(fechaHasta, '%d/%m/%Y')
    
    intenciones = intenciones[intenciones["TipoOrden"]!="LÍMITE"]    
    intenciones["TipoActivo-Mercado"] = intenciones["TipoActivo"] +" "+ intenciones["Mercado"].str.title()
    
    intencionesTotalOps = intenciones.copy()    
    intencionesTotalOps["Completada"] = intencionesTotalOps["Ejecutado"].apply(lambda x: 1 if x== 100 else 0)
    intencionesTotalOps["Activa"] = intencionesTotalOps["Estado"].apply(lambda x: 1 )
    datosActivo = intencionesTotalOps.groupby(["TipoActivo-Mercado"]).sum()  
    datosActivo["PorcentajeEjecutadas"] = datosActivo["Completada"]/ datosActivo["Activa"] 
    datosActivo.loc[datosActivo["PorcentajeEjecutadas"].isna(),"PorcentajeEjecutadas"] = 0   
    
    rangoDatosNumIntenciones = hojaEstadisticas.Range(hojaEstadisticas.Cells(8,2),hojaEstadisticas.Cells(len(datosActivo)+7,3))
    rangoDatosEjecutadas = hojaEstadisticas.Range(hojaEstadisticas.Cells(8,4),hojaEstadisticas.Cells(len(datosActivo)+7,4))
    rangoDatosPorcEjec = hojaEstadisticas.Range(hojaEstadisticas.Cells(8,5),hojaEstadisticas.Cells(len(datosActivo)+7,5))
    hojaEstadisticas.UsedRange.ClearContents()
    rangoDatosNumIntenciones.Value =  datosActivo[["Activa"]].to_records()
    rangoDatosEjecutadas.Value = datosActivo[["Completada"]].to_records(index = False)
    rangoDatosPorcEjec.Value = datosActivo[["PorcentajeEjecutadas"]].to_records(index = False)
    
    hojaEstadisticas.Range("B6").Value = "Total Operaciones"
    hojaEstadisticas.Range("B7").Value = "Tipo Activo y Mercado"
    hojaEstadisticas.Range("C7").Value = "Nro de órdenes"
    hojaEstadisticas.Range("D7").Value = "Ejecutadas"
    hojaEstadisticas.Range("E7").Value = "% Cumplim"
    rangoDatosPorcEjec.NumberFormat = "0.0%"
    
    del datosActivo
    intencionesMuestraValida = intenciones.copy()
    fechaVigenteHasta = pd.to_datetime(intencionesMuestraValida["VigenteHasta"],format="%d/%m/%Y")  
    intencionesMuestraValida = intencionesMuestraValida[~((intencionesMuestraValida["Estado"] =="Cancelada") & (intencionesMuestraValida["Ejecutado"] == 0)) & ~((intencionesMuestraValida["Estado"] != "Ejecutada/Total" ) & (fechaVigenteHasta.apply(lambda x: True if x.date() > fechaHasta.date() else False )))]
    intencionesMuestraValida["Completada"] = intencionesMuestraValida["Ejecutado"].apply(lambda x: 1 if x== 100 else 0)
    intencionesMuestraValida["Activa"] = intencionesMuestraValida["Estado"].apply(lambda x: 1 )
    
    datosActivo = intencionesMuestraValida.groupby(["TipoActivo-Mercado"]).sum() 
    datosActivo["PorcentajeEjecutadas"] = datosActivo["Completada"]/ datosActivo["Activa"] 
    datosActivo.loc[datosActivo["PorcentajeEjecutadas"].isna(),"PorcentajeEjecutadas"] = 0
    rangoDatosNumIntenciones = hojaEstadisticas.Range(hojaEstadisticas.Cells(23,2),hojaEstadisticas.Cells(len(datosActivo)+22,3))
    rangoDatosEjecutadas = hojaEstadisticas.Range(hojaEstadisticas.Cells(23,4),hojaEstadisticas.Cells(len(datosActivo)+22,4))
    rangoDatosPorcEjec = hojaEstadisticas.Range(hojaEstadisticas.Cells(23,5),hojaEstadisticas.Cells(len(datosActivo)+22,5))
   
    rangoDatosNumIntenciones.Value =  datosActivo[["Activa"]].to_records()
    rangoDatosEjecutadas.Value = datosActivo[["Completada"]].to_records(index = False)
    rangoDatosPorcEjec.Value = datosActivo[["PorcentajeEjecutadas"]].to_records(index = False)
  
    hojaEstadisticas.Range("B21").Value = "Muesta válida"
    hojaEstadisticas.Range("B22").Value = "Tipo Activo y Mercado"
    hojaEstadisticas.Range("C22").Value = "Nro de órdenes"
    hojaEstadisticas.Range("D22").Value = "Ejecutadas"
    hojaEstadisticas.Range("E22").Value = "% Cumplim"
    rangoDatosPorcEjec.NumberFormat = "0.0%"
   
    hojaEstadisticas.Visible = True
    # try:
    #     print(hojaEstadisticas.Shapes.Count)
    #     for i in range(1,hojaEstadisticas.Shapes.Count):
    #         hojaEstadisticas.Shapes(i+1).Delete()
            
    # except:
    #     None
    # titulo = "Operaciones por tipo de activo y mercado"    
    #crearGraficoCircular(intencionesAM,rangoDatosNumIntenciones,titulo)
    # hojaEstadisticas.Shapes(hojaEstadisticas.Shapes.Count).Left = hojaEstadisticas.Range("G7").Left
    # hojaEstadisticas.Shapes(hojaEstadisticas.Shapes.Count).Top = hojaEstadisticas.Range("G7").Top
    # hojaEstadisticas.Shapes(hojaEstadisticas.Shapes.Count).Width = 325
    # hojaEstadisticas.Shapes(hojaEstadisticas.Shapes.Count).Height = 325
    hojaEstadisticas.Columns.AutoFit()
    hojaEstadisticas.Visible = False

def crearGraficoCircular(intencionesAM,rangoDatos,titulo):
    chart = intencionesAM.Charts.Add()
    chart.ChartType = definiciones.xlPie
    chart.SetSourceData(rangoDatos)
      
    chart.ChartStyle = 258
    chart.ChartStyle = 259
    chart.ClearToMatchStyle()
    #chart.SetElement(definiciones.msoElementDataLabelBestFit)
    chart.HasTitle = True
    chart.ChartTitle.Text = titulo
    chart.Location(definiciones.xlLocationAutomatic,"Estadisticas")  
   


def obtenerPortafolios():   
    
    portafoliosCRM = definiciones.portafoliosCRM.copy()
    portafoliosCRM.sort_values(["CÓD MUREX"],ascending=True,inplace=True)
    portafolios = portafoliosCRM[portafoliosCRM["ADMINISTRADOR"].isin(["FIDUCIARIA","VALORES"])]
    codigos = portafolios["CÓD MUREX"]    
    codigo =  codigos.tolist()
    codigo = str(codigo)[1:-1] #Para quitar los corchetes
    codigo = codigo.replace(", ",",")
    return codigo.replace("'","")

def portafolioValoresoFidu(portafoliosCRM,portafolio):
    
    administrador = portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,["ADMINISTRADOR"]] 
    if len(administrador)>0:     
        
        if administrador["ADMINISTRADOR"].iloc[0] == "VALORES":
            
            return "VALORES"    
        else:
            return "FIDUCIARIA"
    else:
        
        return "NINGUNO"

 
def obtenerNombrePortafolioPorId(portafolio):
    
    portafoliosCRM = definiciones.portafoliosCRM
    #Es un portafololio de valores o de fiduciaria
    portafolioAdministrador =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    
    if  portafolioAdministrador == "FIDUCIARIA":
        nombrePortafolio = portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,"NOMBRE PORTAFOLIO"].iloc[0]
    elif portafolioAdministrador == "VALORES":
        nombrePortafolio = portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"] == portafolio,"NOMBRE PORTAFOLIO"].iloc[0]
    else:
        nombrePortafolio = "No encontrado"            
    
    return nombrePortafolio


def obtenerCuposcompras(portafolio,macroActivo,nemotecnico):
    
    """
    portafolio: str para aplicar el filtro a la base de cupos
    macroActivo: str para segmentar el proceeso de calculo de la cantidad disponible
    nemotecnico: str para consultarlo en la base de titulos de la administración
    """
    if portafolio == None or nemotecnico == None or macroActivo == None:
        return "SIN DATOS"
    
    portafoliosCRM = definiciones.portafoliosCRM
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria.copy()    
    nemoTitulosValores = definiciones.nemoTitulosValores.copy() 
    cuposFiduciaria = definiciones.cuposFiduciaria.copy()
    cuposValores = definiciones.cuposValores.copy()
    especies = definiciones.especies.copy()
    preciosAccionesTabla = definiciones.precioAcciones.copy()
    
    #EMISOR DEL INVENTARIO DE CUPOS
    emisorCupos = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Emisor cupos"]]
    
    if len(emisorCupos) == 0:
            emisorCupos = ""
    else:
        emisorCupos = emisorCupos["Emisor cupos"].iloc[0]
        if emisorCupos == "NAN":
            emisorCupos = ""
    
   
    
    #EMISOR DEL INVENTARIO DE TITULOS       
    emisorUnificado = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Emisor inventario"]]   
    if len(emisorUnificado) == 0:
        emisorUnificado = ""
    else:
        emisorUnificado = emisorUnificado["Emisor inventario"].iloc[0]  
        if emisorUnificado == "NAN":
            emisorUnificado = ""
    
    #MONEDA DEL TITULOS   
    moneda = especies.loc[especies["Nemo intenciones"] == nemotecnico,"Moneda"]
    if len(moneda) == 0:
        moneda = ""
    else:
        moneda = moneda.iloc[0]
    
        
    
    if administradorPortafolio == "FIDUCIARIA":     
        
        #TOTAL PORTAFOLIO
        saldoPortafolio = definiciones.valorPortafolioFiduciaria.loc[portafolio,"SALDO ABA"]
        print("***********************") 
        print("saldo Portafolio: "+ str(saldoPortafolio))
    
        #CUPO
        if emisorCupos == "":
            porcentajeCupo = pd.DataFrame([])            
        else:   
            cuposFiduciaria["Nombre.1"] = cuposFiduciaria["Nombre.1"].str.replace(".","")          
            porcentajeCupo = cuposFiduciaria[(cuposFiduciaria["MUREX"] == portafolio) & (cuposFiduciaria["Nombre.1"].str.replace(" ","") == emisorCupos.replace(".","").replace(" ",""))]
                      
                
        
        
        #OCUPACION ACTUAL 
        if emisorUnificado == "":
            ocupacionValor = 0
        else:
            #Antes de calcular la ocupación actual del emisor, retiro aquellas operaciones intradia de compra, para que no afecten la ocupación real del emisor en el portafolio
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[~((nemoTitulosFiduciaria["POSICIÓN"]== "VENTA") & (nemoTitulosFiduciaria["Origen Informacion"] =="INTENCIONES VIGENTES"))]
            ocupacionValor = nemoTitulosFiduciaria.loc[(nemoTitulosFiduciaria["Portafolio"] == portafolio) & (nemoTitulosFiduciaria["Emisor Unificado"] == emisorUnificado ) & (~nemoTitulosFiduciaria["Macro Activo"].str.contains("LIQUIDEZ") ),"SALDO ABA"].sum()
           
           
        print("ocupación: " + str(ocupacionValor))   
        #PRECIO ACCION
        especieConsulta = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Especie"]]
        nemoConsulta = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Nemotecnico"]]
        precioAccion_1 = []
        precioAccion_2 = []
        if len(especieConsulta) > 0:
            precioAccion_1 = preciosAccionesTabla.loc[preciosAccionesTabla["Titulo"] == especieConsulta["Especie"].iloc[0],["PrecioAccion"]]
        if len(nemoConsulta) > 0:
            precioAccion_2 = preciosAccionesTabla.loc[preciosAccionesTabla["Titulo"] == nemoConsulta["Nemotecnico"].iloc[0],["PrecioAccion"]]
                  
        if len(precioAccion_1) > 0:
            precioAccion = precioAccion_1["PrecioAccion"].iloc[0]
        elif len(precioAccion_2) > 0:
            precioAccion = precioAccion_2["PrecioAccion"].iloc[0]
        else:
            precioAccion = 0
               
                
    elif administradorPortafolio == "VALORES":
        
        #VALOR PORTAFOLIO
        saldoPortafolio = definiciones.valorPortafolioValores.loc[portafolio,"SALDO ABA"]
        print("***********************") 
        print("saldo Portafolio: "+ str(saldoPortafolio))
        
        #CUPO
        if emisorCupos == "":
            porcentajeCupo = pd.DataFrame([])
        else:    
            cuposValores["Nombre.1"] = cuposValores["Nombre.1"].str.replace(".","")            
            porcentajeCupo = cuposValores[(cuposValores["MUREX"] == portafolio) & (cuposValores["Nombre.1"].str.replace(" ","") == emisorCupos.replace(".","").replace(" ",""))]
           
        
        #OCUPACION ACTUAL        
        if emisorUnificado == "":
            ocupacionValor = 0
        else:
            #Antes de calcular la ocupación actual del emisor, retiro aquellas operaciones intradia de compra, para que no afecten la ocupación real del emisor en el portafolio
            nemoTitulosValores = nemoTitulosValores[~((nemoTitulosValores["POSICIÓN"]== "VENTA") & (nemoTitulosValores["Origen Informacion"] =="INTENCIONES VIGENTES"))]
            ocupacionValor = nemoTitulosValores.loc[(nemoTitulosValores["Portafolio"] == portafolio) & (nemoTitulosValores["Emisor Unificado"] == emisorUnificado) & (~nemoTitulosValores["Macro Activo"].str.contains("LIQUIDEZ")) ,"SALDO ABA"].sum()
        print("ocupación: " + str(ocupacionValor))   
        
        #PRECIO ACCIÓN
        especieConsulta = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Especie"]]
        nemoConsulta = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Nemotecnico"]]
        precioAccion_1 = []
        precioAccion_2 = []
        if len(especieConsulta) > 0:
            precioAccion_1 = preciosAccionesTabla.loc[preciosAccionesTabla["Titulo"] == especieConsulta["Especie"].iloc[0],["PrecioAccion"]]
        if len(nemoConsulta) > 0:
            precioAccion_2 = preciosAccionesTabla.loc[preciosAccionesTabla["Titulo"] == nemoConsulta["Nemotecnico"].iloc[0],["PrecioAccion"]]
                  
        if len(precioAccion_1) > 0:
            precioAccion = precioAccion_1["PrecioAccion"].iloc[0]
        elif len(precioAccion_2) > 0:
            precioAccion = precioAccion_2["PrecioAccion"].iloc[0]
        else:
            precioAccion = 0
                     
    else:
        return "SIN DATOS"
    
    #PORCENTAJE CUPO
   
    isin = especies.loc[especies["Nemo intenciones"] == nemotecnico,"Isin"]
    if len(isin) >0:
        isin = isin.iloc[0]
        if isin == "NAN":
            isin = ""
    else: 
        isin = ""
        
    nemotecnicoOficial = especies.loc[especies["Nemo intenciones"] == nemotecnico,"Nemotecnico"]
    if len(nemotecnicoOficial) >0:
        nemotecnicoOficial = nemotecnicoOficial.iloc[0]
        if nemotecnicoOficial == "NAN":
            nemotecnicoOficial = ""
    else: 
        nemotecnicoOficial = ""
    porcentajeCupo_1 = []
    porcentajeCupo_2 = []
    
    if len(porcentajeCupo ) == 0:
        porcentajeCupo = 0
    else:
        if isin != "":
            porcentajeCupo_1 = porcentajeCupo.loc[porcentajeCupo["ISIN 1"] == isin,["Ocupación Máxima"]]
        if nemotecnicoOficial != "":
            porcentajeCupo_2 = porcentajeCupo.loc[porcentajeCupo["Nemo"] == nemotecnicoOficial,["Ocupación Máxima"]]
        if len(porcentajeCupo_1) > 0:
            porcentajeCupo =  porcentajeCupo_1["Ocupación Máxima"].iloc[0]
        elif len(porcentajeCupo_2) > 0:
            porcentajeCupo =  porcentajeCupo_2["Ocupación Máxima"].iloc[0]
        else:
            porcentajeCupo = porcentajeCupo[porcentajeCupo["Cupo"] == "EMISOR"]
            if len(porcentajeCupo) == 0:
                porcentajeCupo = 0
            else:
                porcentajeCupo = porcentajeCupo["Ocupación Máxima"].iloc[0]  
    print("Porcentaje cupo: " + str(porcentajeCupo))     
    
    if pd.isna(saldoPortafolio):
        cupoDisponible = 0
        return cupoDisponible
    if pd.isna(porcentajeCupo):
        cupoDisponible =0
        return cupoDisponible
    if pd.isna(ocupacionValor):
        cupoDisponible = 0
        return cupoDisponible
    
    cupoDisponible = (saldoPortafolio * porcentajeCupo) - ocupacionValor
    if macroActivo == "Renta Variable":
        if precioAccion == 0 or pd.isna(precioAccion):
            print("Valor Accion: SIN DATOS" )
            cupoDisponible = "SIN DATOS"
        else:
            print("Valor Accion: " +str(precioAccion))
            cupoDisponible = cupoDisponible/precioAccion
    
    if macroActivo in["Deuda Privada","Deuda Pública"]:
        
        if emisorUnificado == "GENERICO DPR":
            return "SIN DATOS"
       
        if moneda == "USD":
            if definiciones.TRM != 0:
                print("USD: " + str(definiciones.TRM))
                cupoDisponible = cupoDisponible/definiciones.TRM
            else:
                cupoDisponible = "SIN DATOS"
        elif moneda == "UVR":
            if definiciones.UVR != 0:
                print("UVR: " + str(definiciones.UVR))
                cupoDisponible = cupoDisponible/definiciones.UVR
            else:
                cupoDisponible = "SIN DATOS"
        elif moneda == "":
            cupoDisponible = "SIN DATOS"
            
        cupoDisponible = cupoDisponible/1000000
    
    if macroActivo == "Fondos":
         if moneda == "USD":
            if definiciones.TRM != 0:                
                cupoDisponible = cupoDisponible/definiciones.TRM
            else:
                cupoDisponible = "SIN DATOS"
    
    return cupoDisponible
    
def obtenerCantidadDisponibleNemos(nemotecnicoInventario,portafolio,macroActivo):
    
    #CANTIDAD DISPONIBLE DE VENTA 
    if nemotecnicoInventario == None or portafolio == None or macroActivo == None:
        return "SIN DATOS"
    
    
    portafoliosCRM = definiciones.portafoliosCRM
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    especies = definiciones.especies.copy()
    
    if  administradorPortafolio == "FIDUCIARIA":
       
        especiesCopia_1 = especies[especies["Especie"] != "NAN"]
        
        especie = especiesCopia_1.loc[especiesCopia_1["Nemo intenciones"] == nemotecnicoInventario,["Especie"]]
        especiesCopia_2 = especies[especies["Nemotecnico"]!= "NAN"]
        nemotecnico = especiesCopia_2.loc[especiesCopia_2["Nemo intenciones"] == nemotecnicoInventario,["Nemotecnico"]]
        if len(nemotecnico) == 0:
            nemotecnico = nemotecnicoInventario
        else:
            nemotecnico = nemotecnico["Nemotecnico"].iloc[0]
        if len(especie) == 0:
            especie = nemotecnicoInventario
        else:
            especie = especie["Especie"].iloc[0]
            
        nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria.copy()
        nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Portafolio"]==portafolio,]
        
        #Para mostrar la cantidad disponible en ventas fiduciaria no tenemos en cuenta los datos de compras de operaciones intradia,
        #para que se muestre la cantida real disponible
        nemoTitulosFiduciaria = nemoTitulosFiduciaria[~((nemoTitulosFiduciaria["POSICIÓN"]== "COMPRA") & (nemoTitulosFiduciaria["Origen Informacion"] =="INTENCIONES VIGENTES"))]
        
        if macroActivo in ["RV INTERNACIONAL","RV LOCAL"]:
            
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"])]             
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[["Especie/Generador","Nominal Remanente"]]
            seleccion = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Especie/Generador"] == especie]
            if len(seleccion) > 0:
                cantidad = seleccion.groupby(["Especie/Generador"]).sum()
                return cantidad["Nominal Remanente"].values[0]
            else:
                return "SIN DATOS"
            
                
        elif macroActivo == "DEUDA PRIVADA":
            
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PRIVADA","RF INTERNACIONAL","DEUDA PÚBLICA"]),]            
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[["Especie/Generador","Nominal Remanente"]]
            seleccion = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Especie/Generador"] == nemotecnicoInventario]
            
            if len(seleccion) > 0:
                cantidad = seleccion.groupby(["Especie/Generador"]).sum()
                return cantidad["Nominal Remanente"].values[0]/1000000
            else:
                return "SIN DATOS"
       
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]            
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[["Nemotécnico","Nominal Remanente"]]
            seleccion = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Nemotécnico"] == nemotecnico]
            
            if len(seleccion) > 0:
                cantidad = seleccion.groupby(["Nemotécnico"]).sum()
                return cantidad["Nominal Remanente"].values[0]/1000000
            else:
                return "SIN DATOS"
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RV INTERNACIONAL","RF INTERNACIONAL"])]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[["Especie/Generador","Vr Mercado Hoy Moneda Empresa"]]
            seleccion = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Especie/Generador"] == especie]
            if len(seleccion) > 0:
                cantidad = seleccion.groupby(["Especie/Generador"]).sum()
                return cantidad["Vr Mercado Hoy Moneda Empresa"].values[0]
        
        elif macroActivo == "SWAP":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"] == macroActivo]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[~nemoTitulosFiduciaria["Contract Id"].isna()]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Contract Id"].apply(lambda contract: esUnNumero(contract))]
            seleccion = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Contract Id"].astype("int64").astype("str").str.strip().apply(lambda nemo: True if nemo in nemotecnico else False),["Nominal Remanente"]]
            if len(seleccion) > 0:
                cantidad = seleccion["Nominal Remanente"].iloc[0]
                return cantidad
            else:
                return "SIN DATOS"
        else:
            return "SIN DATOS"
              

    
    elif administradorPortafolio == "VALORES":
        
        nemoTitulosValores = definiciones.nemoTitulosValores
        nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Portafolio"] == portafolio,]
        #Para mostrar la cantidad disponible en ventas valores no tenemos en cuenta los datos de compras de operaciones intradia,
        #para que se muestre la cantida real disponible
        nemoTitulosValores = nemoTitulosValores[~((nemoTitulosValores["POSICIÓN"]== "COMPRA") & (nemoTitulosValores["Origen Informacion"] =="INTENCIONES VIGENTES"))]
        
        especies = especies[especies["Nemotecnico"]!= "NAN"]
        nemotecnico = especies.loc[especies["Nemo intenciones"] == nemotecnicoInventario,["Nemotecnico"]]
        if len(nemotecnico) == 0:
            nemotecnico = nemotecnicoInventario
        else:
            nemotecnico = nemotecnico["Nemotecnico"].iloc[0]
        
        if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"])]             
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Nemoténico"] == nemotecnico]
                 
            if len(nemoTitulosValores) > 0:
                nemoTitulosValores = nemoTitulosValores.groupby(["Nemoténico"]).sum()
                return nemoTitulosValores["Valor Nominal Actual"].iloc[0]
            else:
                return "SIN DATOS"

            
        elif macroActivo == "DEUDA PRIVADA":
           
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PRIVADA","RF INTERNACIONAL","DEUDA PÚBLICA"]),]            
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Nemoténico"] == nemotecnico]
                 
            if len(nemoTitulosValores) > 0:
                nemoTitulosValores = nemoTitulosValores.groupby(["Nemoténico"]).sum()
                return nemoTitulosValores["Valor Nominal Actual"].iloc[0] /1000000
            else:
                return "SIN DATOS"
                
        
        elif macroActivo == "DEUDA PÚBLICA":
            
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]            
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Nemoténico"] == nemotecnico]       
            if len(nemoTitulosValores) > 0:
                nemoTitulosValores = nemoTitulosValores.groupby(["Nemoténico"]).sum()
                return nemoTitulosValores["Valor Nominal Actual"].iloc[0] /1000000
            else:
                return "SIN DATOS"
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Nemoténico"] == nemotecnico]
            
            if len(nemoTitulosValores) > 0:
                nemoTitulosValores = nemoTitulosValores.groupby(["Nemoténico"]).sum()
                return nemoTitulosValores["Valor VPN Actual"].iloc[0] 
            else:
                return "SIN DATOS"
        else:
            return "SIN DATOS"
       
        
       
    else:
        mostrarMensajeAdvertencia("No se encontró el portafolio : "+ portafolio)
        return "SIN DATOS"
        
    
def obtenerCantidadDisponibleNemosFondos(nemotecnico,portafolio,macroActivo,tipoOperacion):  
    
    if nemotecnico == None or portafolio == None or tipoOperacion == None or macroActivo == None:
        return 0   
   
    cantidadDisponible = obtenerCantidadDisponibleNemos(nemotecnico,portafolio,macroActivo)
    if tipoOperacion == "RETIRO":
        porcentajeProtejido = definiciones.porcentajeProtejidoFondos
        porcentajeProtejido = porcentajeProtejido[porcentajeProtejido["Fondo"].str.strip() == nemotecnico]
        
        if len(porcentajeProtejido) > 0:
            porcentajeProtejido = porcentajeProtejido["% Protegido"].iloc[0]
        else:
            porcentajeProtejido = 0
        cantidadDisponible  = cantidadDisponible - (cantidadDisponible * porcentajeProtejido/100)
    
    return cantidadDisponible
    

def obtenerTodosLosNemosFiduciaria(macroActivo):  

     nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria  
     parametros = definiciones.parametros
     if "RV" in macroActivo:                
         nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"] == macroActivo,]
     
     nemoTitulosFiduciaria = nemoTitulosFiduciaria["Especie/Generador","ISIN","Nominal Remanente"]                                                                 
     nemoTitulosFiduciaria["Especie-ISIN"] = nemoTitulosFiduciaria['Especie/Generador'].str.replace("  "," ") + nemoTitulosFiduciaria['ISIN'].astype('str')
     nemoTitulosFiduciaria = nemoTitulosFiduciaria.groupby(["Especie-ISIN"]).sum()
     nemoTitulosFiduciaria.sort_values(["Especie-ISIN"],ascending=True,inplace=True)
     nemos = str(nemoTitulosFiduciaria.index.tolist())
     nemos = nemos[1:-1]
     nemos = nemos.replace(", ",",")
     nemos = nemos.replace("'","")
     return nemos                

def obtenerNemosdePortafolio(portafolio,macroActivo):
    #Para filtrar los nemos que voy a presentar se deben aplicar los siguinetes filtros
    #ADMINISTRADOR: Fiduciaria o Valores
    #MERCADO: Local o Internacional
    #MACRO ACTIVO
    #COD Murex o Cod porfin portafolio
    portafoliosCRM = definiciones.portafoliosCRM
    parametros = definiciones.parametros
    especies = definiciones.especies.copy()
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    
    if  administradorPortafolio == "FIDUCIARIA":
        portafolioSeleccionado =  portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,] 
    
    elif administradorPortafolio == "VALORES":
        portafolioSeleccionado = portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,]
    
    else:
        mostrarMensajeAdvertencia("No se encontró el portafolio seleccionado: "+ portafolio)
        return "-"
        
    
    
    if  administradorPortafolio == "FIDUCIARIA":
        
        nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria
        nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Portafolio"] == portafolio,]
        
        if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Especie/Generador"].isin(especies.loc[especies["Macro Activo"]==macroActivo,"Especie"].tolist())]
            nemos = nemoTitulosFiduciaria["Especie/Generador"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
                            
                
        elif macroActivo == "DEUDA PRIVADA":           
            
            especiesDeudaPrivada = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PRIVADA"]),].copy()
            especiesRFInternacional = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RF INTERNACIONAL","DEUDA PÚBLICA"]),].copy()
           
            especiesRFInternacional = especiesRFInternacional[especiesRFInternacional["Especie/Generador"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"]),"Especie"].tolist())]
            especiesGenericas = especies.loc[especies["Nemotecnico"] == "GENERICO DPR","Nemo intenciones"]
            
            nemo = pd.concat([especiesDeudaPrivada["Especie/Generador"],especiesRFInternacional["Especie/Generador"]],ignore_index=True)
            nemo.sort_values(ascending=True,inplace=True)
            nemo = especiesGenericas.append(nemo,ignore_index=True)
            nemo.drop_duplicates(inplace=True)
            
            if len(nemo) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
        
            nemo = str(nemo.tolist())
            nemo = nemo[1:-1]
            nemo = nemo.replace(", ",",")
            nemo = nemo.replace("'","")                
            return nemo
            
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Especie/Generador"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"]),"Especie"].tolist())]
            
            nemos = nemoTitulosFiduciaria["Especie/Generador"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RF INTERNACIONAL","RV INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Especie/Generador"].isin(especies.loc[especies["Macro Activo"].isin(["FONDO MUTUO","FIC"]),"Especie"].tolist())]
                      
            nemos = nemoTitulosFiduciaria["Especie/Generador"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "SWAP":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"] == macroActivo,]
                
            nemoTitulosFiduciaria["Contract Id"] = "Id Contract " + nemoTitulosFiduciaria["Contract Id"].astype("int64").astype("str")
            nemos = nemoTitulosFiduciaria["Contract Id"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
            
        
        
        else:
            return "-"
        
        
        if len(nemos) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
        nemos = nemos.apply(lambda x: especies.loc[especies["Especie"]== x,"Nemo intenciones"].iloc[0] if len(especies.loc[especies["Especie"]== x,"Nemo intenciones"]) == 1 else x)
       
        nemos = str(nemos.tolist())
        nemos = nemos[1:-1]
        nemos = nemos.replace(", ",",")
        nemos = nemos.replace("'","")
            
        return nemos
         
       
    
    elif administradorPortafolio == "VALORES":
        
        nemoTitulosValores = definiciones.nemoTitulosValores
        nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Portafolio"] == portafolio,]
        
        if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"]==macroActivo,"Nemotecnico"].tolist())]
            nemos = nemoTitulosValores["Nemoténico"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "DEUDA PRIVADA":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PRIVADA","RF INTERNACIONAL","DEUDA PÚBLICA"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"]),"Nemotecnico"].tolist())]
            especiesGenericas = especies.loc[especies["Nemotecnico"] == "GENERICO DPR","Nemo intenciones"]
            
            nemos = nemoTitulosValores["Nemoténico"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
            nemos = especiesGenericas.append(nemos,ignore_index=True)
            
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"]),"Nemotecnico"].tolist())]
            
            nemos = nemoTitulosValores["Nemoténico"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RF INTERNACIONAL","RV INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"].isin(["FONDO MUTUO","FIC"]),"Nemotecnico"].tolist())]
                      
            nemos = nemoTitulosValores["Nemoténico"].copy()                                                                  
            nemos.drop_duplicates(inplace = True)
            nemos.sort_values(ascending=True,inplace=True)
            
        else:
            return "-"
            
            
        if len(nemos) == 0:
            mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
            return "-"        
        
        nemos = nemos.apply(lambda x: especies.loc[especies["Nemotecnico"]== x,"Nemo intenciones"].iloc[0] if len(especies.loc[especies["Nemotecnico"]== x,"Nemo intenciones"]) == 1 else x)
        nemos = str(nemos.tolist())
        nemos = nemos.replace(", ",",")  
        return nemos.replace("[","").replace("]","").replace("'","")

def obtenerNemosdeMacroactivo(portafolio,macroActivo):
    #Esta función devuelve todos los nemos que pertenecen a determinado macroactivo
    #Es útil para mostrar todas las opciones de compra
    especies = definiciones.especies.copy()    
        
    if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:            
        titulos = especies[especies["Macro Activo"]== macroActivo]
        titulos = titulos["Nemo intenciones"]    
        titulos.drop_duplicates(inplace= True)
        titulos.sort_values(ascending=True,inplace=True)
            
    elif macroActivo == "DEUDA PRIVADA":
        titulos = especies[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"])]
        titulos = titulos["Nemo intenciones"]            
        titulos.drop_duplicates(inplace= True)   
          
    
    elif macroActivo == "DEUDA PÚBLICA":
        titulos = especies[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"])]     
        titulos = titulos["Nemo intenciones"]    
        titulos.drop_duplicates(inplace= True)
        titulos.sort_values(ascending=True,inplace=True)      
    
    elif macroActivo == "PARTICIPACIÓN EN FONDOS":
        titulos = especies.loc[especies["Macro Activo"].isin(["FIC","FONDO MUTUO"])] 
        titulos = titulos["Nemo intenciones"]    
        titulos.drop_duplicates(inplace= True)
        titulos.sort_values(ascending=True,inplace=True)
    else:
        return "-"        
    
    
    if len(titulos) == 0:
        return "-"
    else:               
        titulos = str(titulos.tolist())
        titulos = titulos.replace("[","").replace("]","").replace("'","").replace(", ",",")
        
        return titulos 

def obtenerEmisoresdePortafolio(portafolio,macroActivo):   
    
    portafoliosCRM = definiciones.portafoliosCRM
    parametros = definiciones.parametros
    especies = definiciones.especies.copy()
   
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    
    if  administradorPortafolio == "FIDUCIARIA":
        portafolioSeleccionado =  portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,] 
    
    elif administradorPortafolio == "VALORES":
        portafolioSeleccionado = portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,]
    
    else:
        mostrarMensajeAdvertencia("No se encontró el portafolio seleccionado: "+ portafolio)
        return "-"
        
    
    
    if  administradorPortafolio == "FIDUCIARIA":
        
        nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria
        nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Portafolio"] == portafolio,]
        
        if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"]== macroActivo,"Emisor inventario"].tolist())]
            if len(nemoTitulosFiduciaria) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
            emisores = nemoTitulosFiduciaria["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
                            
                
        elif macroActivo == "DEUDA PRIVADA":
                   
            emisoresDeudaPrivada = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PRIVADA"]),].copy()
            emisoresRFInternacional = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RF INTERNACIONAL","DEUDA PÚBLICA"]),].copy()
            
            emisoresRFInternacional = emisoresRFInternacional[emisoresRFInternacional["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"]),"Emisor inventario"].tolist())]
            
            emisores = pd.concat([emisoresDeudaPrivada["Emisor Unificado"],emisoresRFInternacional["Emisor Unificado"]],ignore_index=True)
            
            if len(emisores) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
        
            emisores.drop_duplicates(inplace=True)
            emisores.sort_values(ascending=True,inplace=True)               
            emisores = str(emisores.tolist())
            emisores = emisores[1:-1]
            emisores = emisores.replace(", ",",")
            emisores = emisores.replace("'","")                
            return emisores
       
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"]),"Emisor inventario"].tolist())]
            if len(nemoTitulosFiduciaria) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
            emisores = nemoTitulosFiduciaria["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RF INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"].isin(["FIC","FONDO MUTUO"]),"Emisor inventario"].tolist())]  
            if len(nemoTitulosFiduciaria) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
            emisores = nemoTitulosFiduciaria["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "SWAP":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"] == macroActivo,]
            if len(nemoTitulosFiduciaria) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
            emisores = nemoTitulosFiduciaria["Emisor / Contraparte"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
            
        
        
        emisores = str(emisores.tolist())
        emisores = emisores[1:-1]
        emisores = emisores.replace(", ",",")
        emisores = emisores.replace("'","")   
        
   
    elif administradorPortafolio == "VALORES":
        
        nemoTitulosValores = definiciones.nemoTitulosValores
        nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Portafolio"] == portafolio,]
        
       
        if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONA","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"] == macroActivo,"Emisor inventario"].tolist())]  
            emisores = nemoTitulosValores["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
            
        elif macroActivo == "DEUDA PRIVADA":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PRIVADA","RF INTERNACIONAL","DEUDA PÚBLICA"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"]),"Emisor inventario"].tolist())]  
            emisores = nemoTitulosValores["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
            emisores = emisores.append(pd.Series(["GENERICO DPR"]),ignore_index= True)
        
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTENRNACIONAL"]),"Emisor inventario"].tolist())]  
            emisores = nemoTitulosValores["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RF INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Emisor Unificado"].isin(especies.loc[especies["Macro Activo"].isin(["FIC","FONDO MUTUO"]),"Emisor inventario"].tolist())]  
            emisores = nemoTitulosValores["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True)
            
        elif macroActivo == "SWAP":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"] == macroActivo,]            
            emisores = nemoTitulosValores["Emisor Unificado"].copy()                                                                 
            emisores.drop_duplicates(inplace = True)
            emisores.sort_values(ascending=True,inplace=True) 
        else:
            return "-"
    
        if len(emisores) == 0:
            mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
            return "-"
        
        
        
        emisores = str(emisores.tolist())
        emisores = emisores[1:-1]
        emisores = emisores.replace(", ",",")
        emisores = emisores.replace("'","") 
        
            
    else:
        emisores = "-"
        
            
    return emisores
    
def obtenerNemosdeEmisor(portafolio,macroActivo,emisor):    
    
    portafoliosCRM = definiciones.portafoliosCRM
    parametros = definiciones.parametros   
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    especies = definiciones.especies.copy()
        
    if  administradorPortafolio == "FIDUCIARIA":
        portafolioSeleccionado =  portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,] 
    
    elif administradorPortafolio == "VALORES":
        portafolioSeleccionado = portafoliosCRM.loc[portafoliosCRM["CÓD MUREX"]== portafolio,]
    
    else:
        #mostrarMensajeAdvertencia("No se encontró el portafolio seleccionado: "+ portafolio)
        return "-"
        
    
    
    if  administradorPortafolio == "FIDUCIARIA":
        
        nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria
        nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Portafolio"] == portafolio,]
        
        if macroActivo in ["RV INTERNACIONAL","RV LOCAL"]:
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Especie/Generador"].isin(especies.loc[especies["Macro Activo"]== macroActivo,"Especie"].tolist())]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Emisor Unificado"] == emisor,]
            
            nemo = nemoTitulosFiduciaria["Especie/Generador"].copy() 
                
        elif macroActivo == "DEUDA PRIVADA":
            especiesDeudaPrivada = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PRIVADA"]),].copy()
            especiesRFInternacional = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["RF INTERNACIONAL","DEUDA PÚBLICA"]),].copy()
           
            especiesDeudaPrivada = especiesDeudaPrivada.loc[especiesDeudaPrivada["Emisor Unificado"] == emisor,]     
            especiesRFInternacional = especiesRFInternacional.loc[especiesRFInternacional["Emisor Unificado"] == emisor,]       
            
            especiesRFInternacional = especiesRFInternacional[especiesRFInternacional["Especie/Generador"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"]),"Especie"].tolist())]
            
            nemo = pd.concat([especiesDeudaPrivada["Especie/Generador"],especiesRFInternacional["Especie/Generador"]],ignore_index=True)
            
            if len(nemo) == 0:
                mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
                return "-"
        
            nemo.drop_duplicates(inplace=True)
            nemo.sort_values(ascending=True,inplace=True)               
            nemo = str(nemo.tolist())
            nemo = nemo[1:-1]
            nemo = nemo.replace(", ",",")
            nemo = nemo.replace("'","")                
            return nemo
            
       
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Especie/Generador"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"]),"Especie"].tolist())]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Emisor Unificado"] == emisor,]
            
            nemo = nemoTitulosFiduciaria["Especie/Generador"].copy()
            
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Especie/Generador"].isin(especies.loc[especies["Macro Activo"].isin(["FIC","FONDO MUTUO"]),"Especie"].tolist())]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Emisor Unificado"] == emisor,]
              
            nemo = nemoTitulosFiduciaria["Especie/Generador"].copy()
           
        
        elif macroActivo == "SWAP":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Macro Activo"] == macroActivo,]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Emisor Unificado"] == emisor,]
                     
            nemoTitulosFiduciaria["Contract Id"] = "Id Contract " + nemoTitulosFiduciaria["Contract Id"].astype("int64").astype("str")
            nemo = nemoTitulosFiduciaria["Contract Id"].copy()
            
        else:
            return "-"
        
        if len(nemo) == 0:
            mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
            return "-"
        
        nemo.drop_duplicates(inplace=True)
        nemo.sort_values(ascending=True,inplace=True)   
        nemo = nemo.apply(lambda x: especies.loc[especies["Especie"]== x,"Nemo intenciones"].iloc[0] if len(especies.loc[especies["Especie"]== x,"Nemo intenciones"]) == 1 else x)
        nemo = str(nemo.tolist())
        nemo = nemo[1:-1]
        nemo = nemo.replace(", ",",")
        nemo = nemo.replace("'","")
            
        return nemo
    
    elif administradorPortafolio == "VALORES":
        
        nemoTitulosValores = definiciones.nemoTitulosValores
        nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Portafolio"] == portafolio,]
       
        
        if macroActivo in ["RV INTERNACIONAL","RV LOCAL"]:
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["RV LOCAL","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"]== macroActivo,"Nemotecnico"].tolist())]
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Emisor Unificado"] == emisor,]
            
           
                
        elif macroActivo == "DEUDA PRIVADA":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PRIVADA","RF INTERNACIONAL","DEUDA PÚBLICA"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"]),"Nemotecnico"].tolist())]
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Emisor Unificado"] == emisor,]
             
            
       
        elif macroActivo == "DEUDA PÚBLICA":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["DEUDA PÚBLICA","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"]),"Nemotecnico"].tolist())]
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Emisor Unificado"] == emisor,]
            
          
            
        
        elif macroActivo == "PARTICIPACIÓN EN FONDOS":
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Macro Activo"].isin(["PARTICIPACIÓN EN FONDOS","RV INTERNACIONAL","RF INTERNACIONAL"]),]
            nemoTitulosValores = nemoTitulosValores[nemoTitulosValores["Nemoténico"].isin(especies.loc[especies["Macro Activo"].isin(["FIC","FONDO MUTUO"]),"Nemotecnico"].tolist())]
            nemoTitulosValores = nemoTitulosValores.loc[nemoTitulosValores["Emisor Unificado"] == emisor,]
              
            
        else:
            return "-"
        
        
        if len(nemoTitulosValores) == 0:
            mostrarMensajeAdvertencia("El portafolio " + portafolio + " no tiene títulos.")
            return "-"
            
        nemo = nemoTitulosValores["Nemoténico"].copy()
        nemo.drop_duplicates(inplace=True)
        nemo.sort_values(ascending=True,inplace=True) 
        nemo = nemo.apply(lambda x: especies.loc[especies["Nemotecnico"]== x,"Nemo intenciones"].iloc[0] if len(especies.loc[especies["Nemotecnico"]== x,"Nemo intenciones"]) == 1 else x)   
        nemo = str(nemo.tolist())
        nemo = nemo[1:-1]
        nemo = nemo.replace(", ",",")
        nemo = nemo.replace("'","")
        return nemo 
    else:
        return "-"
        
def obtenerEmisoresdePortafolioCompras(portafolio, macroActivo):       
    
    #Cuando se selecciona una compra, se va a presentar todos los emisores en lo que el portafolio puede comprar titulos
    portafoliosCRM = definiciones.portafoliosCRM.copy()
    parametros = definiciones.parametros.copy()
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)   
    especies = definiciones.especies.copy()
    
    if  administradorPortafolio == "FIDUCIARIA":
        cuposFiduciaria = definiciones.cuposFiduciaria.copy()
        cuposFiduciaria = cuposFiduciaria[cuposFiduciaria["MUREX"] == portafolio]
        emisores = cuposFiduciaria["Nombre.1"]
        
     
    elif administradorPortafolio == "VALORES":
        cuposValores = definiciones.cuposValores.copy()
        cuposValores = cuposValores[cuposValores["MUREX"]== portafolio]
        emisores = cuposValores["Nombre.1"]        
    
    else:
        #mostrarMensajeAdvertencia("No se encontró el portafolio seleccionado: "+ portafolio)
        return "-"
    
    if macroActivo  in ["RV LOCAL","RV INTERNACIONAL"]:
        emisores = especies.loc[(especies["Emisor cupos"].isin(emisores.tolist())) & (especies["Macro Activo"] == macroActivo),"Emisor inventario"]
    elif macroActivo == "DEUDA PRIVADA":
        emisores = especies.loc[(especies["Emisor cupos"].isin(emisores.tolist())) & (especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"])),"Emisor inventario"]
    elif macroActivo == "DEUDA PÚBLICA":                                   
        emisores = especies.loc[(especies["Emisor cupos"].isin(emisores.tolist())) & (especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"])),"Emisor inventario"]
    elif macroActivo == "PARTICIPACIÓN EN FONDOS":
        emisores = especies.loc[(especies["Emisor cupos"].isin(emisores.tolist())) & (especies["Macro Activo"].isin(["FIC","FONDO MUTUO"])),"Emisor inventario"]
    else:
        return "-"
    
    if len(emisores) == 0:
        return "-"
    else:
        emisores.drop_duplicates(inplace = True)
        emisores.sort_values(ascending = True,inplace = True)
        emisores = str(emisores.tolist())
        emisores = emisores.replace(", ",",")
        emisores = emisores.replace("[","").replace("]","").replace("'","")
        return emisores
        
    
    
    
def obtenerNemosDeEmisorCompras(portafolio,emisor,macroActivo):
    #1. con el nombre del emisor obtengo el nit y el Isin del emisor de la base de cupos
    #2. con la información del emisor busco en el inventario de títulos que nemos tiene asociado ese emisor
    #3. filtro por el macroActivo seleccionado actualmente.
    especies = definiciones.especies.copy()
    if macroActivo in ["RV LOCAL","RV INTERNACIONAL"]:            
        titulos = especies[especies["Macro Activo"]== macroActivo]
            
    elif macroActivo == "DEUDA PRIVADA":
        titulos = especies[especies["Macro Activo"].isin(["DEUDA PRIVADA","DEUDA PRIVADA INTERNACIONAL"])]
    
    elif macroActivo == "DEUDA PÚBLICA":
        titulos = especies[especies["Macro Activo"].isin(["DEUDA PÚBLICA","DEUDA PÚBLICA INTERNACIONAL"])]           
    
    elif macroActivo == "PARTICIPACIÓN EN FONDOS":
        titulos = especies.loc[especies["Macro Activo"].isin(["FIC","FONDO MUTUO"])]
    
    else:
        return "-"    
    
    titulos = titulos[titulos["Emisor inventario"] == emisor]
    titulos = titulos["Nemo intenciones"].copy()
    titulos.drop_duplicates(inplace= True)
    titulos.sort_values(ascending=True,inplace=True)
    
    if len(titulos) == 0:
        return "-"
    else:               
        titulos = str(titulos.tolist())
        titulos = titulos.replace(", ",",")
        titulos = titulos.replace("[","").replace("]","").replace("'","")
        
        return titulos     
    

def obtenerEmisordeNemo(portafolio,macroActivoTitulos,nemotecnico):

    portafoliosCRM = definiciones.portafoliosCRM
    parametros = definiciones.parametros   
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria.copy()
    nemoTitulosValores = definiciones.nemoTitulosValores.copy()
    especies = definiciones.especies.copy()

    if  administradorPortafolio == "FIDUCIARIA":
        if macroActivoTitulos == "SWAP":
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[~nemoTitulosFiduciaria["Contract Id"].isna()]
            nemoTitulosFiduciaria = nemoTitulosFiduciaria[nemoTitulosFiduciaria["Contract Id"].apply(lambda contract: esUnNumero(contract))]
            emisor = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Contract Id"].astype("int64").astype("str").str.strip().apply(lambda nemo: True if nemo in nemotecnico else False),["Emisor Unificado"]]
        
        elif macroActivoTitulos == "DEUDA PRIVADA":
            
            emisor = nemoTitulosFiduciaria.loc[nemoTitulosFiduciaria["Especie/Generador"] == nemotecnico,["Emisor Unificado"]]  
        else:
            emisor = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Emisor inventario"]]
            emisor.rename(columns={"Emisor inventario":"Emisor Unificado"},inplace=True)
                        
        if len(emisor) == 0:
            print("No se encontró emisores para este nemo: " + nemotecnico)
            return "-"
        else:
            return str(emisor["Emisor Unificado"].iloc[0]).strip()
        
    elif administradorPortafolio == "VALORES":
        emisor = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Emisor inventario"]]
        if len(emisor) == 0:
            print("No se encontró emisores para este nemo: " + nemotecnico)
            return "-"
        else:
            return str(emisor["Emisor inventario"].iloc[0]).strip()
    else:
        #mostrarMensajeAdvertencia("No se encontró el portafolio seleccionado: "+ portafolio)
        return "-"
    
def obtenerIndicadorNemo(portafolio,nemotecnico):

    portafoliosCRM = definiciones.portafoliosCRM
    parametros = definiciones.parametros   
    administradorPortafolio =  portafolioValoresoFidu(portafoliosCRM,portafolio)
    especies = definiciones.especies.copy()
    especies = especies[especies["Indicador"]!= "NAN"]
    indicador = especies.loc[especies["Especie"] == nemotecnico,["Indicador"]]
    if len(indicador) == 0:
        indicadores = str(parametros["Valor"]["Indicadores DPR"].split("-"))
        indicadores = indicadores[1:-1]
        indicadores = indicadores.replace(", ",",")
        return indicadores.replace("'","")
    else:
        return indicador["Indicador"].iloc[0]
    

def obtenerEmisordeNemoCompra(portafolio,nemotecnico):   
    
    especies = definiciones.especies.copy()    
    especies = especies.loc[especies["Nemo intenciones"] == nemotecnico,["Emisor inventario"]]   
    
    if len(especies) == 0:
        return "-"
    else:      
        return especies["Emisor inventario"].iloc[0]
    

def completarDatosIntencionRV(intencion,creacionEdicion):
    
    intencionCompleta = pd.DataFrame([])
    especies = definiciones.especiesOriginal.copy()

    
    if creacionEdicion == "CREACION":
        time.sleep(0.2)
        intencionCompleta.loc[1,"Id"] =  int(int(datetime.now().strftime("%y%m%d%H%M%S%f"))/10000)      
        intencionCompleta.loc[1,"Estado"] = "Nueva"
    
    elif creacionEdicion == "EDICION":        
        intencionCompleta.loc[1,"Id"] = intencion["Id"]        
        intencionCompleta.loc[1,"Estado"] = "Modificada"
    
   
    intencionCompleta.loc[1,"FechaIngreso"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"IngresadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"UltimaModificacion"] =datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"ModificadoPor"] = getpass.getuser().upper()    
    intencionCompleta.loc[1,"TipoActivo"] = "Renta Variable"
    intencionCompleta.loc[1,"Mercado"] = intencion["Mercado"]
    intencionCompleta.loc[1,"CodPortafolio"] = intencion["Id Portafolio"]
    intencionCompleta.loc[1,"Portafolio"] = obtenerNombrePortafolioPorId(intencion["Id Portafolio"])
    intencionCompleta.loc[1,"TipoOperacion"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoInstruccionFondo"] = ""
    intencionCompleta.loc[1,"TipoOrden"] = intencion["Tipo orden"]  
    intencionCompleta.loc[1,"Emisor"] = intencion["Emisor"]
    intencionCompleta.loc[1,"Nemotecnico"] = intencion["Nemotécnico"]
    intencionCompleta.loc[1,"Indicador"] = ""
    intencionCompleta.loc[1,"Denominacion"] = "Unidades"
    intencionCompleta.loc[1,"FechaEmision"] = ""
    intencionCompleta.loc[1,"FechaVencimiento"] = ""
    intencionCompleta.loc[1,"TasaFacial"] = ""
    intencionCompleta.loc[1,"Desde"] = ""
    intencionCompleta.loc[1,"Hasta"] = ""
    if intencion["Tipo orden"] == "LÍMITE":        
        intencionCompleta.loc[1,"PrecioLimite"] = intencion["Precio límite"]
    else:
        intencionCompleta.loc[1,"PrecioLimite"] = ""
    intencionCompleta.loc[1,"TasaLimite"] = ""
    intencionCompleta.loc[1,"VigenciaDesde"] = intencion["Vigente desde"]
    intencionCompleta.loc[1,"VigenteHasta"] = intencion["Vigente hasta"]
    intencionCompleta.loc[1,"ComentariosPM"] = intencion["Comentarios PM"]
    intencionCompleta.loc[1,"Trader"] = ""
    intencionCompleta.loc[1,"UltimaModifTrader"] = ""
    intencionCompleta.loc[1,"CantidadTotal"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"CantEjecutada"] = 0
    intencionCompleta.loc[1,"CantPendiente"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"Ejecutado"] = 0
    intencionCompleta.loc[1,"PrecioPromedio"] = ""
    intencionCompleta.loc[1,"ComentariosTrader"] = ""
    intencionCompleta.loc[1,"Prioridad"] = ""
    intencionCompleta.loc[1,"UltimoTrader"] = ""
    intencionCompleta.set_index(["Id"],drop = False,inplace=True)
    return intencionCompleta
    
def completarDatosIntencionDPR(intencion,creacionEdicion):
    
    especies = definiciones.especies.copy()
    intencionCompleta = pd.DataFrame([])

    
    if creacionEdicion == "CREACION":
        time.sleep(0.2)
        intencionCompleta.loc[1,"Id"] =  int(int(datetime.now().strftime("%y%m%d%H%M%S%f"))/10000)      
        intencionCompleta.loc[1,"Estado"] = "Nueva"
    
    elif creacionEdicion == "EDICION":        
        intencionCompleta.loc[1,"Id"] = intencion["Id"]        
        intencionCompleta.loc[1,"Estado"] = "Modificada"
    
   
    intencionCompleta.loc[1,"FechaIngreso"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"IngresadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"UltimaModificacion"] =datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"ModificadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"TipoActivo"] = "Deuda Privada"
    mercado = especies.loc[especies["Nemo intenciones"] == intencion["Nemotécnico"],["Macro Activo","Moneda"]]
    if len(mercado) == 0:
        intencionCompleta.loc[1,"Mercado"] = "LOCAL"
        intencionCompleta.loc[1,"Denominacion"] = "MM"
    else:
        intencionCompleta.loc[1,"Denominacion"] = "MM " + str(mercado["Moneda"].iloc[0])
        if mercado["Macro Activo"].iloc[0] == "DEUDA PRIVADA INTERNACIONAL":
            intencionCompleta.loc[1,"Mercado"] = "INTERNACIONAL"
        else:
            intencionCompleta.loc[1,"Mercado"] = "LOCAL"
    intencionCompleta.loc[1,"CodPortafolio"] = intencion["Id Portafolio"]
    intencionCompleta.loc[1,"Portafolio"] = obtenerNombrePortafolioPorId(intencion["Id Portafolio"])
    intencionCompleta.loc[1,"TipoOperacion"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoInstruccionFondo"] = ""
    intencionCompleta.loc[1,"TipoOrden"] = intencion["Tipo orden"]    
    intencionCompleta.loc[1,"Emisor"] = intencion["Emisor"]
    intencionCompleta.loc[1,"Nemotecnico"] = intencion["Nemotécnico"]
    intencionCompleta.loc[1,"Indicador"] = intencion["Indicador"]    
    intencionCompleta.loc[1,"FechaEmision"] = ""
    intencionCompleta.loc[1,"FechaVencimiento"] = ""
    intencionCompleta.loc[1,"TasaFacial"] = ""
    if intencion["Nemotécnico"] in especies.loc[especies["Nemotecnico"] == "GENERICO DPR","Nemo intenciones"].tolist():
        intencionCompleta.loc[1,"Desde"] = intencion["Desde"]
        intencionCompleta.loc[1,"Hasta"] = intencion["Hasta"]
    else:
        intencionCompleta.loc[1,"Desde"] = ""  
        intencionCompleta.loc[1,"Hasta"] = ""
        
    intencionCompleta.loc[1,"PrecioLimite"] = ""
    if intencion["Tipo orden"] == "LÍMITE":        
        intencionCompleta.loc[1,"TasaLimite"] = intencion["Tasa límite"]
    else:
        intencionCompleta.loc[1,"TasaLimite"] = ""
    intencionCompleta.loc[1,"VigenciaDesde"] = intencion["Vigente desde"]
    intencionCompleta.loc[1,"VigenteHasta"] = intencion["Vigente hasta"]
    intencionCompleta.loc[1,"ComentariosPM"] = intencion["Comentarios PM"]
    intencionCompleta.loc[1,"Trader"] = ""
    intencionCompleta.loc[1,"UltimaModifTrader"] = ""
    intencionCompleta.loc[1,"CantidadTotal"] = intencion["Cantidad(Millones)"]
    intencionCompleta.loc[1,"CantEjecutada"] = 0
    intencionCompleta.loc[1,"CantPendiente"] = intencion["Cantidad(Millones)"]
    intencionCompleta.loc[1,"Ejecutado"] = 0
    intencionCompleta.loc[1,"PrecioPromedio"] = ""
    intencionCompleta.loc[1,"ComentariosTrader"] = ""
    intencionCompleta.loc[1,"Prioridad"] = ""
    intencionCompleta.loc[1,"UltimoTrader"] = ""
    intencionCompleta.set_index(["Id"],drop = False,inplace=True)
    return intencionCompleta
    
def completarDatosIntencionDPU(intencion,creacionEdicion):
    
    especies = definiciones.especies.copy()
    intencionCompleta = pd.DataFrame([])

    
    if creacionEdicion == "CREACION":
        time.sleep(0.2)
        intencionCompleta.loc[1,"Id"] =  int(int(datetime.now().strftime("%y%m%d%H%M%S%f"))/10000)      
        intencionCompleta.loc[1,"Estado"] = "Nueva"
    
    elif creacionEdicion == "EDICION":        
        intencionCompleta.loc[1,"Id"] = intencion["Id"]        
        intencionCompleta.loc[1,"Estado"] = "Modificada"
    
   
    intencionCompleta.loc[1,"FechaIngreso"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"IngresadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"UltimaModificacion"] =datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"ModificadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"TipoActivo"] = "Deuda Pública"    
    intencionCompleta.loc[1,"CodPortafolio"] = intencion["Id Portafolio"]
    intencionCompleta.loc[1,"Portafolio"] = obtenerNombrePortafolioPorId(intencion["Id Portafolio"])
    intencionCompleta.loc[1,"TipoOperacion"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoInstruccionFondo"] = ""
    intencionCompleta.loc[1,"TipoOrden"] = intencion["Tipo orden"] 
    intencionCompleta.loc[1,"Emisor"] = intencion["Emisor"]
    intencionCompleta.loc[1,"Nemotecnico"] = intencion["Nemotécnico"]
    intencionCompleta.loc[1,"Indicador"] = ""
    mercado = especies.loc[especies["Nemo intenciones"] == intencion["Nemotécnico"],["Macro Activo","Moneda"]]
    if len(mercado) == 0:
        intencionCompleta.loc[1,"Mercado"] = "LOCAL"
        intencionCompleta.loc[1,"Denominacion"] = "MM"
    else:
        if mercado["Macro Activo"].iloc[0] == "FUTUROS TES":
            intencionCompleta.loc[1,"Denominacion"] = str(mercado["Moneda"].iloc[0])
        else:
            intencionCompleta.loc[1,"Denominacion"] = "MM " + str(mercado["Moneda"].iloc[0])
        if mercado["Macro Activo"].iloc[0] == "DEUDA PÚBLICA INTERNACIONAL":
            intencionCompleta.loc[1,"Mercado"] = "INTERNACIONAL"
        else:
            intencionCompleta.loc[1,"Mercado"] = "LOCAL" 
        
    intencionCompleta.loc[1,"FechaEmision"] = ""
    intencionCompleta.loc[1,"FechaVencimiento"] = ""
    intencionCompleta.loc[1,"TasaFacial"] = ""
    intencionCompleta.loc[1,"Desde"] = ""
    intencionCompleta.loc[1,"Hasta"] = ""
    intencionCompleta.loc[1,"PrecioLimite"] = ""
    if intencion["Tipo orden"] == "LÍMITE":        
        intencionCompleta.loc[1,"TasaLimite"] = intencion["Tasa límite"]
    else:
        intencionCompleta.loc[1,"TasaLimite"] = ""
    intencionCompleta.loc[1,"VigenciaDesde"] = intencion["Vigente desde"]
    intencionCompleta.loc[1,"VigenteHasta"] = intencion["Vigente hasta"]
    intencionCompleta.loc[1,"ComentariosPM"] = intencion["Comentarios PM"]
    intencionCompleta.loc[1,"Trader"] = ""
    intencionCompleta.loc[1,"UltimaModifTrader"] = ""
    intencionCompleta.loc[1,"CantidadTotal"] = intencion["Cantidad(Millones)"]
    intencionCompleta.loc[1,"CantEjecutada"] = 0
    intencionCompleta.loc[1,"CantPendiente"] = intencion["Cantidad(Millones)"]
    intencionCompleta.loc[1,"Ejecutado"] = 0
    intencionCompleta.loc[1,"PrecioPromedio"] = ""
    intencionCompleta.loc[1,"ComentariosTrader"] = ""
    intencionCompleta.loc[1,"Prioridad"] = ""
    intencionCompleta.loc[1,"UltimoTrader"] = ""
    intencionCompleta.set_index(["Id"],drop = False,inplace=True)
    return intencionCompleta

def completarDatosIntencionFondos(intencion,creacionEdicion):
    
    intencionCompleta = pd.DataFrame([])
    especies = definiciones.especies.copy()

    
    if creacionEdicion == "CREACION":
        time.sleep(0.2)
        intencionCompleta.loc[1,"Id"] =  int(int(datetime.now().strftime("%y%m%d%H%M%S%f"))/10000)      
        intencionCompleta.loc[1,"Estado"] = "Nueva"
    
    elif creacionEdicion == "EDICION":        
        intencionCompleta.loc[1,"Id"] = intencion["Id"]        
        intencionCompleta.loc[1,"Estado"] = "Modificada"
    
   
    intencionCompleta.loc[1,"FechaIngreso"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"IngresadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"UltimaModificacion"] =datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"ModificadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"TipoActivo"] = "Fondos"
    
    intencionCompleta.loc[1,"CodPortafolio"] = intencion["Id Portafolio"]
    intencionCompleta.loc[1,"Portafolio"] = obtenerNombrePortafolioPorId(intencion["Id Portafolio"])
    intencionCompleta.loc[1,"TipoOperacion"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoInstruccionFondo"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoOrden"] = "A MERCADO"   
    intencionCompleta.loc[1,"Emisor"] = intencion["Emisor"]
    intencionCompleta.loc[1,"Nemotecnico"] = intencion["Nemotécnico"]
    intencionCompleta.loc[1,"Indicador"] = ""
    mercado = especies.loc[especies["Nemo intenciones"] == intencion["Nemotécnico"],["Macro Activo","Moneda"]]
    if len(mercado) == 0:
        intencionCompleta.loc[1,"Mercado"] = ""
        intencionCompleta.loc[1,"Denominacion"] = ""
    else:
        intencionCompleta.loc[1,"Denominacion"] = str(mercado["Moneda"].iloc[0])
        if mercado["Macro Activo"].iloc[0] == "FIC":
            intencionCompleta.loc[1,"Mercado"] = "FIC"
        else:
            intencionCompleta.loc[1,"Mercado"] = "FONDO MUTUO"
    
    intencionCompleta.loc[1,"FechaEmision"] = ""
    intencionCompleta.loc[1,"FechaVencimiento"] = ""
    intencionCompleta.loc[1,"TasaFacial"] = ""
    intencionCompleta.loc[1,"Desde"] = ""
    intencionCompleta.loc[1,"Hasta"] = ""
    intencionCompleta.loc[1,"PrecioLimite"] = ""
    intencionCompleta.loc[1,"TasaLimite"] = ""
    intencionCompleta.loc[1,"VigenciaDesde"] = intencion["Vigente desde"]
    intencionCompleta.loc[1,"VigenteHasta"] = intencion["Vigente hasta"]
    intencionCompleta.loc[1,"ComentariosPM"] = intencion["Comentarios PM"]
    intencionCompleta.loc[1,"Trader"] = ""
    intencionCompleta.loc[1,"UltimaModifTrader"] = ""
    intencionCompleta.loc[1,"CantidadTotal"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"CantEjecutada"] = 0
    intencionCompleta.loc[1,"CantPendiente"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"Ejecutado"] = 0
    intencionCompleta.loc[1,"PrecioPromedio"] = ""
    intencionCompleta.loc[1,"ComentariosTrader"] = ""
    intencionCompleta.loc[1,"Prioridad"] = ""
    intencionCompleta.loc[1,"UltimoTrader"] = ""
    intencionCompleta.set_index(["Id"],drop = False,inplace=True)
    return intencionCompleta

def completarDatosIntencionForex(intencion,creacionEdicion):
    
    intencionCompleta = pd.DataFrame([])

    
    if creacionEdicion == "CREACION":
        time.sleep(0.2)
        intencionCompleta.loc[1,"Id"] =  int(int(datetime.now().strftime("%y%m%d%H%M%S%f"))/10000)      
        intencionCompleta.loc[1,"Estado"] = "Nueva"
    
    elif creacionEdicion == "EDICION":        
        intencionCompleta.loc[1,"Id"] = intencion["Id"]        
        intencionCompleta.loc[1,"Estado"] = "Modificada"
    
   
    intencionCompleta.loc[1,"FechaIngreso"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"IngresadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"UltimaModificacion"] =datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"ModificadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"TipoActivo"] = "Forex"
    intencionCompleta.loc[1,"Mercado"] = ""
    intencionCompleta.loc[1,"CodPortafolio"] = intencion["Id Portafolio"]
    intencionCompleta.loc[1,"Portafolio"] = obtenerNombrePortafolioPorId(intencion["Id Portafolio"])
    intencionCompleta.loc[1,"TipoOperacion"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoInstruccionFondo"] = ""
    intencionCompleta.loc[1,"TipoOrden"] = intencion["Tipo orden"]
    intencionCompleta.loc[1,"Emisor"] = ""    
    intencionCompleta.loc[1,"Nemotecnico"] = intencion["Nemotécnico"]      
    intencionCompleta.loc[1,"Indicador"] = ""
    intencionCompleta.loc[1,"Denominacion"] = ""
    intencionCompleta.loc[1,"FechaEmision"] = ""
    intencionCompleta.loc[1,"FechaVencimiento"] = ""
    intencionCompleta.loc[1,"TasaFacial"] = ""
    intencionCompleta.loc[1,"Desde"] = ""
    intencionCompleta.loc[1,"Hasta"] = ""
    if intencion["Tipo orden"] == "LÍMITE":        
        intencionCompleta.loc[1,"PrecioLimite"] = intencion["Precio límite"]
    else:
        intencionCompleta.loc[1,"PrecioLimite"] = ""
    intencionCompleta.loc[1,"TasaLimite"] = ""
    intencionCompleta.loc[1,"VigenciaDesde"] = intencion["Vigente desde"]
    intencionCompleta.loc[1,"VigenteHasta"] = intencion["Vigente hasta"]
    intencionCompleta.loc[1,"ComentariosPM"] = intencion["Comentarios PM"]
    intencionCompleta.loc[1,"Trader"] = ""
    intencionCompleta.loc[1,"UltimaModifTrader"] = ""
    intencionCompleta.loc[1,"CantidadTotal"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"CantEjecutada"] = 0
    intencionCompleta.loc[1,"CantPendiente"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"Ejecutado"] = 0
    intencionCompleta.loc[1,"PrecioPromedio"] = ""
    intencionCompleta.loc[1,"ComentariosTrader"] = ""
    intencionCompleta.loc[1,"Prioridad"] = ""
    intencionCompleta.loc[1,"UltimoTrader"] = ""
    intencionCompleta.set_index(["Id"],drop = False,inplace=True)
    return intencionCompleta

def completarDatosIntencionSwaps(intencion,creacionEdicion):
    
    intencionCompleta = pd.DataFrame([])

    
    if creacionEdicion == "CREACION":
        time.sleep(0.2)
        intencionCompleta.loc[1,"Id"] =  int(int(datetime.now().strftime("%y%m%d%H%M%S%f"))/10000)      
        intencionCompleta.loc[1,"Estado"] = "Nueva"
    
    elif creacionEdicion == "EDICION":        
        intencionCompleta.loc[1,"Id"] = intencion["Id"]        
        intencionCompleta.loc[1,"Estado"] = "Modificada"
    
   
    intencionCompleta.loc[1,"FechaIngreso"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"IngresadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"UltimaModificacion"] =datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
    intencionCompleta.loc[1,"ModificadoPor"] = getpass.getuser().upper()
    intencionCompleta.loc[1,"TipoActivo"] = "Swaps"
    if intencion["Tipo operación"] == "UNWIND": 
        intencionCompleta.loc[1,"Mercado"] = ""
    else:
        intencionCompleta.loc[1,"Mercado"] = intencion["Mercado"]
    intencionCompleta.loc[1,"CodPortafolio"] = intencion["Id Portafolio"]
    intencionCompleta.loc[1,"Portafolio"] = obtenerNombrePortafolioPorId(intencion["Id Portafolio"])
    intencionCompleta.loc[1,"TipoOperacion"] = intencion["Tipo operación"]
    intencionCompleta.loc[1,"TipoInstruccionFondo"] = ""
    intencionCompleta.loc[1,"TipoOrden"] = intencion["Tipo orden"]  
    if intencion["Tipo operación"] == "UNWIND":        
        intencionCompleta.loc[1,"Emisor"] = intencion["Emisor"]
    else:
        intencionCompleta.loc[1,"Emisor"] = ""
    intencionCompleta.loc[1,"Nemotecnico"] = intencion["Nemotécnico"]
    if intencion["Tipo operación"] != "UNWIND":        
        intencionCompleta.loc[1,"Indicador"] = intencion["Indicador"]
    else:
        intencionCompleta.loc[1,"Indicador"] = ""
    if intencion["Tipo operación"] != "UNWIND":
        if intencion["Mercado"] == "LOCAL":
            intencionCompleta.loc[1,"Denominacion"] = "COP"
        elif intencion["Mercado"] == "INTERNACIONAL":
            intencionCompleta.loc[1,"Denominacion"] = "USD"
    else:
        intencionCompleta.loc[1,"Denominacion"] = ""
    
    intencionCompleta.loc[1,"FechaEmision"] = ""
    intencionCompleta.loc[1,"FechaVencimiento"] = ""
    intencionCompleta.loc[1,"TasaFacial"] = ""
    intencionCompleta.loc[1,"Desde"] = ""
    if intencion["Tipo operación"] != "UNWIND":        
        intencionCompleta.loc[1,"Hasta"] = intencion["Hasta"]
    else:
        intencionCompleta.loc[1,"Hasta"] = ""
    intencionCompleta.loc[1,"PrecioLimite"] = ""
    if intencion["Tipo orden"] == "LÍMITE":        
        intencionCompleta.loc[1,"TasaLimite"] = intencion["Tasa límite"]
    else:
        intencionCompleta.loc[1,"TasaLimite"] = ""
    intencionCompleta.loc[1,"VigenciaDesde"] = intencion["Vigente desde"]
    intencionCompleta.loc[1,"VigenteHasta"] = intencion["Vigente hasta"]
    intencionCompleta.loc[1,"ComentariosPM"] = intencion["Comentarios PM"]
    intencionCompleta.loc[1,"Trader"] = ""
    intencionCompleta.loc[1,"UltimaModifTrader"] = ""
    intencionCompleta.loc[1,"CantidadTotal"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"CantEjecutada"] = 0
    intencionCompleta.loc[1,"CantPendiente"] = intencion["Cantidad"]
    intencionCompleta.loc[1,"Ejecutado"] = 0
    intencionCompleta.loc[1,"PrecioPromedio"] = ""
    intencionCompleta.loc[1,"ComentariosTrader"] = ""
    intencionCompleta.loc[1,"Prioridad"] = ""
    intencionCompleta.loc[1,"UltimoTrader"] = ""
    intencionCompleta.set_index(["Id"],drop = False,inplace=True)
    return intencionCompleta
    
    
    
def guardarIntenciones(data,ruta,usuario):
    
    #El nombre del usuario servirá como nombre del archivode Intenciones
    #Esta función se usa para la creación de intenciones
    
    try:
        archivoIntenciones = ruta +"/"+ usuario+".csv"
        data.to_csv(archivoIntenciones,mode = "a", header = False,index=False,sep=",")
        return True
    except:
        return False

def actualizarIntenciones(data,ruta,usuario):
    #El nombre del usuario servirá como nombre del archivo de Intenciones
    #esta función se usa para la edición
    
    try:
        archivoIntenciones = ruta +"/"+ usuario+".csv"
        campos = definiciones.parametros["Valor"]["camposArchivoIntenciones"]
        campos = campos.replace("-FechaIngreso","")
        campos = campos.replace("-IngresadoPor","")
        campos = campos.replace("-CantEjecutada","")
        campos = campos.replace("-ComentariosTrader","")
        campos = campos.replace("-UltimaModifTrader","")
        campos = campos.replace("-Trader","")
        campos = campos.split("-")        
        intenciones = pd.read_csv(archivoIntenciones,sep=",")        
        intenciones.set_index(["Id"],drop = False,inplace=True)
        intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),campos] = data.loc[data["Id"].isin(intenciones["Id"].astype("int64").tolist() ), campos]
        intenciones["Ejecutado"] = intenciones["CantEjecutada"]/intenciones["CantidadTotal"]*100
        intenciones["CantPendiente"] = intenciones["CantidadTotal"] - intenciones["CantEjecutada"]
        intenciones["Estado"] = intenciones[["Estado","Ejecutado"]].apply(lambda intencion: "Ejecutada/Total" if intencion["Ejecutado"] > 99.5 else ("Ejecutada/Parcial" if intencion["Ejecutado"] > 0 else intencion["Estado"]),axis =1 )
        intenciones.to_csv(archivoIntenciones,mode = "w", header = True,index=False,sep=",")
        return True, intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),]
    except:
        return False, pd.DataFrame([])
def renovarIntenciones(data,ruta,usuario):
     #El nombre del usuario servirá como nombre del archivo de Intenciones
    
    try:
        archivoIntenciones = ruta +"/"+ usuario+".csv"
        intenciones = pd.read_csv(archivoIntenciones,sep=",")
        intenciones.set_index(["Id"],drop = False,inplace=True)
        data.set_index(["Id"],drop = False,inplace=True)
        intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),["VigenciaDesde","VigenteHasta"]] = data.loc[data["Id"].isin(intenciones["Id"].astype("int64").tolist() ), ["Vigente desde","Vigente hasta"]].rename(columns={"Vigente desde":"VigenciaDesde","Vigente hasta":"VigenteHasta"})
        intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),"Estado"] = "Renovada"
        intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),"UltimaModificacion"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
        intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),"ModificadoPor"] = usuario
        intenciones.to_csv(archivoIntenciones,mode = "w", header = True,index=False,sep=",")
        return True, intenciones.loc[intenciones["Id"].isin(data["Id"].astype("int64").tolist()),]
    except:
        return False, pd.DataFrame([])
    
def cancelarIntenciones(Ids,ruta,usuario):
    
    try:
        archivoIntenciones = ruta +"/"+ usuario+".csv"
        intenciones = pd.read_csv(archivoIntenciones,sep=",")
        intenciones.set_index(["Id"],drop = False,inplace=True)
        intenciones.loc[intenciones["Id"].isin(Ids),"Estado"] = "Cancelada"
        intenciones.loc[intenciones["Id"].isin(Ids),"UltimaModificacion"] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
        intenciones.loc[intenciones["Id"].isin(Ids),"ModificadoPor"] = getpass.getuser().upper()
        intenciones.to_csv(archivoIntenciones,mode = "w", header = True,index=False,sep=",")
        return True, intenciones.loc[intenciones["Id"].isin(Ids),]
    except:
        
        return False,pd.DataFrame([])
        

def subirCreacion(hojaFormulario,ultFilaIntenciones,camposFormularioCreacion,macroActivo):

    totalColumnas = len(camposFormularioCreacion)
    rangoEncabezado = hojaFormulario.Range(hojaFormulario.Cells(7,2),hojaFormulario.Cells(7,totalColumnas+1))
    rangoDeDatos = hojaFormulario.Range(hojaFormulario.Cells(8,2),hojaFormulario.Cells(ultFilaIntenciones,totalColumnas+1))
        
    data = rangoDeDatos.Value
    intencionesNuevas = pd.DataFrame(data,columns =rangoEncabezado.Value[0])

    if len(intencionesNuevas) == 0:
        mostrarMensajeAdvertencia("No hay intenciones para subir")
        return False, pd.DataFrame([])
    
    #Aqui se realiza la validación de los sobrepasos para los portafolios
    #1.Alguno de los portafolios tiene sobrepasos?   
    tablaSobrepasos = definiciones.tablaSobrepasos.copy()
    sobrepasos = tablaSobrepasos[tablaSobrepasos["PORTAFOLIO"].isin(intencionesNuevas["Id Portafolio"])]
    if len(sobrepasos) > 0:
        
        sobrepasos.loc[sobrepasos["CAUSA"].isna(),"CAUSA"]  = ""
        sobrepasos.loc[sobrepasos["PLAN DE ACCION"].isna(),"PLAN DE ACCION"]  = ""
        root = tk.Tk()
        root.withdraw()  # Hide the main 
        excel = definiciones.excel
        newWb = excel.Workbooks.Add()
        nweSh = newWb.Sheets.Add()
        hoja_sobrepasos = newWb.ActiveSheet
        celdasEncabezado = hoja_sobrepasos.Range("A1:F1")
        celdasEncabezado.Value = sobrepasos[["PORTAFOLIO","FECHA SOBREPASO","TOTAL DIAS TRANSCURRIDOS","MENSAJE","CAUSA","PLAN DE ACCION"]].columns.tolist()
        hoja_sobrepasos.Range(hoja_sobrepasos.Cells(2,1),hoja_sobrepasos.Cells(1+len(sobrepasos),6)).Value =  sobrepasos[["PORTAFOLIO","FECHA SOBREPASO","TOTAL DIAS TRANSCURRIDOS","MENSAJE","CAUSA","PLAN DE ACCION"]].to_records(index=False)
        hoja_sobrepasos.Columns.AutoFit()
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        result = messagebox.askquestion("SOBREPASOS VIGENTES", "Por favor revise la hoja de sobrepasos ya que tiene portafolios con sobrepasos vigentes. ¿Desea realizar alguna gestión antes de guardar las intenciones?")
        
        if result == "yes":
            return False, pd.DataFrame([])
        
        newWb.Close(SaveChanges=False)
    
    #Antes de subir las intenciones se deben validar que estén correctas:        
    validacionCamposVacios = validarIntenciones(intencionesNuevas,macroActivo)
        
    rangoDeDatos.ClearComments()
    rangoDeDatos.Interior.Color = 0xD9D9D9
    #Si hay algun dato por corregir se coloca el comentario del error de validación y se cambia de color la celda 
    #Pintar y comentar campos vacío obligatorios
    
    
    for fila in range(0,len(validacionCamposVacios)):
        for columna in range(0,totalColumnas):
            if validacionCamposVacios.iloc[fila][columna] != "":
                    #pintar y comentar campos con errores de datos
                hojaFormulario.Cells(fila + 8,columna+2).AddComment(validacionCamposVacios.iloc[fila][columna])        
                hojaFormulario.Cells(fila + 8,columna+2).Interior.Color = 0x3755ED
    
    
    #Ahora Se pueden subir las intenciones a la base de datos.
    if  ~validacionCamposVacios.eq("").all().all():
        validacionCamposVaciosCopia = validacionCamposVacios.copy()
        validacionCamposVaciosCopia[validacionCamposVaciosCopia.eq("La cantidad ingresada no puede superar la cantidad disponible")] = ""
        if  ~validacionCamposVaciosCopia.eq("").all().all(): #hay mas errores por corregir, no se puede continuar
            mostrarMensajeAdvertencia("Todavía no está lista la intencion")
            return False, pd.DataFrame([])
        else:
            if mostrarMensajeAdvertenciaSiNo("¿Desea guardar la intención tal y como está ingresada?"):
                mostrarMensajeAdvertencia("Para guardar las intenciones escriba en la consola: 'SI ACEPTO'")
                aceptacion = input("Escriba 'SI ACEPTO': ")
                if aceptacion.strip().upper() == "SI ACEPTO":
                    print("Guardar log")
                else:
                    return False, pd.DataFrame([])
            else:
                return False, pd.DataFrame([])
                
    
    intenciones = pd.DataFrame([],columns =definiciones.parametros["Valor"]["camposArchivoIntenciones"].split("-"))
    for idIntencion in intencionesNuevas.index: 
        if macroActivo == "Renta Variable":
            intenciones = pd.concat([intenciones,completarDatosIntencionRV(intencionesNuevas.loc[idIntencion,],"CREACION")])
        if macroActivo == "Deuda Privada":
            intenciones = pd.concat([intenciones,completarDatosIntencionDPR(intencionesNuevas.loc[idIntencion,],"CREACION")])
        if macroActivo == "Deuda Pública":
            intenciones = pd.concat([intenciones,completarDatosIntencionDPU(intencionesNuevas.loc[idIntencion,],"CREACION")])
        if macroActivo == "Fondos":
            intenciones = pd.concat([intenciones,completarDatosIntencionFondos(intencionesNuevas.loc[idIntencion,],"CREACION")])
        if macroActivo == "Forex":
            intenciones = pd.concat([intenciones,completarDatosIntencionForex(intencionesNuevas.loc[idIntencion,],"CREACION")])
        if macroActivo == "Swaps":
            intenciones = pd.concat([intenciones,completarDatosIntencionSwaps(intencionesNuevas.loc[idIntencion,],"CREACION")])
    
    intenciones = intenciones.where(~intenciones.isna(), other="")
    intenciones.loc[intenciones["TipoOrden"] != "LÍMITE","PrecioLimite"] = ""
    intenciones.loc[intenciones["TipoOrden"] != "LÍMITE","TasaLimite"] = ""

    resultado = guardarIntenciones(intenciones,definiciones.parametros["Valor"]["rutaIntenciones"],getpass.getuser().upper())
    
    #Vamos  a guardar log de las acciones registradas por el usuario
    if resultado == True:            
        try:
            intenciones[validacionCamposVacios.eq("La cantidad ingresada no puede superar la cantidad disponible").T.any().tolist()].apply(lambda x: actualizarLogIntenciones(str(int(x["Id"])),getpass.getuser().upper(),"Creacion de intenciones que no cumplen con la restricción de cupo."),axis=1)
        except:
            mostrarMensajeAdvertencia("Hubo un error guardando datos en el log.")

    return resultado, intenciones
    
def subirEdicion(hojaFormulario,ultFilaIntenciones,camposFormularioEdicion,macroActivo,archivo):

    totalColumnas = len(camposFormularioEdicion)
    rangoEncabezado = hojaFormulario.Range(hojaFormulario.Cells(7,2),hojaFormulario.Cells(7,totalColumnas+1))
    rangoDeDatos = hojaFormulario.Range(hojaFormulario.Cells(8,2),hojaFormulario.Cells(ultFilaIntenciones,totalColumnas+1))
      
    data = rangoDeDatos.Value
    intencionesNuevas = pd.DataFrame(data,columns =rangoEncabezado.Value[0])

    if len(intencionesNuevas) == 0:
            mostrarMensajeAdvertencia("No hay intenciones para subir")
            return False, pd.DataFrame([])
    #Antes de subir las intenciones se deben validar que estén correctas:
        
    validacionCamposVacios = validarIntenciones(intencionesNuevas,macroActivo)

    rangoDeDatos.ClearComments()
    rangoDeDatos.Interior.Color = 0xD9D9D9
    #Si hay algun dato por corregir se coloca el comentario del error de validación y se cambia de color la celda 
    #Pintar y comentar campos vacío obligatorios
    
    for fila in range(0,len(validacionCamposVacios)):
        for columna in range(0,totalColumnas):
            if validacionCamposVacios.iloc[fila][columna] != "":
                    #pintar y comentar campos con errores de datos
                hojaFormulario.Cells(fila + 8,columna+2).AddComment(validacionCamposVacios.iloc[fila][columna])        
                hojaFormulario.Cells(fila + 8,columna+2).Interior.Color = 0x3755ED
    
    
    #Ahora Se pueden subir las intenciones a la base de datos.
    if  ~validacionCamposVacios.eq("").all().all():
        validacionCamposVaciosCopia = validacionCamposVacios.copy()
        validacionCamposVaciosCopia[validacionCamposVaciosCopia.eq("La cantidad ingresada no puede superar la cantidad disponible")] = ""
        if  ~validacionCamposVaciosCopia.eq("").all().all(): #hay mas errores por corregir, no se puede continuar
            mostrarMensajeAdvertencia("Todavía no está lista la intencion")
            return False, pd.DataFrame([])
        else:
            if mostrarMensajeAdvertenciaSiNo("¿Desea guardar la intención tal y como está ingresada?"):
                mostrarMensajeAdvertencia("Para guardar las intenciones escriba en la consola: 'SI ACEPTO'")
                aceptacion = input("Escriba 'SI ACEPTO': ")
                if aceptacion.strip().upper() == "SI ACEPTO":
                    print("Guardar log")
                else:
                    return False, pd.DataFrame([])
            else:
                return False, pd.DataFrame([])       
    
    intenciones = pd.DataFrame([],columns =definiciones.parametros["Valor"]["camposArchivoIntenciones"].split("-"))
    for idIntencion in intencionesNuevas.index: 
        if macroActivo == "Renta Variable":
            intenciones = pd.concat([intenciones,completarDatosIntencionRV(intencionesNuevas.loc[idIntencion,],"EDICION")])
        if macroActivo == "Deuda Privada":
            intenciones = pd.concat([intenciones,completarDatosIntencionDPR(intencionesNuevas.loc[idIntencion,],"EDICION")])
        if macroActivo == "Deuda Pública":
            intenciones = pd.concat([intenciones,completarDatosIntencionDPU(intencionesNuevas.loc[idIntencion,],"EDICION")])
        if macroActivo == "Fondos":
            intenciones = pd.concat([intenciones,completarDatosIntencionFondos(intencionesNuevas.loc[idIntencion,],"EDICION")])
        if macroActivo == "Forex":
            intenciones = pd.concat([intenciones,completarDatosIntencionForex(intencionesNuevas.loc[idIntencion,],"EDICION")])
        if macroActivo == "Swaps":
            intenciones = pd.concat([intenciones,completarDatosIntencionSwaps(intencionesNuevas.loc[idIntencion,],"EDICION")])
    
            
    intenciones = intenciones.where(~intenciones.isna(), other="")
    intenciones.loc[intenciones["TipoOrden"] != "LÍMITE","PrecioLimite"] = ""
    intenciones.loc[intenciones["TipoOrden"] != "LÍMITE","TasaLimite"] = ""
    
    resultado,intenciones =  actualizarIntenciones(intenciones,definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
    if resultado == True:            
        try:
            intenciones[validacionCamposVacios.eq("La cantidad ingresada no puede superar la cantidad disponible").T.any().tolist()].apply(lambda x: actualizarLogIntenciones(str(int(x["Id"])),getpass.getuser().upper(),"Edicion de intenciones que no cumplen con la restricción de cupo."),axis=1)
        except:
            mostrarMensajeAdvertencia("Hubo un error guardando datos en el log.")
    
    return resultado, intenciones


def subirRenovacion(hojaFormulario,ultFilaIntenciones,camposFormularioRenovacion,macroActivo,archivo):   

    totalColumnas = len(camposFormularioRenovacion)
    rangoEncabezado = hojaFormulario.Range(hojaFormulario.Cells(7,2),hojaFormulario.Cells(7,totalColumnas+1))
    rangoDeDatos = hojaFormulario.Range(hojaFormulario.Cells(8,2),hojaFormulario.Cells(ultFilaIntenciones,totalColumnas+1))
        
    data = rangoDeDatos.Value
    intencionesNuevas = pd.DataFrame(data,columns =rangoEncabezado.Value[0])

    if len(intencionesNuevas) == 0:
            mostrarMensajeAdvertencia("No hay intenciones para subir")
            return False, pd.DataFrame([])
    #Antes de subir las intenciones se deben validar que estén correctas:
    #Generales
        #1. QUe ningún campo obligario esté vacío
        
    validacionCamposVacios = intencionesNuevas.isna()
    validacionCamposVacios = validacionCamposVacios.applymap(lambda x: "Celda vacía" if x == True else "")
    #5. Vigencias desde debe ser igual o superior a hoy, si está vacío se coloca hoy. La vigencia hasta debe ser igual o superior a hoy
    
    validacionVigencia = intencionesNuevas[["Vigente desde","Vigente hasta"]].astype('str').apply(lambda x:validarFechaDesdeHastaNueva(x),axis=1)
    
    validacionCamposVacios.loc[list(map(lambda x: not x[0],validacionVigencia)),"Vigente desde"] = "Por favor revise la fecha, no puede ser inferior a hoy."
    validacionCamposVacios.loc[intencionesNuevas["Vigente desde"].isna(),"Vigente desde"] = "Celda Vacía"
    validacionCamposVacios.loc[list(map(lambda x: not x[1],validacionVigencia)),"Vigente hasta"] = "Por favor revise la fecha, no puede ser inferior a la vigencia desde."
    validacionCamposVacios.loc[intencionesNuevas["Vigente hasta"].isna(),"Vigente hasta"] = "Celda Vacía"
    
    
    rangoDeDatos.ClearComments()
    rangoDeDatos.Interior.Color = 0xD9D9D9
    #Si hay algun dato por corregir se coloca el comentario del error de validación y se cambia de color la celda 
    #Pintar y comentar campos vacío obligatorios
    
    for fila in range(0,len(validacionCamposVacios)):
        for columna in range(0,totalColumnas):
            if validacionCamposVacios.iloc[fila][columna] != "":
                #pintar y comentar campos con errores de datos                
                if hojaFormulario.Cells(fila + 8,columna+2).Locked == False:
                    hojaFormulario.Cells(fila + 8,columna+2).AddComment(validacionCamposVacios.iloc[fila][columna])        
                    hojaFormulario.Cells(fila + 8,columna+2).Interior.Color = 0x3755ED
    
    
    #Ahora Se pueden subir las intenciones a la base de datos.
    if  ~validacionCamposVacios.eq("").all().all():
        mostrarMensajeAdvertencia("Todavía no está lista la intencion")
        return False, pd.DataFrame([])
       
    #En intencionesNuevas están los datos listos para actualizar el archivo de intenciones        
    resultado, intenciones = renovarIntenciones(intencionesNuevas,definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
    intenciones = intenciones.where(~intenciones.isna(), other="")
    
    return resultado, intenciones

def subirCancelacion(hojaFormulario,ultFilaIntenciones,macroActivo,archivo):
    
    Ids = hojaFormulario.Range(hojaFormulario.Cells(8,2),hojaFormulario.Cells(ultFilaIntenciones,2)).Value
    if esUnaLista(Ids):
        Ids = list(map(lambda x: int(x[0]),Ids))
    else:
        Ids = [int(Ids)]
         
    
    resultado, intenciones = cancelarIntenciones(Ids,definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
    
    return resultado, intenciones
 
def obtenerNemosForex(tipoOperacion):
    especies = definiciones.especies.copy()
    if tipoOperacion in ["COMPRA SPOT","VENTA SPOT"]:
        nemos = especies.loc[especies["Macro Activo"] == "SPOT FOREX","Nemo intenciones"].tolist()
    elif tipoOperacion in ["COMPRA NDF","VENTA NDF"]:
        nemos = especies.loc[especies["Macro Activo"] == "NDF FOREX","Nemo intenciones"].tolist()
    elif tipoOperacion in ["COMPRA FUTURO", "VENTA FUTURO"]:
        nemos = especies.loc[especies["Macro Activo"] == "FUTUROS FOREX","Nemo intenciones"].tolist()
    elif tipoOperacion in ["COMPRA OPCIONES","VENTA OPCIONES"]: 
        nemos = especies.loc[especies["Macro Activo"] == "OPCIONES FOREX","Nemo intenciones"].tolist()
    else:
        return "-"
    if len(nemos) >0:
        nemos = str(nemos)
        nemos = nemos[1:-1]
        nemos = nemos.replace("'","")
        nemos = nemos.replace(", ",",")
        return nemos
    else:
        return "-"
       
def obtenerFuturosDPU():
    especies = definiciones.especies.copy()
    nemos = especies.loc[especies["Macro Activo"] == "FUTUROS TES","Nemo intenciones"]
    if len(nemos) == 0:
        return "-"
    else:
        nemos = str(nemos.tolist())
        nemos = nemos[1:-1]
        nemos = nemos.replace("'","")
        nemos = nemos.replace(", ",",")
        return nemos
    
     
def guardarIntencionesTraders(data,rutaIntenciones,rutaLogIntenciones,usuario):
    
    archivos = data["IngresadoPor"].unique().tolist()
    data["Id"] = data["Id"].astype("int64")
    data.set_index(["Id"],drop = False,inplace=True)
    
    for archivo in archivos:        
        
        intencionesNuevas = data[data["IngresadoPor"] == archivo]
        archivoIntenciones = rutaIntenciones +"/"+ archivo+".csv"            
        intenciones = pd.read_csv(archivoIntenciones,sep=",")
        intenciones["Id"] = intenciones["Id"].astype("int64")
        intenciones.set_index(["Id"],drop = False,inplace=True)        
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["CantEjecutada"]] = data["CantEjecutada"]
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["CantPendiente"]] = data["CantidadTotal"].astype("float") - data["CantEjecutada"].astype("float")
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["Ejecutado"]] = 100 * (data["CantEjecutada"].astype("float"))/data["CantidadTotal"].astype("float")
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["ComentariosTrader"]] = data["ComentariosTrader"]  
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["UltimaModificacion"]] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")                         
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["ModificadoPor"]] = usuario
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["UltimaModifTrader"]] = datetime.now().strftime("%d/%m/%Y-%H:%M:%S")
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["Trader"]] = usuario
        intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),["Estado"]] = intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),].apply(lambda x: "Ejecutada/Total" if float(x["Ejecutado"])> 99.5 else ( "En proceso" if float(x["Ejecutado"])== 0 else "Ejecutada/Parcial"),axis =1)
        try:
            intenciones.to_csv(archivoIntenciones,mode = "w", header = True,index=False,sep=",") 
        except:
            mostrarMensajeAdvertencia("No fué posible modificar las intenciones.")
        
        #Guardar trazabilidad de las operaciones
        try:
            archivoIntencionesLog = rutaLogIntenciones +"/"+ archivo+".csv" 
            intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"])].to_csv(archivoIntencionesLog,mode = "a", header = False,index=False,sep=",")
        except:
            mostrarMensajeAdvertencia("No fué posible guardar la trazabilidad de las intenciones.")
        #Guardar Log de intecniones
        try:
            intenciones.loc[intenciones["Id"].isin(intencionesNuevas["Id"]),].apply(lambda x: actualizarLogIntenciones(x["Id"],getpass.getuser().upper(),"El Trader realiza actualización de estado en la intención: " + x["Estado"]),axis=1)
        except:
            mostrarMensajeAdvertencia("No fué posible guardar Log de intenciones.")
    return True
       
def validarNemosRV(nemotecnico, IdPortafolio, mercado, tipoOperacion):
    
    if nemotecnico == None or IdPortafolio == None or tipoOperacion == None or mercado == None:
        return False
    
    macroActivoConsulta = "RV " + mercado    
    if tipoOperacion == "COMPRA":
        return True
    else: #VENTA 
        nemosVenta = obtenerNemosdePortafolio(str(IdPortafolio),macroActivoConsulta)
        if nemotecnico in nemosVenta:
            return True
        else:
            return False
  
def validarNemosDPR(nemotecnico,IdPortafolio,tipoOperacion):

    macroActivoConsulta = "DEUDA PRIVADA"   
    
    if nemotecnico == None or IdPortafolio == None or tipoOperacion == None:
        return False
    
    if tipoOperacion == "COMPRA":
        return True
          
    else: #VENTA
        nemosVenta = obtenerNemosdePortafolio(str(IdPortafolio),macroActivoConsulta)
        nemosVenta = nemosVenta 
        if nemotecnico in nemosVenta:
            return True
        else:
            return False
        
def validarNemosDPU(nemotecnico,IdPortafolio,tipoOperacion):

    macroActivoConsulta = "DEUDA PÚBLICA"   
    
    if nemotecnico == None or IdPortafolio == None or tipoOperacion == None:
        return False
    
    if  tipoOperacion =="COMPRA":
        return True  
    elif tipoOperacion == "VENTA":  
    
        nemosVenta = obtenerNemosdePortafolio(str(IdPortafolio),macroActivoConsulta)
        nemosVenta = nemosVenta 
        if nemotecnico in nemosVenta:
            return True
        else:
            return False
    else:
        if nemotecnico in obtenerFuturosDPU():
            return True
        else:
            return False
        
def validarNemosSwaps(nemotecnico,IdPortafolio,tipoOperacion):
    
    if nemotecnico == None or IdPortafolio == None or tipoOperacion == None:
        return False
    
    if tipoOperacion == "VENTA":
        if nemotecnico =="Recibo Tasa Fija-Entrego Tasa Variable":
            return True
        else:
            return False
    elif tipoOperacion == "COMPRA":
        if nemotecnico =="Recibo Variable-Entrego Tasa Fija":
            return True
        else:
            return False
    else:
        if nemotecnico in obtenerNemosdePortafolio(IdPortafolio,"SWAP"):
            return True
        else:
            return False
        
def validarNemosFondos(nemotecnico,IdPortafolio,tipoOperacion):
    
    if tipoOperacion == "APERTURA":
       return True
    else:
        if nemotecnico in obtenerNemosdePortafolio(IdPortafolio,"PARTICIPACIÓN EN FONDOS"):
            return True
        else:
            return False
def validarNemosForex(nemotecnico,tipoOperacion):
    if nemotecnico == None or  tipoOperacion == None:
        return False
    if nemotecnico in obtenerNemosForex(tipoOperacion):
        return True
    else:
        return False
            
def validarPrecioLimite(tipoOrden,precioLimite):   

    if tipoOrden == "LÍMITE":
        if pd.isna(precioLimite):
            return False
        else:
            if esUnNumero(precioLimite):
                return True
            else:
                return False    
    else:
        return True
    
def validarTasaLimite(tipoOrden,tasaLimite):   

    if tipoOrden == "LÍMITE":
        if pd.isna(tasaLimite):
            return False
        else:
            if esUnNumero(tasaLimite):
                return True
            else:
                return False    
    else:
        return True
    
def validarTipoOrden(tipoOrden, tasaPrecioLimite):
    
    if not pd.isna(tasaPrecioLimite) and tasaPrecioLimite != "":
        if tipoOrden != "LÍMITE":
            return False
        else:
            return True
    else:
        return True

def validarCantidad(intencion):
    
    if "Cantidad disponible(Millones)" in intencion.index.tolist():
        cantidadDisponible = intencion["Cantidad disponible(Millones)"]
    else:
        cantidadDisponible = intencion["Cantidad disponible"]
        
    if "Cantidad disponible(Millones)" in intencion.index.tolist():
        cantidad = intencion["Cantidad(Millones)"]
    else:
        cantidad = intencion["Cantidad"]
    
    if cantidad == None or cantidadDisponible == None:
        
        return "Celda Vacía"
    
    if  not esUnNumero(cantidad):
       
        return "No es un número"
    
    if float(cantidad) <= 0:
        
        return "La cantidad no puede ser negativa"    
   
    if esUnNumero(cantidadDisponible):
        if float(cantidad) > float(cantidadDisponible):
            
            return "La cantidad ingresada no puede superar la cantidad disponible"
        else:
            
            return ""
    else:
        
        return ""
        
    

def validarIntenciones(intencionesNuevas,macroActivo):
    
    
    validacionCamposVacios =  intencionesNuevas.isna()
    validacionCamposVacios = validacionCamposVacios.applymap(lambda x: "Celda vacía" if x == True else "")
    especies = definiciones.especies.copy()
    
    #PORTAFOLIO:
    #2. Que el portafolio si exista
    validacionPortafolio = intencionesNuevas[["Id Portafolio"]]
    validacionPortafolio = validacionPortafolio.apply(lambda x: False if "No encontrado"==obtenerNombrePortafolioPorId(x["Id Portafolio"]) else True,axis=1)
   
    #VIGENCIA:
    #3. Vigencias desde debe ser igual o superior a hoy, si está vacío se coloca hoy. La vigencia hasta debe ser igual o superior a hoy
    validacionVigencia = intencionesNuevas[["Vigente desde","Vigente hasta"]].apply(lambda x: validarFechaDesdeHastaNueva(x),axis=1)

    if macroActivo == "Renta Variable": 
        
        #PRECIO LÍMITE: 
        #Si marcó tipo Orden como precio límite debe haber un dato en precio límite
        validacionPrecioLimite = intencionesNuevas.apply(lambda x: validarPrecioLimite(x["Tipo orden"],x["Precio límite"]),axis = 1)
        
        #TIPO ORDEN
        validaciontipoOrden = intencionesNuevas.apply(lambda x: validarTipoOrden(x["Tipo orden"],x["Precio límite"]),axis = 1)
         
        #CANTIDAD:   
        #Que la cantidad ingresada esté dentro del rango permitido
        
        
        validacionCamposVacios["Cantidad"] = intencionesNuevas.apply(lambda x: validarCantidad(x),axis=1)
        
            
        #MERCADO:
        validacionMercado = intencionesNuevas.apply(lambda x: True if x["Mercado"] != None else False,axis =1)
        
        
        if validacionMercado.all():
            
            #NEMOTÉCNICO:
            # Que el nemo si pertenezca al portafolio, esto apllica para algunos mercados RV y solo para ventas.        
            validacionNemotecnicos = intencionesNuevas.apply(lambda x: validarNemosRV(x["Nemotécnico"],x["Id Portafolio"],x["Mercado"],x["Tipo operación"]),axis = 1)
            validacionCamposVacios.loc[validacionNemotecnicos == False,"Nemotécnico"] = "El nemotécnico no pertenece al portafolio"
            validacionCamposVacios.loc[intencionesNuevas["Nemotécnico"].isna(),"Nemotécnico"] = "Celda Vacía"
            
            #CANTIDAD DISPONIBLE:
            #Que la cantidad disponible sea la correcta. Solo se puede revisar si el mercado no es vacío      
            validacionCantidadDisponible = intencionesNuevas.apply(lambda x :True if (x["Tipo operación"] == "VENTA" and x["Cantidad disponible"] == obtenerCantidadDisponibleNemos(x["Nemotécnico"],x["Id Portafolio"],"RV "+ str(x["Mercado"]))) or( x["Tipo operación"] == "COMPRA" and x["Cantidad disponible"] == obtenerCuposcompras(x["Id Portafolio"],"Renta Variable",x["Nemotécnico"]) ) else False,axis =1 )
            validacionCamposVacios.loc[validacionCantidadDisponible == False,"Cantidad disponible"] = "La cantidad disponible no es la real"
            validacionCamposVacios.loc[intencionesNuevas["Cantidad disponible"].isna(),"Cantidad disponible"] = "Celda Vacía"
    
          
        
        validacionCamposVacios.loc[intencionesNuevas["Portafolio"].isna(),"Portafolio"] = "Celda Vacía"
        validacionCamposVacios.loc[:, "Precio límite"] = ""
        validacionCamposVacios.loc[:, "Comentarios PM"] = ""
        validacionCamposVacios.loc[validacionPrecioLimite == False, "Precio límite"] = "Falta ingresar el precio límite"
        validacionCamposVacios.loc[~validaciontipoOrden, "Tipo orden"] = "No está correcto el tipo de orden"
        validacionCamposVacios.loc[:,"Emisor"] = ""
        
        
    
    if macroActivo == "Deuda Privada":
        
        opsGenerica = intencionesNuevas.apply(lambda x: True if x["Nemotécnico"] in especies.loc[especies["Nemotecnico"] == "GENERICO DPR","Nemo intenciones"].tolist()  else False, axis =1)
        #TASA LÍMITE:
        validacionTasaLimite = intencionesNuevas.apply(lambda x: validarTasaLimite(x["Tipo orden"],x["Tasa límite"]),axis = 1)
        
        #TIPO ORDEN
        validaciontipoOrden = intencionesNuevas.apply(lambda x: validarTipoOrden(x["Tipo orden"],x["Tasa límite"]),axis = 1)
        
        #CANTIDAD:
        validacionCamposVacios["Cantidad"] = intencionesNuevas.apply(lambda x: validarCantidad(x),axis=1)
        
        # NEMOTÉCNICO: 
        validacionNemotecnicos = intencionesNuevas.apply(lambda x: validarNemosDPR(x["Nemotécnico"],x["Id Portafolio"],x["Tipo operación"]),axis = 1)
                            
        #CANTIDAD DISPONIBLE:
        # Que la cantidad disponible sea la correcta.        
        validacionCantidadDisponible = intencionesNuevas.apply(lambda x :True if (x["Tipo operación"] == "VENTA" and x["Cantidad disponible(Millones)"] == obtenerCantidadDisponibleNemos(x["Nemotécnico"],x["Id Portafolio"],"DEUDA PRIVADA")) or( x["Tipo operación"] == "COMPRA" and x["Cantidad disponible(Millones)"] == obtenerCuposcompras(x["Id Portafolio"],"Deuda Privada",x["Nemotécnico"]) ) else False,axis =1 )
        
        
        validacionCamposVacios.loc[intencionesNuevas["Portafolio"].isna(),"Portafolio"] = "Celda Vacía"
        validacionCamposVacios.loc[validacionNemotecnicos == False,"Nemotécnico"] = "El nemotécnico no pertenece al portafolio"
        validacionCamposVacios.loc[intencionesNuevas["Nemotécnico"].isna(),"Nemotécnico"] = "Celda Vacía"
        validacionCamposVacios.loc[:, "Tasa límite"] = ""
        validacionCamposVacios.loc[:, "Comentarios PM"] = ""
        validacionCamposVacios.loc[~opsGenerica, "Desde"] = ""
        validacionCamposVacios.loc[~opsGenerica, "Hasta"] = ""
        validacionCamposVacios.loc[opsGenerica,"Emisor"] = ""
        validacionCamposVacios.loc[validacionTasaLimite == False, "Tasa límite"] = "Falta ingresar la tasa límite"
        validacionCamposVacios.loc[~validaciontipoOrden, "Tipo orden"] = "No está correcto el tipo de orden"
        #validacionCamposVacios.loc[~opsGenerica,"Tipo orden"] = ""
        validacionCamposVacios.loc[validacionCantidadDisponible == False,"Cantidad disponible(Millones)"] = "La cantidad disponible no es la real"
        validacionCamposVacios.loc[intencionesNuevas["Cantidad disponible(Millones)"].isna(),"Cantidad disponible(Millones)"] = "Celda Vacía"
        validacionCamposVacios.loc[opsGenerica,"Cantidad disponible(Millones)"] = ""
        validacionCamposVacios.loc[:,"Emisor"] = ""
        validacionCamposVacios.loc[~opsGenerica,"Indicador"] = ""
         
    
    if macroActivo == "Deuda Pública":
        
        #TASA LÍMITE:
        validacionTasaLimite = intencionesNuevas.apply(lambda x: validarTasaLimite(x["Tipo orden"],x["Tasa límite"]),axis = 1)
        
        #TIPO ORDEN
        validaciontipoOrden = intencionesNuevas.apply(lambda x: validarTipoOrden(x["Tipo orden"],x["Tasa límite"]),axis = 1)
        
        #CANTIDAD:
        #Que la cantidad ingresada esté dentro del rango permitido
        validacionCamposVacios["Cantidad"] = intencionesNuevas.apply(lambda x: validarCantidad(x),axis=1)
        
        #NEMOTÉCNICO:    
        # Que el nemo si pertenezca al portafolio, esto apllica para algunos mercados RV y solo para ventas.        
        validacionNemotecnicos = intencionesNuevas.apply(lambda x: validarNemosDPU(x["Nemotécnico"],x["Id Portafolio"],x["Tipo operación"]),axis = 1)
        
        #CANTIDAD DISPONIBLE:
        #Que la cantidad disponible sea la correcta.        
        validacionCantidadDisponible = intencionesNuevas.apply(lambda x :True if ("VENTA" in x["Tipo operación"]  and x["Cantidad disponible(Millones)"] == obtenerCantidadDisponibleNemos(x["Nemotécnico"],x["Id Portafolio"],"DEUDA PÚBLICA")) or("COMPRA" in x["Tipo operación"] and x["Cantidad disponible(Millones)"] == obtenerCuposcompras(x["Id Portafolio"],"Deuda Pública",x["Nemotécnico"]) ) else False,axis =1 )
        validacionCantidadDisponible[intencionesNuevas["Tipo operación"].str.contains("FUTURO")] = True
        
        
        validacionCamposVacios.loc[intencionesNuevas["Portafolio"].isna(),"Portafolio"] = "Celda Vacía"
        validacionCamposVacios.loc[validacionNemotecnicos == False,"Nemotécnico"] = "El nemotécnico no pertenece al portafolio"
        validacionCamposVacios.loc[intencionesNuevas["Nemotécnico"].isna(),"Nemotécnico"] = "Celda Vacía"
        validacionCamposVacios.loc[:, "Tasa límite"] = ""
        validacionCamposVacios.loc[:, "Comentarios PM"] = ""
        validacionCamposVacios.loc[validacionTasaLimite == False, "Tasa límite"] = "Falta ingresar la tasa límite"
        validacionCamposVacios.loc[~validaciontipoOrden, "Tipo orden"] = "No está correcto el tipo de orden"
        validacionCamposVacios.loc[validacionCantidadDisponible == False,"Cantidad disponible(Millones)"] = "La cantidad disponible no es la real"
        validacionCamposVacios.loc[intencionesNuevas["Cantidad disponible(Millones)"].isna(),"Cantidad disponible(Millones)"] = "Celda Vacía"
        validacionCamposVacios.loc[:,"Emisor"] = ""
    
    if macroActivo == "Forex":  
        
        #PRECIO LÍMITE: 
        # Si marcó tipo Orden como precio límite debe haber un dato en precio límite
        validacionPrecioLimite = intencionesNuevas.apply(lambda x: validarPrecioLimite(x["Tipo orden"],x["Precio límite"]),axis = 1)
        
        #TIPO ORDEN
        validaciontipoOrden = intencionesNuevas.apply(lambda x: validarTipoOrden(x["Tipo orden"],x["Precio límite"]),axis = 1)
        
        #NEMOTÉCNICO:
        validacionNemotecnicos = intencionesNuevas.apply(lambda x: validarNemosForex(x["Nemotécnico"],x["Tipo operación"]),axis =1)
        
        #CANTIDAD:
        #La cantidad ingresada debe ser un número
        validacionCantidad = intencionesNuevas.apply(lambda x: True if esUnNumero(x["Cantidad"]) else False, axis =1)
        
        
        validacionCamposVacios.loc[validacionCantidad == False,"Cantidad"] = "La cantidad ingresada debe ser un número"
        validacionCamposVacios.loc[intencionesNuevas["Cantidad"].isna(),"Cantidad"] = "Celda Vacía"
        validacionCamposVacios.loc[:, "Precio límite"] = ""
        validacionCamposVacios.loc[:, "Comentarios PM"] = ""
        validacionCamposVacios.loc[validacionPrecioLimite == False, "Precio límite"] = "Falta ingresar el precio límite"
        validacionCamposVacios.loc[~validaciontipoOrden, "Tipo orden"] = "No está correcto el tipo de orden"
        validacionCamposVacios.loc[:,"Emisor"] = ""
        validacionCamposVacios.loc[validacionNemotecnicos == False, "Nemotécnico"] = "Nemotécnico incorrecto" 
            
    if macroActivo == "Fondos":      
    
        #NEMOTÉCNICO:
        # Que el nemo si pertenezca al portafolio, esto apllica para algunos mercados .        
        validacionNemotecnicos = intencionesNuevas.apply(lambda x: validarNemosFondos(x["Nemotécnico"],x["Id Portafolio"],x["Tipo operación"]),axis = 1)        
        
        #CANTIDAD:
        #Que la cantidad ingresada esté dentro del rango permitido
        validacionCamposVacios["Cantidad"] = intencionesNuevas.apply(lambda x: validarCantidad(x),axis=1)
         
        #CANTIDAD DISPONIBLE:   
        ##Que la cantidad disponible sea la correcta. 
        validacionCantidadDisponible = intencionesNuevas.apply(lambda x :True if ( x["Tipo operación"] in["CANCELACION","RETIRO"]  and x["Cantidad disponible"] == obtenerCantidadDisponibleNemosFondos(x["Nemotécnico"],x["Id Portafolio"],"PARTICIPACIÓN EN FONDOS",x["Tipo operación"])) or(x["Tipo operación"] in ["ADICION","APERTURA"] and x["Cantidad disponible"] == obtenerCuposcompras(x["Id Portafolio"],"Fondos",x["Nemotécnico"]) ) else False,axis =1 )
        
        validacionCamposVacios.loc[validacionNemotecnicos == False,"Nemotécnico"] = "El nemotécnico no pertenece al portafolio"
        validacionCamposVacios.loc[intencionesNuevas["Nemotécnico"].isna(),"Nemotécnico"] = "Celda Vacía"
        validacionCamposVacios.loc[validacionCantidadDisponible == False,"Cantidad disponible"] = "La cantidad disponible no es la real"
        validacionCamposVacios.loc[intencionesNuevas["Cantidad disponible"].isna(),"Cantidad disponible"] = "Celda Vacía"
        validacionCamposVacios.loc[:, "Comentarios PM"] = ""
        validacionCamposVacios.loc[:,"Emisor"] = ""
    
    if macroActivo == "Swaps":
        
        opsUnWind = intencionesNuevas.apply(lambda x: True if x["Tipo operación"] == "UNWIND" else False, axis=1)
        
        #TIPO OPERACIÓN:
        validaciontipoOp = intencionesNuevas.apply(lambda x: True if x["Tipo operación"] != None else False,axis =1)
        
        #TASA LÍMITE:
        # Si marcó tipo Orden como tasa límite debe haber un dato en tasa límite
        validacionTasaLimite = intencionesNuevas.apply(lambda x: validarTasaLimite(x["Tipo orden"],x["Tasa límite"]),axis = 1)
        
        #TIPO ORDEN
        validaciontipoOrden = intencionesNuevas.apply(lambda x: validarTipoOrden(x["Tipo orden"],x["Tasa límite"]),axis = 1)
            
        if validaciontipoOp.all():
            #NEMOTÉCNICO:
            # Que el nemo si pertenezca al portafolio, esto apllica para algunos mercados .        
            validacionNemotecnicos = intencionesNuevas.apply(lambda x: validarNemosSwaps(x["Nemotécnico"],x["Id Portafolio"],x["Tipo operación"]),axis = 1)
            validacionCamposVacios.loc[validacionNemotecnicos == False,"Nemotécnico"] = "El nemotécnico no pertenece al portafolio"
            validacionCamposVacios.loc[intencionesNuevas["Nemotécnico"].isna(),"Nemotécnico"] = "Celda Vacía"
            
            #CANTIDAD:
            #La cantidad ingresada debe ser un número
            #Que la cantidad disponible sea la correcta.                
            validacionCantidad = intencionesNuevas.apply(lambda x :True if (  (x["Tipo operación"] != "UNWIND" or (x["Tipo operación"] == "UNWIND" and x["Cantidad"] == obtenerCantidadDisponibleNemos(x["Nemotécnico"],x["Id Portafolio"],"SWAP")  ) )and   esUnNumero(x["Cantidad"])  ) else False,axis =1 )        
            validacionCamposVacios.loc[validacionCantidad == False,"Cantidad"] = "Existe un problema con el dato ingresado"
            validacionCamposVacios.loc[intencionesNuevas["Cantidad"].isna(),"Cantidad"] = "Celda Vacía"
            
        #Validacion Indicador y plazo hasta
        
        
        validacionCamposVacios.loc[opsUnWind,"Indicador"] = ""
        validacionCamposVacios.loc[opsUnWind,"Hasta"] = ""
        validacionCamposVacios.loc[opsUnWind,"Tipo orden"] = ""
        validacionCamposVacios.loc[opsUnWind,"Mercado"] = ""
        validacionCamposVacios.loc[:,"Emisor"] = ""        
        validacionCamposVacios.loc[:, "Tasa límite"] = ""
        validacionCamposVacios.loc[:, "Comentarios PM"] = ""
        validacionCamposVacios.loc[validacionTasaLimite == False, "Tasa límite"] = "Falta ingresar la tasa límite"
        validacionCamposVacios.loc[~validaciontipoOrden, "Tipo orden"] = "No está correcto el tipo de orden"
        
        
    validacionCamposVacios.loc[list(map(lambda x: not x[0],validacionVigencia)),"Vigente desde"] = "Por favor revise la fecha, no puede ser inferior a hoy"
    validacionCamposVacios.loc[intencionesNuevas["Vigente desde"].isna(),"Vigente desde"] = "Celda Vacía"
    validacionCamposVacios.loc[list(map(lambda x: not x[1],validacionVigencia)),"Vigente hasta"] = "Por favor revise la fecha, no puede ser inferior a la vigencia desde"
    validacionCamposVacios.loc[intencionesNuevas["Vigente hasta"].isna(),"Vigente hasta"] = "Celda Vacía"
    validacionCamposVacios.loc[~validacionPortafolio,"Id Portafolio"] = "El portafolio no existe"
    validacionCamposVacios.loc[:,"Portafolio"] = ""

    
    return validacionCamposVacios
     
    