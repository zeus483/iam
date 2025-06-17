# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 15:59:55 2022

@author: FRCASTRO

******************************************************************************
                        INTENCIONES ASSET MANAGEMENT
******************************************************************************      

FUNCIÓN PRINCIIPAL 
NUEVA VERSION 2.0
"""

#LIBRERIAS
##############################################################################
import win32com.client as win32
import win32com
from datetime import datetime
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
import warnings
import win32timezone
import win32api
import pandas as pd
import definiciones
import funciones
import clases
import shutil
import getpass
import sys
import os
import pythoncom    
import time


#FUNCIÓN PRINCIPAL
##############################################################################

warnings.filterwarnings("ignore")
if win32com.client.gencache.is_readonly == True:
    #allow gencache to create the cached wrapper objects
    win32com.client.gencache.is_readonly = False     
    # under p2exe the call in gencache to __init__() does not happen
    # so we use Rebuild() to force the creation of the gen_py folder
    win32com.client.gencache.Rebuild()    

#Se cargan los insumos los cuales seran guardado en variables globales que están en el módulo de definciones
print("Ha empezado la ejecución de la herramienta de Intenciones de Asset Management")
print("Espere hasta que se carguen todos los datos")
definiciones.crearVariablesDeDatos()
funciones.cargarDatos()
print("Todos los datos han sido cargados.")
print("Espere hasta que se abra el libro de Excel")

#Este dato se usa para garantizar que se est
# á usando la versión correcta
versionActual = '2.1.0'
parametros = definiciones.parametros

if versionActual != parametros["Valor"]["Version"]:
    funciones.mostrarMensajeAdvertencia("No estás usando la última versión, la versión correcta es la " + parametros["Valor"]["Version"])
    sys.exit()

# Crear una copia de la herramienta de intencciones y trabajar desde el local, para lograr que varios usuarios usen la herramienta al mismo tiempo   
if os.path.exists(parametros["Valor"]["rutaAccesoAplicacionExcel"]):
    if os.path.exists(parametros["Valor"]["rutaDestinoAplicacionExcel"]):
        #Eliminamos el archivo copia que ya existe
        os.remove(parametros["Valor"]["rutaDestinoAplicacionExcel"])
    shutil.copy(parametros["Valor"]["rutaAccesoAplicacionExcel"],parametros["Valor"]["rutaDestinoAplicacionExcel"])
else: 
    funciones.mostrarMensajeAdvertencia("No existe el archivo: " + parametros["Valor"]["rutaAccesoAplicacionExcel"])    
    sys.exit()

#Abrir la herramienta desde la copia que se llevó al equipo local
    
miUsuario =  getpass.getuser().upper()
usuarios = definiciones.usuarios
if miUsuario in usuarios.index:
    miRol = usuarios["Rol"][miUsuario]
else:
    funciones.mostrarMensajeAdvertencia("El usuario " + miUsuario + " no está registrado en la base de datos.")
    sys.exit()
    
#Ahora vamos a abrir la herramienta
excel = definiciones.excel
intencionesAM = excel.Workbooks.Open(Filename=parametros["Valor"]["rutaDestinoAplicacionExcel"],ReadOnly=0, UpdateLinks=0)

#Creación de la instancia de las hojas 
hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
hojaMonitor = intencionesAM.Worksheets("Monitor")
hojaFormulario = intencionesAM.Worksheets("Formulario Ordenes")
hojaTrazabilidad = intencionesAM.Worksheets("Trazabilidad")
hojaEjecutar = intencionesAM.Worksheets("Ejecutar Intenciones")
hojaDescargas= intencionesAM.Worksheets("Descargar Intenciones")
hojaEstadisticas = intencionesAM.Worksheets("Estadisticas")
hojaEspecies = intencionesAM.Worksheets("Especies") 
hojaTitulos = intencionesAM.Worksheets("Titulos")


#inicializar las listas desplegables en cada hoja:
#Ordenes
#MacroActivos
macroActivos = funciones.obtenerMacroActivos() 
for macroActivo in macroActivos:
    hojaOrdenes.OLEObjects("lstMacroActivosPM").Object.AddItem(macroActivo)
    hojaDescargas.OLEObjects("lstMacroActivoDescargas").Object.AddItem(macroActivo)    
hojaDescargas.OLEObjects("lstMacroActivoDescargas").Object.AddItem("(Todos)")

#Gerentes de portafolio
gerentes = definiciones.usuarios[definiciones.usuarios["Rol"].isin(["Administrador","PM"])].index.tolist() 
gerentes.append("(Todos)")   
for gerente in gerentes:
    hojaOrdenes.OLEObjects("listaGerentes").Object.AddItem(gerente)
    hojaDescargas.OLEObjects("lstGerentesDescarga").Object.AddItem(gerente)



#Monitor
    #MacroActivo
macroActivos = funciones.obtenerMacroActivosTrader()
for macroActivo in macroActivos:
    hojaMonitor.OLEObjects("lstMacroActivosTrader").Object.AddItem(macroActivo)

    
#Cargar todas la especies a la hoja de Títulos
especies = definiciones.especies.copy()
especies = especies.sort_values("Nemo intenciones")
especies[especies=="NAN"] = ""
hojaTitulos.Range("A1:E1").Value = ["Nemotécnico", "Especie", "Isin","Emisor","Macro Activo"]
hojaTitulos.Range("A2:E" + str(len(especies)+1)).Value = especies[["Nemo intenciones", "Especie", "Isin","Emisor cupos","Macro Activo"]].to_records(index=False)
hojaTitulos.Columns.AutoFit()
    

    

#Configurar eventos para cada botón
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnActualizarVistaIntenciones").Object,clases.BtnActulizarVistaIntencionesEvents)
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnSalir").Object,clases.BtnSalirEvents)
excel_events = win32.WithEvents(hojaMonitor.OLEObjects("btnSalirMonitor").Object,clases.BtnSalirEvents)
excel_events = win32.WithEvents(hojaMonitor.OLEObjects("btnActualizarMonitor").Object,clases.BtnActualizarIntencionesMonitorEvents)
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnCrearIntenciones").Object,clases.BtncrearIntencionesEvents)
excel_events = win32.WithEvents(excel,clases.workbookEvents)
excel_events = win32.WithEvents(hojaFormulario.OLEObjects("btnCancelarCrearIntenciones").Object, clases.BtnCancelarCreacionIntencionesEvents)
excel_events = win32.WithEvents(hojaFormulario.OLEObjects("btnSubirIntenciones").Object,clases.BtnSubirIntencionesEvents )
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnEditarIntenciones").Object, clases.BtnEditarIntencionesEvents )
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnhistoriaIntencion").Object, clases.BtnHistoriaIntencionEvents)
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnCancelarIntencion").Object, clases.BtnCancelarIntencionEvents)
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnRenovarIntencion").Object, clases.BtnRenovarIntencionEvents)
excel_events = win32.WithEvents(hojaTrazabilidad.OLEObjects("btnCerrarTrazabilidad").Object, clases.BtnCerrarTrazabilidadEvents)
excel_events = win32.WithEvents(hojaMonitor.OLEObjects("btnEjecutarIntenciones").Object, clases.btnEjecutarIntencionesEvents)
excel_events = win32.WithEvents(hojaEjecutar.OLEObjects("btnGuardarIntencionesTrader").Object, clases.btnGuardarIntencionesTraderEvents)
excel_events = win32.WithEvents(hojaEjecutar.OLEObjects("btnCancelarIntencionesTrader").Object, clases.btnCancelarIntencionesTraderEvents)
excel_events = win32.WithEvents(hojaMonitor.OLEObjects("btnTrazabilidadTraders").Object, clases.BtnHistoriaIntencionEvents)
excel_events = win32.WithEvents(hojaDescargas.OLEObjects("btnConsultaIntencionesDescargas").Object, clases.BtnConsultarIntencionesEvents)
excel_events = win32.WithEvents(hojaDescargas.OLEObjects("btnDescargarIntenciones").Object, clases.BtnGuardarIntencionesExcelEvents)
excel_events = win32.WithEvents(hojaDescargas.OLEObjects("btnSalirDescargas").Object, clases.BtnSalirDescargasEvents)
excel_events = win32.WithEvents(hojaMonitor.OLEObjects("btnDescargarIntencionesTraders").Object, clases.BtnDescargarIntencionesEvents)
excel_events = win32.WithEvents(hojaOrdenes.OLEObjects("btnDescargarIntencionesGerentes").Object, clases.BtnDescargarIntencionesEvents)
excel_events = win32.WithEvents(hojaDescargas.OLEObjects("btnVerEstadisticas").Object, clases.BtnVerEstadisticasEvents)
excel_events = win32.WithEvents(hojaEstadisticas.OLEObjects("btnCerrarEstadisticas").Object,clases.BtnSalirEstadisticasEvents)



#Configurar roles

if miRol == "PM":
    hojaTrazabilidad.Visible = False
    hojaMonitor.Visible = definiciones.xlVeryHidden
    hojaFormulario.Visible = False
    hojaEjecutar.Visible = False
    hojaOrdenes.Visible = True
    hojaDescargas.Visible = False
    hojaEstadisticas.Visible = False
    hojaEspecies.Visible = definiciones.xlVeryHidden
    hojaTitulos.Visible = True
    hojaOrdenes.OLEObjects("lstMacroActivosPM").Object.Value = "Renta Variable"
    hojaOrdenes.OLEObjects("listaGerentes").Object.Value = miUsuario
    hojaOrdenes.OLEObjects('fechaDesdeIntenciones').Object.Value = datetime.now().strftime("%d/%m/%Y")
    funciones.actualizarVistaGerentes()
    hojaOrdenes.Activate()
    
elif miRol == "Trader":    
    hojaTrazabilidad.Visible = False
    hojaOrdenes.Visible = definiciones.xlVeryHidden
    hojaFormulario.Visible = False
    hojaEjecutar.Visible = False
    hojaMonitor.Visible = True
    hojaDescargas.Visible = False
    hojaEstadisticas.Visible = False
    hojaEspecies.Visible = definiciones.xlVeryHidden
    hojaTitulos.Visible = True
    hojaMonitor.OLEObjects("lstMacroActivosTrader").Object.Value = "Todos"
    hojaMonitor.OLEObjects("txtFechaDesdeMonitor").Object.Value = datetime.now().strftime("%d/%m/%Y")
    funciones.actualizarVistaTraders()
    hojaMonitor.Activate()
    
elif miRol == "Visualizador":
    hojaTrazabilidad.Visible = False
    hojaMonitor.Visible = definiciones.xlVeryHidden
    hojaFormulario.Visible = False
    hojaEjecutar.Visible = False
    hojaOrdenes.Visible = True
    hojaDescargas.Visible = False
    hojaEstadisticas.Visible = False
    hojaEspecies.Visible = definiciones.xlVeryHidden       
    hojaOrdenes.OLEObjects("lstMacroActivosPM").Object.Value = "Renta Variable"
    hojaOrdenes.OLEObjects("listaGerentes").Object.Value = "(Todos)"
    hojaOrdenes.OLEObjects('fechaDesdeIntenciones').Object.Value = datetime.now().strftime("%d/%m/%Y")
    funciones.actualizarVistaGerentes()
    hojaOrdenes.OLEObjects("btnCrearIntenciones").Object.Visible = False
    hojaOrdenes.OLEObjects("btnCancelarIntencion").Object.Visible = False
    hojaOrdenes.OLEObjects("btnRenovarIntencion").Object.Visible = False
    hojaOrdenes.OLEObjects("btnEditarIntenciones").Object.Visible = False
    hojaOrdenes.Activate()
    
else: #Administrador
    hojaOrdenes.Visible = True
    hojaMonitor.Visible = True
    hojaTitulos.Visible = True
    hojaTrazabilidad.Visible = False
    hojaFormulario.Visible = False
    hojaEjecutar.Visible = False
    hojaDescargas.Visible = False
    hojaEstadisticas.Visible = False
    hojaEspecies.Visible = definiciones.xlVeryHidden   
    hojaOrdenes.Range("E4:E5").NumberFormat = "mm/dd/yyyy"
    hojaMonitor.OLEObjects("lstMacroActivosTrader").Object.Value = "Todos"
    hojaMonitor.OLEObjects("txtFechaDesdeMonitor").Object.Value = datetime.now().strftime("%d/%m/%Y")
    funciones.actualizarVistaTraders()
    hojaOrdenes.OLEObjects("lstMacroActivosPM").Object.Value = "Renta Variable"
    hojaOrdenes.OLEObjects("listaGerentes").Object.Value = "(Todos)"
    hojaOrdenes.OLEObjects('fechaDesdeIntenciones').Object.Value = datetime.now().strftime("%d/%m/%Y")
    funciones.actualizarVistaGerentes()
    hojaOrdenes.Activate()
    
    
#Hasta aqui ha terminado todo el proceso de incio, ahora se muestra el excel para 
#ser utilizado por el usuario
excel.WindowState = definiciones.xlMaximized
excel.Visible = True
tiempoInicial = datetime.now()
tiempoInicialEjecucion = datetime.now()

#Este ciclo sin fin mantiene activa la ejecución a menos que pase alguno de los siguientes casos:
#1.El libro de Excel sea cerrado
#2. Se termine el tiempo asignado de sesión activa

while definiciones.keepOpen:
    
    tiempoFinal = datetime.now()
    tiempoTranscurrido = tiempoFinal - tiempoInicial
    
    #Si se cumple la condición se actualiza de manera automática los datos que ven los traders y gerentes
    if tiempoTranscurrido.total_seconds()/60 >= definiciones.parametros["Valor"]["tiempoActualizacion"]:
        try:
            if miRol == "PM":
                funciones.actualizarVistaGerentes()
            elif miRol == "Trader":
                funciones.actualizarVistaTraders()
            else:
                funciones.actualizarVistaTraders()
                funciones.actualizarVistaGerentes()
        except:
            None
        tiempoInicial = datetime.now()
    
    #La herramienta se cierra luego de estar abierta mas de x horas: normalmente 12 horas
    tiempoOperacionHerramienta = datetime.now() - tiempoInicialEjecucion
    
    if tiempoOperacionHerramienta.total_seconds()/3600 >= definiciones.parametros["Valor"]["tiempoApagar"]:
        intencionesAM.Close(SaveChanges=False)
        definiciones.keepOpen = False
        
    #Si el archivo de excel donde se montan las intencciones se cierra entonces se termina la ejecución del Script   
    try:
        librosAbiertos = definiciones.excel.Workbooks.Count
    except:
        librosAbiertos = 2
    if librosAbiertos == 0:
        definiciones.keepOpen = False    
        
        
    time.sleep(0.1)   
    pythoncom.PumpWaitingMessages()

definiciones.excel.Quit()
definiciones.excel = None     
print("Ejecución terminada satisfactoriamente")    

    
