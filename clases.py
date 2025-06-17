# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 16:03:35 2022

@author: FRCASTRO
******************************************************************************
                        INTENCIONES ASSET MANAGEMENT
******************************************************************************

"""
import win32com.client as win32
import definiciones
import funciones
import pandas as pd
import getpass
from datetime import date,datetime
from pathlib import Path


class BtnSalirEvents: 

    def OnClick(self,*args):
    
       excel = definiciones.excel
       intencionesAM= win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
       intencionesAM.Saved = True
       intencionesAM.Close(True)
       
       
class BtnCancelarCreacionIntencionesEvents:
    
    def OnClick(self,*args):
        excel = win32.DispatchEx("Excel.Application") 
        intencionesAM= win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        primeraFilaTabla = 7
        hojaFormulario = intencionesAM.Worksheets("Formulario Ordenes")
        hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
        funciones.limpiarHojaIntencionesPM(hojaFormulario, primeraFilaTabla)
        hojaFormulario.Visible = False
        hojaOrdenes.Activate()
        
class BtnActulizarVistaIntencionesEvents:    
  
    def OnClick(self,*args):
        funciones.actualizarVistaGerentes()
        
        
        
class BtnActualizarIntencionesMonitorEvents:
    
    def OnClick(self, *args):
       
        funciones.actualizarVistaTraders()
 
class btnCancelarIntencionesTraderEvents:
    
    def OnClick(self,*args):
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaMonitor = intencionesAM.Worksheets("Monitor")
        hojaEjecutar = intencionesAM.Worksheets("Ejecutar Intenciones")
        hojaMonitor.Visible= True
        hojaEjecutar.Visible= False
        filaDesde = 7
        funciones.limpiarHojaIntencionesPM(hojaEjecutar, filaDesde)
        
        hojaMonitor.Activate()
        
class btnGuardarIntencionesTraderEvents:
    
    def OnClick(self,*args):       
         
        #Obtener la intenciones ingresadas
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaEjecucion = intencionesAM.Worksheets("Ejecutar Intenciones")
        hojaEjecucion.Unprotect()
        hojaEjecucion.Range("B:Z").NumberFormat = "@" 
        ultFilaIntenciones = hojaEjecucion.Cells(hojaEjecucion.Rows.Count, "B").End(definiciones.xlUp).Row
        numeroIntenciones = ultFilaIntenciones -7
        
        nombreDatos = definiciones.parametros["Valor"]["ColumnasEjecutarIntencionesArchivo"].split("-")  
        totalColumnas = len(nombreDatos)
        
        rangoDeDatos = hojaEjecucion.Range(hojaEjecucion.Cells(8,2),hojaEjecucion.Cells(ultFilaIntenciones,totalColumnas+1))
            
        data = rangoDeDatos.Value
        intencionesNuevas = pd.DataFrame(data,columns = nombreDatos)
        
        #Se validan los datos
        #1. La cantidad ejecutada no puede ser superiror a la cantidad total
        validacionCantidadEjecutada = intencionesNuevas.loc[:,["CantEjecutada"]].apply(lambda x: funciones.esUnNumero(x),axis=1)
        if validacionCantidadEjecutada.all():
            validacionCantidadEjecutada = intencionesNuevas[["CantidadTotal","CantEjecutada"]].apply(lambda x: True if float(x["CantidadTotal"])>= float(x["CantEjecutada"]) else False,axis=1)
            if ~validacionCantidadEjecutada.all():
                #mostrar errores
                rangoDeDatos.ClearComments()
                rangoDeDatos.Interior.Color = 0xD9D9D9    
                for fila in range(0,len(validacionCantidadEjecutada)):                
                    if validacionCantidadEjecutada.loc[fila,]==False:
                        #pintar y comentar campos con errores de datos
                        hojaEjecucion.Cells(fila + 8,nombreDatos.index("CantEjecutada")+2).AddComment("La cantidad ejecutada no puede ser superior a la cantidad total.")        
                        hojaEjecucion.Cells(fila + 8,nombreDatos.index("CantEjecutada")+2).Interior.Color = 0x3755ED
                hojaEjecucion.Protect()
                return
                
        else:
            #mostrar erroresy
            rangoDeDatos.ClearComments()
            rangoDeDatos.Interior.Color = 0xD9D9D9    
            for fila in range(0,len(validacionCantidadEjecutada)):                
                if validacionCantidadEjecutada.loc[fila,]==False:
                    #pintar y comentar campos con errores de datos
                    hojaEjecucion.Cells(fila + 8,nombreDatos.index("CantEjecutada")+2).AddComment("Esto no es un número.")        
                    hojaEjecucion.Cells(fila + 8,nombreDatos.index("CantEjecutada")+2).Interior.Color = 0x3755ED
            hojaEjecucion.Protect()
            return
        
        
        #Una vez validados los datos se suben las intenciones
        #1. Se calcula el porcentaje de la cantidad total ya ejecutada
        #2. Se asigna el Trader a la intención
        #3. Se asigna la ultima fecha de moficación
        
        rutaIntenciones = definiciones.parametros["Valor"]["rutaIntenciones"]
        rutaLogIntenciones = definiciones.parametros["Valor"]["rutaIntencionesTrazabilidad"]
        usuario = getpass.getuser().upper()
        funciones.guardarIntencionesTraders(intencionesNuevas,rutaIntenciones,rutaLogIntenciones,usuario)
        
        hojaMonitor = intencionesAM.Worksheets("Monitor")
        hojaMonitor.Visible = True        
        hojaEjecucion.Visible= False
        filaDesde = 7
        funciones.limpiarHojaIntencionesPM(hojaEjecucion, filaDesde)
        
        hojaMonitor.Activate()
        funciones.actualizarVistaTraders()
    
class btnEjecutarIntencionesEvents:
    
    def OnClick(self,*args):
        #Se van a cargar en la hoja de ejecutar las intenciones que se seleccionaron en la hoja Monitor
        #Solo se pueden traer intenciones no sean: Cancelada, vencida o Ejecutada/Total
        #Se deben tomar los ids seleccionados
        #Luego se toma a que gerentes pertenecen las intenciones
        #luego se cargan todas las intenciones de los gerentes pedidos
        #Después se filtra por los Ids seleccionados
        #Finalmente se colocan en la hoja 
        #Y se les da el formato correspondiente
        miUsuario =  getpass.getuser().upper()
        miRol = definiciones.usuarios["Rol"][miUsuario]
        if miRol == "PM":
            funciones.mostrarMensajeAdvertencia("Los gerentes no pueden ejecutar intenciones")
            return
        excel = definiciones.excel
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaMonitor = intencionesAM.Worksheets("Monitor")
        ultimaFilaMonitor = hojaMonitor.Cells(hojaMonitor.Rows.Count, "B").End(definiciones.xlUp).Row
       
        filasParaEjecutar = excel.Selection.Rows.Count
        primeraFilaSel = excel.Selection.Row
        ultimaFilaSel = primeraFilaSel + filasParaEjecutar -1
        
        idsParaEjecutar = hojaMonitor.Range(hojaMonitor.Cells(primeraFilaSel,2),hojaMonitor.Cells(ultimaFilaSel,2)).Value
        
        columnaGerentes = 22
        listaGerentes = hojaMonitor.Range(hojaMonitor.Cells(primeraFilaSel,columnaGerentes),hojaMonitor.Cells(ultimaFilaSel,columnaGerentes)).Value
        
        if primeraFilaSel <8 or ultimaFilaSel >ultimaFilaMonitor:
            funciones.mostrarMensajeAdvertencia("El rango seleccionado no es correcto.")
            return
        if funciones.esUnaLista(idsParaEjecutar):
            archivo = list(dict.fromkeys(list(map(lambda x: x[0] +".csv",listaGerentes))))
            idsParaEjecutar = list(map(lambda x: int(x[0]),idsParaEjecutar))
        else:
            archivo = [listaGerentes +".csv"]
            idsParaEjecutar = [idsParaEjecutar]
        
        intenciones = funciones.obtenerIntenciones(definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
        intenciones["Id"] = intenciones["Id"].astype('int64')
        intenciones = intenciones[intenciones["Id"].isin(idsParaEjecutar)]
        intenciones = intenciones[~intenciones["Estado"].isin(["Cancelada","Vencida"])]
        if len(intenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para ejecutar.")
            return
        
        hojaEjecutar = intencionesAM.Worksheets("Ejecutar Intenciones")   
        primeraFilaTabla = 7
        funciones.limpiarHojaIntencionesPM(hojaEjecutar, primeraFilaTabla)
        nombreDatos = definiciones.parametros["Valor"]["ColumnasEjecutarIntencionesArchivo"].split("-")          
        encabezado = definiciones.parametros["Valor"]["ColumnasEjecutarIntenciones"].split("-")
        columnaCantidad = encabezado.index("Cantidad") +2
        columnaCantidadEjecutada = encabezado.index("Cantidad Ejecutada") +2

                
        celdasEncabezado = hojaEjecutar.Range(hojaEjecutar.Cells(7,2),hojaEjecutar.Cells(7,1 + len(encabezado)))
        celdasEncabezado.Value = encabezado
        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        hojaEjecutar.Visible = True    
        hojaEjecutar.Range("B:Z").NumberFormat = "@" 
        hojaEjecutar.Range("B:B").NumberFormat = "0"
        hojaEjecutar.Columns(columnaCantidad).NumberFormat = "#,##0.00" #$10.00
        hojaEjecutar.Columns(columnaCantidadEjecutada).NumberFormat = "#,##0.00" #$10.00
        #Dar a las columnas de fecha el formato de texto en excel
        hojaEjecutar.Activate()
        
        #colocar los datos en la hoja del formulario incluyendo el Id, de acá en adelante ya está lista la programación, falta ajustra el guardar los datos cuando sean una edición
        intenciones = intenciones.where(~intenciones.isna(), other="")
        
        
        hojaEjecutar.Range(hojaEjecutar.Cells(8,2),hojaEjecutar.Cells(7 + len(intenciones),len(encabezado)+1)).Value = intenciones[nombreDatos].to_records(index=False)
       
        

        #Bloquaer hoja para que no se agreguen mas intenciones
        #Colocar los Ids y no permitir modificar esta columna

        hojaEjecutar.Range(hojaEjecutar.Cells(8,encabezado.index("Cantidad")+2),hojaEjecutar.Cells(7 + len(intenciones),encabezado.index("Comentarios Trader")+2)).Locked = False
        hojaEjecutar.Columns.AutoFit()
        hojaEjecutar.Protect()
        hojaMonitor.Visible= False
        
       
class BtncrearIntencionesEvents:
    
    def OnClick(self, *args):
        
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])        
        macroActivo = intencionesAM.Worksheets("Mis Órdenes").OLEObjects("lstMacroActivosPM").Object.Value
        
        if macroActivo =="Renta Variable":            
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionRV"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad") + 2
            columnaCantidadDisponible = camposMostrar.index("Cantidad disponible") + 2
        elif macroActivo == "Deuda Privada":            
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionDPr"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad(Millones)") + 2
            columnaCantidadDisponible = camposMostrar.index("Cantidad disponible(Millones)") +2
        elif macroActivo == "Deuda Pública":
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionDPu"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad(Millones)") +2
            columnaCantidadDisponible = camposMostrar.index("Cantidad disponible(Millones)") +2
        elif macroActivo == "Fondos":
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionFondos"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad") +2
            columnaCantidadDisponible = camposMostrar.index("Cantidad disponible")+2
        elif macroActivo == "Forex":
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionForex"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad")+2
            columnaCantidadDisponible = 102
        elif macroActivo == "Liquidez":
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionLiquidez"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad") +2
            columnaCantidadDisponible = camposMostrar.index("Cantidad disponible")+2
        elif macroActivo == "Swaps":
            camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCreacionSwaps"].split("-")
            columnaCantidad = camposMostrar.index("Cantidad")+2
            columnaCantidadDisponible = 102
        else:
            funciones.mostrarMensajeAdvertencia("Primero Seleccione un mercado.")
            return
        
       
        if macroActivo in ["Renta Variable","Deuda Privada","Deuda Pública"]:  
            #Se traen las operaciones intradia
            intencionesVigentes = funciones.cargarOperacionesVigentes() 
            #Se eliminan la últmia revisión de operacioens intradia
            definiciones.nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria[definiciones.nemoTitulosFiduciaria["Origen Informacion"]!="INTENCIONES VIGENTES"]        
            definiciones.nemoTitulosValores = definiciones.nemoTitulosValores[definiciones.nemoTitulosValores["Origen Informacion"]!="INTENCIONES VIGENTES"]        
            if len(intencionesVigentes) > 0:            
                #FIDUCIARIA
                #Se actualzian las operaciones Intradia
                definiciones.nemoTitulosFiduciaria = funciones.agregarintencionesVigentesFidu(definiciones.nemoTitulosFiduciaria,intencionesVigentes)
                #definiciones.nemoTitulosFiduciaria.to_excel("C:/Users/frcastro/downloads/inventario_titulos_fidu.xlsx")
                #VALORES
                #Se actualzian las operaciones Intradia
                definiciones.nemoTitulosValores = funciones.agregarintencionesVigentesValores(definiciones.nemoTitulosValores,intencionesVigentes)
                #definiciones.nemoTitulosValores.to_excel("C:/Users/frcastro/downloads/inventario_titulos_valores.xlsx")

            
        formularioOrdenes = intencionesAM.Worksheets("Formulario Ordenes")  
        primeraFilaTabla = 7
        funciones.limpiarHojaIntencionesPM(formularioOrdenes, primeraFilaTabla) 
        formularioOrdenes.Range("A1").Value = "CREACION"
        formularioOrdenes.Range("B1").Value = getpass.getuser().upper()
        
        celdasEncabezado = formularioOrdenes.Range(formularioOrdenes.Cells(7,2),formularioOrdenes.Cells(7,1 + len(camposMostrar)))
        celdasEncabezado.Value = camposMostrar
        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        formularioOrdenes.Visible = True    
        formularioOrdenes.Range("A:Z").NumberFormat = "@" #Formato Texto
        formularioOrdenes.Columns(columnaCantidadDisponible).NumberFormat = "#,##0.00" #$10.00
        formularioOrdenes.Columns(columnaCantidad).NumberFormat = "#,##0.00" #$10.00
        #Dar a las columnas de fecha el formato de texto en excel
        formularioOrdenes.Cells(8,camposMostrar.index("Vigente desde")+2).Value = datetime.now().strftime("%d/%m/%Y")
        formularioOrdenes.Activate()
            
class BtnSubirIntencionesEvents:
    
    def OnClick(self, *args):
       
        #Obtener la intenciones ingresadas
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
        hojaFormulario = intencionesAM.Worksheets("Formulario Ordenes")
        hojaFormulario.Unprotect()# puede venir portegida desde la edición o la renovación
        hojaFormulario.Range("A:Z").NumberFormat = "@"
        creacionEdicion = hojaFormulario.Range("A1").Value
        macroActivo = hojaOrdenes.OLEObjects("lstMacroActivosPM").Object.Value
        ultFilaIntenciones = hojaFormulario.Cells(hojaFormulario.Rows.Count, "B").End(definiciones.xlUp).Row
        numeroIntenciones = ultFilaIntenciones -7
        archivo = hojaFormulario.Range("B1").Value 
        
        if numeroIntenciones == 0:
            funciones.mostrarMensajeAdvertencia("No hay intenciones para subir")
            return
        
        if creacionEdicion == "CREACION":
            if macroActivo == "Renta Variable":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionRV"].split("-")
            if macroActivo == "Deuda Privada":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionDPr"].split("-")
            if macroActivo == "Deuda Pública":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionDPu"].split("-")
            if macroActivo == "Fondos":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionFondos"].split("-")
            if macroActivo == "Forex":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionForex"].split("-")
            if macroActivo == "Liquidez":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionLiquidez"].split("-")
            if macroActivo == "Swaps":
                camposFormularioCreacion = definiciones.parametros["Valor"]["ColumnasFormularioCreacionSwaps"].split("-")
           
            resultado, intenciones = funciones.subirCreacion(hojaFormulario,ultFilaIntenciones,camposFormularioCreacion,macroActivo)
        elif creacionEdicion == "EDICION":
            if macroActivo == "Renta Variable":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioRV"].split("-")
            if macroActivo == "Deuda Privada":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioDPr"].split("-")
            if macroActivo == "Deuda Pública":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioDPu"].split("-")
            if macroActivo == "Fondos":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioFondos"].split("-")
            if macroActivo == "Forex":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioForex"].split("-")
            if macroActivo == "Liquidez":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioLiquidez"].split("-")
            if macroActivo == "Swaps":
                camposFormularioEdicion = definiciones.parametros["Valor"]["ColumnasFormularioSwaps"].split("-")
               
           
            resultado, intenciones = funciones.subirEdicion(hojaFormulario,ultFilaIntenciones,camposFormularioEdicion,macroActivo,archivo)
        elif creacionEdicion == "RENOVACION":
            
            camposFormularioRenovacion = definiciones.parametros["Valor"]["ColumnasFormularioRenovacion"].split("-")            
            resultado, intenciones = funciones.subirRenovacion(hojaFormulario,ultFilaIntenciones,camposFormularioRenovacion,macroActivo,archivo)  
            
        elif creacionEdicion == "CANCELACION":
            resultado, intenciones = funciones.subirCancelacion(hojaFormulario,ultFilaIntenciones,macroActivo,archivo)
            
        if resultado == False:
            funciones.mostrarMensajeAdvertencia("No fué posible guardar las intenciones")
        else:
            #Faltar guardar la trazabilidad de la intencion
            resultado = funciones.guardarIntenciones(intenciones,definiciones.parametros["Valor"]["rutaIntencionesTrazabilidad"],archivo)
            if resultado == False: 
                funciones.mostrarMensajeAdvertencia("No se pudo guardar trazabilidad de las intenciones")
                return
            else:
                #Vamos  a guardar log de las acciones registradas por el usuario
                try:
                    intenciones.apply(lambda x: funciones.actualizarLogIntenciones(str(int(x["Id"])),getpass.getuser().upper(),creacionEdicion.title() + " de intenciones."),axis=1)
                except:
                    funciones.mostrarMensajeAdvertencia("Hubo un error guardando datos en el log.")
                #Ahora se limpia la hoja de intenciones
                primeraFilaTabla = 1
                funciones.limpiarHojaIntencionesPM(hojaFormulario, primeraFilaTabla)     
                funciones.mostrarMensajeAdvertencia("Las intenciones se guardaron exitosamente")      
                hojaFormulario.Visible = False
                hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
                hojaOrdenes.Activate()
                funciones.actualizarVistaGerentes()
        #Tener cuidado con duplicar intenciones     

class BtnEditarIntencionesEvents:
    
    def OnClick(self, *args):
        
        #Se tienen las siguientes reglas para editar Intenciones
        #1. Se deben seleccionar las intenciones que se desean editar
        #2. Solo se pueden editar intenciones propias
        #3. solo se pueden editar intencionees no vencidas
        #4 Tienen que ser de un mismo mercado
        #Pasos
        #1. Se obtienen los Id de las intenciones seleccionadas
        #2. Se carga el formulario del mercado correspondiente
        #3. Se colocan las intenciones a editar y se coloca el ID al inicio
        #4. Luego de dar guardar, se validan los datos ingresados en el formulario
        #5. Se modifican los datos que no son ingresados por el usuario 
        #6. Las intenciones del archivo csv de ese PM se cargan a un DF y se modifica el registro completo, usando el Id como llave
        #7. Los datos se pegan completos al csv reemplazando el archivo original
        #8. Se guarda la trazabilidad de las intenciones.
        excel = definiciones.excel
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
        ultimaFilaOrdenes = hojaOrdenes.Cells(hojaOrdenes.Rows.Count, "B").End(definiciones.xlUp).Row
        macroActivo = hojaOrdenes.OLEObjects("lstMacroActivosPM").Object.Value  
       
        filasParaEditar = excel.Selection.Rows.Count
        primeraFilaSel = excel.Selection.Row
        ultimaFilaSel = primeraFilaSel + filasParaEditar -1
        
        idsParaEditar = hojaOrdenes.Range(hojaOrdenes.Cells(primeraFilaSel,2),hojaOrdenes.Cells(ultimaFilaSel,2)).Value
        if primeraFilaSel <8 or ultimaFilaSel >ultimaFilaOrdenes:
            funciones.mostrarMensajeAdvertencia("El rango seleccionado no es correcto.")
            return
    
        
        
        #Traer los datos de la intención desde el archivo de intenciones, si no se selecciona ningún nombre o Todos los gerentes, 
        #entonces se trae solo los datos del gerente que está operando la herramienta
        miUsuario =  getpass.getuser().upper()
        gerenteAConsultar = hojaOrdenes.OLEObjects('listaGerentes').Object.Value
        if gerenteAConsultar in ["","(Todos)"]:
            archivo = [miUsuario+".csv"]
        else:
            archivo = [gerenteAConsultar + ".csv"]
    
        intenciones = funciones.obtenerIntenciones(definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
        intenciones["Id"] = intenciones["Id"].astype('int64')
        #Solo los administradores podrán editar intenciones que ya están EN proceso o Ejecutadas parcialmente
        
        intenciones = intenciones[intenciones["Estado"].isin(["Nueva","Modificada","Renovada","En proceso","Ejecutada/Parcial","Ejecutada/Total"])] 
    
        if len(intenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para editar.")
            return
        
        intenciones = intenciones[intenciones["TipoActivo"]== macroActivo]
        if len(intenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para editar.")
            return
        
        if funciones.esUnaLista(idsParaEditar):
            idsParaEditar =  list(map(lambda x: int(x[0]),idsParaEditar))
        else:
            idsParaEditar = [idsParaEditar]
                
        #Traer las operaciones seleccionadas
        intencionesParaEditar = intenciones.loc[intenciones["Id"].isin(idsParaEditar)]
        if len(intencionesParaEditar) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para editar.")
            return
        
        #elegir las columnas para editar
        if macroActivo =="Renta Variable":            
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionRV"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos]
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioRV"].split("-")
            columnaCantidadDisponible = encabezado.index("Cantidad") +2
            columnaCantidad = encabezado.index("Cantidad disponible") +2
        elif macroActivo == "Deuda Privada":
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionDPr"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos] 
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioDPr"].split("-")
            columnaCantidadDisponible = encabezado.index("Cantidad(Millones)") +2
            columnaCantidad = encabezado.index("Cantidad disponible(Millones)") +2
        elif macroActivo == "Deuda Pública":
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionDPu"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos] 
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioDPu"].split("-")
            columnaCantidadDisponible = encabezado.index("Cantidad(Millones)") +2
            columnaCantidad = encabezado.index("Cantidad disponible(Millones)") +2
        elif macroActivo == "Fondos":
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionFondos"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos]  
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioFondos"].split("-")
            columnaCantidadDisponible = encabezado.index("Cantidad") +2
            columnaCantidad = encabezado.index("Cantidad disponible") +2
        elif macroActivo == "Forex":
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionForex"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos]  
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioForex"].split("-")
            columnaCantidadDisponible = 102
            columnaCantidad = encabezado.index("Cantidad") +2
        elif macroActivo == "Liquidez":
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionLiquidez"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos]  
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioLiquidez"].split("-")
            columnaCantidadDisponible = encabezado.index("Cantidad") +2
            columnaCantidad = encabezado.index("Cantidad disponible") +2
        elif macroActivo == "Swaps":
            nombreDatos = definiciones.parametros["Valor"]["ColumnasFormularioEdicionSwaps"].split("-")
            intencionesParaEditar = intencionesParaEditar.loc[:,nombreDatos]  
            encabezado = definiciones.parametros["Valor"]["ColumnasFormularioSwaps"].split("-")
            columnaCantidadDisponible = encabezado.index("Cantidad") +2
            columnaCantidad = 102
        else:
            funciones.mostrarMensajeAdvertencia("Primero seleccione un mercado.")
            return
        
        if macroActivo in ["Renta Variable","Deuda Privada","Deuda Pública"]: 
            #Se traen las intenciones que aún están vigentes 
            intencionesVigentes = funciones.cargarOperacionesVigentes()
            #Se eliminan la últmia revisión de intenciones vigentes
            definiciones.nemoTitulosFiduciaria = definiciones.nemoTitulosFiduciaria[definiciones.nemoTitulosFiduciaria["Origen Informacion"]!="INTENCIONES VIGENTES"]        
            definiciones.nemoTitulosValores = definiciones.nemoTitulosValores[definiciones.nemoTitulosValores["Origen Informacion"]!="INTENCIONES VIGENTES"]        
            if len(intencionesVigentes) > 0:
                
                #FIDUCIARIA
                #Se actualzian las operaciones Intradia
                definiciones.nemoTitulosFiduciaria = funciones.agregarintencionesVigentesFidu(definiciones.nemoTitulosFiduciaria,intencionesVigentes)
                
                #VALORES
                #Se actualzian las operaciones Intradia
                definiciones.nemoTitulosValores = funciones.agregarintencionesVigentesValores(definiciones.nemoTitulosValores,intencionesVigentes)

        
        formularioOrdenes = intencionesAM.Worksheets("Formulario Ordenes")     
        primeraFilaTabla = 7
        funciones.limpiarHojaIntencionesPM(formularioOrdenes, primeraFilaTabla)
        formularioOrdenes.Range("A1").Value = "EDICION"
        formularioOrdenes.Range("B1").Value = archivo[0].split(".")[0]
        
        celdasEncabezado = formularioOrdenes.Range(formularioOrdenes.Cells(7,2),formularioOrdenes.Cells(7,1 + len(encabezado)))
        celdasEncabezado.Value = encabezado
        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        formularioOrdenes.Visible = True    
        formularioOrdenes.Range("C:Z").NumberFormat = "@" 
        formularioOrdenes.Range("B:B").NumberFormat = "0"
        formularioOrdenes.Columns(columnaCantidadDisponible).NumberFormat = "#,##0.00" #$10.00
        formularioOrdenes.Columns(columnaCantidad).NumberFormat = "#,##0.00" #$10.00
        #Dar a las columnas de fecha el formato de texto en excel
        formularioOrdenes.Activate()
        
        #colocar los datos en la hoja del formulario incluyendo el Id, de acá en adelante ya está lista la programación, falta ajustra el guardar los datos cuando sean una edición
        intencionesParaEditar = intencionesParaEditar.where(~intencionesParaEditar.isna(), other="")
        
        if "Cantidad disponible(Millones)" in encabezado:
            celdaSinDatos = encabezado.index("Cantidad disponible(Millones)")
            formularioOrdenes.Range(formularioOrdenes.Cells(8,2),formularioOrdenes.Cells(7 + len(intencionesParaEditar),celdaSinDatos +1)).Value = intencionesParaEditar[nombreDatos[0:celdaSinDatos+1]].to_records(index=False)
            formularioOrdenes.Range(formularioOrdenes.Cells(8,celdaSinDatos + 3),formularioOrdenes.Cells(7 + len(intencionesParaEditar),len(encabezado)+1)).Value = intencionesParaEditar[nombreDatos[celdaSinDatos:]].to_records(index= False)
        elif "Cantidad disponible" in encabezado:
            celdaSinDatos = encabezado.index("Cantidad disponible")
            formularioOrdenes.Range(formularioOrdenes.Cells(8,2),formularioOrdenes.Cells(7 + len(intencionesParaEditar),celdaSinDatos +1)).Value = intencionesParaEditar[nombreDatos[0:celdaSinDatos+1]].to_records(index=False)
            formularioOrdenes.Range(formularioOrdenes.Cells(8,celdaSinDatos + 3),formularioOrdenes.Cells(7 + len(intencionesParaEditar),len(encabezado)+1)).Value = intencionesParaEditar[nombreDatos[celdaSinDatos:]].to_records(index= False)            
        else:
           formularioOrdenes.Range(formularioOrdenes.Cells(8,2),formularioOrdenes.Cells(7 + len(intencionesParaEditar),len(encabezado)+1)).Value = intencionesParaEditar[nombreDatos].to_records(index=False)
           
        
        intenciones.where(~intenciones.isna(), other="")

        #Bloquaer hoja para que no se agreguen mas intenciones
        #Colocar los Ids y no permitir modificar esta columna
        formularioOrdenes.Range(formularioOrdenes.Cells(8,3),formularioOrdenes.Cells(7 + len(intencionesParaEditar),13)).Locked = False
        formularioOrdenes.Columns.AutoFit()
        
class BtnDescargarIntencionesEvents:
    
    def OnClick(self, *args):
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaOrigen = intencionesAM.ActiveSheet
        hojaDescargas = intencionesAM.Worksheets("Descargar Intenciones")
        hojaDescargas.Range("A1").Value  = hojaOrigen.Name
        hojaDescargas.OLEObjects('txtFechaDesdeDescargas').Object.Value = datetime(datetime.now().year,datetime.now().month,1).strftime("%d/%m/%Y")
        hojaDescargas.OLEObjects("txtFechaHastaDescargas").Object.Value = datetime.now().strftime("%d/%m/%Y")
        
        hojaDescargas.OLEObjects("lstMacroActivoDescargas").Object.Value = "(Todos)"
        hojaDescargas.OLEObjects("lstGerentesDescarga").Object.Value = getpass.getuser().upper()
        hojaDescargas.Visible = True
        hojaDescargas.Activate()
        hojaOrigen.Visible = False

class BtnConsultarIntencionesEvents:

    def OnClick(self,*args):        
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaDescargas = intencionesAM.Worksheets("Descargar Intenciones")
        hojaEstadisticas = intencionesAM.Worksheets("Estadisticas")
        filaDesde = 7
        hojaDescargas.UsedRange.ClearContents()
        intenciones = funciones.obtenerIntencionesParaDescargar(hojaDescargas)
        if len(intenciones) == 0:
            funciones.mostrarMensajeAdvertencia("La consulta no retornó ningún resultado")
            return
        funciones.graficarEstadisticasIntenciones(intenciones,hojaEstadisticas,hojaDescargas,intencionesAM)
        intenciones["FechaIngreso"] = intenciones["FechaIngreso"].astype("str")
        celdasEncabezado = hojaDescargas.Range(hojaDescargas.Cells(7,2),hojaDescargas.Cells(7,len(intenciones.columns)+1))
        celdasEncabezado.Value = intenciones.columns.tolist()
        hojaDescargas.Range(hojaDescargas.Cells(8,2),hojaDescargas.Cells(len(intenciones)+7,1 + intenciones.shape[1])).Value = intenciones.to_records(index=False) 

        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        hojaDescargas.Range("B:B").NumberFormat = "0"
        hojaDescargas.Range("C:AZ").NumberFormat = "@"
        hojaDescargas.Columns.AutoFit() 
        
        
        
class BtnGuardarIntencionesExcelEvents:
    
    def OnClick(self,*args):
        
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaDescargas = intencionesAM.Worksheets("Descargar Intenciones")
        intenciones = funciones.obtenerIntencionesParaDescargar(hojaDescargas)
        rutaDescargas = str(Path.home()/"Downloads") + "\Intenciones " + datetime.now().strftime("%d-%m-%Y") +".xlsx"
        if len(intenciones) == 0:
            funciones.mostrarMensajeAdvertencia("No hay datos para guardar")
            return
        try:
            intenciones.to_excel(rutaDescargas,index=False)
            funciones.mostrarMensajeAdvertencia("Se guardaron las intenciones en la siguiente ubicación: " + rutaDescargas)
        except:
            funciones.mostrarMensajeAdvertencia("No fué posible descargar las intenciones")
            
class BtnSalirDescargasEvents:       
    def OnClick(self,*args):
           
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaDescargas = intencionesAM.Worksheets("Descargar Intenciones")
        if hojaDescargas.Range("A1").Value == "Monitor":
            hojaDestino = intencionesAM.Worksheets("Monitor")
        else:
            hojaDestino = intencionesAM.Worksheets("Mis Órdenes")       
        
        hojaDescargas.UsedRange.ClearContents()
        hojaDestino.Visible = True
        hojaDescargas.Visible= False        
        hojaDestino.Activate()

class BtnSalirEstadisticasEvents:
    def OnClick(self, *args):
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaDescargas = intencionesAM.Worksheets("Descargar Intenciones")
        hojaEstadisticas = intencionesAM.Worksheets("Estadisticas")        
        hojaDescargas.Visible = True
        hojaEstadisticas.Visible= False        
        hojaDescargas.Activate()

class BtnVerEstadisticasEvents:  

    def OnClick(self, *args):  
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaDescargas = intencionesAM.Worksheets("Descargar Intenciones")
        hojaEstadisticas = intencionesAM.Worksheets("Estadisticas")      
        hojaEstadisticas.Visible= True  
        hojaDescargas.Visible = False              
        hojaEstadisticas.Activate()

class BtnHistoriaIntencionEvents:

    def OnClick(self,*args):
        
        #Solo se puede seleccionar una intención para mostrar la trazabilidad
        excel = definiciones.excel
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaOrigen = intencionesAM.ActiveSheet
        hojaTrazabilidad = intencionesAM.Worksheets("Trazabilidad")
        hojaTrazabilidad.Range("A1").Value = hojaOrigen.Name
        ultimaFilaOrdenes = hojaOrigen.Cells(hojaOrigen.Rows.Count, "B").End(definiciones.xlUp).Row
          
        #filaSel = excel.Selection.Row
        
        #idSel = hojaOrigen.Range(hojaOrigen.Cells(filaSel,2),hojaOrigen.Cells(filaSel,2)).Value
        
        filasSeleccionadas = excel.Selection.Rows.Count
        primeraFilaSel = excel.Selection.Row
        ultimaFilaSel = primeraFilaSel + filasSeleccionadas -1
        
        idSel = hojaOrigen.Range(hojaOrigen.Cells(primeraFilaSel,2),hojaOrigen.Cells(ultimaFilaSel,2)).Value
        
        if primeraFilaSel <8 or primeraFilaSel >ultimaFilaOrdenes:
            funciones.mostrarMensajeAdvertencia("El rango seleccionado no es correcto.")
            return
        
        if hojaOrigen.Name == "Mis Órdenes":
            colNombreArchivo = definiciones.parametros["Valor"]["nombreCamposVerGerentes"].split("-").index("IngresadoPor")+2
        else: #Monitor
            colNombreArchivo = 22 #Columna del gerente de portafolio
        
        listaGerentes = hojaOrigen.Range(hojaOrigen.Cells(primeraFilaSel,colNombreArchivo),hojaOrigen.Cells(ultimaFilaSel,colNombreArchivo)).Value
        
        if funciones.esUnaLista(idSel):
            archivo = list(dict.fromkeys(list(map(lambda x: x[0] +".csv",listaGerentes))))
            idSel = list(map(lambda x: int(x[0]),idSel))
        else:
            archivo = [listaGerentes +".csv"]
            idSel = [idSel]  
        
        
        intenciones = funciones.obtenerIntenciones(definiciones.parametros["Valor"]["rutaIntencionesTrazabilidad"],archivo)
       
        #Traer las operaciones seleccionadas
        intenciones["Id"] = intenciones["Id"].astype('int64')
        intenciones["UltimaModificacion_datetime"] = pd.to_datetime(intenciones["UltimaModificacion"],format="%d/%m/%Y-%H:%M:%S")
        historicoIntenciones = intenciones[intenciones["Id"].isin(idSel)]
        
        if len(historicoIntenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay datos para mostrar.")
            return
        
        historicoIntenciones = historicoIntenciones.sort_values(by=['Id','UltimaModificacion_datetime'],ascending=[False,True])
        camposMostrar = definiciones.parametros["Valor"]["nombreCamposVerTrazabilidad"].split("-")
        titulosTabla = definiciones.parametros["Valor"]["tituloCamposVerTrazabilidad"].split("-")
        celdasEncabezado = hojaTrazabilidad.Range(hojaTrazabilidad.Cells(7,2),hojaTrazabilidad.Cells(7,1 + len(titulosTabla)))
        celdasEncabezado.Value = titulosTabla
        historicoIntenciones = historicoIntenciones.loc[:,camposMostrar]
        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1     
        #Dar a las columnas de fecha el formato de texto en excel
        hojaTrazabilidad.Range("C:Z").NumberFormat = "@" #Formato Texto, solo aplica para RV
        hojaTrazabilidad.Range("B:B").NumberFormat = "0"
        
        #colocar los datos en la hoja 
        historicoIntenciones = historicoIntenciones.where(~historicoIntenciones.isna(), other="")
        hojaTrazabilidad.Range(hojaTrazabilidad.Cells(8,2),hojaTrazabilidad.Cells(7 + len(historicoIntenciones),1+ len(camposMostrar))).Value = historicoIntenciones[camposMostrar].to_records(index=False)
        hojaTrazabilidad.Visible = True
        hojaTrazabilidad.Activate()
        #Bloquaer hoja       
        hojaTrazabilidad.Protect()
        
        
        
        
        
        
class BtnCancelarIntencionEvents:
    
    def OnClick(self,*args):
        
        excel = definiciones.excel
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
        ultimaFilaOrdenes = hojaOrdenes.Cells(hojaOrdenes.Rows.Count, "B").End(definiciones.xlUp).Row
       
        filasParaEditar = excel.Selection.Rows.Count
        primeraFilaSel = excel.Selection.Row
        ultimaFilaSel = primeraFilaSel + filasParaEditar -1
        
        idsParaCancelar = hojaOrdenes.Range(hojaOrdenes.Cells(primeraFilaSel,2),hojaOrdenes.Cells(ultimaFilaSel,2)).Value
        if primeraFilaSel <8 or ultimaFilaSel >ultimaFilaOrdenes:
            funciones.mostrarMensajeAdvertencia("El rango seleccionado no es correcto.")
            return
        
        #Traer los datos de la intención desde el archivo de intenciones, si no se selecciona ningún nombre o Todos los gerentes, 
        #entonces se trae solo los datos del gerente que está operando la herramienta
        miUsuario =  getpass.getuser().upper()
        gerenteAConsultar = hojaOrdenes.OLEObjects('listaGerentes').Object.Value
        if gerenteAConsultar in ["","(Todos)"]:
            archivo = [miUsuario+".csv"]
        else:
            archivo = [gerenteAConsultar + ".csv"]
        
        intenciones = funciones.obtenerIntenciones(definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
        if len(intenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones.")
            return
        #************
        #
        
        intenciones = intenciones[intenciones["Estado"].isin(["Nueva","Modificada","Renovada","Vencida","En proceso","Ejecutada/Parcial","Ejecutada/Total"])] 
        intenciones["Id"] = intenciones["Id"].astype('int64')
       
        #***********
        if len(intenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para cancelar.")
            return
        
        if funciones.esUnaLista(idsParaCancelar):
            idsParaCancelar =  list(map(lambda x: int(x[0]),idsParaCancelar))
        else:
            idsParaCancelar = [idsParaCancelar]
            
        #Traer las operaciones seleccionadas
        intencionesParaCancelar = intenciones.loc[intenciones["Id"].isin(idsParaCancelar)]
        if len(intencionesParaCancelar) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para cancelar.")
            return
        
               
        #elegir las columnas para editar
        
        formularioOrdenes = intencionesAM.Worksheets("Formulario Ordenes")        
        formularioOrdenes.Range("A1").Value = "CANCELACION"
        formularioOrdenes.Range("B1").Value = archivo[0].split(".")[0]
        camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioCancelacion"].split("-")
        celdasEncabezado = formularioOrdenes.Range(formularioOrdenes.Cells(7,2),formularioOrdenes.Cells(7,1 + len(camposMostrar)))
        celdasEncabezado.Value = camposMostrar
        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        formularioOrdenes.Visible = True    
        formularioOrdenes.Range("C:Z").NumberFormat = "@" #Formato Texto
        formularioOrdenes.Range("B:B").NumberFormat = "0"
        #Dar a las columnas de fecha el formato de texto en excel
        formularioOrdenes.Activate()
        
        #colocar los datos en la hoja del formulario incluyendo el Id, de acá en adelante ya está lista la programación, falta ajustra el guardar los datos cuando sean una edición
        camposTablaCancelar = definiciones.parametros["Valor"]["ColumnasFormularioCancelacionArchivo"].split("-")
        intencionesParaCancelar = intencionesParaCancelar.where(~intencionesParaCancelar.isna(), other="")
        formularioOrdenes.Range(formularioOrdenes.Cells(8,2),formularioOrdenes.Cells(7 + len(intencionesParaCancelar),len(camposTablaCancelar)+1)).Value = intencionesParaCancelar[camposTablaCancelar].to_records(index=False)
        
        #Bloquaer hoja para que no se agreguen mas intenciones
        #Colocar los Ids y no permitir modificar esta columna        
        formularioOrdenes.Columns.AutoFit()
        formularioOrdenes.Protect()
            

class BtnRenovarIntencionEvents:
    
    def OnClick(self,*args):
        
        excel = definiciones.excel
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaOrdenes = intencionesAM.Worksheets("Mis Órdenes")
        ultimaFilaOrdenes = hojaOrdenes.Cells(hojaOrdenes.Rows.Count, "B").End(definiciones.xlUp).Row
       
        filasParaEditar = excel.Selection.Rows.Count
        primeraFilaSel = excel.Selection.Row
        ultimaFilaSel = primeraFilaSel + filasParaEditar -1
        
        idsParaEditar = hojaOrdenes.Range(hojaOrdenes.Cells(primeraFilaSel,2),hojaOrdenes.Cells(ultimaFilaSel,2)).Value
        if primeraFilaSel <8 or ultimaFilaSel >ultimaFilaOrdenes:
            funciones.mostrarMensajeAdvertencia("El rango seleccionado no es correcto.")
            return
            
        #Traer los datos de la intención desde el archivo de intenciones, si no se selecciona ningún nombre o Todos los gerentes, 
        #entonces se trae solo los datos del gerente que está operando la herramienta
        miUsuario =  getpass.getuser().upper()
        gerenteAConsultar = hojaOrdenes.OLEObjects('listaGerentes').Object.Value
        if gerenteAConsultar in ["","(Todos)"]:
            archivo = [miUsuario+".csv"]
        else:
            archivo = [gerenteAConsultar + ".csv"]
        intenciones = funciones.obtenerIntenciones(definiciones.parametros["Valor"]["rutaIntenciones"],archivo)
        intenciones["Id"] = intenciones["Id"].astype('int64')
        if len(intenciones) ==0:
           funciones.mostrarMensajeAdvertencia("No hay intenciones.") 
           return
       
        intenciones = intenciones[intenciones["Estado"].isin(["Vencida","Cancelada"])]
        if len(intenciones) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para renovar.")
            return
        
        if funciones.esUnaLista(idsParaEditar):
            idsParaEditar =  list(map(lambda x: int(x[0]),idsParaEditar))
        else:
            idsParaEditar = [idsParaEditar]
        
        #Traer las operaciones seleccionadas
        intencionesParaEditar = intenciones.loc[intenciones["Id"].isin(idsParaEditar)]
        if len(intencionesParaEditar) == 0 :
            funciones.mostrarMensajeAdvertencia("No hay intenciones para renovar.")
            return
        
        intencionesParaEditar = intencionesParaEditar.loc[:,definiciones.parametros["Valor"]["ColumnasFormularioRenovacionArchivo"].split("-")]
        
        #elegir las columnas para editar
        
        
        
        formularioOrdenes = intencionesAM.Worksheets("Formulario Ordenes")  
        primeraFilaTabla  = 1
        funciones.limpiarHojaIntencionesPM(formularioOrdenes, primeraFilaTabla)    
        formularioOrdenes.Range("A1").Value = "RENOVACION"
        formularioOrdenes.Range("B1").Value = archivo[0].split(".")[0]

        camposMostrar = definiciones.parametros["Valor"]["ColumnasFormularioRenovacion"].split("-")
        celdasEncabezado = formularioOrdenes.Range(formularioOrdenes.Cells(7,2),formularioOrdenes.Cells(7,1 + len(camposMostrar)))
        celdasEncabezado.Value = camposMostrar
        #Dar formato a la nueva tabla
        celdasEncabezado.Font.Name ="CIBFont Sans"
        celdasEncabezado.Font.Size = 11
        celdasEncabezado.Font.ThemeColor = 1
        celdasEncabezado.Font.Bold = True
        celdasEncabezado.Interior.Color = 0x2C2A29
        celdasEncabezado.ColumnWidth = 20
        celdasEncabezado.Borders(definiciones.xlInsideVertical).ThemeColor = 1
        formularioOrdenes.Visible = True    
        formularioOrdenes.Range("C:Z").NumberFormat = "@" #Formato Texto
        formularioOrdenes.Range("B:B").NumberFormat = "0"
        #Dar a las columnas de fecha el formato de texto en excel
        formularioOrdenes.Activate()
        
        #colocar los datos en la hoja del formulario incluyendo el Id, de acá en adelante ya está lista la programación, falta ajustra el guardar los datos cuando sean una edición
        camposTablaEdicion = definiciones.parametros["Valor"]["ColumnasFormularioRenovacionArchivo"].split("-")
        intencionesParaEditar = intencionesParaEditar.where(~intencionesParaEditar.isna(), other="")
        formularioOrdenes.Range(formularioOrdenes.Cells(8,2),formularioOrdenes.Cells(7 + len(intencionesParaEditar),len(camposTablaEdicion)+1)).Value = intencionesParaEditar[camposTablaEdicion].to_records(index=False)
        
        #Bloquaer hoja para que no se agreguen mas intenciones
        #Colocar los Ids y no permitir modificar esta columna
        formularioOrdenes.Range(formularioOrdenes.Cells(8,camposTablaEdicion.index("VigenciaDesde")+2),formularioOrdenes.Cells(7 + len(intencionesParaEditar),camposTablaEdicion.index("VigenciaDesde")+2)).Locked = False
        formularioOrdenes.Range(formularioOrdenes.Cells(8,camposTablaEdicion.index("VigenteHasta")+2),formularioOrdenes.Cells(7 + len(intencionesParaEditar),camposTablaEdicion.index("VigenteHasta")+2)).Locked = False
        formularioOrdenes.Columns.AutoFit()
        formularioOrdenes.Protect()
            
        
class BtnCerrarTrazabilidadEvents:

    def OnClick(self,*args):
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        hojaTrazabilidad = intencionesAM.Worksheets("Trazabilidad")
        if hojaTrazabilidad.Range("A1").Value == "Monitor":
            hojaDestino = intencionesAM.Worksheets("Monitor")
        else:
            hojaDestino = intencionesAM.Worksheets("Mis Órdenes")
        
        hojaTrazabilidad = intencionesAM.Worksheets("Trazabilidad")
        hojaTrazabilidad.Unprotect()
        hojaTrazabilidad.Visible= False
        hojaTrazabilidad.UsedRange.ClearContents()
        hojaDestino.Visible = True
        hojaDestino.Activate()



class workbookEvents:

    def OnSheetSelectionChange(self,*args):
        
        hoja = args[0]
        rango = args[1]        
        intencionesAM = win32.GetObject(definiciones.parametros["Valor"]["rutaDestinoAplicacionExcel"])
        
        
        #Esto aplica solo para la hoja donde se ingresan las intenciones
        if hoja.Name == "Formulario Ordenes":
            
            creacionEdicion = hoja.Range("A1").Value
            if creacionEdicion in ["CANCELACION","RENOVACION"]:
                return
            
            macroActivo = intencionesAM.Worksheets("Mis Órdenes").OLEObjects("lstMacroActivosPM").Object.Value
           
            
            if macroActivo == "Renta Variable":
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionRV"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = encabezado.index("Mercado") +2
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colEmisor =encabezado.index("Emisor")+2
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colCantDisponible = encabezado.index("Cantidad disponible") +2                 
                colIndicador = 1000
                colTipoOrden = encabezado.index("Tipo orden") +2
                colDenominacion = 1000
                colDesde = 1000
                colHasta = 1000
                ###Opciones para seleccionar
                opcionesMercado = definiciones.parametros["Valor"]["Mercados RV"]
                opcionesDenominacion = ""
                opcionesIndicador = ""
                opcionesTipoOperacion =  definiciones.parametros["Valor"]["Tipo Operacion RV"] 
                opcionesTipoOrden = definiciones.parametros["Valor"]["Tipo Orden RV"]
                opcionesDesde = ""
                opcionesHasta = ""
            
            if macroActivo == "Deuda Privada":
                
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionDPr"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = 1000
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colEmisor =encabezado.index("Emisor")+2
                colCantDisponible = encabezado.index("Cantidad disponible(Millones)") +2 
                colIndicador = encabezado.index("Indicador") +2
                colDenominacion = 1000
                colDesde = encabezado.index("Desde") +2
                colHasta = encabezado.index("Hasta") +2
                colTipoOrden = encabezado.index("Tipo orden") +2
                ###Opciones para seleccionar
                opcionesMercado = ""
                opcionesDenominacion = ""
                opcionesIndicador = definiciones.parametros["Valor"]["Indicadores DPR"]
                opcionesTipoOperacion =  definiciones.parametros["Valor"]["Tipo Operacion DPR"]  
                opcionesTipoOrden = definiciones.parametros["Valor"]["Tipo Orden DPR"]
                opcionesDesde = definiciones.parametros["Valor"]["Desde DPR"]
                opcionesHasta = definiciones.parametros["Valor"]["Hasta DPR"] 
                
            if macroActivo == "Deuda Pública":
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionDPu"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = 1000
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colEmisor =encabezado.index("Emisor")+2
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colCantDisponible = encabezado.index("Cantidad disponible(Millones)") +2 
                colIndicador = 1000
                colDenominacion = 1000
                colDesde = 1000
                colHasta = 1000
                colTipoOrden = encabezado.index("Tipo orden") +2
                ###Opciones para seleccionar
                opcionesMercado = ""
                opcionesDenominacion = ""
                opcionesIndicador = ""
                opcionesTipoOperacion =  definiciones.parametros["Valor"]["Tipo Operacion DPU"]
                opcionesTipoOrden = definiciones.parametros["Valor"]["Tipo Orden DPU"]
                opcionesDesde = ""
                opcionesHasta = ""
                
            if macroActivo == "Fondos":
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionFondos"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = 1000
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colEmisor =encabezado.index("Emisor")+2
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colCantDisponible = encabezado.index("Cantidad disponible") +2 
                colIndicador = 1000
                colDenominacion = 1000
                colDesde =  0
                colHasta = 1000
                colTipoOrden = 1000
                ###Opciones para seleccionar
                opcionesMercado = ""
                opcionesDenominacion = ""
                opcionesIndicador = ""
                opcionesTipoOperacion =  definiciones.parametros["Valor"]["Tipo Operacion Fondos"]  
                opcionesTipoOrden = ""
                opcionesDesde = ""
                opcionesHasta = ""
                
            if macroActivo == "Forex":
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionForex"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = 1000
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colEmisor = 1000
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colCantDisponible = 1000
                colIndicador = 1000
                colDenominacion = 1000
                colDesde = 1000
                colHasta = 1000
                colTipoOrden = encabezado.index("Tipo orden") +2
                ###Opciones para seleccionar
                opcionesMercado = ""
                opcionesDenominacion = ""
                opcionesIndicador = ""
                opcionesTipoOperacion =  definiciones.parametros["Valor"]["Tipo Operacion Forex"]
                opcionesTipoOrden = definiciones.parametros["Valor"]["Tipo Orden Forex"]
                opcionesDesde = ""
                opcionesHasta = ""
                
            if macroActivo == "Liquidez":
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionLiquidez"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = 1000
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colEmisor = 1000
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colCantDisponible = encabezado.index("Cantidad disponible") +2 
                colIndicador = 1000
                colDenominacion = 1000
                colDesde = encabezado.index("Desde") +2
                colHasta = encabezado.index("Hasta") +2
                colTipoOrden = encabezado.index("Tipo orden") +2
                ###Opciones para seleccionar
                opcionesMercado = ""
                opcionesDenominacion = ""
                opcionesIndicador = ""
                opcionesTipoOperacion =  "REPO ACTIVO,REPO PASIVO,SIMULTANEA ACTIVA,SIMULTANEA PASIVA"  
                opcionesTipoOrden = "LIMITE,A MERCADO"
                opcionesDesde = "0 DIAS,30 DIAS,60 DIAS,90 DIAS,180 DIAS,270 DIAS,1 AÑO,1.5 AÑOS,2 AÑOS,3 AÑOS,4 AÑOS,5 AÑOS,6 AÑOS,7 AÑOS,8 AÑOS,9 AÑOS"
                opcionesHasta = "0 DIAS,30 DIAS,60 DIAS,90 DIAS,180 DIAS,270 DIAS,1 AÑO,1.5 AÑOS,2 AÑOS,3 AÑOS,4 AÑOS,5 AÑOS,6 AÑOS,7 AÑOS,8 AÑOS,9 AÑOS"
                
            if macroActivo == "Swaps":
                encabezado = definiciones.parametros["Valor"]["ColumnasFormularioCreacionSwaps"].split("-")
                colIDPortafolio = encabezado.index("Id Portafolio") +2
                colPortafolio = encabezado.index("Portafolio") +2
                colMercado = encabezado.index("Mercado") +2
                colTipoOperacion = encabezado.index("Tipo operación") +2
                colEmisor = encabezado.index("Emisor") +2
                colNemotecnico = encabezado.index("Nemotécnico") +2
                colCantDisponible = encabezado.index("Cantidad") +2
                colIndicador = encabezado.index("Indicador") + 2
                colDenominacion = 1000
                colDesde = 1000
                colHasta = encabezado.index("Hasta") +2
                colTipoOrden = encabezado.index("Tipo orden") +2
                ###Opciones para seleccionar
                opcionesMercado = definiciones.parametros["Valor"]["Mercados Swaps"]
                opcionesDenominacion = ""
                opcionesIndicador = ""
                opcionesTipoOperacion =  definiciones.parametros["Valor"]["Tipo Operacion Swaps"] 
                opcionesTipoOrden = definiciones.parametros["Valor"]["Tipo Orden Swaps"]
                opcionesDesde = ""
                opcionesHasta = definiciones.parametros["Valor"]["Hasta Swaps"]
                
            if creacionEdicion == "EDICION":
                colIDPortafolio = colIDPortafolio +1
                colPortafolio = colPortafolio +1
                colMercado = colMercado +1
                colTipoOperacion = colTipoOperacion +1
                colEmisor = colEmisor +1
                colNemotecnico = colNemotecnico +1
                colCantDisponible = colCantDisponible +1
                colIndicador = colIndicador +1
                colTipoOrden = colTipoOrden +1
                colDenominacion = colDenominacion +1
                colDesde = colDesde +1
                colHasta =  colHasta  +1
            
            ultFilaPortafolio = hoja.Cells(hoja.Rows.Count, colIDPortafolio).End(definiciones.xlUp).Row
            portafolio = hoja.Cells(rango.Row,colIDPortafolio).Value            
            mercado = hoja.Cells(rango.Row,colMercado).Value
            tipoOperacion =  hoja.Cells(rango.Row,colTipoOperacion).Value
            nemotecnico = hoja.Cells(rango.Row,colNemotecnico).Value 
            emisor = hoja.Cells(rango.Row,colEmisor).Value
            if portafolio != None:
                portafolio = portafolio.strip().upper()
            if mercado != None:
                mercado = mercado.strip().upper()
            if tipoOperacion != None:
                tipoOperacion = tipoOperacion.strip().upper()
            if nemotecnico != None:
                nemotecnico = nemotecnico.strip().upper()
            if emisor != None:
                emisor = emisor.strip().upper()
                      
            if creacionEdicion == "EDICION": 
                if rango.Column == 1 or rango.Row> ultFilaPortafolio or rango.Row <8:
                    hoja.Protect()                    
                    return  
                else:
                    hoja.Unprotect()
            elif creacionEdicion == "RENOVACION":
                if rango.Locked == True:
                    hoja.Protect()
                    return
                else:
                    hoja.Unprotect()
            elif creacionEdicion == "CANCELACION":
                return

            hoja.Columns.AutoFit()  
            if emisor == None:
                hoja.Columns(hoja.Cells(8,colEmisor).Column).ColumnWidth = 30
            if nemotecnico == None:
                hoja.Columns(hoja.Cells(8,colNemotecnico).Column).ColumnWidth = 30
            
            #SELECCIÓN DE DATOS
            
            
            #ID PORTAFOLIO
            if creacionEdicion == "CREACION":
                
                if ultFilaPortafolio + 1  == rango.Row and rango.Column == colIDPortafolio and hoja.Range(rango.Address).Value == None :     
                    Formula1 = funciones.obtenerPortafolios() 
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
            else:#EDICION                
                if rango.Column == colIDPortafolio:
                    Formula1 = funciones.obtenerPortafolios() 
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                
            #NOMBRE PORTAFOLIO
            if portafolio != None and rango.Column == colPortafolio and  rango.Row >7:
                hoja.Range(rango.Address).Value = funciones.obtenerNombrePortafolioPorId(portafolio)
                return
            
            #MERCADO
            if portafolio != None and rango.Column == colMercado and rango.Row >7:
                Formula1 = opcionesMercado   
                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                    if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                else:
                    hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
            
            #TIPO OPERACIÓN
            if portafolio != None and rango.Column == colTipoOperacion and rango.Row >7:
                Formula1 = opcionesTipoOperacion 
                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                    if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                else:
                    hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
            #Selección de nemotécnicos: Solo si ya ha seleccionado un portafolio y mercado
                    
            #EMISOR
            if portafolio != None  and tipoOperacion != None and rango.Column == colEmisor and rango.Row >7: 
                if macroActivo == "Renta Variable":
                    if mercado!= None:
                        macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                        if tipoOperacion == "VENTA":
                            if nemotecnico == None:
                                Formula1 = funciones.obtenerEmisoresdePortafolio(portafolio,macroActivoTitulos)   
                            else:
                                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                    hoja.Range(rango.Address).Validation.Delete()
                                hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemo(portafolio,macroActivoTitulos,nemotecnico)
                                return
                        else: #COMPRA     
                            if nemotecnico == None:                                         
                                Formula1 = funciones.obtenerEmisoresdePortafolioCompras(portafolio,macroActivoTitulos)
                            else:
                                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                    hoja.Range(rango.Address).Validation.Delete()
                                hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemoCompra(portafolio,nemotecnico)
                                return
                        
                        if funciones.celdaConValidador(hoja.Range(rango.Address)):
                            if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                                hoja.Range(rango.Address).ClearContents()
                                hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                        else:
                            hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                        
                elif macroActivo  == "Deuda Privada":   
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    
                    if tipoOperacion == "VENTA":                        
                        if nemotecnico == None:
                            Formula1 = funciones.obtenerEmisoresdePortafolio(portafolio,macroActivoTitulos)   
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemo(portafolio,macroActivoTitulos,nemotecnico)
                            return
                    else: #COMPRA                        
                       
                        if nemotecnico == None:                                         
                            Formula1 = funciones.obtenerEmisoresdePortafolioCompras(portafolio,macroActivoTitulos)
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemoCompra(portafolio,nemotecnico)
                            return
                       
                       
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                        
                        if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                            hoja.Range(rango.Address).ClearContents()
                            hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                
                elif macroActivo == "Deuda Pública":
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    
                    
                    if tipoOperacion == "VENTA" :
                        if nemotecnico == None:
                            Formula1 = funciones.obtenerEmisoresdePortafolio(portafolio,macroActivoTitulos)   
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemo(portafolio,macroActivoTitulos,nemotecnico)
                            return 
                    elif tipoOperacion == "COMPRA": #COMPRA
                        if nemotecnico == None:                                         
                            Formula1 = funciones.obtenerEmisoresdePortafolioCompras(portafolio,macroActivoTitulos)
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemoCompra(portafolio,nemotecnico)
                            return
                    else:
                        if funciones.celdaConValidador(hoja.Range(rango.Address)):
                            hoja.Range(rango.Address).Validation.Delete()
                        hoja.Range(rango.Address).Value  = "FUTUROS"
                        return
                        
                        
                        
                        
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                        
                        if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                            hoja.Range(rango.Address).ClearContents()
                            hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                
                elif macroActivo == "Fondos":
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    if tipoOperacion == "APERTURA":
                        
                        if nemotecnico == None:                                         
                            Formula1 = funciones.obtenerEmisoresdePortafolioCompras(portafolio,macroActivoTitulos)
                            
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemoCompra(portafolio,nemotecnico)
                            return                        
                    elif tipoOperacion == "ADICION":
                        if nemotecnico != None:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemoCompra(portafolio,nemotecnico)
                            return 
                        else:
                            return

                    else:
                        if nemotecnico == None:
                            Formula1 = funciones.obtenerEmisoresdePortafolio(portafolio,macroActivoTitulos)   
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemo(portafolio,macroActivoTitulos,nemotecnico)
                            return 
                         
                    
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                            
                        if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                            hoja.Range(rango.Address).ClearContents()
                            hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                            
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                        
                
                elif macroActivo == "Swaps":
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    
                    if tipoOperacion == "UNWIND":                        
                        if nemotecnico == None:
                            Formula1 = funciones.obtenerEmisoresdePortafolio(portafolio,macroActivoTitulos)   
                        else:
                            if funciones.celdaConValidador(hoja.Range(rango.Address)):
                                hoja.Range(rango.Address).Validation.Delete()
                            hoja.Range(rango.Address).Value  = funciones.obtenerEmisordeNemo(portafolio,macroActivoTitulos,nemotecnico)
                            return
                    else:
                        hoja.Range(rango.Address).Value = "-"
                        return
                    
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                            
                        if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                            hoja.Range(rango.Address).ClearContents()
                            hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                            
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                 
            #NEMOTECNICO
            
            if portafolio != None  and tipoOperacion != None and rango.Column == colNemotecnico and rango.Row >7:  
                
                if macroActivo   == "Renta Variable":                    
                    if mercado!= None:
                        macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado) 
                        if tipoOperacion == "VENTA":
                            if emisor == None:
                                Formula1 = funciones.obtenerNemosdePortafolio(portafolio, macroActivoTitulos)                                
                            else:
                                Formula1 = funciones.obtenerNemosdeEmisor(portafolio,macroActivoTitulos,emisor)   
                        else: #COMPRA
                            if emisor == None:                                
                                Formula1 = funciones.obtenerNemosdeMacroactivo(portafolio,macroActivoTitulos)                                
                                
                            else:                                            
                                Formula1 = funciones.obtenerNemosDeEmisorCompras(portafolio,emisor,macroActivoTitulos)
                    else:
                        return
                    
                elif macroActivo  == "Deuda Privada":   
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    
                    if tipoOperacion == "VENTA":  
                        if emisor == None:
                            Formula1 = funciones.obtenerNemosdePortafolio(portafolio, macroActivoTitulos)                                
                        else:
                            Formula1 = funciones.obtenerNemosdeEmisor(portafolio,macroActivoTitulos,emisor)                      
                      
                    else: #COMPRA  
                        if emisor == None:                                
                            Formula1 = funciones.obtenerNemosdeMacroactivo(portafolio,macroActivoTitulos)                                
                        else:                                            
                            Formula1 = funciones.obtenerNemosDeEmisorCompras(portafolio,emisor,macroActivoTitulos)                        
                    
                elif macroActivo == "Deuda Pública":
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    
                    
                    if tipoOperacion == "VENTA":
                        if emisor == None:
                            Formula1 = funciones.obtenerNemosdePortafolio(portafolio, macroActivoTitulos)                                
                        else:
                            Formula1 = funciones.obtenerNemosdeEmisor(portafolio,macroActivoTitulos,emisor)
                    elif tipoOperacion == "COMPRA":
                        if emisor == None:                                
                            Formula1 = funciones.obtenerNemosdeMacroactivo(portafolio,macroActivoTitulos)                                
                        else:                                            
                            Formula1 = funciones.obtenerNemosDeEmisorCompras(portafolio,emisor,macroActivoTitulos)
                    else: #FUTUROS
                        Formula1 = funciones.obtenerFuturosDPU()   
                       
                elif macroActivo == "Fondos":
                    macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                    if tipoOperacion == "APERTURA":
                        if emisor == None:                                
                            Formula1 = funciones.obtenerNemosdeMacroactivo(portafolio,macroActivoTitulos)                                
                        else:                                            
                            Formula1 = funciones.obtenerNemosDeEmisorCompras(portafolio,emisor,macroActivoTitulos)

                    elif tipoOperacion == "ADICION":
                        Formula1 = funciones.obtenerNemosdePortafolio(portafolio, macroActivoTitulos)
                    else:                         
                        if emisor == None:
                            Formula1 = funciones.obtenerNemosdePortafolio(portafolio, macroActivoTitulos)                                
                        else:
                            Formula1 = funciones.obtenerNemosdeEmisor(portafolio,macroActivoTitulos,emisor)
                            
                 
                elif macroActivo == "Forex":
                    Formula1 = funciones.obtenerNemosForex(tipoOperacion)  
                    
                
                elif macroActivo == "Swaps":
                    if tipoOperacion == "UNWIND":
                        macroActivoTitulos = "SWAP"
                        if emisor == None:
                            Formula1 = funciones.obtenerNemosdePortafolio(portafolio, macroActivoTitulos)                                
                        else:
                            Formula1 = funciones.obtenerNemosdeEmisor(portafolio,macroActivoTitulos,emisor)
                        
                    
                    elif tipoOperacion == "VENTA":
                        Formula1 = "Recibo Tasa Fija-Entrego Tasa Variable"
                        
                    else:
                        Formula1 = "Recibo Variable-Entrego Tasa Fija"
                
                else:
                    return
                
                datos = Formula1.split(",")  
                #datos = list(map(lambda x: x.strip(),datos))
                             
                if len(datos) > 8: 
                    hojaEspecies = intencionesAM.Worksheets("Especies")
                    Formula1 = funciones.crearValidadorConFiltro (hojaEspecies,datos,rango)
                    
                if nemotecnico == None: #Hay datos en la celda de nemotécnico?                       
                    if funciones.celdaConValidador(hoja.Range(rango.Address)): #Ya existe un validador
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                        hoja.Range(rango.Address).Validation.ShowError = False #Frank
                    else:                            
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                        hoja.Range(rango.Address).Validation.ShowError = False #Frank  
                        
                 
                return  
            

                    
                        
                            
                 
            #CANTIDAD DISPONIBLE   
            if nemotecnico != None and rango.Column == colCantDisponible  and rango.Row >7 :
                macroActivoTitulos = funciones.traducirMacroActivo(macroActivo,mercado)
                
                if macroActivo == "Renta Variable":
                    if tipoOperacion == "VENTA":
                        hoja.Range(rango.Address).Value = funciones.obtenerCantidadDisponibleNemos(nemotecnico,portafolio,macroActivoTitulos)
                    else:
                        hoja.Range(rango.Address).Value = funciones.obtenerCuposcompras(portafolio,macroActivo,nemotecnico)
                
                elif macroActivo == "Deuda Privada":
                    if tipoOperacion == "VENTA":
                        hoja.Range(rango.Address).Value = funciones.obtenerCantidadDisponibleNemos(nemotecnico,portafolio,macroActivoTitulos)
                    else:
                        hoja.Range(rango.Address).Value = funciones.obtenerCuposcompras(portafolio,macroActivo,nemotecnico)
                
                elif macroActivo == "Deuda Pública":
                    if tipoOperacion == "VENTA":
                        hoja.Range(rango.Address).Value = funciones.obtenerCantidadDisponibleNemos(nemotecnico,portafolio,macroActivoTitulos)                        
                    elif tipoOperacion == "COMPRA":
                        hoja.Range(rango.Address).Value = funciones.obtenerCuposcompras(portafolio,macroActivo,nemotecnico)
                    else:
                        hoja.Range(rango.Address).Value = "SIN DATOS"
                    
                elif macroActivo == "Fondos":
                    if tipoOperacion in ["RETIRO","CANCELACION"]:
                        hoja.Range(rango.Address).Value = funciones.obtenerCantidadDisponibleNemosFondos(nemotecnico,portafolio,macroActivoTitulos,tipoOperacion)
                    else:
                        hoja.Range(rango.Address).Value = funciones.obtenerCuposcompras(portafolio,macroActivo,nemotecnico)  
                
                elif macroActivo == "Swaps":
                    if tipoOperacion == "UNWIND": #Solo aplica para Swaps UNWIND, se colocará la cantidad disponible en el campo cantidad
                        hoja.Range(rango.Address).ClearContents()
                        hoja.Range(rango.Address).Value = funciones.obtenerCantidadDisponibleNemos(nemotecnico,portafolio,macroActivoTitulos)
                    
                else:
                    hoja.Range(rango.Address).Value = "SIN DATOS"        
                        
                return    
            
            #TIPO ORDEN
            if rango.Column== colTipoOrden and rango.Row >7 :
                Formula1 = opcionesTipoOrden   
                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                    if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                else:
                    hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
            
            #DESDE
            if portafolio != None and rango.Column == colDesde and rango.Row >7:
                Formula1 = opcionesDesde 
                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                    if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                else:
                    hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
            
            #HASTA
            if portafolio != None and rango.Column == colHasta and rango.Row >7:
                Formula1 = opcionesHasta 
                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                    if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                else:
                    hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
            #INDICADOR
            if portafolio != None and nemotecnico != None and rango.Column == colIndicador and rango.Row >7:
                if macroActivo == "Deuda Privada":
                    especies = definiciones.especies.copy()
                    if nemotecnico in especies.loc[especies["Nemotecnico"] == "GENERICO DPR","Nemo intenciones"].tolist():
                        Formula1 = opcionesIndicador   
                    elif nemotecnico == "-":
                        Formula1 = "-"
                    else:                        
                        Formula1 = funciones.obtenerIndicadorNemo(portafolio,nemotecnico)
                        if Formula1 == "NAN":
                            Formula1 = opcionesIndicador                            
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                        if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                            hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
            
                if macroActivo == "Swaps" and mercado != None:
                    if mercado =="LOCAL":
                        Formula1 = definiciones.parametros["Valor"]["Indicadores Local Swaps"]
                    else:
                        Formula1 = definiciones.parametros["Valor"]["Indicadores Internacional Swaps"]                   
                    if funciones.celdaConValidador(hoja.Range(rango.Address)):
                        if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                            hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                    else:
                        hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
                    
                
            #DENOMINACIÓN
            if portafolio != None and rango.Column == colDenominacion and rango.Row >7:
                Formula1 = opcionesDenominacion  
                if funciones.celdaConValidador(hoja.Range(rango.Address)):
                    if Formula1 != hoja.Range(rango.Address).Validation.Formula1:
                        hoja.Range(rango.Address).Validation.Modify(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                else:
                    hoja.Range(rango.Address).Validation.Add(definiciones.xlValidateList,definiciones.xlValidAlertStop,definiciones.xlBetween,Formula1)
                return
            
            #Muestra del inventario
        #Esto aplica solo para la hoja donde los Traders editan intenciones
        if hoja.Name == "Ejecutar Intenciones": 
            hoja.Unprotect()
            colComentariosTrader = definiciones.parametros["Valor"]["ColumnasEjecutarIntenciones"].split("-").index("Comentarios Trader") + 2
            opcionesComentariosTrader =  definiciones.parametros["Valor"]["comentariosTrader"]    
            if rango.Locked == True:
                    hoja.Protect()
                    return
            
            
            
        