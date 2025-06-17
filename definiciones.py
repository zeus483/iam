# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 17:05:23 2022

@author: FRCASTRO
"""
import pandas as pd

ruta = open('//sbbogscl0/VP Asset Management/Gcia Renta Variable/Ordenes/IntencionesAM/Codigo/Codigo Python/ruta.txt','r')
lineaDeParametros = ruta.readline() 
ruta.close()
archivoParametros = lineaDeParametros.split("|")[0]
rutaCupos = lineaDeParametros.split("|")[1]
xlUp = -4162
xlValidateList = 3
xlValidAlertStop = 1
xlBetween = 1
xlInsideVertical = 11
xlMaximized = -4137
xlCellTypeVisible = 12
xlPie = 5
xlVeryHidden = 2
xlLocationAsObject = 2
xlLocationAutomatic = 3

msoElementDataLabelBestFit = 210

def crearVariablesDeDatos():
    global parametros, portafoliosCRM, nemoTitulosFiduciaria, nemoTitulosValores,usuarios,  keepOpen,especies,nitEmisoresLocales,porcentajeProtejidoFondos, valorPortafolioFiduciaria,  valorPortafolioValores, cuposValores, cuposFiduciaria, precioAcciones, UVR,TRM,nombreYcodigosPortValores, tablaSobrepasos, operacionesPorCumplirFiduciaria
    global fechaInvenFidu, fechaInvenVal, fechaOpPendientesFidu, fechaOpPendientesValores, fechaSeguimientoLimites, fechaSobrepasoLimites
    
    

