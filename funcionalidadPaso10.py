import pandas as pd
#esta funcion es provicional y no es la real es solo para probar el funcionamiento del paso 10
def obtenerUltimoArchivoCarpeta(ruta):
    import os
    archivos = [f for f in os.listdir(ruta) if os.path.isfile(os.path.join(ruta, f))]
    archivos.sort(reverse=True)  # Ordenar por fecha, asumiendo que el nombre del archivo contiene la fecha
    return os.path.join(ruta, archivos[0]) if archivos else None
#nueva version de la funcion cargarOperacionesPorCumplirFiduciaria
def cargarOperacionesPorCumplirFiduciaria(ruta_insumo,nemoTitulosFiduciaria):
    
    archivo_operaciones_por_cumplir_fiduciaria = obtenerUltimoArchivoCarpeta(ruta_insumo)        
    fechaOpPendientesFidu = archivo_operaciones_por_cumplir_fiduciaria.split(".")
    fechaOpPendientesFidu = fechaOpPendientesFidu[0][-8:]
    operaciones_por_cumplir_fiduciaria = pd.read_excel(archivo_operaciones_por_cumplir_fiduciaria)
    if len(operaciones_por_cumplir_fiduciaria) > 0:
        operaciones_por_cumplir_fiduciaria = operaciones_por_cumplir_fiduciaria.rename(columns={
            'Moneda instrumento': 'MON',
            'Cantidad / Valor nominal': 'VR NOMINAL ACTUAL',
            'Tipo operacion': 'TRANSACCIO',
            'Nombre portafolio': 'POR',
            'Descripcion instrumento': 'ESPECIE',
            'Macro activo': 'MC AT',
        })
        operaciones_por_cumplir_fiduciaria = operaciones_por_cumplir_fiduciaria[operaciones_por_cumplir_fiduciaria["Mod"]!= "---"]
        operaciones_por_cumplir_fiduciaria.columns =  operaciones_por_cumplir_fiduciaria.columns.str.strip().str.upper()
        operaciones_por_cumplir_fiduciaria =  limpiarDatos(operaciones_por_cumplir_fiduciaria,colTrim=["TRANSACCIO","POR","MON","ESPECIE"],colUpper=["TRANSACCIO","POR","MON","ESPECIE"],colFloat=["VR NOMINAL ACTUAL"])
        operaciones_por_cumplir_fiduciaria["PORTAFOLIO"] = operaciones_por_cumplir_fiduciaria["POR"]

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

