// @ts-ignore: XLSX viene del CDN
declare var XLSX: any;

// Declaración para que TypeScript permita exponer funciones en window
interface Window {
    LeerExcelDesdeArchivo: typeof LeerExcelDesdeArchivo;
    AgregarCodigosMurex: typeof AgregarCodigosMurex;
    CargarPortafoliosCRM: typeof CargarPortafoliosCRM;
}

// Interfaces
interface OpcionesLectura {
    hoja?: string;
    columnasDeseadas: string[];
    filasAsaltadas?: number;
}

interface MapeoOydaMurex {
    codigo_oyd: string;
    codigo_murex: string;
}

// Función: Leer Excel desde archivo
async function LeerExcelDesdeArchivo(
    archivo: File,
    opciones: OpcionesLectura
): Promise<any[]> {
    const data = await archivo.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const nombreHoja = opciones.hoja || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[nombreHoja];

    if (!worksheet) {
        throw new Error(`La hoja "${nombreHoja}" no existe en el archivo.`);
    }

    const datos = XLSX.utils.sheet_to_json(worksheet, {
        range: opciones.filasAsaltadas ?? 0,
        defval: '',
    });

    const filtrado = datos.map((fila: any) => {
        const resultado: any = {};
        for (const col of opciones.columnasDeseadas) {
            resultado[col] = fila[col];
        }
        return resultado;
    });

    return filtrado;
}

// Función: Agregar códigos Murex
async function AgregarCodigosMurex(
    datosExcel: any[],
    endpointBackend: string
): Promise<any[]> {
    const response = await fetch(endpointBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'consultaoydmurex' })
    });

    if (!response.ok) {
        throw new Error('Error al consultar el backend para obtener el mapeo de códigos.');
    }

    const mapeo: MapeoOydaMurex[] = await response.json();
    const mapa = new Map(mapeo.map(item => [item.codigo_oyd.toString(), item.codigo_murex]));

    const resultado = datosExcel.map(fila => {
        let codMurex = '';
        if (fila["ADMINISTRADOR"] === "VALORES") {
            const codOyd = fila["CÓD. CONT. | OYD | PERSHING"]?.toString().trim();
            if (mapa.has(codOyd)) {
                codMurex = mapa.get(codOyd) || '-';
            }
        }
        return { ...fila, "CÓDIGO MUREX": codMurex };
    });

    return resultado;
}
//funcion agregar codigo murex para consolidado valores
async function AgregarCodigosMurexConsolidadoFidu(
    datosExcel: any[],
    endpointBackend: string
): Promise<any[]> {
    const response = await fetch(endpointBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'consultaoydmurex' })
    });
    if (!response.ok) {
        throw new Error('Error al consultar el backend para obtener el mapeo de códigos.');
    }
    const mapeo: MapeoOydaMurex[] = await response.json();
    const mapa = new Map(mapeo.map(item => [item.codigo_oyd.toString(), item.codigo_murex]));
    const resultado = datosExcel.map(fila => {
        let codMurex = '';
        const codOyd = fila["Código OyD"]?.toString().trim();
        if (mapa.has(codOyd)) {
            codMurex = mapa.get(codOyd) || '-';
        }
        return { ...fila, "Portafolio": codMurex };
    });
    //ahora nos quedamos solo con los que en portafolio tengan dato valido
    const response2 = await fetch(endpointBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'obtenerportafoliosvalorescrm' })
    });
    if (!response2.ok) {
        throw new Error('Error al consultar el backend para obtener los portafolios de valores.');
    }
    const portafoliosPermitidos: string[] = await response2.json();
    const portafoliosSet = new Set(portafoliosPermitidos.map(p => p.toString().trim()));
    const resultado2 = resultado.filter(fila => {
        const portafolio = fila["Portafolio"]?.toString().trim();
        return portafolio && portafoliosSet.has(portafolio);
    });
    return resultado2
}
//Funcion para agregar la columna MUREX con los codigos de murex a base cupos valores
async function AgregarCodigosMurexCuposValores(
    datosExcel: any[],
    endpointBackend: string
): Promise<any[]> {
    const response = await fetch(endpointBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'consultaoydmurex' })
    });
    if (!response.ok) {
        throw new Error('Error al consultar el backend para obtener el mapeo de códigos.');
    }
    const mapeo: MapeoOydaMurex[] = await response.json();
    const mapa = new Map(mapeo.map(item => [item.codigo_oyd.toString(), item.codigo_murex]));
    const resultado = datosExcel.map(fila => {
        let codMurex = '';
        const codOyd = fila["OyD"]?.toString().trim();
        if (mapa.has(codOyd)) {
            codMurex = mapa.get(codOyd) || '-';
        }
        return { ...fila, "MUREX": codMurex };
    });
    //ahora nos quedamos solo con los que en MUREX tengan dato valido
    const response2 = await fetch(endpointBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'obtenerportafoliosvalorescrm' })
    });
    if (!response2.ok) {
        throw new Error('Error al consultar el backend para obtener los portafolios de valores.');
    }
    const portafoliosPermitidos: string[] = await response2.json();
    const portafoliosSet = new Set(portafoliosPermitidos.map(p => p.toString().trim()));
    const resultado2 = resultado.filter(fila => {
        const portafolio = fila["MUREX"]?.toString().trim();
        return portafolio && portafoliosSet.has(portafolio);
    });
    return resultado2;
    
}
// Función principal: flujo completo
async function CargarPortafoliosCRM(
    archivo: File,
    urlBackend: string
): Promise<void> {
    const datos = await LeerExcelDesdeArchivo(archivo, {
        hoja: 'Listado PD',
        columnasDeseadas: [
            'ADMINISTRADOR',
            'CÓD. CONT. | OYD | PERSHING',
            'NOMBRE PORTAFOLIO'
        ],
        filasAsaltadas: 4
    });

    const datosConMurex = await AgregarCodigosMurex(
        datos,
        urlBackend
    );

    const respClean = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'limpiartablaportafoliocrm' })
    });

    const resCleanJson = await respClean.json();
    if (resCleanJson.status !== 'ok') {
        throw new Error('El backend no respondió ok al limpiar la tabla.');
    }

    for (let i = 0; i < datosConMurex.length; i += 100) {
        const lote = datosConMurex.slice(i, i + 100);
        const respLote = await fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'carguecrm', data: lote })
        });

        const resLoteJson = await respLote.json();
        if (resLoteJson.status !== 'ok') {
            console.error(`Error al enviar el lote ${i / 100 + 1}`);
            break;
        }
    }

    console.log('Carga finalizada con éxito.');
}
// async function cargarConsolidadoFidu(archivo: File, urlBackend: string): Promise<void> {
//     const columnas = [
//         "Macro Activo",
//         "ISIN",
//         "Emisor / Contraparte",
//         "Especie/Generador",
//         "Nemotécnico",
//         "Emisor Unificado",
//         "SALDO Macro Activo",
//         "SALDO ABA",
//         "Nominal Remanente",
//         "Vr Mercado Hoy Moneda Empresa"
//     ];
//     const columnasNumericas = [
//         "SALDO Macro Activo",
//         "SALDO ABA",
//         "Nominal Remanente",
//         "Vr Mercado Hoy Moneda Empresa"
//     ];
//     const datos = await LeerExcelDesdeArchivo(archivo, {
//         hoja: 'Base_Consolidado',
//         columnasDeseadas: columnas,
//         filasAsaltadas: 2
//     });

//     // Normalizar filas
//     const normalizados = datos.map((fila: any) => {
//         const filaNormalizada: any = {};

//         for (const key of columnas) {
//             const valor = (fila[key] ?? "").toString().trim();

//             if (columnasNumericas.includes(key)) {
//                 // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
//                 const num = parseFloat(valor.replace(",", "."));
//                 filaNormalizada[key] = isNaN(num) ? null : num;
//             } else {
//                 filaNormalizada[key] = valor.toUpperCase();
//             }
//         }

//         filaNormalizada["origen informacion"] = "INVENTARIO TITULOS";

//         return filaNormalizada;
//     });
//     // Limpiar tabla en el backend
//     const respClean = await fetch(urlBackend, {
//         method: 'POST',
//         headers: { 'Content-Type': 'application/json' },
//         body: JSON.stringify({ id: 'limpiartablaconsolidadofidu' })
//     });

//     const resCleanJson = await respClean.json();
//     if (resCleanJson.status !== 'ok') {
//         throw new Error('El backend no respondió ok al limpiar Consolidado Fidu.');
//     }
//     // Enviar por bloques
//     const bloque = 400;
//     for (let i = 0; i < normalizados.length; i += bloque) {
//         const lote = normalizados.slice(i, i + bloque);
//         const respLote = await fetch(urlBackend, {
//             method: 'POST',
//             headers: { 'Content-Type': 'application/json' },
//             body: JSON.stringify({ id: 'cargueconsolidadofidu', data: lote })
//         });

//         const resLoteJson = await respLote.json();
//         if (resLoteJson.status !== 'ok') {
//             console.error(`Error al enviar el lote ${i / bloque + 1}`);
//             break;
//         }
//     }

//     console.log('Carga Consolidado Fidu finalizada con éxito.');

// }
async function cargarConsolidadoFidu(archivo: File, urlBackend: string): Promise<void> {
    const columnas = [
        "Macro Activo",
        "ISIN",
        "Emisor / Contraparte",
        "Especie/Generador",
        "Nemotécnico",
        "Emisor Unificado",
        "SALDO Macro Activo",
        "SALDO ABA",
        "Nominal Remanente",
        "Vr Mercado Hoy Moneda Empresa",
        "Portafolio"
    ];

    const columnasNumericas = [
        "SALDO Macro Activo",
        "SALDO ABA",
        "Nominal Remanente",
        "Vr Mercado Hoy Moneda Empresa"
    ];

    const datos = await LeerExcelDesdeArchivo(archivo, {
        hoja: 'Base_Consolidado',
        columnasDeseadas: columnas,
        filasAsaltadas: 2
    });

    // Validar columnas presentes
    const columnasArchivo = Object.keys(datos[0] || {});
    const faltantes = columnas.filter(c => !columnasArchivo.includes(c));
    if (faltantes.length > 0) {
        throw new Error(`Columnas faltantes en el Excel: ${faltantes.join(', ')}`);
    }

    // Normalizar filas
    const normalizados = datos.map((fila: any) => {
        const filaNormalizada: any = {};
        for (const key of columnas) {
            const valor = (fila[key] ?? "").toString().trim();
            if (columnasNumericas.includes(key)) {
                const num = parseFloat(valor.replace(",", "."));
                filaNormalizada[key] = isNaN(num) ? null : num;
            } else {
                filaNormalizada[key] = valor.toUpperCase();
            }
        }
        filaNormalizada["origen informacion"] = "INVENTARIO TITULOS";
        return filaNormalizada;
    });

    // Limpiar tabla
    const respClean = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'limpiartablaconsolidadofidu' })
    });
    const resCleanJson = await respClean.json();
    if (resCleanJson.status !== 'ok') {
        throw new Error('El backend no respondió ok al limpiar Consolidado Fidu.');
    }

    // Envío por bloques con reintento
    const bloque = 1000;
    for (let i = 0; i < normalizados.length; i += bloque) {
        const lote = normalizados.slice(i, i + bloque);
        const exito = await intentarEnvioLote(lote, urlBackend, i / bloque + 1);
        if (!exito) break;
    }

    console.log('Carga Consolidado Fidu finalizada con éxito.');
}

// Función auxiliar para reintentos
async function intentarEnvioLote(lote: any[], urlBackend: string, numeroLote: number, intentos = 3): Promise<boolean> {
    for (let intento = 1; intento <= intentos; intento++) {
        try {
            const resp = await fetch(urlBackend, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: 'cargueconsolidadofidu', data: lote })
            });
            const json = await resp.json();
            if (json.status === 'ok') {
                console.log(`✅ Lote ${numeroLote} enviado correctamente.`);
                return true;
            } else {
                throw new Error(`Backend respondió: ${json.status}`);
            }
        } catch (e) {
            console.warn(`⚠️ Error al enviar lote ${numeroLote}, intento ${intento}: ${e}`);
            await new Promise(r => setTimeout(r, 1000 * intento)); // backoff creciente
        }
    }
    console.error(`❌ No se pudo enviar el lote ${numeroLote} después de ${intentos} intentos.`);
    return false;
}

async function cargarConsolidadoValores(archivo: File, urlBackend: string): Promise<void> {
    console.log('Iniciando carga de Consolidado Valores...');
    const columnas = [
        "Macro Activo","Isin","Nemoténico","Emisor Unificado","Nombre Emisor","SALDO Macro Activo","SALDO ABA","Valor Nominal Actual","Codigo OyD","Fecha"
    ];
    const columnasNumericas = [
        "SALDO Macro Activo",
        "SALDO ABA",
        "Valor Nominal Actual"
    ];
    const datos = await LeerExcelDesdeArchivo(archivo, {
        hoja: 'Reporte_Consolidado',
        columnasDeseadas: columnas,
        filasAsaltadas: 8
    });
    // Normalizar filas
    const normalizados = datos.map((fila: any) => {
        const filaNormalizada: any = {};

        for (const key of columnas) {
            const valor = (fila[key] ?? "").toString().trim();

            if (columnasNumericas.includes(key)) {
                // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
                const num = parseFloat(valor.replace(",", "."));
                filaNormalizada[key] = isNaN(num) ? null : num;
            } else {
                filaNormalizada[key] = valor.toUpperCase();
            }
        }

        filaNormalizada["origen informacion"] = "INVENTARIO TITULOS";

        return filaNormalizada;
    });
    //Agregar lo de murex cod
    const datosConMurex = await AgregarCodigosMurexConsolidadoFidu(
        normalizados,
        urlBackend
    );
    const respClean = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'limpiartablaconsolidadovalores' })
    });
    const resCleanJson = await respClean.json();
    if (resCleanJson.status !== 'ok') {
        throw new Error('El backend no respondió ok al limpiar la tabla de Consolidado Valores.');
    }
    for (let i = 0; i < datosConMurex.length; i += 1000) {
        const lote = datosConMurex.slice(i, i + 1000);
        const respLote = await fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'cargueconsolidadovalores', data: lote })
        });

        const resLoteJson = await respLote.json();
        if (resLoteJson.status !== 'ok') {
            console.error(`Error al enviar el lote ${i / 1000 + 1}`);
            break;
        }
    }
    console.log('Carga Consolidado Valores finalizada con éxito.');


}
async function cargueBaseCuposFidu(archivo: File, urlBackend: string): Promise<void> {
    const columnas = [
        "Entidad","MUREX","Nombre_1","Cupo","Nemo","ISIN 1","Ocupación Máxima"
    ];
    const columnasNumericas = [
        "Ocupación Máxima"
    ];
    const datos = await LeerExcelDesdeArchivo(archivo, {
        hoja: 'BDCupos',
        columnasDeseadas: columnas,
        filasAsaltadas: 2
    });
        // Filtrar en la columna "Entidad" para que solo contenga "Fiduciaria"
    const datosFiltrados = datos.filter(fila => fila["Entidad"] === "Fiduciaria");
    // Normalizar filas
    const normalizados = datosFiltrados.map((fila: any) => {
        const filaNormalizada: any = {};

        for (const key of columnas) {
            const valor = (fila[key] ?? "").toString().trim();

            if (columnasNumericas.includes(key)) {
                // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
                const num = parseFloat(valor.replace(",", "."));
                filaNormalizada[key] = isNaN(num) ? null : num;
            } else {
                filaNormalizada[key] = valor.toUpperCase();
            }
        }

        return filaNormalizada;
    }
    );

    // Limpiar tabla en el backend
    const respClean = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'limpiartablacuposfidu' })
    });
    const resCleanJson = await respClean.json();
    if (resCleanJson.status !== 'ok') {
        throw new Error('El backend no respondió ok al limpiar la tabla de Cupos Fiduciarios.');
    }
    // Enviar por bloques
    const bloque = 1000;
    for (let i = 0; i < normalizados.length; i += bloque) {
        const lote = normalizados.slice(i, i + bloque);
        const respLote = await fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'carguecuposfidu', data: lote })
        });

        const resLoteJson = await respLote.json();
        if (resLoteJson.status !== 'ok') {
            console.error(`Error al enviar el lote ${i / bloque + 1}`);
            break;
        }
    }
}
async function cargueBaseCuposValores(archivo: File, urlBackend: string): Promise<void> {
    const columnas = [ "Entidad","Nombre_1","Cupo","Nemo","ISIN 1","Ocupación Máxima","OyD"];
    const columnasNumericas = [
        "Ocupación Máxima"
    ];
    const datos = await LeerExcelDesdeArchivo(archivo, {
        hoja: 'BDCupos',
        columnasDeseadas: columnas,
        filasAsaltadas: 2
    });
    // Filtrar en la columna "Entidad" para que solo contenga "Valores"
    const datosFiltrados = datos.filter(fila => fila["Entidad"] === "Valores");
    // Normalizar filas
    const normalizados = datosFiltrados.map((fila: any) => {
        const filaNormalizada: any = {};

        for (const key of columnas) {
            const valor = (fila[key] ?? "").toString().trim();

            if (columnasNumericas.includes(key)) {
                // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
                const num = parseFloat(valor.replace(",", "."));
                filaNormalizada[key] = isNaN(num) ? null : num;
            } else {
                filaNormalizada[key] = valor.toUpperCase();
            }
        }

        return filaNormalizada;
    }
    );
    // Agregar códigos Murex
    const datosConMurex = await AgregarCodigosMurexCuposValores(
        normalizados,
        urlBackend
    );
    
    // Limpiar tabla en el backend
    const respClean = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'limpiartablacuposvalores' })
    });
    const resCleanJson = await respClean.json();
    if (resCleanJson.status !== 'ok') {
        throw new Error('El backend no respondió ok al limpiar la tabla de Cupos Valores.');
    }
    // Enviar por bloques
    const bloque = 1000;
    for (let i = 0; i < datosConMurex.length; i += bloque) {
        const lote = datosConMurex.slice(i, i + bloque);
        const respLote = await fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'carguecuposvalores', data: lote })
        });

        const resLoteJson = await respLote.json();
        if (resLoteJson.status !== 'ok') {
            console.error(`Error al enviar el lote ${i / bloque + 1}`);
            break;
        }
    }    
}
//ahora implementare lo que es la logica para poder cargar las operaciones pendientes y poderlo registar
async function cargarOperacionesPendientesFidu(archivo: File, urlBackend: string): Promise<void> {
    const columnas = [
        "Moneda instrumento", "Tipo operacion", "Cantidad / Valor nominal",
        "Nombre portafolio", "Descripcion instrumento", "ISIN instrumento", "Fecha"
    ];
    const columnasNumericas = ["Cantidad / Valor nominal"];

    const datos = await LeerExcelDesdeArchivo(archivo, {
        hoja: 'Sheet1',
        columnasDeseadas: columnas,
        filasAsaltadas: 0
    });

    // Normalizar datos
    const normalizados = datos.map((fila: any) => {
        const filaNormalizada: any = {};
        for (const key of columnas) {
            const valor = (fila[key] ?? "").toString().trim();
            if (columnasNumericas.includes(key)) {
                const num = parseFloat(valor.replace(",", "."));
                filaNormalizada[key] = isNaN(num) ? 0 : num;
            } else {
                filaNormalizada[key] = valor.toUpperCase();
            }
        }
        return filaNormalizada;
    });

    // Traer datos de cruce (mercado y remanente)
    const response = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'traerdatoscrucependientesfidu' })
    });
    if (!response.ok) throw new Error('Error al consultar el backend para obtener los datos de cruce pendientes.');

    const datosCruce: any[] = (await response.json()).filter(
        (item: { ISIN?: string }) => item.ISIN && item.ISIN.trim() !== ""
    );

    const mapaCruce = new Map(datosCruce.map(item => [
        item.ISIN.toString().trim(),
        {
            "Vr Mercado Hoy Moneda Empresa": item.vrmercadohoymonedaempresa,
            "Nominal Remanente": item["Nominal Remanente"]
        }
    ]));

    // Enriquecer datos
    const datosConCruce = normalizados.map(fila => {
        const isin = fila["ISIN instrumento"]?.toString().trim();
        const datosAdicionales = isin && mapaCruce.has(isin)
            ? mapaCruce.get(isin)!
            : { "Vr Mercado Hoy Moneda Empresa": 0, "Nominal Remanente": 0 };

        return {
            ...fila,
            ...datosAdicionales
        };
    });

    // Agregar columna PRECIO
    const datosConPrecio = datosConCruce.map(fila => {
        const vrMercado = parseFloat(fila["Vr Mercado Hoy Moneda Empresa"]);
        const nominal = parseFloat(fila["Nominal Remanente"]);
        const precio = !isNaN(vrMercado) && !isNaN(nominal) && nominal !== 0
            ? parseFloat((vrMercado / nominal).toFixed(2))
            : 0;
        return { ...fila, "PRECIO": precio };
    });

    // Agregar columna Valor mercado
    const datosConValorMercado = datosConPrecio.map(fila => {
        const precio = parseFloat(fila["PRECIO"]);
        const cantidad = parseFloat(fila["Cantidad / Valor nominal"]);
        const valorMercado = !isNaN(precio) && !isNaN(cantidad)
            ? parseFloat((cantidad * precio).toFixed(2))
            : 0;
        return { ...fila, "Valor mercado": valorMercado };
    });

    // Ajustar signo de cantidad si es VENTA
    const datosFinales = datosConValorMercado.map(fila => {
        const tipoOperacion = fila["Tipo operacion"]?.toUpperCase();
        const cantidad = parseFloat(fila["Cantidad / Valor nominal"]);
        return {
            ...fila,
            "Cantidad / Valor nominal": tipoOperacion === "V" ? -Math.abs(cantidad) : Math.abs(cantidad),
            "Tipo operacion": tipoOperacion
        };
    });

    // Agregar Macro Activo
    const datosConMacroActivo = datosFinales.map(fila => {
        const moneda = fila["Moneda instrumento"];
        const macroActivo = moneda === "COP" ? "RV LOCAL" : "RV INTERNACIONAL";
        return { ...fila, "Macro Activo": macroActivo };
    });

    // Traer emisores del backend
    const responseEmisores = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'traerdatoscrucependientesfiud2' })
    });
    if (!responseEmisores.ok) throw new Error('Error al consultar el backend para obtener los emisores.');

    const datosEmisores: any[] = (await responseEmisores.json()).filter(
        (item: { ISIN?: string }) => item.ISIN && item.ISIN.trim() !== ""
    );

    const mapaEmisores = new Map(datosEmisores.map(item => [
        item.ISIN.toString().trim(),
        {
            emisor: item["Emisor / Contraparte"],
            emisorUnificado: item["Emisor Unificado"]
        }
    ]));

    const datosConEmisores = datosConMacroActivo.map(fila => {
        const isin = fila["ISIN instrumento"]?.toString().trim();
        const emisores = isin && mapaEmisores.has(isin)
            ? mapaEmisores.get(isin)!
            : { emisor: "", emisorUnificado: "" };
        return {
            ...fila,
            "Emisor / Contraparte": emisores.emisor,
            "Emisor Unificado": emisores.emisorUnificado
        };
    });

    // SALDO Macro Activo y ABA
    const datosConSaldos = datosConEmisores.map(fila => {
        const tipoOperacion = fila["Tipo operacion"];
        const valorMercado = parseFloat(fila["Valor mercado"]);
        const saldo = tipoOperacion === "V" ? -Math.abs(valorMercado) : Math.abs(valorMercado);
        return {
            ...fila,
            "SALDO Macro Activo": saldo,
            "SALDO ABA": saldo
        };
    });

    // Agregar Nemotécnico vacío
    const datosConNemotecnico = datosConSaldos.map(fila => ({
        ...fila,
        "Nemotécnico": ""
    }));

    // Filtrar columnas finales
    const columnasFinales = [
        "Nombre portafolio", "Descripcion instrumento", "Emisor / Contraparte",
        "Emisor Unificado", "Nemotécnico", "ISIN instrumento",
        "Macro Activo", "SALDO Macro Activo", "SALDO ABA", "Nominal Remanente"
    ];

    const datosFinalesConColumnas = datosConNemotecnico.map(fila => {
        const resultado: any = {};
        const isin = fila["ISIN instrumento"];
        for (const col of columnasFinales) {
            const valor = fila[col];
            resultado[col] = (isin && isin !== "")
                ? (valor ?? (typeof valor === "number" ? 0 : ""))
                : (typeof valor === "number" ? 0 : "");
        }
        return resultado;
    });

    if (datosFinalesConColumnas.length === 0) return;

    // Limpiar tabla en el backend
    const respClean = await fetch(urlBackend, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: 'limpiartablacargueoperacionespendientesfidu' })
    });
    const resCleanJson = await respClean.json();
    if (resCleanJson.status !== 'ok') {
        throw new Error('El backend no respondió ok al limpiar la tabla.');
    }

    // Enviar datos por bloques
    const bloque = 300;
    for (let i = 0; i < datosFinalesConColumnas.length; i += bloque) {
        const lote = datosFinalesConColumnas.slice(i, i + bloque);
        const respLote = await fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'cargueoperacionespendientesfidu', data: lote })
        });
        const resLoteJson = await respLote.json();
        if (resLoteJson.status !== 'ok') {
            console.error(`Error al enviar el lote ${i / bloque + 1}`);
            break;
        }
    }
}

// Exponer funciones al objeto window
window.LeerExcelDesdeArchivo = LeerExcelDesdeArchivo;
window.AgregarCodigosMurex = AgregarCodigosMurex;
window.CargarPortafoliosCRM = CargarPortafoliosCRM;
window.cargarConsolidadoFidu = cargarConsolidadoFidu;
window.AgregarCodigosMurexConsolidadoFidu = AgregarCodigosMurexConsolidadoFidu;
window.AgregarCodigosMurexCuposValores = AgregarCodigosMurexCuposValores;
window.cargarConsolidadoValores = cargarConsolidadoValores;
window.cargueBaseCuposFidu = cargueBaseCuposFidu;
window.cargueBaseCuposValores = cargueBaseCuposValores;
