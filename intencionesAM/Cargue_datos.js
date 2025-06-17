"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// Función: Leer Excel desde archivo
function LeerExcelDesdeArchivo(archivo, opciones) {
    return __awaiter(this, void 0, void 0, function* () {
        var _a;
        const data = yield archivo.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const nombreHoja = opciones.hoja || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[nombreHoja];
        if (!worksheet) {
            throw new Error(`La hoja "${nombreHoja}" no existe en el archivo.`);
        }
        const datos = XLSX.utils.sheet_to_json(worksheet, {
            range: (_a = opciones.filasAsaltadas) !== null && _a !== void 0 ? _a : 0,
            defval: '',
        });
        const filtrado = datos.map((fila) => {
            const resultado = {};
            for (const col of opciones.columnasDeseadas) {
                resultado[col] = fila[col];
            }
            return resultado;
        });
        return filtrado;
    });
}
// Función: Agregar códigos Murex
function AgregarCodigosMurex(datosExcel, endpointBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const response = yield fetch(endpointBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'consultaoydmurex' })
        });
        if (!response.ok) {
            throw new Error('Error al consultar el backend para obtener el mapeo de códigos.');
        }
        const mapeo = yield response.json();
        const mapa = new Map(mapeo.map(item => [item.codigo_oyd.toString(), item.codigo_murex]));
        const resultado = datosExcel.map(fila => {
            var _a;
            let codMurex = '';
            if (fila["ADMINISTRADOR"] === "VALORES") {
                const codOyd = (_a = fila["CÓD. CONT. | OYD | PERSHING"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
                if (mapa.has(codOyd)) {
                    codMurex = mapa.get(codOyd) || '-';
                }
            }
            return Object.assign(Object.assign({}, fila), { "CÓDIGO MUREX": codMurex });
        });
        return resultado;
    });
}
//funcion agregar codigo murex para consolidado valores
function AgregarCodigosMurexConsolidadoFidu(datosExcel, endpointBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const response = yield fetch(endpointBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'consultaoydmurex' })
        });
        if (!response.ok) {
            throw new Error('Error al consultar el backend para obtener el mapeo de códigos.');
        }
        const mapeo = yield response.json();
        const mapa = new Map(mapeo.map(item => [item.codigo_oyd.toString(), item.codigo_murex]));
        const resultado = datosExcel.map(fila => {
            var _a;
            let codMurex = '';
            const codOyd = (_a = fila["Código OyD"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
            if (mapa.has(codOyd)) {
                codMurex = mapa.get(codOyd) || '-';
            }
            return Object.assign(Object.assign({}, fila), { "Portafolio": codMurex });
        });
        //ahora nos quedamos solo con los que en portafolio tengan dato valido
        const response2 = yield fetch(endpointBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'obtenerportafoliosvalorescrm' })
        });
        if (!response2.ok) {
            throw new Error('Error al consultar el backend para obtener los portafolios de valores.');
        }
        const portafoliosPermitidos = yield response2.json();
        const portafoliosSet = new Set(portafoliosPermitidos.map(p => p.toString().trim()));
        const resultado2 = resultado.filter(fila => {
            var _a;
            const portafolio = (_a = fila["Portafolio"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
            return portafolio && portafoliosSet.has(portafolio);
        });
        return resultado2;
    });
}
//Funcion para agregar la columna MUREX con los codigos de murex a base cupos valores
function AgregarCodigosMurexCuposValores(datosExcel, endpointBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const response = yield fetch(endpointBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'consultaoydmurex' })
        });
        if (!response.ok) {
            throw new Error('Error al consultar el backend para obtener el mapeo de códigos.');
        }
        const mapeo = yield response.json();
        const mapa = new Map(mapeo.map(item => [item.codigo_oyd.toString(), item.codigo_murex]));
        const resultado = datosExcel.map(fila => {
            var _a;
            let codMurex = '';
            const codOyd = (_a = fila["OyD"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
            if (mapa.has(codOyd)) {
                codMurex = mapa.get(codOyd) || '-';
            }
            return Object.assign(Object.assign({}, fila), { "MUREX": codMurex });
        });
        //ahora nos quedamos solo con los que en MUREX tengan dato valido
        const response2 = yield fetch(endpointBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'obtenerportafoliosvalorescrm' })
        });
        if (!response2.ok) {
            throw new Error('Error al consultar el backend para obtener los portafolios de valores.');
        }
        const portafoliosPermitidos = yield response2.json();
        const portafoliosSet = new Set(portafoliosPermitidos.map(p => p.toString().trim()));
        const resultado2 = resultado.filter(fila => {
            var _a;
            const portafolio = (_a = fila["MUREX"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
            return portafolio && portafoliosSet.has(portafolio);
        });
        return resultado2;
    });
}
// Función principal: flujo completo
function CargarPortafoliosCRM(archivo, urlBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const datos = yield LeerExcelDesdeArchivo(archivo, {
            hoja: 'Listado PD',
            columnasDeseadas: [
                'ADMINISTRADOR',
                'CÓD. CONT. | OYD | PERSHING',
                'NOMBRE PORTAFOLIO'
            ],
            filasAsaltadas: 4
        });
        const datosConMurex = yield AgregarCodigosMurex(datos, urlBackend);
        const respClean = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'limpiartablaportafoliocrm' })
        });
        const resCleanJson = yield respClean.json();
        if (resCleanJson.status !== 'ok') {
            throw new Error('El backend no respondió ok al limpiar la tabla.');
        }
        for (let i = 0; i < datosConMurex.length; i += 100) {
            const lote = datosConMurex.slice(i, i + 100);
            const respLote = yield fetch(urlBackend, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: 'carguecrm', data: lote })
            });
            const resLoteJson = yield respLote.json();
            if (resLoteJson.status !== 'ok') {
                console.error(`Error al enviar el lote ${i / 100 + 1}`);
                break;
            }
        }
        console.log('Carga finalizada con éxito.');
    });
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
function cargarConsolidadoFidu(archivo, urlBackend) {
    return __awaiter(this, void 0, void 0, function* () {
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
        const datos = yield LeerExcelDesdeArchivo(archivo, {
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
        const normalizados = datos.map((fila) => {
            var _a;
            const filaNormalizada = {};
            for (const key of columnas) {
                const valor = ((_a = fila[key]) !== null && _a !== void 0 ? _a : "").toString().trim();
                if (columnasNumericas.includes(key)) {
                    const num = parseFloat(valor.replace(",", "."));
                    filaNormalizada[key] = isNaN(num) ? null : num;
                }
                else {
                    filaNormalizada[key] = valor.toUpperCase();
                }
            }
            filaNormalizada["origen informacion"] = "INVENTARIO TITULOS";
            return filaNormalizada;
        });
        // Limpiar tabla
        const respClean = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'limpiartablaconsolidadofidu' })
        });
        const resCleanJson = yield respClean.json();
        if (resCleanJson.status !== 'ok') {
            throw new Error('El backend no respondió ok al limpiar Consolidado Fidu.');
        }
        // Envío por bloques con reintento
        const bloque = 1000;
        for (let i = 0; i < normalizados.length; i += bloque) {
            const lote = normalizados.slice(i, i + bloque);
            const exito = yield intentarEnvioLote(lote, urlBackend, i / bloque + 1);
            if (!exito)
                break;
        }
        console.log('Carga Consolidado Fidu finalizada con éxito.');
    });
}
// Función auxiliar para reintentos
function intentarEnvioLote(lote_1, urlBackend_1, numeroLote_1) {
    return __awaiter(this, arguments, void 0, function* (lote, urlBackend, numeroLote, intentos = 3) {
        for (let intento = 1; intento <= intentos; intento++) {
            try {
                const resp = yield fetch(urlBackend, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ id: 'cargueconsolidadofidu', data: lote })
                });
                const json = yield resp.json();
                if (json.status === 'ok') {
                    console.log(`✅ Lote ${numeroLote} enviado correctamente.`);
                    return true;
                }
                else {
                    throw new Error(`Backend respondió: ${json.status}`);
                }
            }
            catch (e) {
                console.warn(`⚠️ Error al enviar lote ${numeroLote}, intento ${intento}: ${e}`);
                yield new Promise(r => setTimeout(r, 1000 * intento)); // backoff creciente
            }
        }
        console.error(`❌ No se pudo enviar el lote ${numeroLote} después de ${intentos} intentos.`);
        return false;
    });
}
function cargarConsolidadoValores(archivo, urlBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        console.log('Iniciando carga de Consolidado Valores...');
        const columnas = [
            "Macro Activo", "Isin", "Nemoténico", "Emisor Unificado", "Nombre Emisor", "SALDO Macro Activo", "SALDO ABA", "Valor Nominal Actual", "Codigo OyD", "Fecha"
        ];
        const columnasNumericas = [
            "SALDO Macro Activo",
            "SALDO ABA",
            "Valor Nominal Actual"
        ];
        const datos = yield LeerExcelDesdeArchivo(archivo, {
            hoja: 'Reporte_Consolidado',
            columnasDeseadas: columnas,
            filasAsaltadas: 8
        });
        // Normalizar filas
        const normalizados = datos.map((fila) => {
            var _a;
            const filaNormalizada = {};
            for (const key of columnas) {
                const valor = ((_a = fila[key]) !== null && _a !== void 0 ? _a : "").toString().trim();
                if (columnasNumericas.includes(key)) {
                    // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
                    const num = parseFloat(valor.replace(",", "."));
                    filaNormalizada[key] = isNaN(num) ? null : num;
                }
                else {
                    filaNormalizada[key] = valor.toUpperCase();
                }
            }
            filaNormalizada["origen informacion"] = "INVENTARIO TITULOS";
            return filaNormalizada;
        });
        //Agregar lo de murex cod
        const datosConMurex = yield AgregarCodigosMurexConsolidadoFidu(normalizados, urlBackend);
        const respClean = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'limpiartablaconsolidadovalores' })
        });
        const resCleanJson = yield respClean.json();
        if (resCleanJson.status !== 'ok') {
            throw new Error('El backend no respondió ok al limpiar la tabla de Consolidado Valores.');
        }
        for (let i = 0; i < datosConMurex.length; i += 1000) {
            const lote = datosConMurex.slice(i, i + 1000);
            const respLote = yield fetch(urlBackend, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: 'cargueconsolidadovalores', data: lote })
            });
            const resLoteJson = yield respLote.json();
            if (resLoteJson.status !== 'ok') {
                console.error(`Error al enviar el lote ${i / 1000 + 1}`);
                break;
            }
        }
        console.log('Carga Consolidado Valores finalizada con éxito.');
    });
}
function cargueBaseCuposFidu(archivo, urlBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const columnas = [
            "Entidad", "MUREX", "Nombre_1", "Cupo", "Nemo", "ISIN 1", "Ocupación Máxima"
        ];
        const columnasNumericas = [
            "Ocupación Máxima"
        ];
        const datos = yield LeerExcelDesdeArchivo(archivo, {
            hoja: 'BDCupos',
            columnasDeseadas: columnas,
            filasAsaltadas: 2
        });
        // Filtrar en la columna "Entidad" para que solo contenga "Fiduciaria"
        const datosFiltrados = datos.filter(fila => fila["Entidad"] === "Fiduciaria");
        // Normalizar filas
        const normalizados = datosFiltrados.map((fila) => {
            var _a;
            const filaNormalizada = {};
            for (const key of columnas) {
                const valor = ((_a = fila[key]) !== null && _a !== void 0 ? _a : "").toString().trim();
                if (columnasNumericas.includes(key)) {
                    // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
                    const num = parseFloat(valor.replace(",", "."));
                    filaNormalizada[key] = isNaN(num) ? null : num;
                }
                else {
                    filaNormalizada[key] = valor.toUpperCase();
                }
            }
            return filaNormalizada;
        });
        // Limpiar tabla en el backend
        const respClean = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'limpiartablacuposfidu' })
        });
        const resCleanJson = yield respClean.json();
        if (resCleanJson.status !== 'ok') {
            throw new Error('El backend no respondió ok al limpiar la tabla de Cupos Fiduciarios.');
        }
        // Enviar por bloques
        const bloque = 1000;
        for (let i = 0; i < normalizados.length; i += bloque) {
            const lote = normalizados.slice(i, i + bloque);
            const respLote = yield fetch(urlBackend, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: 'carguecuposfidu', data: lote })
            });
            const resLoteJson = yield respLote.json();
            if (resLoteJson.status !== 'ok') {
                console.error(`Error al enviar el lote ${i / bloque + 1}`);
                break;
            }
        }
    });
}
function cargueBaseCuposValores(archivo, urlBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const columnas = ["Entidad", "Nombre_1", "Cupo", "Nemo", "ISIN 1", "Ocupación Máxima", "OyD"];
        const columnasNumericas = [
            "Ocupación Máxima"
        ];
        const datos = yield LeerExcelDesdeArchivo(archivo, {
            hoja: 'BDCupos',
            columnasDeseadas: columnas,
            filasAsaltadas: 2
        });
        // Filtrar en la columna "Entidad" para que solo contenga "Valores"
        const datosFiltrados = datos.filter(fila => fila["Entidad"] === "Valores");
        // Normalizar filas
        const normalizados = datosFiltrados.map((fila) => {
            var _a;
            const filaNormalizada = {};
            for (const key of columnas) {
                const valor = ((_a = fila[key]) !== null && _a !== void 0 ? _a : "").toString().trim();
                if (columnasNumericas.includes(key)) {
                    // Convertir a número flotante (reemplaza comas por puntos si vienen en formato LATAM)
                    const num = parseFloat(valor.replace(",", "."));
                    filaNormalizada[key] = isNaN(num) ? null : num;
                }
                else {
                    filaNormalizada[key] = valor.toUpperCase();
                }
            }
            return filaNormalizada;
        });
        // Agregar códigos Murex
        const datosConMurex = yield AgregarCodigosMurexCuposValores(normalizados, urlBackend);
        // Limpiar tabla en el backend
        const respClean = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'limpiartablacuposvalores' })
        });
        const resCleanJson = yield respClean.json();
        if (resCleanJson.status !== 'ok') {
            throw new Error('El backend no respondió ok al limpiar la tabla de Cupos Valores.');
        }
        // Enviar por bloques
        const bloque = 1000;
        for (let i = 0; i < datosConMurex.length; i += bloque) {
            const lote = datosConMurex.slice(i, i + bloque);
            const respLote = yield fetch(urlBackend, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: 'carguecuposvalores', data: lote })
            });
            const resLoteJson = yield respLote.json();
            if (resLoteJson.status !== 'ok') {
                console.error(`Error al enviar el lote ${i / bloque + 1}`);
                break;
            }
        }
    });
}
//ahora implementare lo que es la logica para poder cargar las operaciones pendientes y poderlo registar
function cargarOperacionesPendientesFidu(archivo, urlBackend) {
    return __awaiter(this, void 0, void 0, function* () {
        const columnas = [
            "Moneda instrumento", "Tipo operacion", "Cantidad / Valor nominal",
            "Nombre portafolio", "Descripcion instrumento", "ISIN instrumento", "Fecha"
        ];
        const columnasNumericas = ["Cantidad / Valor nominal"];
        const datos = yield LeerExcelDesdeArchivo(archivo, {
            hoja: 'Sheet1',
            columnasDeseadas: columnas,
            filasAsaltadas: 0
        });
        // Normalizar datos
        const normalizados = datos.map((fila) => {
            var _a;
            const filaNormalizada = {};
            for (const key of columnas) {
                const valor = ((_a = fila[key]) !== null && _a !== void 0 ? _a : "").toString().trim();
                if (columnasNumericas.includes(key)) {
                    const num = parseFloat(valor.replace(",", "."));
                    filaNormalizada[key] = isNaN(num) ? 0 : num;
                }
                else {
                    filaNormalizada[key] = valor.toUpperCase();
                }
            }
            return filaNormalizada;
        });
        // Traer datos de cruce (mercado y remanente)
        const response = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'traerdatoscrucependientesfidu' })
        });
        if (!response.ok)
            throw new Error('Error al consultar el backend para obtener los datos de cruce pendientes.');
        const datosCruce = (yield response.json()).filter((item) => item.ISIN && item.ISIN.trim() !== "");
        const mapaCruce = new Map(datosCruce.map(item => [
            item.ISIN.toString().trim(),
            {
                "Vr Mercado Hoy Moneda Empresa": item.vrmercadohoymonedaempresa,
                "Nominal Remanente": item["Nominal Remanente"]
            }
        ]));
        // Enriquecer datos
        const datosConCruce = normalizados.map(fila => {
            var _a;
            const isin = (_a = fila["ISIN instrumento"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
            const datosAdicionales = isin && mapaCruce.has(isin)
                ? mapaCruce.get(isin)
                : { "Vr Mercado Hoy Moneda Empresa": 0, "Nominal Remanente": 0 };
            return Object.assign(Object.assign({}, fila), datosAdicionales);
        });
        // Agregar columna PRECIO
        const datosConPrecio = datosConCruce.map(fila => {
            const vrMercado = parseFloat(fila["Vr Mercado Hoy Moneda Empresa"]);
            const nominal = parseFloat(fila["Nominal Remanente"]);
            const precio = !isNaN(vrMercado) && !isNaN(nominal) && nominal !== 0
                ? parseFloat((vrMercado / nominal).toFixed(2))
                : 0;
            return Object.assign(Object.assign({}, fila), { "PRECIO": precio });
        });
        // Agregar columna Valor mercado
        const datosConValorMercado = datosConPrecio.map(fila => {
            const precio = parseFloat(fila["PRECIO"]);
            const cantidad = parseFloat(fila["Cantidad / Valor nominal"]);
            const valorMercado = !isNaN(precio) && !isNaN(cantidad)
                ? parseFloat((cantidad * precio).toFixed(2))
                : 0;
            return Object.assign(Object.assign({}, fila), { "Valor mercado": valorMercado });
        });
        // Ajustar signo de cantidad si es VENTA
        const datosFinales = datosConValorMercado.map(fila => {
            var _a;
            const tipoOperacion = (_a = fila["Tipo operacion"]) === null || _a === void 0 ? void 0 : _a.toUpperCase();
            const cantidad = parseFloat(fila["Cantidad / Valor nominal"]);
            return Object.assign(Object.assign({}, fila), { "Cantidad / Valor nominal": tipoOperacion === "V" ? -Math.abs(cantidad) : Math.abs(cantidad), "Tipo operacion": tipoOperacion });
        });
        // Agregar Macro Activo
        const datosConMacroActivo = datosFinales.map(fila => {
            const moneda = fila["Moneda instrumento"];
            const macroActivo = moneda === "COP" ? "RV LOCAL" : "RV INTERNACIONAL";
            return Object.assign(Object.assign({}, fila), { "Macro Activo": macroActivo });
        });
        // Traer emisores del backend
        const responseEmisores = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'traerdatoscrucependientesfiud2' })
        });
        if (!responseEmisores.ok)
            throw new Error('Error al consultar el backend para obtener los emisores.');
        const datosEmisores = (yield responseEmisores.json()).filter((item) => item.ISIN && item.ISIN.trim() !== "");
        const mapaEmisores = new Map(datosEmisores.map(item => [
            item.ISIN.toString().trim(),
            {
                emisor: item["Emisor / Contraparte"],
                emisorUnificado: item["Emisor Unificado"]
            }
        ]));
        const datosConEmisores = datosConMacroActivo.map(fila => {
            var _a;
            const isin = (_a = fila["ISIN instrumento"]) === null || _a === void 0 ? void 0 : _a.toString().trim();
            const emisores = isin && mapaEmisores.has(isin)
                ? mapaEmisores.get(isin)
                : { emisor: "", emisorUnificado: "" };
            return Object.assign(Object.assign({}, fila), { "Emisor / Contraparte": emisores.emisor, "Emisor Unificado": emisores.emisorUnificado });
        });
        // SALDO Macro Activo y ABA
        const datosConSaldos = datosConEmisores.map(fila => {
            const tipoOperacion = fila["Tipo operacion"];
            const valorMercado = parseFloat(fila["Valor mercado"]);
            const saldo = tipoOperacion === "V" ? -Math.abs(valorMercado) : Math.abs(valorMercado);
            return Object.assign(Object.assign({}, fila), { "SALDO Macro Activo": saldo, "SALDO ABA": saldo });
        });
        // Agregar Nemotécnico vacío
        const datosConNemotecnico = datosConSaldos.map(fila => (Object.assign(Object.assign({}, fila), { "Nemotécnico": "" })));
        // Filtrar columnas finales
        const columnasFinales = [
            "Nombre portafolio", "Descripcion instrumento", "Emisor / Contraparte",
            "Emisor Unificado", "Nemotécnico", "ISIN instrumento",
            "Macro Activo", "SALDO Macro Activo", "SALDO ABA", "Nominal Remanente"
        ];
        const datosFinalesConColumnas = datosConNemotecnico.map(fila => {
            const resultado = {};
            const isin = fila["ISIN instrumento"];
            for (const col of columnasFinales) {
                const valor = fila[col];
                resultado[col] = (isin && isin !== "")
                    ? (valor !== null && valor !== void 0 ? valor : (typeof valor === "number" ? 0 : ""))
                    : (typeof valor === "number" ? 0 : "");
            }
            return resultado;
        });
        if (datosFinalesConColumnas.length === 0)
            return;
        // Limpiar tabla en el backend
        const respClean = yield fetch(urlBackend, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: 'limpiartablacargueoperacionespendientesfidu' })
        });
        const resCleanJson = yield respClean.json();
        if (resCleanJson.status !== 'ok') {
            throw new Error('El backend no respondió ok al limpiar la tabla.');
        }
        // Enviar datos por bloques
        const bloque = 300;
        for (let i = 0; i < datosFinalesConColumnas.length; i += bloque) {
            const lote = datosFinalesConColumnas.slice(i, i + bloque);
            const respLote = yield fetch(urlBackend, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: 'cargueoperacionespendientesfidu', data: lote })
            });
            const resLoteJson = yield respLote.json();
            if (resLoteJson.status !== 'ok') {
                console.error(`Error al enviar el lote ${i / bloque + 1}`);
                break;
            }
        }
    });
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

