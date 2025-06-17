export async function listarMacroActivosTrader() {
    const params = new URLSearchParams();
    params.append('id', 'listarmacroactivostrader');
    const resp = await fetch('vistaTrader.asp', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: params.toString() });
    if (!resp.ok)
        throw new Error('Error consultando macroactivos');
    return resp.json();
}
export async function obtenerIntencionesTrader(mercado, fecha, pagina = 1, cantidad = 50) {
    const params = new URLSearchParams();
    params.append('id', 'obtenerintencionestrader');
    params.append('mercado', mercado);
    params.append('fecha', fecha);
    params.append('pagina', pagina.toString());
    params.append('cantidad', cantidad.toString());
    const resp = await fetch('vistaTrader.asp', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: params.toString() });
    if (!resp.ok)
        throw new Error('Error obteniendo datos');
    return resp.json();
}
