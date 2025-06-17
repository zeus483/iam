export async function listarMacroActivosTrader(): Promise<string[]> {
  const params = new URLSearchParams();
  params.append('id','listarmacroactivostrader');
  const resp = await fetch('vistaTrader.asp',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:params.toString()});
  if(!resp.ok) throw new Error('Error consultando macroactivos');
  return resp.json();
}

export async function obtenerIntencionesTrader(mercado: string, fecha: string, pagina:number=1,cantidad:number=50): Promise<any[]> {
  const params = new URLSearchParams();
  params.append('id','obtenerintencionestrader');
  params.append('mercado',mercado);
  params.append('fecha',fecha);
  params.append('pagina',pagina.toString());
  params.append('cantidad',cantidad.toString());
  const resp = await fetch('vistaTrader.asp',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:params.toString()});
  if(!resp.ok) throw new Error('Error obteniendo datos');
  return resp.json();
}

