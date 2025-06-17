<!--#include file="Jason2ASP/JSON_UTIL_0.1.1.asp"-->
<!--#include file="Jason2ASP/jsonObject.class.asp"-->
<!--#include file="Jason2ASP/JSON_2.0.4.asp"-->
<!--#include file="conexion.asp"-->
<%
Response.ContentType = "application/json"
Session.LCID = 2057
Response.Buffer = True
Response.Clear
Server.ScriptTimeout = 600

Dim conn, rs, sql, firstItem, json
Set conn = ConectarBD()

Dim action
action = LCase(Trim(Request.Form("id")))
If IsNull(Request.Form("id")) Then
    Response.Status = "400 Bad Request"
    Response.Write "{\"error\":\"No se recibió parámetro 'id'\"}"
    Response.End
End If

Function SanitizeSql(input)
    SanitizeSql = Replace(input, "'", "''")
End Function

Select Case action
    Case "listarmacroactivostrader"
        json = "[""Deuda Privada"",""Deuda Pública"",""RF Internacional"",""Forex"",""RV y Fondos"",""Fondos"",""Swaps"",""Todos""]"
        Response.Write json
        conn.Close: Set conn = Nothing

    Case "obtenerintencionestrader"
        Dim mercado, fecha, pagina, cantidad, offset, cmd
        mercado = SanitizeSql(Request.Form("mercado"))
        fecha = Request.Form("fecha")
        pagina = Request.Form("pagina")
        cantidad = Request.Form("cantidad")
        If IsEmpty(pagina) Or Not IsNumeric(pagina) Or pagina < 1 Then pagina = 1
        If IsEmpty(cantidad) Or Not IsNumeric(cantidad) Or cantidad < 1 Then cantidad = 50
        offset = (pagina - 1) * cantidad
        sql = "SELECT * FROM vista_intenciones_completa WHERE " & _
              "(CAST(""UltimaModificacion"" AS date) >= '" & fecha & "' " & _
              "OR (""Estado"" NOT IN ('Ejecutada/Total', 'Cancelada') AND CAST(""VigenteHasta"" AS date) >= '" & fecha & "') " & _
              "OR ""Estado"" = 'Vencida')"

        If Len(mercado) > 0 Then
            Select Case LCase(mercado)
                Case "rv y fondos"
                    sql = sql & " AND ""TipoActivo"" IN ('Renta Variable','Fondos') " & _
                          "AND ""Mercado"" IN ('LOCAL','INTERNACIONAL','FONDO MUTUO')"
                Case "deuda privada"
                    sql = sql & " AND ""TipoActivo"" = 'Deuda Privada' AND ""Mercado"" = 'LOCAL'"
                Case "deuda pública"
                    sql = sql & " AND ""TipoActivo"" = 'Deuda Pública' AND ""Mercado"" = 'LOCAL'"
                Case "fondos"
                    sql = sql & " AND ""TipoActivo"" = 'Fondos' AND ""Mercado"" = 'FIC'"
                Case "forex"
                    sql = sql & " AND ""TipoActivo"" = 'Forex'"
                Case "swaps"
                    sql = sql & " AND ""TipoActivo"" = 'Swaps'"
                Case "rf internacional"
                    sql = sql & " AND ""TipoActivo"" IN ('Deuda Privada','Deuda Pública') AND ""Mercado"" = 'INTERNACIONAL'"
                Case "todos"
                    sql = sql
            End Select
       End If

        sql = sql & " AND CAST(""VigenciaDesde"" AS date) <= CURRENT_DATE" & _
              " ORDER BY ""Id"" DESC LIMIT " & cantidad & " OFFSET " & offset & ";"

        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.CommandTimeout = 300
        Set rs = cmd.Execute

        json = "["
        firstItem = True
        Do While Not rs.EOF
            If Not firstItem Then
                json = json & ","
            Else
                firstItem = False
            End If
            json = json & "{" & _
                """Id"":""" & rs("Id") & """," & _
                """FechaIngreso"":""" & rs("FechaIngreso") & """," & _
                """IngresadoPor"":""" & rs("IngresadoPor") & """," & _
                """UltimaModificacion"":""" & rs("UltimaModificacion") & """," & _
                """ModificadoPor"":""" & rs("ModificadoPor") & """," & _
                """TipoActivo"":""" & rs("TipoActivo") & """," & _
                """Mercado"":""" & rs("Mercado") & """," & _
                """CodPortafolio"":""" & rs("CodPortafolio") & """," & _
                """Portafolio"":""" & rs("Portafolio") & """," & _
                """TipoOperacion"":""" & rs("TipoOperacion") & """," & _
                """TipoOrden"":""" & rs("TipoOrden") & """," & _
                """Emisor"":""" & rs("Emisor") & """," & _
                """Nemotecnico"":""" & rs("Nemotecnico") & """," & _
                """Indicador"":""" & rs("Indicador") & """," & _
                """Denominacion"":""" & rs("Denominacion") & """," & _
                """Desde"":""" & rs("Desde") & """," & _
                """Hasta"":""" & rs("Hasta") & """," & _
                """PrecioLimite"":""" & rs("PrecioLimite") & """," & _
                """TasaLimite"":""" & rs("TasaLimite") & """," & _
                """VigenciaDesde"":""" & rs("VigenciaDesde") & """," & _
                """VigenteHasta"":""" & rs("VigenteHasta") & """," & _
                """ComentariosPM"":""" & rs("ComentariosPM") & """," & _
                """Estado"":""" & rs("Estado") & """," & _
                """Trader"":""" & rs("Trader") & """," & _
                """CantEjecutada"":""" & rs("CantEjecutada") & """," & _
                """CantidadTotal"":""" & rs("CantidadTotal") & """," & _
                """CantPendiente"":""" & rs("CantPendiente") & """," & _
                """Ejecutado"":""" & rs("Ejecutado") & """," & _
                """ComentariosTrader"":""" & rs("ComentariosTrader") & """" & _
            "}"
            rs.MoveNext
        Loop
        json = json & "]"
        Response.Write json
        rs.Close: Set rs = Nothing
        conn.Close: Set conn = Nothing

    Case Else
        Response.Status = "400 Bad Request"
        Response.Write "{\"error\":\"Acción desconocida: " & action & "\"}"
End Select
Session.Abandon
%>
