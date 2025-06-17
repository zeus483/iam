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
'Conexion a la BD
Dim conn, rs,sql,firstItem,json
Set conn = ConectarBD()
'Leemos la accion del usuario
Dim action 
action = LCase(Trim(Request.Form("id")))
If IsNull(Request.Form("id")) Then
    Response.Status = "400 Bad Request"
    Response.Write "{""error"":""No se recibió parámetro 'id'""}"
    Response.End
End If
Function SanitizeSql(input)
    SanitizeSql = Replace(input, "'", "''") ' Doble comillas simples para SQL
End Function
'Segmentamos en case para cada funcionalidad
Select Case action
    Case "listargerentesportafolios"
        sql = "SELECT Usuario FROM Usuarios WHERE Perfil = 8"
        Set rs = conn.Execute(sql)
        Response.Write "["
        firstItem = True
        Do While Not rs.EOF
            If Not firstItem Then Response.Write ","
            firstItem = False
            Response.Write "{""Usuario"":""" &_
                Replace(rs("Usuario"), """","\""") & """}"
            rs.MoveNext
        Loop
        Response.Write "]"
        rs.Close: Set rs = Nothing
        conn.Close: Set conn= Nothing
    Case "obtenerintencionesgerentesparavisualizar"
        Dim fecha, gerente, cmd, pagina, cantidad, offset
        'Obtenemos los parametros de la peticion
        fecha = Request.Form("fecha")
        pagina = Request.Form("pagina")
        cantidad = Request.Form("cantidad")
        If IsEmpty(pagina) Or Not IsNumeric(pagina) Or pagina < 1 Then pagina = 1
        If IsEmpty(cantidad) Or Not IsNumeric(cantidad) Or cantidad < 1 Then cantidad = 20

        offset = (pagina - 1) * cantidad
        gerente = SanitizeSql(Request.Form("gerente"))
        'cambiar por la vista completa en lugar de la limitada en produccion
        sql = "SELECT * FROM vista_intenciones_completa WHERE " & _
            "(CAST(""UltimaModificacion"" AS date) >= '" & fecha & "' " & _
            "OR (""Estado"" NOT IN ('Ejecutada/Total', 'Cancelada') AND CAST(""VigenteHasta"" AS date) >= '" & fecha & "') " & _
            "OR ""Estado"" = 'Vencida')"

        If Not IsEmpty(gerente) And Len(Trim(gerente)) > 0 Then
            sql = sql & " AND LOWER(""Trader"") = LOWER('" & gerente & "')"
        End If
        sql = sql & " ORDER BY ""Id"" DESC LIMIT " & cantidad & " OFFSET " & offset & ";"
        'sql = sql & " ORDER BY ""Id"" DESC;"
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.CommandTimeout = 300 'en segundos
        Set rs = cmd.Execute
        'Set rs = conn.Execute(sql)
        
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
    Case "listarportafolios"
        sql = "SELECT DISTINCT ""CÓD MUREX"" AS IdPortafolio, ""NOMBRE PORTAFOLIO"" AS Nombre " & _
            "FROM intportafolios WHERE ""ADMINISTRADOR"" IN ('FIDUCIARIA', 'VALORES') ORDER BY ""CÓD MUREX"";"
        Set rs = conn.Execute(sql)
        Response.Write "["
        firstItem = True
        Do While Not rs.EOF
            If Not firstItem Then Response.Write "," Else firstItem = False
            Response.Write "{""IdPortafolio"":""" & rs("IdPortafolio") & """,""Nombre"":""" & Replace(rs("Nombre"), """", "\""") & """}"
            rs.MoveNext
        Loop
        Response.Write "]"
        rs.Close: Set rs = Nothing
        conn.Close: Set conn = Nothing

    Case "listarespeciesfidu"
        Dim idPortaFidu
        idPortaFidu = Request.Form("id_portafolio")
        
        If idPortaFidu = "" Then
            Response.Write "[]"
            Response.End
        End If

        sql = "SELECT DISTINCT ""Especie/Generador"" AS Especie " & _
            "FROM consolidadofidu " & _
            "WHERE ""Portafolio"" = '" & Replace(idPortaFidu, "'", "''") & "' " & _
            "AND ""Especie/Generador"" IS NOT NULL " & _
            "ORDER BY ""Especie/Generador"";"
        'Response.Write sql
        Set rs = conn.Execute(sql)
        
        Response.Write "["
        firstItem = True
        Do While Not rs.EOF
            If Not firstItem Then Response.Write "," Else firstItem = False
            Response.Write "{""Especie"":""" & Replace(rs("Especie"), """", "\""") & """}"
            rs.MoveNext
        Loop
        Response.Write "]"
        rs.Close: Set rs = Nothing
        conn.Close: Set conn = Nothing
    Case "listarnemosfiduporespecie"
        Dim especie
        especie = Request.Form("especie")
        
        If especie = "" Then
            Response.Write "[]"
            Response.End
        End If

        sql = "SELECT DISTINCT ""Nemotécnico"" AS Nemo " & _
            "FROM consolidadofidu " & _
            "WHERE ""Especie/Generador"" = '" & Replace(especie, "'", "''") & "' " & _
            "AND ""Nemotécnico"" IS NOT NULL " & _
            "ORDER BY ""Nemotécnico"";"

        Set rs = conn.Execute(sql)
        Response.Write "["
        firstItem = True
        Do While Not rs.EOF
            If Not firstItem Then Response.Write "," Else firstItem = False
            Response.Write "{""Nemo"":""" & Replace(rs("Nemo"), """", "\""") & """}"
            rs.MoveNext
        Loop
        Response.Write "]"
        rs.Close: Set rs = Nothing
        conn.Close: Set conn = Nothing
    Case "listarnemosfidusinespecie"
        Dim idPorta
        idPorta = Request.Form("id_portafolio")
        
        If idPorta = "" Then
            Response.Write "[]"
            Response.End
        End If

        sql = "SELECT DISTINCT ""Nemotécnico"" AS Nemo " & _
            "FROM consolidadofidu " & _
            "WHERE ""Portafolio"" = '" & Replace(idPorta, "'", "''") & "' " & _
            "AND ""Nemotécnico"" IS NOT NULL " & _
            "ORDER BY ""Nemotécnico"";"

        Set rs = conn.Execute(sql)
        Response.Write "["
        firstItem = True
        Do While Not rs.EOF
            If Not firstItem Then Response.Write "," Else firstItem = False
            Response.Write "{""Nemo"":""" & Replace(rs("Nemo"), """", "\""") & """}"
            rs.MoveNext
        Loop
        Response.Write "]"
        rs.Close: Set rs = Nothing
        conn.Close: Set conn = Nothing

    Case Else
        Response.Status = "400 Bad Request"
        Response.Write "{""error"":""Acción desconocida: " & action & """}"
    End Select
Session.Abandon
%>