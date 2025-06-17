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

Dim conn, rs, sql, i, item, rawJson, payload, id, data, responseText,murex
Set conn = Nothing
Set rs = Nothing
responseText = "[]"

On Error Resume Next

' Crear objeto JSON
Set JSON = New JSONobject

' Función para decodificar el cuerpo JSON enviado por fetch
Function BinaryToText(bin, charset)
    Dim stream: Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 'binary
    stream.Open
    stream.Write bin
    stream.Position = 0
    stream.Type = 2 'text
    stream.Charset = charset
    BinaryToText = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function

' Leer el JSON del cuerpo
rawJson = BinaryToText(Request.BinaryRead(Request.TotalBytes), "utf-8")
Set payload = JSON.parse(rawJson)
id = LCase(payload("id"))

Set conn = ConectarBD()
If id = "consultaoydmurex" Then
    sql = "SELECT codigo_oyd, codigo_murex FROM murextooyd WHERE codigo_oyd IS NOT NULL AND codigo_murex IS NOT NULL;"
    Set rs = conn.Execute(sql)

    responseText = "["
    Dim firstItem: firstItem = True

    Do While Not rs.EOF
        If Not firstItem Then responseText = responseText & "," Else firstItem = False
        responseText = responseText & "{" & _
            """codigo_oyd"":""" & Replace(rs("codigo_oyd"), """", "\""") & """," & _
            """codigo_murex"":""" & Replace(rs("codigo_murex"), """", "\""") & """" & _
        "}"
        rs.MoveNext
    Loop
    responseText = responseText & "]"

ElseIf id = "limpiartablaportafoliocrm" Then
    sql = "DELETE FROM portafolioscrm;"
    conn.Execute(sql)
    responseText = "{""status"":""ok""}"

ElseIf id = "carguecrm" Then
    Set data = payload("data")
    If IsEmpty(data) Or data.length = 0 Then
        responseText = "{""status"":""sin datos""}"
    Else
        For i = 0 To data.length - 1
            Set item = data(i)
            If IsNull(item("CÓDIGO MUREX")) Or item("CÓD MUREX") = "" Then
                murex = "-"
            Else
                murex = item("CÓDIGO MUREX")
            End If
            sql = "INSERT INTO portafolioscrm (""administrador"", ""oyd_codigo"", ""codigo_murex"", ""nombre_portafolio"") VALUES (" & _
                "'" & Replace(item("ADMINISTRADOR"), "'", "''") & "'," & _
                "'" & Replace(item("CÓD. CONT. | OYD | PERSHING"), "'", "''") & "'," & _
                "'" & Replace(murex, "'", "''") & "'," & _
                "'" & Replace(item("NOMBRE PORTAFOLIO"), "'", "''") & "'" & _
            ");"
            conn.Execute(sql)
            If Err.Number <> 0 Then
                responseText = "{""status"":""error"",""fila"":" & (i+1) & ",""detalle"":""" & Replace(Err.Description, """", "'") & """}"
                Exit For
            End If
        Next
        If Left(responseText, 1) <> "{" Then
            responseText = "{""status"":""ok""}"
        End If
    End If
ElseIf id = "limpiartablaconsolidadofidu" Then
    sql = "DELETE FROM nemotitulosfiduciaria WHERE ""Origen Informacion"" = 'INVENTARIO TITULOS';"
    conn.Execute(sql)
    responseText = "{""status"":""ok""}"
ElseIf id = "cargueconsolidadofidu" Then
    
    On Error Resume Next
    Set data = payload("data")
    If IsEmpty(data) Or data.length = 0 Then
        responseText = "{""status"":""sin datos""}"
    Else
        For i = 0 To data.length - 1
            Set item = data(i)
            sql = "INSERT INTO nemotitulosfiduciaria (""Macro Activo"", ""ISIN"", ""vrmercadohoymonedaempresa"", ""Emisor / Contraparte"", ""Especie/Generador"", ""Nemotécnico"", ""Emisor Unificado"", ""SALDO Macro Activo"", ""SALDO ABA"", ""Nominal Remanente"", ""Origen Informacion"", ""Portafolio"") VALUES (" & _
                "'" & Replace(item("Macro Activo"), "'", "''") & "'," & _
                "'" & Replace(item("ISIN"), "'", "''") & "'," & _
                item("Vr Mercado Hoy Moneda Empresa") & "," & _
                "'" & Replace(item("Emisor / Contraparte"), "'", "''") & "'," & _
                "'" & Replace(item("Especie/Generador"), "'", "''") & "'," & _
                "'" & Replace(item("Nemotécnico"), "'", "''") & "'," & _
                "'" & Replace(item("Emisor Unificado"), "'", "''") & "'," & _
                item("SALDO Macro Activo") & "," & _
                item("SALDO ABA") & "," & _
                item("Nominal Remanente") & "," & _
                "'INVENTARIO TITULOS'," & _
                "'" & Replace(item("Portafolio"), "'", "''") & "'" & _
            ");"
            conn.Execute(sql)
            If Err.Number <> 0 Then
                responseText = "{""status"":""error"",""fila"":" & (i+1) & ",""detalle"":""" & Replace(Err.Description, """", "'") & """}"
                Exit For
            End If
        Next

        ' Siempre devuelve al menos un estado si no hubo errores
        If Left(responseText, 1) <> "{" Then
            responseText = "{""status"":""ok""}"
        End If
    End If
ElseIf id = "obtenerportafoliosvalorescrm" Then
    sql = "SELECT DISTINCT codigo_murex FROM portafolioscrm WHERE administrador IN ('VALORES');"
    Set rs = conn.Execute(sql)
    responseText = "["
    firstItem = True
    Do While Not rs.EOF
        If Not firstItem Then
            responseText = responseText & ","
        Else
            firstItem = False
        End If
        responseText = responseText & """" & Replace(rs("codigo_murex"), """", "\""") & """"
        rs.MoveNext
    Loop
    responseText = responseText & "]"
ElseIf id = "limpiartablaconsolidadovalores" Then
    sql = "DELETE FROM nemotitulosvalores WHERE ""Origen Informacion"" = 'INVENTARIO TITULOS';"
    conn.Execute(sql)
    responseText = "{""status"":""ok""}"

ElseIf id = "cargueconsolidadovalores" Then
    Set data = payload("data")
    If IsEmpty(data) Or data.length = 0 Then
        responseText = "{""status"":""sin datos""}"
    Else
        For i = 0 To data.length - 1
            Set item = data(i)
            sql = "INSERT INTO nemotitulosvalores (""Macro Activo"", ""Isin"", ""Nemoténico"", ""Emisor Unificado"", ""Nombre Emisor"",""SALDO Macro Activo"",""SALDO ABA"",""Valor Nominal Actual"",""Codigo OyD"",""Fecha"",""Portafolio"",""POSICIÓN"",""Origen Informacion"") VALUES (" & _
                "'" & Replace(item("Macro Activo"), "'", "''") & "'," & _
                "'" & Replace(item("Isin"), "'", "''") & "'," & _
                "'" & Replace(item("Nemoténico"), "'", "''") & "'," & _
                "'" & Replace(item("Emisor Unificado"), "'", "''") & "'," & _
                "'" & Replace(item("Nombre Emisor"), "'", "''") & "'," & _
                item("SALDO Macro Activo") & "," & _
                item("SALDO ABA") & "," & _
                item("Valor Nominal Actual") & "," & _
                "'" & Replace(item("Código OyD"), "'", "''") & "'," & _
                "'" & FormatDateTime(item("Fecha"), vbShortDate) & "'," & _
                "'" & Replace(item("Portafolio"), "'", "''") & "'," & _
                "''," & _
                "'INVENTARIO TITULOS'" & _
            ");"
            conn.Execute(sql)
            If Err.Number <> 0 Then
                responseText = "{""status"":""error"",""fila"":" & (i+1) & ",""detalle"":""" & Replace(Err.Description, """", "'") & """}"
                Exit For
            End If
        Next
        If Left(responseText, 1) <> "{" Then
            responseText = "{""status"":""ok""}"
        End If
    End If
ElseIf id = "limpiartablacuposfidu" Then
    sql = "DELETE FROM cuposfiduciaria;"
    conn.Execute(sql)
    responseText = "{""status"":""ok""}"
ElseIf id = "carguecuposfidu" Then
    Set data = payload("data")
    If IsEmpty(data) Or data.length = 0 Then
        responseText = "{""status"":""sin datos""}"
    Else
        For i = 0 To data.length - 1
            Set item = data(i)
            sql = "INSERT INTO cuposfiduciaria (""Entidad"", ""MUREX"", ""Nombre"", ""Cupo"", ""Nemo"",""ISIN 1"",""Ocupación Máxima"") VALUES (" & _
                "'" & Replace(item("Entidad"), "'", "''") & "'," & _
                "'" & Replace(item("MUREX"), "'", "''") & "'," & _
                "'" & Replace(item("Nombre_1"), "'", "''") & "'," & _
                "'" & Replace(item("Cupo"), "'", "''") & "'," & _
                "'" & Replace(item("Nemo"), "'", "''") & "'," & _
                "'" & Replace(item("ISIN 1"), "'", "''") & "'," & _
                item("Ocupación Máxima")  & _
            ");"
            conn.Execute(sql)
            If Err.Number <> 0 Then
                responseText = "{""status"":""error"",""fila"":" & (i+1) & ",""detalle"":""" & Replace(Err.Description, """", "'") & """}"
                Exit For
            End If
        Next
        If Left(responseText, 1) <> "{" Then
            responseText = "{""status"":""ok""}"
        End If
    End If
ElseIf id = "limpiartablacuposvalores" Then
    sql = "DELETE FROM cuposvalores;"
    conn.Execute(sql)
    responseText = "{""status"":""ok""}"
ElseIf id = "carguecuposvalores" Then
    Set data = payload("data")
    If IsEmpty(data) Or data.length = 0 Then
        responseText = "{""status"":""sin datos""}"
    Else
        For i = 0 To data.length - 1
            Set item = data(i)
            sql = "INSERT INTO cuposvalores (""Entidad"", ""MUREX"", ""Nombre"", ""Cupo"", ""Nemo"", ""ISIN 1"", ""Ocupación Máxima"", ""OyD"") VALUES (" & _
                "'" & Replace(item("Entidad"), "'", "''") & "'," & _
                "'" & Replace(item("MUREX"), "'", "''") & "'," & _
                "'" & Replace(item("Nombre_1"), "'", "''") & "'," & _
                "'" & Replace(item("Cupo"), "'", "''") & "'," & _
                "'" & Replace(item("Nemo"), "'", "''") & "'," & _
                "'" & Replace(item("ISIN 1"), "'", "''") & "'," & _
                item("Ocupación Máxima") & "," & _
                "'" & Replace(item("OyD"), "'", "''") & "'" & _
            ");"
            conn.Execute(sql)
            If Err.Number <> 0 Then
                responseText = "{""status"":""error"",""fila"":" & (i+1) & ",""detalle"":""" & Replace(Err.Description, """", "'") & """}"
                Exit For
            End If
        Next
        If Left(responseText, 1) <> "{" Then
            responseText = "{""status"":""ok""}"
        End If
    End If
ElseIf id = "traerdatoscrucependientesfidu" Then
    sql = "SELECT DISTINCT ON (""ISIN"") ""ISIN"",""vrmercadohoymonedaempresa"",""Nominal Remanente"" FROM nemotitulosfiduciaria WHERE ""ISIN"" IS NOT NULL AND ""ISIN"" != '' ORDER BY ""ISIN"";"
    Set rs = conn.Execute(sql)
    responseText = "["
    firstItem = True
    Do While Not rs.EOF
        If Not firstItem Then
            responseText = responseText & ","
        Else
            firstItem = False
        End If

        ' Validaciones
        Dim valorMercado, nominalRemanente
        valorMercado = rs("vrmercadohoymonedaempresa")
        If IsNull(valorMercado) Or Trim(valorMercado) = "" Then
            valorMercado = 0
        End If

        nominalRemanente = rs("Nominal Remanente")
        If IsNull(nominalRemanente) Or Trim(nominalRemanente) = "" Then
            nominalRemanente = 0
        End If

        responseText = responseText & "{" & _
            """ISIN"":""" & Replace(rs("ISIN"), """", "\""") & """," & _
            """vrmercadohoymonedaempresa"":" & valorMercado & "," & _
            """Nominal Remanente"":" & nominalRemanente & _
        "}"
        rs.MoveNext
    Loop
    responseText = responseText & "]"

ElseIf id = "traerdatoscrucependientesfiud2" Then
    sql = "SELECT DISTINCT ON (""ISIN"") ""ISIN"",""Emisor / Contraparte"",""Emisor Unificado"" FROM nemotitulosfiduciaria WHERE  ""ISIN"" IS NOT NULL  AND ""ISIN"" != '' ORDER BY  ""ISIN"";"
    Set rs = conn.Execute(sql)
    responseText = "["
    firstItem = True
    Do While Not rs.EOF
        If Not firstItem Then
            responseText = responseText & ","
        Else
            firstItem = False
        End If
        responseText = responseText & "{" & _
            """ISIN"":""" & Replace(rs("ISIN"), """", "\""") & """," & _
            """Emisor / Contraparte"":""" & Replace(rs("Emisor / Contraparte"), """", "\""") & """," & _
            """Emisor Unificado"":""" & Replace(rs("Emisor Unificado"), """", "\""") & """" & _
        "}"
        rs.MoveNext
    Loop
    responseText = responseText & "]"
ElseIf id = "limpiartablacargueoperacionespendientesfidu" Then
    sql = "DELETE FROM nemotitulosfiduciaria WHERE ""Origen Informacion"" = 'OPERACIONES PENDIENTES';"
    conn.Execute(sql)
    responseText = "{""status"":""ok""}"

ElseIf id = "cargueoperacionespendientesfidu" Then
    'cargar las operaciones pendientes vienes del front con estos nombres las columasn "Portafolio",
            ' "Descripcion instrumento",
            ' "Emisor / Contraparte",
            ' "Emisor Unificado",
            ' "Nemotécnico",
            ' "ISIN instrumento",
            ' "Macro Activo",
            ' "SALDO Macro Activo",
            ' "SALDO ABA",
            ' "Nominal Remanente"
    Set data = payload("data")
    If IsEmpty(data) Or data.length = 0 Then
        responseText = "{""status"":""sin datos""}"
    Else
        For i = 0 To data.length - 1
            Set item = data(i)
            sql = "INSERT INTO nemotitulosfiduciaria (""Portafolio"", ""Especie/Generador"", ""Emisor / Contraparte"", ""Emisor Unificado"", ""Nemotécnico"", ""ISIN"", ""Macro Activo"", ""SALDO Macro Activo"", ""SALDO ABA"", ""Nominal Remanente"", ""Origen Informacion"") VALUES (" & _
                "'" & Replace(item("Nombre portafolio"), "'", "''") & "'," & _
                "'" & Replace(item("Descripcion instrumento"), "'", "''") & "'," & _
                "'" & Replace(item("Emisor / Contraparte"), "'", "''") & "'," & _
                "'" & Replace(item("Emisor Unificado"), "'", "''") & "'," & _
                "'" & Replace(item("Nemotécnico"), "'", "''") & "'," & _
                "'" & Replace(item("ISIN instrumento"), "'", "''") & "'," & _
                "'" & Replace(item("Macro Activo"), "'", "''") & "'," & _
                item("SALDO Macro Activo") & "," & _
                item("SALDO ABA") & "," & _
                item("Nominal Remanente") & "," & _
                "'OPERACIONES PENDIENTES'" & _
            ");"
            conn.Execute(sql)
            If Err.Number <> 0 Then
                responseText = "{""status"":""error"",""fila"":" & (i+1) & ",""detalle"":""" & Replace(Err.Description, """", "'") & """}"
                Exit For
            End If
        Next
        If Left(responseText, 1) <> "{" Then
            responseText = "{""status"":""ok""}"
        End If
    End If
Else
    responseText = "[]"
End If

' Cierre seguro de recursos
If Not rs Is Nothing Then rs.Close: Set rs = Nothing
If Not conn Is Nothing Then conn.Close: Set conn = Nothing
Set payload = Nothing
Set JSON = Nothing

' Respuesta final
Response.Write responseText
Session.Abandon
Response.End
%>
