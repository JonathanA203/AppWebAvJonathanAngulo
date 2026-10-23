<%
'=====================================================================
'  PoliMusic :: acceso a datos (ADO + SQL Server)
'  Todas las consultas usan parametros para evitar inyeccion SQL.
'=====================================================================
Dim g_conn, g_connOpen
g_connOpen = False

Function GetConn()
    If Not g_connOpen Then
        Set g_conn = Server.CreateObject("ADODB.Connection")
        g_conn.ConnectionString = ConnString()
        g_conn.ConnectionTimeout = 15
        g_conn.Open
        g_connOpen = True
    End If
    Set GetConn = g_conn
End Function

Sub CloseConn()
    On Error Resume Next
    If g_connOpen Then
        g_conn.Close
        Set g_conn = Nothing
        g_connOpen = False
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------
'  Construccion de comandos parametrizados
'---------------------------------------------------------------------
Function BuildCmd(sql, params)
    Dim cmd, i
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = GetConn()
    cmd.CommandType = 1          ' adCmdText
    cmd.CommandTimeout = 30
    cmd.CommandText = sql
    If IsArray(params) Then
        For i = 0 To UBound(params)
            Call AddParam(cmd, params(i))
        Next
    End If
    Set BuildCmd = cmd
End Function

Sub AddParam(cmd, v)
    Dim t
    t = VarType(v)
    Select Case t
        Case 0, 1                       ' vbEmpty / vbNull
            cmd.Parameters.Append cmd.CreateParameter("p", 202, 1, 4000, Null)
        Case 2, 3                       ' vbInteger / vbLong
            cmd.Parameters.Append cmd.CreateParameter("p", 3, 1, , CLng(v))
        Case 4, 5, 6, 14                ' Single / Double / Currency / Decimal
            cmd.Parameters.Append cmd.CreateParameter("p", 5, 1, , CDbl(v))
        Case 7                          ' vbDate
            cmd.Parameters.Append cmd.CreateParameter("p", 135, 1, , CDate(v))
        Case 11                         ' vbBoolean
            cmd.Parameters.Append cmd.CreateParameter("p", 11, 1, , v)
        Case Else                       ' texto
            cmd.Parameters.Append cmd.CreateParameter("p", 202, 1, 4000, CStr(v))
    End Select
End Sub

' Devuelve un Recordset (siempre usar Set)
Function DbQuery(sql, params)
    Dim cmd
    Set cmd = BuildCmd(sql, params)
    Set DbQuery = cmd.Execute
End Function

' Devuelve el primer valor de la primera fila (o Null)
Function DbScalar(sql, params)
    Dim rs
    DbScalar = Null
    Set rs = DbQuery(sql, params)
    If rs.State <> 0 Then
        If Not rs.EOF Then DbScalar = rs(0).Value
        rs.Close
    End If
    Set rs = Nothing
End Function

' Ejecuta una sentencia sin devolver datos
Sub DbExec(sql, params)
    Dim cmd
    Set cmd = BuildCmd(sql, params)
    cmd.Execute
    Set cmd = Nothing
End Sub

' Ejecuta y devuelve True si no hubo error (util para DELETE con FK)
Function DbTry(sql, params)
    Dim cmd
    DbTry = True
    On Error Resume Next
    Set cmd = BuildCmd(sql, params)
    cmd.Execute
    If Err.Number <> 0 Then
        DbTry = False
        Err.Clear
    End If
    Set cmd = Nothing
    On Error GoTo 0
End Function
%>
