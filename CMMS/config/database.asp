<%
' CMMS - Configuración de Base de Datos (GENERADO POR INSTALADOR)
' Generado: Manual para corrección de error
Const DB_TYPE     = "sqlserver"
Const DB_SERVER   = "localhost"
Const DB_NAME     = "CMMS_DB"
Const DB_USER     = "cmms_user"
Const DB_PASS     = "password"
Const DB_PORT     = 1433
Const DB_PROVIDER = "SQLNCLI11"
Const DB_TIMEOUT  = 30
Const DB_APP_NAME = "CMMS_System"

Function GetConnectionString()
    Select Case DB_TYPE
        Case "sqlserver"
            GetConnectionString = "Provider=" & DB_PROVIDER & ";Server=" & DB_SERVER & ";Database=" & DB_NAME & ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";Connect Timeout=" & DB_TIMEOUT & ";Application Name=" & DB_APP_NAME & ";"
        Case "mysql"
            GetConnectionString = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=" & DB_SERVER & ";Port=" & DB_PORT & ";Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & ";Option=3;"
        Case "sqlite"
            GetConnectionString = "Driver={SQLite3 ODBC Driver};Database=" & DB_FILE & ";"
        Case Else
            GetConnectionString = ""
    End Select
End Function

Function GetConnection()
    Dim oConn
    Set oConn = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    oConn.Open GetConnectionString()
    If Err.Number <> 0 Then
        Response.Write "<div style='color:red;font-family:monospace;padding:20px'>ERROR BD: " & Err.Description & "</div>"
        Response.End
    End If
    On Error GoTo 0
    Set GetConnection = oConn
End Function

Sub CloseConnection(oConn)
    If Not IsNull(oConn) And Not IsEmpty(oConn) Then
        If oConn.State = 1 Then oConn.Close
        Set oConn = Nothing
    End If
End Sub
%>
