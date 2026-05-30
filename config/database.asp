<%
' CMMS - Configuración de Base de Datos (GENERADO POR INSTALADOR)
' Generado: 5/30/2026 10:25:16 AM
Const DB_TYPE     = "sqlite"
Const DB_SERVER   = ","
Const DB_NAME     = "CMMS_DB, cmms_db"
Const DB_USER     = "admin,"
Const DB_PASS     = "admin,"
Const DB_PORT     = 0
Const DB_PROVIDER = ""
Const DB_TIMEOUT  = 30
Const DB_APP_NAME = "CMMS_System"
Const DB_FILE     = "C:\inetpub\wwwroot\CMMS\cmms.db"

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
