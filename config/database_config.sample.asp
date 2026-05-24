<%
' =============================================================================
' CMMS - Configuración de Base de Datos (MUESTRA)
' Copiar este archivo a: config/database.asp
' Generado automáticamente por el instalador
' =============================================================================

' --- Tipo de base de datos: sqlserver | mysql | sqlite ---
Const DB_TYPE     = "sqlserver"

' --- SQL Server ---
Const DB_SERVER   = "NOMBRE_SERVIDOR"    ' Ej: SQLSERVER01, .\SQLEXPRESS, 192.168.1.10
Const DB_NAME     = "CMMS_DB"           ' Nombre de la base de datos
Const DB_USER     = "cmms_user"         ' Usuario SQL Server
Const DB_PASS     = "TU_CONTRASENA"     ' Contraseña SQL Server
Const DB_PORT     = 1433               ' Puerto (default 1433)
Const DB_PROVIDER = "SQLNCLI11"        ' Proveedor: SQLNCLI11 (SQL Server Native Client 11)

' --- Opciones de conexión ---
Const DB_TIMEOUT  = 30                 ' Timeout de conexión en segundos
Const DB_APP_NAME = "CMMS_System"      ' Nombre de la aplicación en el servidor SQL

' =============================================================================
' NO MODIFICAR DEBAJO DE ESTA LÍNEA
' =============================================================================

' Cadena de conexión SQL Server
Function GetConnectionString()
    Select Case DB_TYPE
        Case "sqlserver"
            GetConnectionString = "Provider=" & DB_PROVIDER & ";" & _
                                  "Server=" & DB_SERVER & ";" & _
                                  "Database=" & DB_NAME & ";" & _
                                  "UID=" & DB_USER & ";" & _
                                  "PWD=" & DB_PASS & ";" & _
                                  "Connect Timeout=" & DB_TIMEOUT & ";" & _
                                  "Application Name=" & DB_APP_NAME & ";"
        Case Else
            GetConnectionString = ""
    End Select
End Function

' Obtener conexión activa
Function GetConnection()
    Dim oConn
    Set oConn = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    oConn.Open GetConnectionString()
    If Err.Number <> 0 Then
        Response.Write "<div style='color:red;font-family:monospace;padding:20px'>"
        Response.Write "<strong>ERROR DE CONEXIÓN A BASE DE DATOS:</strong><br>"
        Response.Write Err.Description & " (Código: " & Err.Number & ")"
        Response.Write "<br><br>Verifique config/database.asp o ejecute el instalador."
        Response.Write "</div>"
        Response.End
    End If
    On Error GoTo 0
    Set GetConnection = oConn
End Function

' Cerrar conexión de forma segura
Sub CloseConnection(oConn)
    If Not IsNull(oConn) And Not IsEmpty(oConn) Then
        If oConn.State = 1 Then oConn.Close
        Set oConn = Nothing
    End If
End Sub
%>
