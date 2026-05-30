<%
' =============================================================================
' CMMS - Funciones Utilitarias Globales (functions.asp)
' Requiere: config/database.asp
' =============================================================================

' --- FUNCIONES AUXILIARES ---

' Función IIf (Immediate If) - Retorna un valor basado en una condición
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

' --- FUNCIONES COMPATIBLES MULTI-BD ---

' Obtener función de fecha actual según tipo de BD
Function GetDateSQL()
    Dim dbType
    dbType = LCase(GetDBType())
    Select Case dbType
        Case "mysql"
            GetDateSQL = "NOW()"
        Case "sqlite"
            GetDateSQL = "datetime('now')"
        Case Else  ' sqlserver
            GetDateSQL = "GETDATE()"
    End Select
End Function

' Obtener último ID insertado según tipo de BD
Function GetLastInsertID(oConn)
    Dim dbType, rs, lastId
    dbType = LCase(GetDBType())
    Select Case dbType
        Case "mysql"
            Set rs = oConn.Execute("SELECT LAST_INSERT_ID() AS id")
        Case "sqlite"
            Set rs = oConn.Execute("SELECT last_insert_rowid() AS id")
        Case Else  ' sqlserver
            Set rs = oConn.Execute("SELECT SCOPE_IDENTITY() AS id")
    End Select
    lastId = rs("id")
    rs.Close
    Set rs = Nothing
    GetLastInsertID = lastId
End Function

' Obtener tipo de base de datos actual
Function GetDBType()
    On Error Resume Next
    GetDBType = DB_TYPE
    If Err.Number <> 0 Then GetDBType = "sqlserver"
    On Error GoTo 0
End Function

' --- HASHING SHA-256 (via .NET Framework COM, disponible en WS2012 R2+) ---
Function SHA256Hash(strText)
    On Error Resume Next
    Dim oSHA, oEncoding, oBytes, oHash, sHex, i
    Set oSHA      = CreateObject("System.Security.Cryptography.SHA256Managed")
    Set oEncoding = CreateObject("System.Text.UTF8Encoding")
    oBytes = oEncoding.GetBytes_4(strText)
    oHash  = oSHA.ComputeHash_2(oBytes)
    sHex = ""
    For i = 0 To UBound(oHash)
        sHex = sHex & Right("0" & Hex(oHash(i)), 2)
    Next
    If Err.Number <> 0 Then
        ' Fallback básico si .NET COM no disponible
        SHA256Hash = MD5Fallback(strText)
    Else
        SHA256Hash = LCase(sHex)
    End If
    Set oSHA      = Nothing
    Set oEncoding = Nothing
    On Error GoTo 0
End Function

' Fallback MD5 simple vía ADODB.Stream (menos seguro pero funcional)
Function MD5Fallback(str)
    Dim oMD5, oBytes, oHash, i, sHex
    On Error Resume Next
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    Set oEncoding = CreateObject("System.Text.UTF8Encoding")
    oBytes = oEncoding.GetBytes_4(str)
    oHash  = oMD5.ComputeHash_2(oBytes)
    sHex = ""
    For i = 0 To UBound(oHash)
        sHex = sHex & Right("0" & Hex(oHash(i)), 2)
    Next
    MD5Fallback = LCase(sHex)
    Set oMD5 = Nothing
    On Error GoTo 0
End Function

' Hashear contraseña con salt
Function HashPassword(plainPassword, salt)
    HashPassword = SHA256Hash(plainPassword & salt)
End Function

' Generar salt aleatorio
Function GenerateSalt()
    Randomize
    Dim chars, i, salt
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    salt  = ""
    For i = 1 To 16
        salt = salt & Mid(chars, Int(Rnd * Len(chars)) + 1, 1)
    Next
    GenerateSalt = salt
End Function

' --- SANITIZACIÓN SQL ---
' Escapa comillas simples para prevenir SQL Injection
' NOTA: Se recomienda usar ADODB.Command con parámetros donde sea posible
Function SafeStr(s)
    If IsNull(s) Or IsEmpty(s) Then
        SafeStr = ""
    Else
        SafeStr = Replace(CStr(s), "'", "''")
    End If
End Function

Function SafeInt(v)
    If IsNumeric(v) Then
        SafeInt = CLng(v)
    Else
        SafeInt = 0
    End If
End Function

Function SafeDecimal(v)
    If IsNumeric(v) Then
        SafeDecimal = CDbl(v)
    Else
        SafeDecimal = 0
    End If
End Function

' --- VALIDACIÓN ---
Function IsValidEmail(email)
    Dim re
    Set re = New RegExp
    re.Pattern = "^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
    IsValidEmail = re.Test(email)
    Set re = Nothing
End Function

Function IsValidDate(d)
    IsValidDate = IsDate(d)
End Function

' --- GENERACIÓN DE CÓDIGOS ÚNICOS ---
Function GenerateWOCode()
    ' Formato: WO-2026-00001
    Dim oConn, oRS, nextNum
    Set oConn = GetConnection()
    Set oRS   = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_work_orders")
    nextNum   = oRS("cnt") + 1
    GenerateWOCode = "WO-" & Year(Now) & "-" & Right("00000" & nextNum, 5)
    oRS.Close
    Set oRS   = Nothing
    Set oConn = Nothing
End Function

Function GenerateAssetCode(prefix)
    Dim oConn, oRS, nextNum
    If prefix = "" Then prefix = "EQ"
    Set oConn = GetConnection()
    Set oRS   = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_assets")
    nextNum   = oRS("cnt") + 1
    GenerateAssetCode = prefix & "-" & Right("0000" & nextNum, 4)
    oRS.Close
    Set oRS   = Nothing
    Set oConn = Nothing
End Function

Function GenerateInventoryCode()
    Dim oConn, oRS, nextNum
    Set oConn = GetConnection()
    Set oRS   = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_inventory")
    nextNum   = oRS("cnt") + 1
    GenerateInventoryCode = "INV-" & Right("00000" & nextNum, 5)
    oRS.Close
    Set oRS   = Nothing
    Set oConn = Nothing
End Function

' --- FORMATO DE FECHAS ---
Function FormatDateShort(d)
    If IsNull(d) Or IsEmpty(d) Or d = "" Then
        FormatDateShort = "-"
    Else
        FormatDateShort = Format2D(Day(d)) & "/" & Format2D(Month(d)) & "/" & Year(d)
    End If
End Function

Function FormatDateTime(d)
    If IsNull(d) Or IsEmpty(d) Or d = "" Then
        FormatDateTime = "-"
    Else
        FormatDateTime = Format2D(Day(d)) & "/" & Format2D(Month(d)) & "/" & Year(d) & " " & Format2D(Hour(d)) & ":" & Format2D(Minute(d))
    End If
End Function

Function Format2D(n)
    Format2D = Right("0" & n, 2)
End Function

Function FormatMoney(amount)
    If IsNull(amount) Or IsEmpty(amount) Then
        FormatMoney = "$0.00"
    Else
        FormatMoney = "$" & FormatNumber(CDbl(amount), 2)
    End If
End Function

' Tiempo relativo (hace N minutos / horas / días)
Function TimeAgo(dt)
    If IsNull(dt) Or IsEmpty(dt) Or dt = "" Then
        TimeAgo = "-"
        Exit Function
    End If
    Dim diff, mins, hrs, days
    diff = DateDiff("n", dt, Now())  ' diferencia en minutos
    If diff < 1 Then
        TimeAgo = "Ahora mismo"
    ElseIf diff < 60 Then
        TimeAgo = "Hace " & diff & " min"
    ElseIf diff < 1440 Then
        hrs = Int(diff / 60)
        TimeAgo = "Hace " & hrs & " h"
    Else
        days = Int(diff / 1440)
        TimeAgo = "Hace " & days & " días"
    End If
End Function

' --- PAGINACIÓN ---
Function GetCurrentPage()
    Dim p
    p = Request.QueryString("page")
    If Not IsNumeric(p) Or p < 1 Then p = 1
    GetCurrentPage = CInt(p)
End Function

Function GetPerPage()
    Dim pp
    pp = Request.QueryString("pp")
    If Not IsNumeric(pp) Or pp < 5 Then pp = 25
    If pp > 200 Then pp = 200
    GetPerPage = CInt(pp)
End Function

' Genera SQL de paginación para SQL Server (OFFSET/FETCH NEXT)
Function PaginateSQL(sql, page, perPage)
    Dim offset
    offset = (page - 1) * perPage
    PaginateSQL = sql & " OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"
End Function

' Genera HTML de controles de paginación
Function PaginationHTML(totalRecords, currentPage, perPage, baseUrl)
    Dim totalPages, i, html, sep
    If perPage = 0 Then perPage = 25
    totalPages = Int((totalRecords - 1) / perPage) + 1
    If totalPages < 1 Then totalPages = 1

    ' Agrega ? o & según la URL base
    If InStr(baseUrl, "?") > 0 Then
        sep = "&"
    Else
        sep = "?"
    End If

    html = "<div class='pagination'>"
    html = html & "<span class='pag-info'>Página " & currentPage & " de " & totalPages & " (" & totalRecords & " registros)</span>"
    html = html & "<div class='pag-btns'>"

    If currentPage > 1 Then
        html = html & "<a href='" & baseUrl & sep & "page=1' class='pag-btn' title='Primera'>&laquo;</a>"
        html = html & "<a href='" & baseUrl & sep & "page=" & (currentPage - 1) & "' class='pag-btn'>&lsaquo;</a>"
    End If

    ' Ventana de páginas
    Dim startPage, endPage
    startPage = currentPage - 2
    endPage   = currentPage + 2
    If startPage < 1 Then startPage = 1
    If endPage > totalPages Then endPage = totalPages

    For i = startPage To endPage
        If i = currentPage Then
            html = html & "<span class='pag-btn active'>" & i & "</span>"
        Else
            html = html & "<a href='" & baseUrl & sep & "page=" & i & "' class='pag-btn'>" & i & "</a>"
        End If
    Next

    If currentPage < totalPages Then
        html = html & "<a href='" & baseUrl & sep & "page=" & (currentPage + 1) & "' class='pag-btn'>&rsaquo;</a>"
        html = html & "<a href='" & baseUrl & sep & "page=" & totalPages & "' class='pag-btn' title='Última'>&raquo;</a>"
    End If

    html = html & "</div></div>"
    PaginationHTML = html
End Function

' --- LOG DE ACTIVIDAD ---
Sub LogActivity(userId, action, description, entityType, entityId)
    On Error Resume Next
    Dim oConn, oCmd
    Set oConn = GetConnection()
    Set oCmd  = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandText = "INSERT INTO cmms_activity_logs (user_id, action, description, entity_type, entity_id, ip_address, user_agent, created_at) " & _
                       "VALUES (?, ?, ?, ?, ?, ?, ?, GETDATE())"
    oCmd.Parameters.Append oCmd.CreateParameter("@user_id",     203, 1, 4,   userId)
    oCmd.Parameters.Append oCmd.CreateParameter("@action",      200, 1, 255, action)
    oCmd.Parameters.Append oCmd.CreateParameter("@desc",        200, 1, -1,  description)
    oCmd.Parameters.Append oCmd.CreateParameter("@etype",       200, 1, 100, entityType)
    oCmd.Parameters.Append oCmd.CreateParameter("@eid",         203, 1, 4,   entityId)
    oCmd.Parameters.Append oCmd.CreateParameter("@ip",          200, 1, 45,  Request.ServerVariables("REMOTE_ADDR"))
    oCmd.Parameters.Append oCmd.CreateParameter("@ua",          200, 1, 255, Left(Request.ServerVariables("HTTP_USER_AGENT"), 255))
    oCmd.Execute
    Set oCmd  = Nothing
    CloseConnection oConn
    On Error GoTo 0
End Sub

' --- NOTIFICACIONES ---
Sub CreateNotification(userId, title, message, notifType, entityType, entityId)
    On Error Resume Next
    Dim oConn, oCmd
    Set oConn = GetConnection()
    Set oCmd  = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandText = "INSERT INTO cmms_notifications (user_id, title, message, type, entity_type, entity_id, created_at) " & _
                       "VALUES (?, ?, ?, ?, ?, ?, GETDATE())"
    oCmd.Parameters.Append oCmd.CreateParameter("@uid",   203, 1, 4,   userId)
    oCmd.Parameters.Append oCmd.CreateParameter("@title", 200, 1, 255, title)
    oCmd.Parameters.Append oCmd.CreateParameter("@msg",   200, 1, -1,  message)
    oCmd.Parameters.Append oCmd.CreateParameter("@type",  200, 1, 20,  notifType)
    oCmd.Parameters.Append oCmd.CreateParameter("@etype", 200, 1, 100, entityType)
    oCmd.Parameters.Append oCmd.CreateParameter("@eid",   203, 1, 4,   entityId)
    oCmd.Execute
    Set oCmd  = Nothing
    CloseConnection oConn
    On Error GoTo 0
End Sub

' Contar notificaciones no leídas para un usuario
Function CountUnreadNotifications(userId)
    Dim oConn, oRS
    Set oConn = GetConnection()
    Set oRS   = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_notifications WHERE user_id = " & SafeInt(userId) & " AND is_read = 0")
    CountUnreadNotifications = oRS("cnt")
    oRS.Close
    Set oRS   = Nothing
    CloseConnection oConn
End Function

' --- RESPUESTA JSON ---
Function JsonResponse(success, data, message)
    Dim json
    If success Then
        json = "{""success"":true,""message"":""" & EscapeJson(message) & """,""data"":" & data & "}"
    Else
        json = "{""success"":false,""message"":""" & EscapeJson(message) & """,""data"":null}"
    End If
    JsonResponse = json
End Function

Function EscapeJson(s)
    If IsNull(s) Or IsEmpty(s) Then
        EscapeJson = ""
        Exit Function
    End If
    s = Replace(s, "\",  "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(9),  "\t")
    EscapeJson = s
End Function

' Converter RecordSet a JSON array
Function RSToJson(oRS)
    If oRS.EOF And oRS.BOF Then
        RSToJson = "[]"
        Exit Function
    End If
    Dim json, fields, i, fieldCount
    fieldCount = oRS.Fields.Count
    json = "["
    Dim firstRow : firstRow = True
    Do While Not oRS.EOF
        If Not firstRow Then json = json & ","
        json = json & "{"
        For i = 0 To fieldCount - 1
            If i > 0 Then json = json & ","
            json = json & """" & oRS.Fields(i).Name & """:"
            If IsNull(oRS.Fields(i).Value) Then
                json = json & "null"
            ElseIf IsNumeric(oRS.Fields(i).Value) And _
                   oRS.Fields(i).Type <> 202 And _
                   oRS.Fields(i).Type <> 203 And _
                   oRS.Fields(i).Type <> 200 Then
                json = json & oRS.Fields(i).Value
            Else
                json = json & """" & EscapeJson(CStr(oRS.Fields(i).Value)) & """"
            End If
        Next
        json = json & "}"
        firstRow = False
        oRS.MoveNext
    Loop
    json = json & "]"
    RSToJson = json
End Function

' --- CSRF TOKEN ---
Function GetCSRFToken()
    If Session("csrf_token") = "" Then
        Randomize
        Dim token, i
        Dim chars : chars = "abcdefghijklmnopqrstuvwxyz0123456789"
        token = ""
        For i = 1 To 32
            token = token & Mid(chars, Int(Rnd * Len(chars)) + 1, 1)
        Next
        Session("csrf_token") = token
    End If
    GetCSRFToken = Session("csrf_token")
End Function

Function ValidateCSRF()
    Dim formToken, sessionToken
    formToken   = Request.Form("csrf_token")
    sessionToken = Session("csrf_token")
    If formToken = "" Or sessionToken = "" Or formToken <> sessionToken Then
        ValidateCSRF = False
    Else
        ValidateCSRF = True
    End If
End Function

' HTML del campo oculto CSRF
Function CSRFField()
    CSRFField = "<input type='hidden' name='csrf_token' value='" & GetCSRFToken() & "'>"
End Function

' --- HTML HELPERS ---
Function HtmlEncode(s)
    If IsNull(s) Or IsEmpty(s) Then
        HtmlEncode = ""
    Else
        HtmlEncode = Server.HTMLEncode(CStr(s))
    End If
End Function

Function NullStr(v)
    If IsNull(v) Or IsEmpty(v) Then
        NullStr = ""
    Else
        NullStr = CStr(v)
    End If
End Function

Function NullInt(v)
    If IsNull(v) Or IsEmpty(v) Then
        NullInt = 0
    Else
        NullInt = CLng(v)
    End If
End Function

' Obtener query string como integer seguro
Function QSInt(param)
    Dim v : v = Request.QueryString(param)
    If IsNumeric(v) Then
        QSInt = CLng(v)
    Else
        QSInt = 0
    End If
End Function

' Obtener form field como string seguro
Function FormStr(param)
    FormStr = Trim(Request.Form(param))
End Function

' Obtener form field como integer seguro
Function QSIntForm(param)
    Dim v : v = Request.Form(param)
    If IsNumeric(v) Then
        QSIntForm = CLng(v)
    Else
        QSIntForm = 0
    End If
End Function

' --- MANEJO GLOBAL DE ERRORES ---
Const DEBUG_MODE = True ' Cambiar a False en producción

Sub CheckError(contextMsg)
    If Err.Number <> 0 Then
        Dim errNum, errDesc, errSrc
        errNum = Err.Number
        errDesc = Err.Description
        errSrc = Err.Source
        Err.Clear()
        
        Response.Clear()
        Response.Write "<div style='font-family:sans-serif;max-width:800px;margin:40px auto;padding:20px;border:1px solid #f87171;border-radius:8px;background:#fef2f2;color:#991b1b;'>"
        Response.Write "<h2 style='margin-top:0'>Error de Sistema</h2>"
        
        If DEBUG_MODE Or IsAdmin() Then
            Response.Write "<p><strong>Contexto:</strong> " & HtmlEncode(contextMsg) & "</p>"
            Response.Write "<p><strong>Error [" & errNum & "]:</strong> " & HtmlEncode(errDesc) & "</p>"
            Response.Write "<p><strong>Origen:</strong> " & HtmlEncode(errSrc) & "</p>"
            Response.Write "<hr style='border:none;border-top:1px solid #fca5a5;margin:20px 0'>"
            Response.Write "<p style='font-size:12px;color:#b91c1c'><em>Esta información técnica se muestra porque el sistema está en modo DEBUG o usted es administrador.</em></p>"
        Else
            Response.Write "<p>Ha ocurrido un problema al procesar su solicitud. El error ha sido registrado y el equipo técnico ha sido notificado.</p>"
            Response.Write "<p>Intente nuevamente más tarde o contacte a soporte técnico indicando el módulo en el que se encontraba.</p>"
        End If
        
        Response.Write "</div>"
        
        ' Registrar error silenciosamente si existe la DB (intentar)
        On Error Resume Next
        Dim oConn, sqlE
        Set oConn = GetConnection()
        sqlE = "INSERT INTO cmms_activity_logs (action, description, entity_type, ip_address, created_at) VALUES ('SYSTEM_ERROR', '" & Replace(Left(contextMsg & " | " & errDesc, 500), "'", "''") & "', 'system', '" & Request.ServerVariables("REMOTE_ADDR") & "', GETDATE())"
        oConn.Execute(sqlE)
        On Error GoTo 0
        
        Response.End
    End If
End Sub

' Badge HTML para estado de órdenes de trabajo
Function WOStatusBadge(status)
    Dim cls, lbl
    Select Case LCase(status)
        Case "open"        : cls = "badge-info"    : lbl = "Abierta"
        Case "in_progress" : cls = "badge-primary"  : lbl = "En Progreso"
        Case "pending"     : cls = "badge-warning"  : lbl = "Pendiente"
        Case "completed"   : cls = "badge-success"  : lbl = "Completada"
        Case "cancelled"   : cls = "badge-danger"   : lbl = "Cancelada"
        Case Else          : cls = "badge-secondary": lbl = status
    End Select
    WOStatusBadge = "<span class='badge " & cls & "'>" & lbl & "</span>"
End Function

Function WOPriorityBadge(priority)
    Dim cls, lbl
    Select Case LCase(priority)
        Case "low"    : cls = "badge-secondary" : lbl = "Baja"
        Case "medium" : cls = "badge-info"      : lbl = "Media"
        Case "high"   : cls = "badge-warning"   : lbl = "Alta"
        Case "urgent" : cls = "badge-danger"    : lbl = "Urgente"
        Case Else     : cls = "badge-secondary" : lbl = priority
    End Select
    WOPriorityBadge = "<span class='badge " & cls & "'>" & lbl & "</span>"
End Function

Function AssetStatusBadge(status)
    Dim cls, lbl
    Select Case LCase(status)
        Case "operational" : cls = "badge-success"   : lbl = "Operativo"
        Case "maintenance" : cls = "badge-warning"   : lbl = "Mantenimiento"
        Case "down"        : cls = "badge-danger"    : lbl = "Fuera de Servicio"
        Case "retired"     : cls = "badge-secondary" : lbl = "Retirado"
        Case Else          : cls = "badge-secondary" : lbl = status
    End Select
    AssetStatusBadge = "<span class='badge " & cls & "'>" & lbl & "</span>"
End Function

Function CriticalityBadge(crit)
    Dim cls, lbl
    Select Case LCase(crit)
        Case "low"      : cls = "badge-success"  : lbl = "Baja"
        Case "medium"   : cls = "badge-info"     : lbl = "Media"
        Case "high"     : cls = "badge-warning"  : lbl = "Alta"
        Case "critical" : cls = "badge-danger"   : lbl = "Crítica"
        Case Else       : cls = "badge-secondary": lbl = crit
    End Select
    CriticalityBadge = "<span class='badge " & cls & "'>" & lbl & "</span>"
End Function

' --- REDIRECT HELPERS ---
Sub RedirectTo(url)
    Response.Redirect url
End Sub

' Alert de éxito en la siguiente página via Session
Sub SetFlashMessage(msgType, message)
    Session("flash_type") = msgType
    Session("flash_msg")  = message
End Sub

Function GetFlashMessage()
    Dim msgType, message, html
    msgType = Session("flash_type")
    message = Session("flash_msg")
    If message = "" Then
        GetFlashMessage = ""
        Exit Function
    End If
    Session("flash_type") = ""
    Session("flash_msg")  = ""
    html = "<div class='alert alert-" & msgType & " alert-dismissible' role='alert'>"
    html = html & "<span>" & HtmlEncode(message) & "</span>"
    html = html & "<button class='alert-close' onclick='this.parentElement.remove()'>×</button>"
    html = html & "</div>"
    GetFlashMessage = html
End Function
%>
