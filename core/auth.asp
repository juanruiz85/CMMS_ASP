<%
' =============================================================================
' CMMS - Sistema de Autenticación y Control de Acceso (auth.asp)
' Requiere: config/database.asp, core/functions.asp, core/i18n.asp
' =============================================================================

' Roles del sistema
Const ROLE_ADMIN      = "admin"
Const ROLE_SUPERVISOR = "supervisor"
Const ROLE_TECHNICIAN = "technician"
Const ROLE_VIEWER     = "viewer"

' Timeout de sesión en minutos
Const SESSION_TIMEOUT = 480  ' 8 horas

' Ruta de login
Const LOGIN_URL = "/CMMS/login.asp"

' -----------------------------------------------------------
' CheckAuth - Verificar que el usuario está autenticado
' Si no está autenticado, redirige al login
' -----------------------------------------------------------
Sub CheckAuth()
    ' Verificar sesión activa
    If Session("user_id") = "" Or Session("user_id") = 0 Then
        ' Guardar URL actual para redirigir después del login
        Session("redirect_after_login") = Request.ServerVariables("URL") & _
            IIf(Request.QueryString <> "", "?" & Request.QueryString, "")
        Response.Redirect LOGIN_URL & "?expired=1"
        Response.End
    End If

    ' Verificar timeout de sesión
    If Session("last_activity") <> "" Then
        Dim minutesInactive
        minutesInactive = DateDiff("n", Session("last_activity"), Now())
        If minutesInactive > SESSION_TIMEOUT Then
            Session.Abandon
            Response.Redirect LOGIN_URL & "?expired=1"
            Response.End
        End If
    End If

    ' Actualizar timestamp de actividad
    Session("last_activity") = Now()
End Sub

' -----------------------------------------------------------
' CheckRole - Verificar que el usuario tiene uno de los roles permitidos
' roles: string delimitado por coma, ej: "admin,supervisor"
' -----------------------------------------------------------
Sub CheckRole(allowedRoles)
    CheckAuth
    Dim userRole, roleArray, i, allowed
    userRole  = LCase(Session("user_role"))
    roleArray = Split(LCase(allowedRoles), ",")
    allowed   = False
    For i = 0 To UBound(roleArray)
        If Trim(roleArray(i)) = userRole Then
            allowed = True
            Exit For
        End If
    Next
    If Not allowed Then
        Response.Redirect "/CMMS/index.asp?access_denied=1"
        Response.End
    End If
End Sub

' -----------------------------------------------------------
' HasRole - Retorna True si el usuario tiene el rol indicado
' -----------------------------------------------------------
Function HasRole(roleName)
    Dim userRole
    userRole = LCase(Session("user_role"))
    If roleName = "" Then
        HasRole = True
    ElseIf userRole = ROLE_ADMIN Then
        HasRole = True  ' Admin siempre tiene acceso
    Else
        HasRole = (userRole = LCase(roleName))
    End If
End Function

' Verificar si es admin
Function IsAdmin()
    IsAdmin = (LCase(Session("user_role")) = ROLE_ADMIN)
End Function

' Verificar si es admin o supervisor
Function IsSupervisorOrAdmin()
    IsSupervisorOrAdmin = (LCase(Session("user_role")) = ROLE_ADMIN Or _
                           LCase(Session("user_role")) = ROLE_SUPERVISOR)
End Function

' -----------------------------------------------------------
' DoLogin - Autenticar usuario y crear sesión
' Retorna: True si autenticación exitosa, False si no
' -----------------------------------------------------------
Function DoLogin(username, plainPassword)
    Dim oConn, oCmd, oRS, success
    success = False

    On Error Resume Next
    Set oConn = GetConnection()
    If Err.Number <> 0 Then
        DoLogin = False
        Exit Function
    End If

    Set oCmd = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandText = "SELECT id, username, password, password_salt, role, first_name, last_name, " & _
                       "email, status, avatar, department, dashboard_config " & _
                       "FROM cmms_users WHERE username = ? AND status = 'active'"
    oCmd.Parameters.Append oCmd.CreateParameter("@username", 200, 1, 50, Trim(username))
    Set oRS = oCmd.Execute()

    If Not oRS.EOF Then
        Dim storedHash, storedSalt, inputHash
        storedHash = NullStr(oRS("password"))
        storedSalt = NullStr(oRS("password_salt"))
        inputHash  = HashPassword(plainPassword, storedSalt)

        If inputHash = storedHash Then
            ' Credenciales correctas: crear sesión
            Session("user_id")        = NullInt(oRS("id"))
            Session("user_name")      = NullStr(oRS("username"))
            Session("user_role")      = NullStr(oRS("role"))
            Session("user_firstname") = NullStr(oRS("first_name"))
            Session("user_lastname")  = NullStr(oRS("last_name"))
            Session("user_email")     = NullStr(oRS("email"))
            Session("user_avatar")    = NullStr(oRS("avatar"))
            Session("user_dept")      = NullStr(oRS("department"))
            Session("last_activity")  = Now()

            ' Cargar preferencia de idioma del usuario (desde settings)
            Dim dashConfig : dashConfig = NullStr(oRS("dashboard_config"))
            If InStr(dashConfig, """lang"":""en""") > 0 Then
                Session("user_lang") = "en"
            Else
                Session("user_lang") = "es"
            End If

            ' Actualizar último login en BD (compatible multi-BD)
            Dim oCmd2
            Set oCmd2 = Server.CreateObject("ADODB.Command")
            Set oCmd2.ActiveConnection = oConn
            oCmd2.CommandText = "UPDATE cmms_users SET last_login = " & GetDateSQL() & " WHERE id = ?"
            oCmd2.Parameters.Append oCmd2.CreateParameter("@id", 203, 1, 4, Session("user_id"))
            oCmd2.Execute
            Set oCmd2 = Nothing

            ' Log de actividad
            LogActivity Session("user_id"), "login", "Inicio de sesión desde " & Request.ServerVariables("REMOTE_ADDR"), "user", Session("user_id")

            success = True
        End If
    End If

    oRS.Close
    Set oRS   = Nothing
    Set oCmd  = Nothing
    CloseConnection oConn
    On Error GoTo 0

    DoLogin = success
End Function

' -----------------------------------------------------------
' DoLogout - Cerrar sesión del usuario
' -----------------------------------------------------------
Sub DoLogout()
    If Session("user_id") <> "" And Session("user_id") <> 0 Then
        LogActivity Session("user_id"), "logout", "Cierre de sesión", "user", Session("user_id")
    End If
    Session.Abandon
End Sub

' -----------------------------------------------------------
' GetCurrentUser - Obtener datos del usuario actual de la sesión
' Retorna un diccionario con los datos del usuario
' -----------------------------------------------------------
Function GetCurrentUser()
    Dim u
    Set u = Server.CreateObject("Scripting.Dictionary")
    u.Add "id",        Session("user_id")
    u.Add "username",  Session("user_name")
    u.Add "role",      Session("user_role")
    u.Add "firstname", Session("user_firstname")
    u.Add "lastname",  Session("user_lastname")
    u.Add "email",     Session("user_email")
    u.Add "avatar",    Session("user_avatar")
    u.Add "dept",      Session("user_dept")
    u.Add "fullname",  Trim(Session("user_firstname") & " " & Session("user_lastname"))
    Set GetCurrentUser = u
End Function

' Nombre completo del usuario actual
Function CurrentUserFullName()
    Dim fn, ln
    fn = Trim(Session("user_firstname"))
    ln = Trim(Session("user_lastname"))
    If fn = "" And ln = "" Then
        CurrentUserFullName = Session("user_name")
    Else
        CurrentUserFullName = Trim(fn & " " & ln)
    End If
End Function

' ID del usuario actual
Function CurrentUserId()
    CurrentUserId = SafeInt(Session("user_id"))
End Function

' Rol del usuario actual
Function CurrentUserRole()
    CurrentUserRole = LCase(NullStr(Session("user_role")))
End Function

' -----------------------------------------------------------
' API Token Authentication (para endpoints REST)
' -----------------------------------------------------------
Function ValidateAPIToken(apiKey)
    If apiKey = "" Then
        ValidateAPIToken = False
        Exit Function
    End If

    Dim oConn, oRS, valid
    valid = False
    Set oConn = GetConnection()

    ' El API key se almacena como SHA256 en cmms_settings
    ' Formato clave: "api_key_<username>"
    Set oRS = oConn.Execute("SELECT value FROM cmms_settings WHERE key_name = 'api_token_" & SafeStr(apiKey) & "'")
    If Not oRS.EOF Then
        ' Token válido: cargar usuario asociado
        valid = True
    End If

    oRS.Close
    Set oRS   = Nothing
    CloseConnection oConn
    ValidateAPIToken = valid
End Function

' -----------------------------------------------------------
' CheckAPIAuth - Para páginas de la API REST
' Retorna el user_id si autenticado, 0 si no
' -----------------------------------------------------------
Function CheckAPIAuth()
    Dim apiKey
    ' Buscar en Header Authorization: Bearer TOKEN
    apiKey = Request.ServerVariables("HTTP_AUTHORIZATION")
    If Left(apiKey, 7) = "Bearer " Then
        apiKey = Mid(apiKey, 8)
    Else
        apiKey = Request.QueryString("api_key")
    End If

    If apiKey = "" Then
        CheckAPIAuth = 0
        Exit Function
    End If

    Dim oConn, oRS
    Set oConn = GetConnection()
    Set oRS   = oConn.Execute("SELECT u.id FROM cmms_users u " & _
                              "INNER JOIN cmms_settings s ON s.key_name = 'api_key_' + CAST(u.id AS VARCHAR) " & _
                              "WHERE s.value = '" & SafeStr(apiKey) & "' AND u.status = 'active'")
    If Not oRS.EOF Then
        CheckAPIAuth = NullInt(oRS("id"))
    Else
        CheckAPIAuth = 0
    End If

    oRS.Close
    Set oRS   = Nothing
    CloseConnection oConn
End Function
%>
