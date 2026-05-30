<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
If Not IsAdmin() Then
    SetFlashMessage "danger", "Solo los administradores pueden gestionar usuarios."
    RedirectTo "/CMMS/index.asp"
End If

Dim PageTitle : PageTitle = "Usuario"
Dim PageModule : PageModule = "users"

Dim oConn : Set oConn = GetConnection()
Dim itemId : itemId = QSInt("id")
Dim isEdit : isEdit = (itemId > 0)

Dim uUser, uEmail, uFirst, uLast, uDept, uPhone, uRole, uStatus

If isEdit Then
    Dim rs
    Set rs = oConn.Execute("SELECT * FROM cmms_users WHERE id=" & itemId)
    If Not rs.EOF Then
        uUser = rs("username")
        uEmail = rs("email")
        uFirst = rs("first_name")
        uLast = rs("last_name")
        uDept = NullStr(rs("department"))
        uPhone = NullStr(rs("phone"))
        uRole = rs("role")
        uStatus = rs("status")
        PageTitle = "Editar Usuario: " & uUser
    End If
    rs.Close : Set rs = Nothing
Else
    PageTitle = "Nuevo Usuario"
    uRole = "technician"
    uStatus = "active"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME") & "?id=" & itemId
    End If

    uUser = Trim(Request.Form("username"))
    uEmail = Trim(Request.Form("email"))
    uFirst = Trim(Request.Form("first_name"))
    uLast = Trim(Request.Form("last_name"))
    uDept = Trim(Request.Form("department"))
    uPhone = Trim(Request.Form("phone"))
    uRole = Trim(Request.Form("role"))
    uStatus = Trim(Request.Form("status"))
    Dim uPass : uPass = Trim(Request.Form("password"))

    Dim cmd, sql
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = oConn
    
    If isEdit Then
        sql = "UPDATE cmms_users SET username=?, email=?, first_name=?, last_name=?, department=?, phone=?, role=?, status=?, updated_at=GETDATE() WHERE id=?"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@user", 200, 1, 50, uUser)
        cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, uEmail)
        cmd.Parameters.Append cmd.CreateParameter("@first", 200, 1, 50, uFirst)
        cmd.Parameters.Append cmd.CreateParameter("@last", 200, 1, 50, uLast)
        cmd.Parameters.Append cmd.CreateParameter("@dept", 200, 1, 50, uDept)
        cmd.Parameters.Append cmd.CreateParameter("@phone", 200, 1, 50, uPhone)
        cmd.Parameters.Append cmd.CreateParameter("@role", 200, 1, 20, uRole)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, uStatus)
        cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , itemId)
        cmd.Execute
        
        ' Update password if provided
        If uPass <> "" Then
            Dim newSalt, newHash, cmdPass
            Randomize
            Dim saltChars : saltChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            Dim j, saltStr : saltStr = ""
            For j = 1 To 16
                saltStr = saltStr & Mid(saltChars, Int(Rnd * Len(saltChars)) + 1, 1)
            Next
            newSalt = saltStr
            newHash = SHA256Hash(uPass & newSalt)
            
            Set cmdPass = Server.CreateObject("ADODB.Command")
            Set cmdPass.ActiveConnection = oConn
            cmdPass.CommandText = "UPDATE cmms_users SET password=?, password_salt=? WHERE id=?"
            cmdPass.Parameters.Append cmdPass.CreateParameter("@pass", 200, 1, 64, newHash)
            cmdPass.Parameters.Append cmdPass.CreateParameter("@salt", 200, 1, 50, newSalt)
            cmdPass.Parameters.Append cmdPass.CreateParameter("@id", 3, 1, , itemId)
            cmdPass.Execute
            Set cmdPass = Nothing
        End If
        
        LogActivity CurrentUserId(), "UPDATE_USER", "Actualizó usuario: " & uUser, "users", itemId
    Else
        ' Generate salt and hash for new user
        Dim salt, passHash
        Randomize
        saltStr = ""
        For j = 1 To 16
            saltStr = saltStr & Mid(saltChars, Int(Rnd * Len(saltChars)) + 1, 1)
        Next
        salt = saltStr
        passHash = SHA256Hash(uPass & salt)
        
        sql = "INSERT INTO cmms_users (username, email, password, password_salt, first_name, last_name, department, phone, role, status, created_at, updated_at) " & _
              "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, GETDATE(), GETDATE())"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@user", 200, 1, 50, uUser)
        cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, uEmail)
        cmd.Parameters.Append cmd.CreateParameter("@pass", 200, 1, 64, passHash)
        cmd.Parameters.Append cmd.CreateParameter("@salt", 200, 1, 50, salt)
        cmd.Parameters.Append cmd.CreateParameter("@first", 200, 1, 50, uFirst)
        cmd.Parameters.Append cmd.CreateParameter("@last", 200, 1, 50, uLast)
        cmd.Parameters.Append cmd.CreateParameter("@dept", 200, 1, 50, uDept)
        cmd.Parameters.Append cmd.CreateParameter("@phone", 200, 1, 50, uPhone)
        cmd.Parameters.Append cmd.CreateParameter("@role", 200, 1, 20, uRole)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, uStatus)
        cmd.Execute
        
        LogActivity CurrentUserId(), "CREATE_USER", "Creó usuario: " & uUser, "users", 0
    End If
    
    SetFlashMessage "success", "Usuario guardado correctamente."
    RedirectTo "/CMMS/modules/users/index.asp"
End If

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/users/index.asp" class="btn btn-outline">Cancelar</a>
  </div>
</div>

<div class="card" style="max-width:800px;margin:0 auto">
  <div class="card-body">
    <form method="POST" action="form.asp?id=<%= itemId %>" data-validate>
      <%= CSRFField() %>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Información Personal</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Nombre <span class="required">*</span></label>
          <input type="text" name="first_name" class="form-control" required value="<%= HtmlEncode(uFirst) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Apellido <span class="required">*</span></label>
          <input type="text" name="last_name" class="form-control" required value="<%= HtmlEncode(uLast) %>">
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Correo Electrónico <span class="required">*</span></label>
          <input type="email" name="email" class="form-control" required value="<%= HtmlEncode(uEmail) %>" data-type="email">
        </div>
        <div class="form-group">
          <label class="form-label">Teléfono</label>
          <input type="text" name="phone" class="form-control" value="<%= HtmlEncode(uPhone) %>">
        </div>
      </div>
      
      <div class="form-group">
        <label class="form-label">Departamento</label>
        <input type="text" name="department" class="form-control" value="<%= HtmlEncode(uDept) %>" placeholder="Ej: Mantenimiento Eléctrico">
      </div>

      <h4 style="margin:24px 0 16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Credenciales y Acceso</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Usuario (Login) <span class="required">*</span></label>
          <input type="text" name="username" class="form-control" required value="<%= HtmlEncode(uUser) %>">
        </div>
        <div class="form-group">
          <label class="form-label"><%= IIf(isEdit, "Nueva Contraseña (dejar vacío para mantener actual)", "Contraseña <span class='required'>*</span>") %></label>
          <input type="password" name="password" class="form-control" <%= IIf(isEdit, "", "required") %>>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Rol del Sistema <span class="required">*</span></label>
          <select name="role" class="form-control" required>
            <option value="viewer" <%= IIf(uRole="viewer", "selected", "") %>>Solo Lectura</option>
            <option value="technician" <%= IIf(uRole="technician", "selected", "") %>>Técnico</option>
            <option value="supervisor" <%= IIf(uRole="supervisor", "selected", "") %>>Supervisor</option>
            <option value="admin" <%= IIf(uRole="admin", "selected", "") %>>Administrador</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Estado <span class="required">*</span></label>
          <select name="status" class="form-control" required>
            <option value="active" <%= IIf(uStatus="active", "selected", "") %>>Activo</option>
            <option value="inactive" <%= IIf(uStatus="inactive", "selected", "") %>>Inactivo</option>
          </select>
        </div>
      </div>

      <div style="margin-top:24px;text-align:right">
        <button type="submit" class="btn btn-primary">Guardar Usuario</button>
      </div>
    </form>
  </div>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
