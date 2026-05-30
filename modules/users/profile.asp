<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = "Mi Perfil"
Dim PageModule : PageModule = ""

Dim oConn : Set oConn = GetConnection()
Dim userId : userId = CurrentUserId()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME")
    End If

    Dim action : action = Request.Form("action")
    
    If action = "update_profile" Then
        Dim uFirst : uFirst = Trim(Request.Form("first_name"))
        Dim uLast  : uLast  = Trim(Request.Form("last_name"))
        Dim uEmail : uEmail = Trim(Request.Form("email"))
        Dim uPhone : uPhone = Trim(Request.Form("phone"))
        Dim uLang  : uLang  = Trim(Request.Form("language"))
        
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = oConn
        cmd.CommandText = "UPDATE cmms_users SET first_name=?, last_name=?, email=?, phone=?, updated_at=GETDATE() WHERE id=" & userId
        cmd.Parameters.Append cmd.CreateParameter("@first", 200, 1, 50, uFirst)
        cmd.Parameters.Append cmd.CreateParameter("@last", 200, 1, 50, uLast)
        cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, uEmail)
        cmd.Parameters.Append cmd.CreateParameter("@phone", 200, 1, 50, uPhone)
        cmd.Execute
        
        Session("user_firstname") = uFirst
        Session("user_lastname")  = uLast
        Session("user_lang")      = uLang
        ' Aquí se podría actualizar el dashboard_config para guardar el idioma
        
        SetFlashMessage "success", "Perfil actualizado correctamente."
        
    ElseIf action = "change_password" Then
        Dim currentPass : currentPass = Trim(Request.Form("current_password"))
        Dim newPass     : newPass     = Trim(Request.Form("new_password"))
        
        Dim rsPass, cmdPassUpd
        Set rsPass = oConn.Execute("SELECT password, password_salt FROM cmms_users WHERE id=" & userId)
        If Not rsPass.EOF Then
            Dim currHash : currHash = SHA256Hash(currentPass & rsPass("password_salt"))
            If currHash = rsPass("password") Then
                Dim newSalt, newHash
                Randomize
                Dim saltChars : saltChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
                Dim j, saltStr : saltStr = ""
                For j = 1 To 16
                    saltStr = saltStr & Mid(saltChars, Int(Rnd * Len(saltChars)) + 1, 1)
                Next
                newSalt = saltStr
                newHash = SHA256Hash(newPass & newSalt)
                
                Set cmdPassUpd = Server.CreateObject("ADODB.Command")
                Set cmdPassUpd.ActiveConnection = oConn
                cmdPassUpd.CommandText = "UPDATE cmms_users SET password=?, password_salt=? WHERE id=?"
                cmdPassUpd.Parameters.Append cmdPassUpd.CreateParameter("@pass", 200, 1, 64, newHash)
                cmdPassUpd.Parameters.Append cmdPassUpd.CreateParameter("@salt", 200, 1, 50, newSalt)
                cmdPassUpd.Parameters.Append cmdPassUpd.CreateParameter("@id", 3, 1, , userId)
                cmdPassUpd.Execute
                Set cmdPassUpd = Nothing
                
                SetFlashMessage "success", "Contraseña cambiada correctamente."
            Else
                SetFlashMessage "danger", "La contraseña actual es incorrecta."
            End If
        End If
        rsPass.Close : Set rsPass = Nothing
    End If
    
    RedirectTo Request.ServerVariables("SCRIPT_NAME")
End If

Dim rsUser
Set rsUser = oConn.Execute("SELECT * FROM cmms_users WHERE id=" & userId)

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
  </div>
</div>

<div class="dashboard-grid">
  <div class="col-8">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Información Personal</h3>
      </div>
      <div class="card-body">
        <form method="POST" action="profile.asp" data-validate>
          <input type="hidden" name="action" value="update_profile">
          <%= CSRFField() %>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Nombre <span class="required">*</span></label>
              <input type="text" name="first_name" class="form-control" required value="<%= HtmlEncode(rsUser("first_name")) %>">
            </div>
            <div class="form-group">
              <label class="form-label">Apellido <span class="required">*</span></label>
              <input type="text" name="last_name" class="form-control" required value="<%= HtmlEncode(rsUser("last_name")) %>">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Correo Electrónico <span class="required">*</span></label>
              <input type="email" name="email" class="form-control" required value="<%= HtmlEncode(rsUser("email")) %>" data-type="email">
            </div>
            <div class="form-group">
              <label class="form-label">Teléfono</label>
              <input type="text" name="phone" class="form-control" value="<%= HtmlEncode(NullStr(rsUser("phone"))) %>">
            </div>
          </div>
          
          <h4 style="margin:24px 0 16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Preferencias</h4>
          
          <div class="form-group" style="max-width:300px">
            <label class="form-label">Idioma de la Interfaz</label>
            <select name="language" class="form-control">
              <option value="es" <%= IIf(Session("user_lang")="es" Or Session("user_lang")="", "selected", "") %>>Español</option>
              <option value="en" <%= IIf(Session("user_lang")="en", "selected", "") %>>English</option>
            </select>
          </div>

          <div style="margin-top:24px">
            <button type="submit" class="btn btn-primary">Actualizar Perfil</button>
          </div>
        </form>
      </div>
    </div>
  </div>

  <div class="col-4">
    <div class="card">
      <div class="card-header">
        <h3 class="card-title">Seguridad</h3>
      </div>
      <div class="card-body">
        <form method="POST" action="profile.asp" data-validate>
          <input type="hidden" name="action" value="change_password">
          <%= CSRFField() %>

          <div class="form-group">
            <label class="form-label">Contraseña Actual <span class="required">*</span></label>
            <input type="password" name="current_password" class="form-control" required>
          </div>
          <div class="form-group">
            <label class="form-label">Nueva Contraseña <span class="required">*</span></label>
            <input type="password" name="new_password" class="form-control" required>
          </div>
          <div style="margin-top:24px">
            <button type="submit" class="btn btn-secondary w-full">Cambiar Contraseña</button>
          </div>
        </form>
      </div>
    </div>
  </div>
</div>

<% 
rsUser.Close : Set rsUser = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
