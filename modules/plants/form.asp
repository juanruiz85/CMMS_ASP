<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
If Not IsSupervisorOrAdmin() Then
    SetFlashMessage "danger", "No tiene permisos para acceder a esta página."
    RedirectTo "/CMMS/index.asp"
End If

Dim PageTitle : PageTitle = "Planta"
Dim PageModule : PageModule = "plants"

Dim oConn : Set oConn = GetConnection()
Dim itemId : itemId = QSInt("id")
Dim isEdit : isEdit = (itemId > 0)

Dim pCode, pName, pDesc, pAddress, pCity, pState, pCountry, pZip, pManagerId, pPhone, pEmail, pStatus

If isEdit Then
    Dim rs
    Set rs = oConn.Execute("SELECT * FROM cmms_plants WHERE id=" & itemId)
    If Not rs.EOF Then
        pCode = rs("code")
        pName = rs("name")
        pDesc = NullStr(rs("description"))
        pAddress = NullStr(rs("address"))
        pCity = NullStr(rs("city"))
        pState = NullStr(rs("state"))
        pCountry = NullStr(rs("country"))
        pZip = NullStr(rs("postal_code"))
        pManagerId = NullInt(rs("manager_id"))
        pPhone = NullStr(rs("phone"))
        pEmail = NullStr(rs("email"))
        pStatus = rs("status")
        PageTitle = "Editar Planta: " & pName
    End If
    rs.Close : Set rs = Nothing
Else
    PageTitle = "Nueva Planta"
    pStatus = "active"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME") & "?id=" & itemId
    End If

    pCode = Trim(Request.Form("code"))
    pName = Trim(Request.Form("name"))
    pDesc = Trim(Request.Form("description"))
    pAddress = Trim(Request.Form("address"))
    pCity = Trim(Request.Form("city"))
    pState = Trim(Request.Form("state"))
    pCountry = Trim(Request.Form("country"))
    pZip = Trim(Request.Form("postal_code"))
    pManagerId = Request.Form("manager_id")
    If pManagerId = "" Then pManagerId = "NULL"
    pPhone = Trim(Request.Form("phone"))
    pEmail = Trim(Request.Form("email"))
    pStatus = Trim(Request.Form("status"))

    Dim cmd, sql
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = oConn
    
    If isEdit Then
        sql = "UPDATE cmms_plants SET code=?, name=?, description=?, address=?, city=?, state=?, country=?, postal_code=?, manager_id=" & pManagerId & ", phone=?, email=?, status=?, updated_at=GETDATE() WHERE id=?"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, pCode)
        cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 100, pName)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 500, pDesc)
        cmd.Parameters.Append cmd.CreateParameter("@addr", 200, 1, 200, pAddress)
        cmd.Parameters.Append cmd.CreateParameter("@city", 200, 1, 100, pCity)
        cmd.Parameters.Append cmd.CreateParameter("@state", 200, 1, 100, pState)
        cmd.Parameters.Append cmd.CreateParameter("@country", 200, 1, 100, pCountry)
        cmd.Parameters.Append cmd.CreateParameter("@zip", 200, 1, 20, pZip)
        cmd.Parameters.Append cmd.CreateParameter("@phone", 200, 1, 50, pPhone)
        cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, pEmail)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, pStatus)
        cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , itemId)
        cmd.Execute
        LogActivity CurrentUserId(), "UPDATE_PLANT", "Actualizó planta: " & pName, "plants", itemId
    Else
        sql = "INSERT INTO cmms_plants (code, name, description, address, city, state, country, postal_code, manager_id, phone, email, status, created_at, updated_at) " & _
              "VALUES (?, ?, ?, ?, ?, ?, ?, ?, " & pManagerId & ", ?, ?, ?, GETDATE(), GETDATE())"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, pCode)
        cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 100, pName)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 500, pDesc)
        cmd.Parameters.Append cmd.CreateParameter("@addr", 200, 1, 200, pAddress)
        cmd.Parameters.Append cmd.CreateParameter("@city", 200, 1, 100, pCity)
        cmd.Parameters.Append cmd.CreateParameter("@state", 200, 1, 100, pState)
        cmd.Parameters.Append cmd.CreateParameter("@country", 200, 1, 100, pCountry)
        cmd.Parameters.Append cmd.CreateParameter("@zip", 200, 1, 20, pZip)
        cmd.Parameters.Append cmd.CreateParameter("@phone", 200, 1, 50, pPhone)
        cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, pEmail)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, pStatus)
        cmd.Execute
        LogActivity CurrentUserId(), "CREATE_PLANT", "Creó nueva planta: " & pName, "plants", 0
    End If
    
    SetFlashMessage "success", "Datos guardados correctamente."
    RedirectTo "/CMMS/modules/plants/index.asp"
End If

' Users list for manager select
Dim rsUsers
Set rsUsers = oConn.Execute("SELECT id, first_name + ' ' + last_name AS full_name FROM cmms_users WHERE status = 'active' ORDER BY first_name")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <div class="page-subtitle">Complete los campos para guardar la información.</div>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/plants/index.asp" class="btn btn-outline">Cancelar</a>
  </div>
</div>

<div class="card" style="max-width:800px;margin:0 auto">
  <div class="card-body">
    <form method="POST" action="form.asp?id=<%= itemId %>" data-validate>
      <%= CSRFField() %>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Información General</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Código <span class="required">*</span></label>
          <input type="text" name="code" class="form-control" required value="<%= HtmlEncode(pCode) %>" placeholder="Ej: PL-001">
        </div>
        <div class="form-group">
          <label class="form-label">Nombre <span class="required">*</span></label>
          <input type="text" name="name" class="form-control" required value="<%= HtmlEncode(pName) %>">
        </div>
      </div>

      <div class="form-group">
        <label class="form-label">Descripción</label>
        <textarea name="description" class="form-control"><%= HtmlEncode(pDesc) %></textarea>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Ubicación y Contacto</h4>
      <div class="form-group">
        <label class="form-label">Dirección</label>
        <input type="text" name="address" class="form-control" value="<%= HtmlEncode(pAddress) %>">
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Ciudad</label>
          <input type="text" name="city" class="form-control" value="<%= HtmlEncode(pCity) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Estado / Región</label>
          <input type="text" name="state" class="form-control" value="<%= HtmlEncode(pState) %>">
        </div>
      </div>
      
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">País</label>
          <input type="text" name="country" class="form-control" value="<%= HtmlEncode(pCountry) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Código Postal</label>
          <input type="text" name="postal_code" class="form-control" value="<%= HtmlEncode(pZip) %>">
        </div>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Administración</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Responsable (Manager)</label>
          <select name="manager_id" class="form-control">
            <option value="">Seleccione un usuario...</option>
            <% Do While Not rsUsers.EOF %>
            <option value="<%= rsUsers("id") %>" <%= IIf(pManagerId = rsUsers("id"), "selected", "") %>><%= HtmlEncode(rsUsers("full_name")) %></option>
            <% rsUsers.MoveNext : Loop %>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Estado</label>
          <select name="status" class="form-control" required>
            <option value="active" <%= IIf(pStatus="active", "selected", "") %>>Activo</option>
            <option value="inactive" <%= IIf(pStatus="inactive", "selected", "") %>>Inactivo</option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Teléfono</label>
          <input type="text" name="phone" class="form-control" value="<%= HtmlEncode(pPhone) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Correo de contacto</label>
          <input type="email" name="email" class="form-control" value="<%= HtmlEncode(pEmail) %>" data-type="email">
        </div>
      </div>

      <div style="margin-top:24px;text-align:right">
        <button type="submit" class="btn btn-primary">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>
          Guardar Datos
        </button>
      </div>

    </form>
  </div>
</div>

<%
rsUsers.Close : Set rsUsers = Nothing
CloseConnection(oConn)
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
