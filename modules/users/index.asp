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

Dim PageTitle : PageTitle = "Usuarios"
Dim PageModule : PageModule = "users"

Dim oConn : Set oConn = GetConnection()

' Action: Delete (Desactivar)
If Request.QueryString("action") = "delete" And IsAdmin() Then
    Dim delId : delId = QSInt("id")
    If delId > 0 And delId <> CurrentUserId() Then
        Dim cmdDel
        Set cmdDel = Server.CreateObject("ADODB.Command")
        Set cmdDel.ActiveConnection = oConn
        cmdDel.CommandText = "UPDATE cmms_users SET status='inactive' WHERE id=?"
        cmdDel.Parameters.Append cmdDel.CreateParameter("@id", 3, 1, , delId)
        cmdDel.Execute
        Set cmdDel = Nothing
        LogActivity CurrentUserId(), "DELETE_USER", "Usuario desactivado ID: " & delId, "users", delId
        SetFlashMessage "success", "Usuario desactivado."
        RedirectTo "/CMMS/modules/users/index.asp"
    End If
End If

' Filters
Dim filterQuery  : filterQuery  = Trim(Request.QueryString("q"))
Dim filterRole   : filterRole   = Trim(Request.QueryString("role"))
Dim filterStatus : filterStatus = Trim(Request.QueryString("status"))

Dim sqlWhere : sqlWhere = " WHERE 1=1 "
If filterQuery <> "" Then
    sqlWhere = sqlWhere & " AND (u.username LIKE '%" & Replace(filterQuery, "'", "''") & "%' OR u.email LIKE '%" & Replace(filterQuery, "'", "''") & "%' OR u.first_name + ' ' + u.last_name LIKE '%" & Replace(filterQuery, "'", "''") & "%') "
End If
If filterRole <> "" Then
    sqlWhere = sqlWhere & " AND u.role = '" & Replace(filterRole, "'", "''") & "' "
End If
If filterStatus <> "" Then
    sqlWhere = sqlWhere & " AND u.status = '" & Replace(filterStatus, "'", "''") & "' "
End If

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = 20 ' Hardcoded to 20 for users
Dim offset      : offset = (currentPage - 1) * perPage

' Count total
Dim totalRows : totalRows = 0
Dim rsTotal
Set rsTotal = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_users u " & sqlWhere)
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing

' Get data
Dim sql
sql = "SELECT u.* FROM cmms_users u " & _
      sqlWhere & _
      "ORDER BY u.first_name, u.last_name " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Dim rsUsers
Set rsUsers = oConn.Execute(sql)

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Gestión de usuarios y accesos del sistema</p>
  </div>
  <div class="page-actions">
    <% If IsAdmin() Then %>
    <a href="/CMMS/modules/users/form.asp" class="btn btn-primary">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="8.5" cy="7" r="4"/><line x1="20" y1="8" x2="20" y2="14"/><line x1="23" y1="11" x2="17" y2="11"/></svg>
      Nuevo Usuario
    </a>
    <% End If %>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="index.asp">
      <div class="search-wrap" style="flex:2">
        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" name="q" class="form-control search-input" placeholder="Buscar nombre, usuario o email..." value="<%= HtmlEncode(filterQuery) %>">
      </div>
      <div class="search-wrap" style="flex:1;min-width:140px">
        <select name="role" class="form-control">
          <option value="">Cualquier Rol</option>
          <option value="admin" <%= IIf(filterRole="admin", "selected", "") %>>Administrador</option>
          <option value="supervisor" <%= IIf(filterRole="supervisor", "selected", "") %>>Supervisor</option>
          <option value="technician" <%= IIf(filterRole="technician", "selected", "") %>>Técnico</option>
          <option value="viewer" <%= IIf(filterRole="viewer", "selected", "") %>>Solo lectura</option>
        </select>
      </div>
      <div class="search-wrap" style="flex:1;min-width:140px">
        <select name="status" class="form-control">
          <option value="">Cualquier estado</option>
          <option value="active" <%= IIf(filterStatus="active", "selected", "") %>>Activo</option>
          <option value="inactive" <%= IIf(filterStatus="inactive", "selected", "") %>>Inactivo</option>
        </select>
      </div>
      <button type="submit" class="btn btn-outline">Filtrar</button>
      <% If filterQuery <> "" Or filterRole <> "" Or filterStatus <> "" Then %>
      <a href="index.asp" class="btn btn-ghost">Limpiar</a>
      <% End If %>
    </form>
  </div>

  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Usuario</th>
          <th>Contacto</th>
          <th>Departamento</th>
          <th>Rol</th>
          <th>Estado</th>
          <th>Último Login</th>
          <% If IsAdmin() Then %><th style="text-align:right">Acciones</th><% End If %>
        </tr>
      </thead>
      <tbody>
        <% If rsUsers.EOF Then %>
        <tr><td colspan="7" class="table-empty">No se encontraron usuarios</td></tr>
        <% Else %>
        <% Do While Not rsUsers.EOF 
            Dim iniName : iniName = UCase(Left(rsUsers("first_name"),1) & Left(rsUsers("last_name"),1))
        %>
        <tr>
          <td>
            <div style="display:flex;align-items:center;gap:12px">
              <div class="user-avatar"><%= HtmlEncode(iniName) %></div>
              <div>
                <div style="font-weight:600;color:var(--text-primary)"><%= HtmlEncode(rsUsers("first_name") & " " & rsUsers("last_name")) %></div>
                <div style="font-size:11px;color:var(--text-muted)">@<%= HtmlEncode(rsUsers("username")) %></div>
              </div>
            </div>
          </td>
          <td>
            <div><%= HtmlEncode(rsUsers("email")) %></div>
            <div style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(NullStr(rsUsers("phone"))) %></div>
          </td>
          <td><%= HtmlEncode(NullStr(rsUsers("department"))) %></td>
          <td>
            <%
            Dim rColor : rColor = "badge-muted"
            Select Case rsUsers("role")
                Case "admin": rColor = "badge-danger"
                Case "supervisor": rColor = "badge-warning"
                Case "technician": rColor = "badge-primary"
                Case "viewer": rColor = "badge-info"
            End Select
            %>
            <span class="badge <%= rColor %> no-dot"><%= UCase(Left(rsUsers("role"),1)) & Mid(rsUsers("role"),2) %></span>
          </td>
          <td>
            <% If rsUsers("status") = "active" Then %>
              <span class="badge badge-success">Activo</span>
            <% Else %>
              <span class="badge badge-danger">Inactivo</span>
            <% End If %>
          </td>
          <td style="font-size:12px"><%= FormatDateShort(rsUsers("last_login")) %></td>
          
          <% If IsAdmin() Then %>
          <td style="text-align:right;white-space:nowrap">
            <a href="/CMMS/modules/users/form.asp?id=<%= rsUsers("id") %>" class="btn btn-ghost btn-icon" title="Editar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></a>
            <% If rsUsers("id") <> CurrentUserId() Then %>
            <a href="?action=delete&id=<%= rsUsers("id") %>" class="btn btn-ghost btn-icon text-danger" data-confirm="¿Está seguro de desactivar este usuario?" title="Desactivar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg></a>
            <% End If %>
          </td>
          <% End If %>
        </tr>
        <% 
        rsUsers.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsUsers.Close : Set rsUsers = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "?q=" & Server.URLEncode(filterQuery) & "&role=" & filterRole & "&status=" & filterStatus) %>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
