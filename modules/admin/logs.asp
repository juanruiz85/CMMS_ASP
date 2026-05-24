<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
If Not IsSupervisorOrAdmin() Then
    SetFlashMessage "danger", "No tiene permisos para ver logs."
    RedirectTo "/CMMS/index.asp"
End If

Dim PageTitle : PageTitle = "Logs de Actividad"
Dim PageModule : PageModule = "admin"

Dim oConn : Set oConn = GetConnection()

' Filters
Dim filterUser : filterUser = QSInt("user_id")
Dim filterMod  : filterMod  = Trim(Request.QueryString("module"))
Dim filterDate : filterDate = Trim(Request.QueryString("date"))

Dim sqlWhere : sqlWhere = " WHERE 1=1 "
If filterUser > 0 Then
    sqlWhere = sqlWhere & " AND l.user_id = " & filterUser
End If
If filterMod <> "" Then
    sqlWhere = sqlWhere & " AND l.entity_type = '" & Replace(filterMod, "'", "''") & "' "
End If
If filterDate <> "" Then
    sqlWhere = sqlWhere & " AND CAST(l.created_at AS DATE) = '" & Replace(filterDate, "'", "''") & "' "
End If

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = 50
Dim offset      : offset = (currentPage - 1) * perPage

' Count
Dim totalRows : totalRows = 0
Dim rsTotal
Set rsTotal = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_activity_logs l " & sqlWhere)
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing

' Data
Dim sql
sql = "SELECT l.*, u.first_name + ' ' + u.last_name AS user_name, u.username " & _
      "FROM cmms_activity_logs l " & _
      "LEFT JOIN cmms_users u ON u.id = l.user_id " & _
      sqlWhere & _
      "ORDER BY l.created_at DESC " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Dim rsLogs
Set rsLogs = oConn.Execute(sql)

' Get Users for filter
Dim rsUsers
Set rsUsers = oConn.Execute("SELECT id, first_name + ' ' + last_name AS name FROM cmms_users ORDER BY first_name")
%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Registro de auditoría del sistema</p>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="logs.asp">
      <div class="search-wrap" style="flex:1">
        <select name="user_id" class="form-control">
          <option value="">Cualquier Usuario</option>
          <% Do While Not rsUsers.EOF %>
          <option value="<%= rsUsers("id") %>" <%= IIf(filterUser=rsUsers("id"), "selected", "") %>><%= HtmlEncode(rsUsers("name")) %></option>
          <% rsUsers.MoveNext : Loop %>
        </select>
      </div>
      <div class="search-wrap" style="flex:1">
        <select name="module" class="form-control">
          <option value="">Cualquier Módulo</option>
          <option value="auth" <%= IIf(filterMod="auth", "selected", "") %>>Autenticación</option>
          <option value="plants" <%= IIf(filterMod="plants", "selected", "") %>>Plantas</option>
          <option value="assets" <%= IIf(filterMod="assets", "selected", "") %>>Equipos</option>
          <option value="work_orders" <%= IIf(filterMod="work_orders", "selected", "") %>>Órdenes de Trabajo</option>
          <option value="inventory" <%= IIf(filterMod="inventory", "selected", "") %>>Inventario</option>
          <option value="users" <%= IIf(filterMod="users", "selected", "") %>>Usuarios</option>
        </select>
      </div>
      <div class="search-wrap" style="flex:1">
        <input type="date" name="date" class="form-control" value="<%= HtmlEncode(filterDate) %>">
      </div>
      <button type="submit" class="btn btn-outline">Filtrar</button>
      <% If filterUser > 0 Or filterMod <> "" Or filterDate <> "" Then %>
      <a href="logs.asp" class="btn btn-ghost">Limpiar</a>
      <% End If %>
    </form>
  </div>

  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Fecha / Hora</th>
          <th>Usuario</th>
          <th>Acción</th>
          <th>Detalles</th>
          <th>Módulo</th>
          <th>IP</th>
        </tr>
      </thead>
      <tbody>
        <% If rsLogs.EOF Then %>
        <tr><td colspan="6" class="table-empty">No se encontraron registros de actividad</td></tr>
        <% Else %>
        <% Do While Not rsLogs.EOF %>
        <tr>
          <td style="white-space:nowrap;font-size:12px;color:var(--text-muted)"><%= FormatDateShort(rsLogs("created_at")) %> <%= FormatDateTime(rsLogs("created_at"), 4) %></td>
          <td>
            <% If IsNull(rsLogs("user_name")) Then %>
              <span style="color:var(--text-muted);font-style:italic">Sistema</span>
            <% Else %>
              <div style="font-weight:500"><%= HtmlEncode(rsLogs("user_name")) %></div>
            <% End If %>
          </td>
          <td><span class="badge badge-muted"><%= HtmlEncode(rsLogs("action")) %></span></td>
          <td style="max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="<%= HtmlEncode(rsLogs("details")) %>"><%= HtmlEncode(rsLogs("details")) %></td>
          <td><%= HtmlEncode(rsLogs("entity_type")) %></td>
          <td style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(NullStr(rsLogs("ip_address"))) %></td>
        </tr>
        <% 
        rsLogs.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsLogs.Close : Set rsLogs = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "?user_id=" & filterUser & "&module=" & Server.URLEncode(filterMod) & "&date=" & Server.URLEncode(filterDate)) %>
</div>

<% 
rsUsers.Close : Set rsUsers = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
