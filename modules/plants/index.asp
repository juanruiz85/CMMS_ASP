<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = T("nav_plants")
Dim PageModule : PageModule = "plants"

Dim oConn : Set oConn = GetConnection()

' Action: Delete
If Request.QueryString("action") = "delete" And IsSupervisorOrAdmin() Then
    Dim delId : delId = QSInt("id")
    If delId > 0 Then
        oConn.Execute("UPDATE cmms_plants SET status='inactive' WHERE id=" & delId)
        LogActivity CurrentUserId(), "DELETE_PLANT", "Planta desactivada ID: " & delId, "plants", delId
        SetFlashMessage "success", "Planta desactivada correctamente."
        RedirectTo "/CMMS/modules/plants/index.asp"
    End If
End If

' Filters
Dim filterQuery : filterQuery = Trim(Request.QueryString("q"))
Dim sqlWhere : sqlWhere = " WHERE 1=1 "
If filterQuery <> "" Then
    sqlWhere = sqlWhere & " AND (p.name LIKE '%" & Replace(filterQuery, "'", "''") & "%' OR p.code LIKE '%" & Replace(filterQuery, "'", "''") & "%') "
End If

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = GetPerPage()
Dim offset      : offset = (currentPage - 1) * perPage

' Count total
Dim totalRows : totalRows = 0
Dim rsTotal
Set rsTotal = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_plants p " & sqlWhere)
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing

' Get data
Dim sql
sql = "SELECT p.*, u.first_name + ' ' + u.last_name AS manager_name, " & _
      "(SELECT COUNT(*) FROM cmms_assets WHERE plant_id = p.id) AS asset_count " & _
      "FROM cmms_plants p " & _
      "LEFT JOIN cmms_users u ON u.id = p.manager_id " & _
      sqlWhere & _
      "ORDER BY p.name " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Dim rsPlants
Set rsPlants = oConn.Execute(sql)
%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Gestión de instalaciones y plantas físicas</p>
  </div>
  <div class="page-actions">
    <% If IsSupervisorOrAdmin() Then %>
    <a href="/CMMS/modules/plants/form.asp" class="btn btn-primary">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
      Nueva Planta
    </a>
    <% End If %>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="index.asp">
      <div class="search-wrap">
        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" name="q" class="form-control search-input" placeholder="Buscar por nombre o código..." value="<%= HtmlEncode(filterQuery) %>">
      </div>
      <button type="submit" class="btn btn-outline">Filtrar</button>
      <% If filterQuery <> "" Then %>
      <a href="index.asp" class="btn btn-ghost">Limpiar</a>
      <% End If %>
    </form>
  </div>

  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Código</th>
          <th class="sortable">Nombre</th>
          <th>Ciudad, País</th>
          <th>Responsable</th>
          <th style="text-align:center">Equipos</th>
          <th>Estado</th>
          <% If IsSupervisorOrAdmin() Then %><th style="text-align:right">Acciones</th><% End If %>
        </tr>
      </thead>
      <tbody>
        <% If rsPlants.EOF Then %>
        <tr><td colspan="7" class="table-empty">No se encontraron plantas</td></tr>
        <% Else %>
        <% Do While Not rsPlants.EOF %>
        <tr>
          <td class="bold"><%= HtmlEncode(rsPlants("code")) %></td>
          <td><%= HtmlEncode(rsPlants("name")) %></td>
          <td><%= HtmlEncode(NullStr(rsPlants("city"))) %>, <%= HtmlEncode(NullStr(rsPlants("country"))) %></td>
          <td><%= HtmlEncode(NullStr(rsPlants("manager_name"))) %></td>
          <td style="text-align:center"><span class="badge badge-muted no-dot"><%= rsPlants("asset_count") %></span></td>
          <td>
            <% If rsPlants("status") = "active" Then %>
              <span class="badge badge-success">Activo</span>
            <% Else %>
              <span class="badge badge-danger">Inactivo</span>
            <% End If %>
          </td>
          <% If IsSupervisorOrAdmin() Then %>
          <td style="text-align:right">
            <a href="/CMMS/modules/plants/form.asp?id=<%= rsPlants("id") %>" class="btn btn-ghost btn-icon" title="Editar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></a>
            <a href="?action=delete&id=<%= rsPlants("id") %>" class="btn btn-ghost btn-icon text-danger" data-confirm="¿Está seguro de desactivar esta planta?" title="Eliminar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg></a>
          </td>
          <% End If %>
        </tr>
        <% 
        rsPlants.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsPlants.Close : Set rsPlants = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "?q=" & Server.URLEncode(filterQuery)) %>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
