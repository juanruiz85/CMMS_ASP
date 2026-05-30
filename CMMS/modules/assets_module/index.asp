<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = T("nav_assets")
Dim PageModule : PageModule = "assets"

Dim oConn : Set oConn = GetConnection()

' Action: Delete
If Request.QueryString("action") = "delete" And IsSupervisorOrAdmin() Then
    Dim delId : delId = QSInt("id")
    If delId > 0 Then
        Dim cmdDel
        Set cmdDel = Server.CreateObject("ADODB.Command")
        Set cmdDel.ActiveConnection = oConn
        cmdDel.CommandText = "UPDATE cmms_assets SET status='retired' WHERE id=?"
        cmdDel.Parameters.Append cmdDel.CreateParameter("@id", 3, 1, , delId)
        cmdDel.Execute
        Set cmdDel = Nothing
        LogActivity CurrentUserId(), "DELETE_ASSET", "Equipo retirado ID: " & delId, "assets", delId
        SetFlashMessage "success", "Equipo marcado como retirado."
        RedirectTo "/CMMS/modules/assets_module/index.asp"
    End If
End If

' Filters
Dim filterQuery  : filterQuery  = Trim(Request.QueryString("q"))
Dim filterPlant  : filterPlant  = QSInt("plant_id")
Dim filterStatus : filterStatus = Trim(Request.QueryString("status"))

' Build WHERE clause with parameters
Dim sqlWhere : sqlWhere = " WHERE 1=1 "
Dim whereParams : whereParams = ""

If filterQuery <> "" Then
    sqlWhere = sqlWhere & " AND (a.name LIKE ? OR a.code LIKE ? OR a.category LIKE ?) "
    whereParams = whereParams & "%" & filterQuery & "%|" & "%" & filterQuery & "%|" & "%" & filterQuery & "%|"
End If
If filterPlant > 0 Then
    sqlWhere = sqlWhere & " AND a.plant_id = ? "
    whereParams = whereParams & filterPlant & "|"
End If
If filterStatus <> "" Then
    sqlWhere = sqlWhere & " AND a.status = ? "
    whereParams = whereParams & filterStatus & "|"
End If

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = GetPerPage()
Dim offset      : offset = (currentPage - 1) * perPage

' Count total with parameters
Dim totalRows : totalRows = 0
Dim cmdTotal, rsTotal
Set cmdTotal = Server.CreateObject("ADODB.Command")
cmdTotal.ActiveConnection = oConn
cmdTotal.CommandText = "SELECT COUNT(*) AS cnt FROM cmms_assets a " & sqlWhere

' Add parameters for count
Dim iParam, paramArr
If whereParams <> "" Then
    paramArr = Split(whereParams, "|")
    For iParam = 0 To UBound(paramArr) - 1
        If IsNumeric(paramArr(iParam)) Then
            cmdTotal.Parameters.Append cmdTotal.CreateParameter("@p" & iParam, 3, 1, , paramArr(iParam))
        Else
            cmdTotal.Parameters.Append cmdTotal.CreateParameter("@p" & iParam, 200, 1, 500, paramArr(iParam))
        End If
    Next
End If

Set rsTotal = cmdTotal.Execute()
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing
Set cmdTotal = Nothing

' Get data with parameters
Dim sql, cmdAssets, rsAssets
sql = "SELECT a.*, p.name AS plant_name " & _
      "FROM cmms_assets a " & _
      "LEFT JOIN cmms_plants p ON p.id = a.plant_id " & _
      sqlWhere & _
      "ORDER BY a.name " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Set cmdAssets = Server.CreateObject("ADODB.Command")
cmdAssets.ActiveConnection = oConn
cmdAssets.CommandText = sql

' Re-add parameters for main query
If whereParams <> "" Then
    paramArr = Split(whereParams, "|")
    For iParam = 0 To UBound(paramArr) - 1
        If IsNumeric(paramArr(iParam)) Then
            cmdAssets.Parameters.Append cmdAssets.CreateParameter("@p" & iParam, 3, 1, , paramArr(iParam))
        Else
            cmdAssets.Parameters.Append cmdAssets.CreateParameter("@p" & iParam, 200, 1, 500, paramArr(iParam))
        End If
    Next
End If

Set rsAssets = cmdAssets.Execute()
Set cmdAssets = Nothing

' Get plants for filter
Dim rsPlantsFilter
Set rsPlantsFilter = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Gestión de activos físicos y equipos</p>
  </div>
  <div class="page-actions">
    <% If IsSupervisorOrAdmin() Then %>
    <a href="/CMMS/modules/assets_module/form.asp" class="btn btn-primary">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
      Nuevo Equipo
    </a>
    <% End If %>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="index.asp">
      <div class="search-wrap" style="flex:2">
        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" name="q" class="form-control search-input" placeholder="Buscar nombre, código o categoría..." value="<%= HtmlEncode(filterQuery) %>">
      </div>
      <div class="search-wrap" style="flex:1;min-width:150px">
        <select name="plant_id" class="form-control">
          <option value="">Todas las Plantas</option>
          <% Do While Not rsPlantsFilter.EOF %>
          <option value="<%= rsPlantsFilter("id") %>" <%= IIf(filterPlant=rsPlantsFilter("id"), "selected", "") %>><%= HtmlEncode(rsPlantsFilter("name")) %></option>
          <% rsPlantsFilter.MoveNext : Loop %>
        </select>
      </div>
      <div class="search-wrap" style="flex:1;min-width:120px">
        <select name="status" class="form-control">
          <option value="">Cualquier estado</option>
          <option value="operational" <%= IIf(filterStatus="operational", "selected", "") %>>Operativo</option>
          <option value="maintenance" <%= IIf(filterStatus="maintenance", "selected", "") %>>En Mantenimiento</option>
          <option value="down" <%= IIf(filterStatus="down", "selected", "") %>>Fuera de servicio</option>
          <option value="retired" <%= IIf(filterStatus="retired", "selected", "") %>>Retirado</option>
        </select>
      </div>
      <button type="submit" class="btn btn-outline">Filtrar</button>
      <% If filterQuery <> "" Or filterPlant > 0 Or filterStatus <> "" Then %>
      <a href="index.asp" class="btn btn-ghost">Limpiar</a>
      <% End If %>
    </form>
  </div>

  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Código</th>
          <th class="sortable">Equipo</th>
          <th>Planta</th>
          <th>Categoría</th>
          <th>Criticidad</th>
          <th>Estado</th>
          <% If IsSupervisorOrAdmin() Then %><th style="text-align:right">Acciones</th><% End If %>
        </tr>
      </thead>
      <tbody>
        <% If rsAssets.EOF Then %>
        <tr><td colspan="7" class="table-empty">No se encontraron equipos</td></tr>
        <% Else %>
        <% Do While Not rsAssets.EOF %>
        <tr>
          <td class="bold"><a href="/CMMS/modules/assets_module/detail.asp?id=<%= rsAssets("id") %>"><%= HtmlEncode(rsAssets("code")) %></a></td>
          <td>
            <div><%= HtmlEncode(rsAssets("name")) %></div>
            <div style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(NullStr(rsAssets("location"))) %></div>
          </td>
          <td><%= HtmlEncode(NullStr(rsAssets("plant_name"))) %></td>
          <td><%= HtmlEncode(NullStr(rsAssets("category"))) %></td>
          <td><%= CriticalityBadge(rsAssets("criticality")) %></td>
          <td><%= AssetStatusBadge(rsAssets("status")) %></td>
          <% If IsSupervisorOrAdmin() Then %>
          <td style="text-align:right;white-space:nowrap">
            <a href="/CMMS/modules/assets_module/form.asp?id=<%= rsAssets("id") %>" class="btn btn-ghost btn-icon" title="Editar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></a>
            <a href="?action=delete&id=<%= rsAssets("id") %>" class="btn btn-ghost btn-icon text-danger" data-confirm="¿Está seguro de retirar este equipo?" title="Retirar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg></a>
          </td>
          <% End If %>
        </tr>
        <% 
        rsAssets.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsAssets.Close : Set rsAssets = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "?q=" & Server.URLEncode(filterQuery) & "&plant_id=" & filterPlant & "&status=" & filterStatus) %>
</div>

<% 
rsPlantsFilter.Close : Set rsPlantsFilter = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
