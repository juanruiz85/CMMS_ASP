<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = T("nav_inventory")
Dim PageModule : PageModule = "inventory"

Dim oConn : Set oConn = GetConnection()

' Action: Delete
If Request.QueryString("action") = "delete" And IsSupervisorOrAdmin() Then
    Dim delId : delId = QSInt("id")
    If delId > 0 Then
        oConn.Execute("UPDATE cmms_inventory SET status='inactive' WHERE id=" & delId)
        LogActivity CurrentUserId(), "DELETE_INVENTORY", "Artículo inactivo ID: " & delId, "inventory", delId
        SetFlashMessage "success", "Artículo marcado como inactivo."
        RedirectTo "/CMMS/modules/inventory/index.asp"
    End If
End If

' Filters
Dim filterQuery  : filterQuery  = Trim(Request.QueryString("q"))
Dim filterPlant  : filterPlant  = QSInt("plant_id")
Dim filterStatus : filterStatus = Trim(Request.QueryString("status"))

Dim sqlWhere : sqlWhere = " WHERE 1=1 "
If filterQuery <> "" Then
    sqlWhere = sqlWhere & " AND (i.name LIKE '%" & Replace(filterQuery, "'", "''") & "%' OR i.code LIKE '%" & Replace(filterQuery, "'", "''") & "%' OR i.part_number LIKE '%" & Replace(filterQuery, "'", "''") & "%') "
End If
If filterPlant > 0 Then
    sqlWhere = sqlWhere & " AND i.plant_id = " & filterPlant
End If
If filterStatus <> "" Then
    sqlWhere = sqlWhere & " AND i.status = '" & Replace(filterStatus, "'", "''") & "' "
End If

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = GetPerPage()
Dim offset      : offset = (currentPage - 1) * perPage

' Count total
Dim totalRows : totalRows = 0
Dim rsTotal
Set rsTotal = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_inventory i " & sqlWhere)
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing

' Get data
Dim sql
sql = "SELECT i.*, p.name AS plant_name, ISNULL((SELECT SUM(quantity) FROM cmms_inventory_stock WHERE inventory_id = i.id), 0) AS total_stock " & _
      "FROM cmms_inventory i " & _
      "LEFT JOIN cmms_plants p ON p.id = i.plant_id " & _
      sqlWhere & _
      "ORDER BY i.name " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Dim rsInv
Set rsInv = oConn.Execute(sql)

' Get plants for filter
Dim rsPlantsFilter
Set rsPlantsFilter = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")

' KPIs
Dim kpiTotal : kpiTotal = 0
Dim kpiLow : kpiLow = 0
Dim kpiOut : kpiOut = 0

Dim rsKpi
Set rsKpi = oConn.Execute("SELECT i.id, i.reorder_point, ISNULL((SELECT SUM(quantity) FROM cmms_inventory_stock WHERE inventory_id = i.id), 0) AS stk FROM cmms_inventory i WHERE i.status = 'active'")
Do While Not rsKpi.EOF
    kpiTotal = kpiTotal + 1
    If rsKpi("stk") = 0 Then
        kpiOut = kpiOut + 1
    ElseIf rsKpi("stk") <= rsKpi("reorder_point") Then
        kpiLow = kpiLow + 1
    End If
    rsKpi.MoveNext
Loop
rsKpi.Close : Set rsKpi = Nothing

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Gestión de repuestos y materiales</p>
  </div>
  <div class="page-actions">
    <% If IsSupervisorOrAdmin() Then %>
    <a href="/CMMS/modules/inventory/form.asp" class="btn btn-primary">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
      Nuevo Artículo
    </a>
    <% End If %>
  </div>
</div>

<div class="stats-grid" style="margin-bottom:var(--space-lg)">
  <div class="stat-card primary">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiTotal %>">0</div>
      <div class="stat-label">Total Artículos</div>
    </div>
  </div>
  <div class="stat-card warning">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiLow %>">0</div>
      <div class="stat-label">Stock Bajo</div>
    </div>
  </div>
  <div class="stat-card danger">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiOut %>">0</div>
      <div class="stat-label">Sin Stock</div>
    </div>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="index.asp">
      <div class="search-wrap" style="flex:2">
        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" name="q" class="form-control search-input" placeholder="Buscar nombre, código o # parte..." value="<%= HtmlEncode(filterQuery) %>">
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
          <option value="active" <%= IIf(filterStatus="active", "selected", "") %>>Activo</option>
          <option value="inactive" <%= IIf(filterStatus="inactive", "selected", "") %>>Inactivo</option>
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
          <th class="sortable">Artículo</th>
          <th>Planta</th>
          <th>Categoría</th>
          <th style="text-align:right">Stock</th>
          <th style="text-align:right">Costo Unit.</th>
          <th>Estado</th>
          <% If IsSupervisorOrAdmin() Then %><th style="text-align:right">Acciones</th><% End If %>
        </tr>
      </thead>
      <tbody>
        <% If rsInv.EOF Then %>
        <tr><td colspan="8" class="table-empty">No se encontraron artículos</td></tr>
        <% Else %>
        <% Do While Not rsInv.EOF 
             Dim stk : stk = rsInv("total_stock")
             Dim rp  : rp  = rsInv("reorder_point")
             Dim stkColor : stkColor = ""
             If stk <= 0 Then
                stkColor = "color:var(--danger);font-weight:700"
             ElseIf stk <= rp Then
                stkColor = "color:var(--warning);font-weight:700"
             End If
        %>
        <tr>
          <td class="bold"><%= HtmlEncode(rsInv("code")) %></td>
          <td>
            <div style="font-weight:500;color:var(--text-primary)"><%= HtmlEncode(rsInv("name")) %></div>
            <div style="font-size:11px;color:var(--text-muted)">PN: <%= HtmlEncode(NullStr(rsInv("part_number"))) %></div>
          </td>
          <td><%= HtmlEncode(NullStr(rsInv("plant_name"))) %></td>
          <td><%= HtmlEncode(NullStr(rsInv("category"))) %></td>
          <td style="text-align:right;<%= stkColor %>">
            <%= stk %> <%= HtmlEncode(rsInv("unit_of_measure")) %>
            <% If stk <= rp And stk > 0 Then %><div style="font-size:10px">Stock Bajo</div><% End If %>
            <% If stk <= 0 Then %><div style="font-size:10px">Sin Stock</div><% End If %>
          </td>
          <td style="text-align:right"><%= FormatMoney(rsInv("unit_cost")) %></td>
          <td>
            <% If rsInv("status") = "active" Then %>
              <span class="badge badge-success">Activo</span>
            <% Else %>
              <span class="badge badge-danger">Inactivo</span>
            <% End If %>
          </td>
          <% If IsSupervisorOrAdmin() Then %>
          <td style="text-align:right;white-space:nowrap">
            <a href="/CMMS/modules/inventory/form.asp?id=<%= rsInv("id") %>" class="btn btn-ghost btn-icon" title="Editar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></a>
            <a href="?action=delete&id=<%= rsInv("id") %>" class="btn btn-ghost btn-icon text-danger" data-confirm="¿Está seguro de desactivar este artículo?" title="Desactivar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg></a>
          </td>
          <% End If %>
        </tr>
        <% 
        rsInv.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsInv.Close : Set rsInv = Nothing %>
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
