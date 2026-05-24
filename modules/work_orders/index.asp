<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = T("nav_work_orders")
Dim PageModule : PageModule = "work_orders"

Dim oConn : Set oConn = GetConnection()

' Action: Delete
If Request.QueryString("action") = "delete" And IsSupervisorOrAdmin() Then
    Dim delId : delId = QSInt("id")
    If delId > 0 Then
        oConn.Execute("UPDATE cmms_work_orders SET status='cancelled' WHERE id=" & delId)
        LogActivity CurrentUserId(), "DELETE_WO", "Orden de trabajo cancelada ID: " & delId, "work_orders", delId
        SetFlashMessage "success", "Orden de Trabajo cancelada."
        RedirectTo "/CMMS/modules/work_orders/index.asp"
    End If
End If

' Filters
Dim filterQuery  : filterQuery  = Trim(Request.QueryString("q"))
Dim filterType   : filterType   = Trim(Request.QueryString("type"))
Dim filterStatus : filterStatus = Trim(Request.QueryString("status"))
Dim filterPlant  : filterPlant  = QSInt("plant_id")

Dim sqlWhere : sqlWhere = " WHERE 1=1 "
If filterQuery <> "" Then
    sqlWhere = sqlWhere & " AND (wo.title LIKE '%" & Replace(filterQuery, "'", "''") & "%' OR wo.code LIKE '%" & Replace(filterQuery, "'", "''") & "%') "
End If
If filterType <> "" Then
    sqlWhere = sqlWhere & " AND wo.type = '" & Replace(filterType, "'", "''") & "' "
End If
If filterStatus <> "" Then
    sqlWhere = sqlWhere & " AND wo.status = '" & Replace(filterStatus, "'", "''") & "' "
End If
If filterPlant > 0 Then
    sqlWhere = sqlWhere & " AND wo.plant_id = " & filterPlant
End If

' KPI Stats
Dim kpiOpen : kpiOpen = 0
Dim kpiInProgress : kpiInProgress = 0
Dim kpiPending : kpiPending = 0
Dim kpiCompleted : kpiCompleted = 0

Dim rsKpi
Set rsKpi = oConn.Execute("SELECT status, COUNT(*) AS cnt FROM cmms_work_orders GROUP BY status")
Do While Not rsKpi.EOF
    Select Case rsKpi("status")
        Case "open"        kpiOpen = rsKpi("cnt")
        Case "in_progress" kpiInProgress = rsKpi("cnt")
        Case "pending"     kpiPending = rsKpi("cnt")
        Case "completed"   kpiCompleted = rsKpi("cnt")
    End Select
    rsKpi.MoveNext
Loop
rsKpi.Close : Set rsKpi = Nothing

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = GetPerPage()
Dim offset      : offset = (currentPage - 1) * perPage

' Count total
Dim totalRows : totalRows = 0
Dim rsTotal
Set rsTotal = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_work_orders wo " & sqlWhere)
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing

' Get data
Dim sql
sql = "SELECT wo.*, a.name AS asset_name, p.name AS plant_name, u.first_name + ' ' + u.last_name AS assigned_name " & _
      "FROM cmms_work_orders wo " & _
      "LEFT JOIN cmms_assets a ON a.id = wo.asset_id " & _
      "LEFT JOIN cmms_plants p ON p.id = wo.plant_id " & _
      "LEFT JOIN cmms_users u ON u.id = wo.assigned_to_id " & _
      sqlWhere & _
      "ORDER BY wo.created_at DESC " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Dim rsWO
Set rsWO = oConn.Execute(sql)

' Filter Plants
Dim rsPlantsFilter
Set rsPlantsFilter = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")
%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Gestión y seguimiento de trabajos de mantenimiento</p>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/work_orders/form.asp" class="btn btn-primary">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
      Nueva Orden de Trabajo
    </a>
  </div>
</div>

<div class="stats-grid" style="margin-bottom:var(--space-lg)">
  <div class="stat-card primary">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiOpen %>">0</div>
      <div class="stat-label">Abiertas</div>
    </div>
  </div>
  <div class="stat-card warning">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiInProgress %>">0</div>
      <div class="stat-label">En Progreso</div>
    </div>
  </div>
  <div class="stat-card secondary">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiPending %>">0</div>
      <div class="stat-label">Pendientes</div>
    </div>
  </div>
  <div class="stat-card success">
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= kpiCompleted %>">0</div>
      <div class="stat-label">Completadas</div>
    </div>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="index.asp">
      <div class="search-wrap" style="flex:2">
        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" name="q" class="form-control search-input" placeholder="Buscar título o código..." value="<%= HtmlEncode(filterQuery) %>">
      </div>
      <div class="search-wrap" style="flex:1;min-width:140px">
        <select name="type" class="form-control">
          <option value="">Cualquier tipo</option>
          <option value="preventive" <%= IIf(filterType="preventive", "selected", "") %>>Preventivo</option>
          <option value="corrective" <%= IIf(filterType="corrective", "selected", "") %>>Correctivo</option>
          <option value="predictive" <%= IIf(filterType="predictive", "selected", "") %>>Predictivo</option>
          <option value="emergency" <%= IIf(filterType="emergency", "selected", "") %>>Emergencia</option>
        </select>
      </div>
      <div class="search-wrap" style="flex:1;min-width:140px">
        <select name="status" class="form-control">
          <option value="">Cualquier estado</option>
          <option value="open" <%= IIf(filterStatus="open", "selected", "") %>>Abierta</option>
          <option value="in_progress" <%= IIf(filterStatus="in_progress", "selected", "") %>>En Progreso</option>
          <option value="pending" <%= IIf(filterStatus="pending", "selected", "") %>>Pendiente</option>
          <option value="completed" <%= IIf(filterStatus="completed", "selected", "") %>>Completada</option>
          <option value="cancelled" <%= IIf(filterStatus="cancelled", "selected", "") %>>Cancelada</option>
        </select>
      </div>
      <div class="search-wrap" style="flex:1;min-width:140px">
        <select name="plant_id" class="form-control">
          <option value="">Todas las Plantas</option>
          <% Do While Not rsPlantsFilter.EOF %>
          <option value="<%= rsPlantsFilter("id") %>" <%= IIf(filterPlant=rsPlantsFilter("id"), "selected", "") %>><%= HtmlEncode(rsPlantsFilter("name")) %></option>
          <% rsPlantsFilter.MoveNext : Loop %>
        </select>
      </div>
      <button type="submit" class="btn btn-outline">Filtrar</button>
      <% If filterQuery <> "" Or filterType <> "" Or filterStatus <> "" Or filterPlant > 0 Then %>
      <a href="index.asp" class="btn btn-ghost">Limpiar</a>
      <% End If %>
    </form>
  </div>

  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Código</th>
          <th>Título / Equipo</th>
          <th>Planta</th>
          <th>Tipo</th>
          <th>Prioridad</th>
          <th>Estado</th>
          <th>Asignado a</th>
          <th>Programada</th>
          <th style="text-align:right">Acciones</th>
        </tr>
      </thead>
      <tbody>
        <% If rsWO.EOF Then %>
        <tr><td colspan="9" class="table-empty">No se encontraron órdenes de trabajo</td></tr>
        <% Else %>
        <% Do While Not rsWO.EOF %>
        <tr>
          <td class="bold"><a href="/CMMS/modules/work_orders/detail.asp?id=<%= rsWO("id") %>"><%= HtmlEncode(rsWO("code")) %></a></td>
          <td>
            <div style="font-weight:500;color:var(--text-primary)"><%= HtmlEncode(rsWO("title")) %></div>
            <div style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(NullStr(rsWO("asset_name"))) %></div>
          </td>
          <td><%= HtmlEncode(NullStr(rsWO("plant_name"))) %></td>
          <td>
            <%
            Dim typeTxt : typeTxt = "Desconocido"
            Select Case rsWO("type")
                Case "preventive": typeTxt = "Preventivo"
                Case "corrective": typeTxt = "Correctivo"
                Case "predictive": typeTxt = "Predictivo"
                Case "emergency":  typeTxt = "Emergencia"
            End Select
            %><span class="badge badge-muted"><%= typeTxt %></span>
          </td>
          <td><%= WOPriorityBadge(rsWO("priority")) %></td>
          <td><%= WOStatusBadge(rsWO("status")) %></td>
          <td>
            <% If IsNull(rsWO("assigned_name")) Then %>
              <span style="color:var(--text-muted);font-style:italic">Sin asignar</span>
            <% Else %>
              <div style="display:flex;align-items:center;gap:6px">
                <div class="user-avatar" style="width:20px;height:20px;font-size:9px"><%= UCase(Left(rsWO("assigned_name"),1)) %></div>
                <%= HtmlEncode(rsWO("assigned_name")) %>
              </div>
            <% End If %>
          </td>
          <td><%= FormatDateShort(rsWO("scheduled_start")) %></td>
          <td style="text-align:right;white-space:nowrap">
            <a href="/CMMS/modules/work_orders/detail.asp?id=<%= rsWO("id") %>" class="btn btn-ghost btn-icon" title="Ver Detalle"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></a>
            <% If IsSupervisorOrAdmin() Then %>
            <a href="/CMMS/modules/work_orders/form.asp?id=<%= rsWO("id") %>" class="btn btn-ghost btn-icon" title="Editar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></a>
            <a href="?action=delete&id=<%= rsWO("id") %>" class="btn btn-ghost btn-icon text-danger" data-confirm="¿Está seguro de cancelar esta orden?" title="Cancelar"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg></a>
            <% End If %>
          </td>
        </tr>
        <% 
        rsWO.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsWO.Close : Set rsWO = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "?q=" & Server.URLEncode(filterQuery) & "&type=" & filterType & "&status=" & filterStatus & "&plant_id=" & filterPlant) %>
</div>

<% 
rsPlantsFilter.Close : Set rsPlantsFilter = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
