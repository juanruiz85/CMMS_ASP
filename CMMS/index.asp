<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = T("nav_dashboard")
Dim PageModule : PageModule = "dashboard"

Dim oConn : Set oConn = GetConnection()
Dim oRS

' KPI Stats
Dim totalPlants : totalPlants = 0
Dim totalAssets : totalAssets = 0
Dim totalWO     : totalWO = 0
Dim openWO      : openWO = 0

Set oRS = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_plants WHERE status = 'active'")
If Not oRS.EOF Then totalPlants = oRS("cnt")
oRS.Close

Set oRS = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_assets WHERE status != 'retired'")
If Not oRS.EOF Then totalAssets = oRS("cnt")
oRS.Close

Set oRS = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_work_orders")
If Not oRS.EOF Then totalWO = oRS("cnt")
oRS.Close

Set oRS = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_work_orders WHERE status IN ('open', 'in_progress', 'pending')")
If Not oRS.EOF Then openWO = oRS("cnt")
oRS.Close

' Recent Work Orders
Dim recentWO
Set recentWO = oConn.Execute("SELECT TOP 5 wo.id, wo.code, wo.title, wo.status, wo.priority, wo.created_at, a.name AS asset_name FROM cmms_work_orders wo LEFT JOIN cmms_assets a ON a.id = wo.asset_id ORDER BY wo.created_at DESC")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= T("nav_dashboard") %></h1>
    <p class="page-subtitle">Bienvenido, <%= HtmlEncode(CurrentUserFullName()) %></p>
  </div>
</div>

<div class="stats-grid">
  <div class="stat-card primary">
    <div class="stat-icon primary">
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 21h18M5 21V7l7-4 7 4v14M9 21V11h6v10"/></svg>
    </div>
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= totalPlants %>">0</div>
      <div class="stat-label">Plantas Activas</div>
    </div>
  </div>
  
  <div class="stat-card secondary">
    <div class="stat-icon secondary">
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.07 4.93a10 10 0 0 1 0 14.14M4.93 4.93a10 10 0 0 0 0 14.14M16.24 7.76a6 6 0 0 1 0 8.49M7.76 7.76a6 6 0 0 0 0 8.49"/></svg>
    </div>
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= totalAssets %>">0</div>
      <div class="stat-label">Equipos</div>
    </div>
  </div>

  <div class="stat-card warning">
    <div class="stat-icon warning">
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="1"/><line x1="9" y1="12" x2="15" y2="12"/><line x1="9" y1="16" x2="13" y2="16"/></svg>
    </div>
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= openWO %>">0</div>
      <div class="stat-label">OTs Abiertas</div>
    </div>
  </div>

  <div class="stat-card success">
    <div class="stat-icon success">
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
    </div>
    <div class="stat-info">
      <div class="stat-value" data-counter="<%= totalWO %>">0</div>
      <div class="stat-label">Total OTs</div>
    </div>
  </div>
</div>

<div class="dashboard-grid">
  <div class="col-8">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Órdenes de Trabajo Recientes</h3>
      </div>
      <div class="card-body" style="padding:0">
        <div class="table-responsive">
          <table class="table">
            <thead>
              <tr>
                <th>Código</th>
                <th>Título</th>
                <th>Equipo</th>
                <th>Prioridad</th>
                <th>Estado</th>
                <th>Fecha</th>
              </tr>
            </thead>
            <tbody>
              <% If recentWO.EOF Then %>
              <tr><td colspan="6" class="table-empty">No hay órdenes de trabajo recientes</td></tr>
              <% Else %>
              <% Do While Not recentWO.EOF %>
              <tr>
                <td class="bold"><a href="/CMMS/modules/work_orders/detail.asp?id=<%= recentWO("id") %>"><%= HtmlEncode(recentWO("code")) %></a></td>
                <td><%= HtmlEncode(recentWO("title")) %></td>
                <td><%= HtmlEncode(NullStr(recentWO("asset_name"))) %></td>
                <td><%= WOPriorityBadge(recentWO("priority")) %></td>
                <td><%= WOStatusBadge(recentWO("status")) %></td>
                <td><%= FormatDateShort(recentWO("created_at")) %></td>
              </tr>
              <% 
              recentWO.MoveNext
              Loop
              %>
              <% End If %>
              <% recentWO.Close : Set recentWO = Nothing %>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <div class="col-4">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Acciones Rápidas</h3>
      </div>
      <div class="card-body">
        <div class="quick-actions">
          <a href="/CMMS/modules/work_orders/form.asp" class="quick-action">
            <div class="quick-action-icon" style="color:var(--primary)">+</div>
            <div class="quick-action-label">Nueva OT</div>
          </a>
          <% If IsSupervisorOrAdmin() Then %>
          <a href="/CMMS/modules/assets_module/form.asp" class="quick-action">
            <div class="quick-action-icon" style="color:var(--secondary)">+</div>
            <div class="quick-action-label">Nuevo Equipo</div>
          </a>
          <% End If %>
        </div>
      </div>
    </div>
  </div>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
