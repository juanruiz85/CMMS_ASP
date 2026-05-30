<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = "Reportes"
Dim PageModule : PageModule = "reports"

Dim oConn : Set oConn = GetConnection()

' Get basic stats for the dashboard-like report view
Dim totalPlants : totalPlants = 0
Dim totalAssets : totalAssets = 0
Dim totalWO     : totalWO = 0
Dim completedWO : completedWO = 0

Dim rsStats
Set rsStats = oConn.Execute("SELECT (SELECT COUNT(*) FROM cmms_plants) as p, (SELECT COUNT(*) FROM cmms_assets) as a, (SELECT COUNT(*) FROM cmms_work_orders) as w, (SELECT COUNT(*) FROM cmms_work_orders WHERE status='completed') as c")
If Not rsStats.EOF Then
    totalPlants = rsStats("p")
    totalAssets = rsStats("a")
    totalWO = rsStats("w")
    completedWO = rsStats("c")
End If
rsStats.Close : Set rsStats = Nothing

' WO by type
Dim woByType : woByType = "[]"
Dim woByLbl : woByLbl = "[]"
Dim lblStr : lblStr = ""
Dim valStr : valStr = ""
Set rsStats = oConn.Execute("SELECT type, COUNT(*) as cnt FROM cmms_work_orders GROUP BY type")
Do While Not rsStats.EOF
    lblStr = lblStr & "'" & rsStats("type") & "',"
    valStr = valStr & rsStats("cnt") & ","
    rsStats.MoveNext
Loop
If Len(lblStr) > 0 Then woByLbl = "[" & Left(lblStr, Len(lblStr)-1) & "]"
If Len(valStr) > 0 Then woByType = "[" & Left(valStr, Len(valStr)-1) & "]"
rsStats.Close : Set rsStats = Nothing

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Análisis e indicadores clave de rendimiento (KPIs)</p>
  </div>
</div>

<div class="dashboard-grid">
  <!-- General Stats -->
  <div class="col-12">
    <div class="stats-grid">
      <div class="stat-card primary">
        <div class="stat-info">
          <div class="stat-value" data-counter="<%= totalWO %>">0</div>
          <div class="stat-label">Total OTs</div>
        </div>
      </div>
      <div class="stat-card success">
        <div class="stat-info">
          <div class="stat-value" data-counter="<%= completedWO %>">0</div>
          <div class="stat-label">OTs Completadas</div>
        </div>
      </div>
      <div class="stat-card warning">
        <div class="stat-info">
          <div class="stat-value"><%= IIf(totalWO>0, Round((completedWO/totalWO)*100, 1), 0) %>%</div>
          <div class="stat-label">Tasa de Finalización</div>
        </div>
      </div>
    </div>
  </div>

  <!-- Reports List -->
  <div class="col-8">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Reportes Disponibles</h3>
      </div>
      <div class="card-body">
        
        <div style="display:flex;flex-direction:column;gap:16px">
          
          <div style="border:1px solid var(--border-subtle);border-radius:8px;padding:16px;display:flex;align-items:center;justify-content:space-between">
            <div>
              <h4 style="font-size:16px;margin-bottom:4px">Reporte de Órdenes de Trabajo</h4>
              <p style="font-size:13px;color:var(--text-muted)">Exportación detallada de todas las órdenes de trabajo con filtros por fecha y planta.</p>
            </div>
            <a href="export.asp?type=wo" class="btn btn-outline" target="_blank">Exportar CSV</a>
          </div>

          <div style="border:1px solid var(--border-subtle);border-radius:8px;padding:16px;display:flex;align-items:center;justify-content:space-between">
            <div>
              <h4 style="font-size:16px;margin-bottom:4px">Inventario Actual</h4>
              <p style="font-size:13px;color:var(--text-muted)">Listado completo de artículos, stock disponible, costos y puntos de reorden.</p>
            </div>
            <a href="export.asp?type=inventory" class="btn btn-outline" target="_blank">Exportar CSV</a>
          </div>

          <div style="border:1px solid var(--border-subtle);border-radius:8px;padding:16px;display:flex;align-items:center;justify-content:space-between">
            <div>
              <h4 style="font-size:16px;margin-bottom:4px">Catálogo de Equipos</h4>
              <p style="font-size:13px;color:var(--text-muted)">Listado maestro de todos los activos, ubicaciones y criticidad.</p>
            </div>
            <a href="export.asp?type=assets" class="btn btn-outline" target="_blank">Exportar CSV</a>
          </div>

        </div>

      </div>
    </div>
  </div>

  <!-- Charts (Mock) -->
  <div class="col-4">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">OTs por Tipo</h3>
      </div>
      <div class="card-body" style="display:flex;align-items:center;justify-content:center;min-height:250px">
        <div style="width:200px;height:200px;border-radius:50%;background:conic-gradient(var(--primary) 0% 40%, var(--secondary) 40% 70%, var(--warning) 70% 90%, var(--danger) 90% 100%);display:flex;align-items:center;justify-content:center">
          <div style="width:120px;height:120px;background:var(--bg-card);border-radius:50%"></div>
        </div>
      </div>
    </div>
  </div>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
