<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = "Solicitudes de Trabajo"
Dim PageModule : PageModule = "work_requests"

Dim oConn : Set oConn = GetConnection()

' Filters
Dim filterStatus : filterStatus = Trim(Request.QueryString("status"))
Dim filterPlant  : filterPlant  = QSInt("plant_id")

' Build WHERE clause with parameters
Dim sqlWhere : sqlWhere = " WHERE 1=1 "
Dim paramCount : paramCount = 0
Dim whereParams : whereParams = ""

If Not IsSupervisorOrAdmin() Then
    ' Los usuarios normales solo ven sus propias solicitudes
    sqlWhere = sqlWhere & " AND r.requested_by_id = ? "
    paramCount = paramCount + 1
    whereParams = whereParams & CurrentUserId() & "|"
End If

If filterStatus <> "" Then
    sqlWhere = sqlWhere & " AND r.status = ? "
    paramCount = paramCount + 1
    whereParams = whereParams & filterStatus & "|"
End If

If filterPlant > 0 Then
    sqlWhere = sqlWhere & " AND r.plant_id = ? "
    paramCount = paramCount + 1
    whereParams = whereParams & filterPlant & "|"
End If

' Paginacion
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = GetPerPage()
Dim offset      : offset = (currentPage - 1) * perPage

' Count with parameters
Dim totalRows : totalRows = 0
Dim cmdTotal, rsTotal
Set cmdTotal = Server.CreateObject("ADODB.Command")
cmdTotal.ActiveConnection = oConn
cmdTotal.CommandText = "SELECT COUNT(*) AS cnt FROM cmms_work_requests r " & sqlWhere

' Add parameters for count
Dim iParam, paramArr
If whereParams <> "" Then
    paramArr = Split(whereParams, "|")
    For iParam = 0 To UBound(paramArr) - 1
        If iParam < 1 Or IsNumeric(paramArr(iParam)) Then
            cmdTotal.Parameters.Append cmdTotal.CreateParameter("@p" & iParam, 3, 1, , paramArr(iParam))
        Else
            cmdTotal.Parameters.Append cmdTotal.CreateParameter("@p" & iParam, 200, 1, 255, paramArr(iParam))
        End If
    Next
End If

Set rsTotal = cmdTotal.Execute()
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing
Set cmdTotal = Nothing

' Get data with parameters
Dim sql, cmdReq, rsReq
sql = "SELECT r.*, p.name AS plant_name, a.name AS asset_name, u.first_name + ' ' + u.last_name AS requested_by, wo.code AS wo_code " & _
      "FROM cmms_work_requests r " & _
      "LEFT JOIN cmms_plants p ON p.id = r.plant_id " & _
      "LEFT JOIN cmms_assets a ON a.id = r.asset_id " & _
      "LEFT JOIN cmms_users u ON u.id = r.requested_by_id " & _
      "LEFT JOIN cmms_work_orders wo ON wo.id = r.work_order_id " & _
      sqlWhere & _
      "ORDER BY r.created_at DESC " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Set cmdReq = Server.CreateObject("ADODB.Command")
cmdReq.ActiveConnection = oConn
cmdReq.CommandText = sql

' Re-add parameters for main query
If whereParams <> "" Then
    paramArr = Split(whereParams, "|")
    For iParam = 0 To UBound(paramArr) - 1
        If iParam < 1 Or IsNumeric(paramArr(iParam)) Then
            cmdReq.Parameters.Append cmdReq.CreateParameter("@p" & iParam, 3, 1, , paramArr(iParam))
        Else
            cmdReq.Parameters.Append cmdReq.CreateParameter("@p" & iParam, 200, 1, 255, paramArr(iParam))
        End If
    Next
End If

Set rsReq = cmdReq.Execute()
Set cmdReq = Nothing

Dim rsPlants
Set rsPlants = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")
%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Bandeja de solicitudes de mantenimiento</p>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/work_requests/form.asp" class="btn btn-primary">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
      Nueva Solicitud
    </a>
  </div>
</div>

<div class="card">
  <div class="card-body" style="padding-bottom:0">
    <form class="filters-bar" method="GET" action="index.asp">
      <div class="search-wrap" style="flex:1">
        <select name="status" class="form-control">
          <option value="">Cualquier estado</option>
          <option value="pending" <%= IIf(filterStatus="pending", "selected", "") %>>Pendiente de Aprobación</option>
          <option value="approved" <%= IIf(filterStatus="approved", "selected", "") %>>Aprobada (OT Generada)</option>
          <option value="rejected" <%= IIf(filterStatus="rejected", "selected", "") %>>Rechazada</option>
        </select>
      </div>
      <div class="search-wrap" style="flex:1">
        <select name="plant_id" class="form-control">
          <option value="">Todas las Plantas</option>
          <% Do While Not rsPlants.EOF %>
          <option value="<%= rsPlants("id") %>" <%= IIf(filterPlant=rsPlants("id"), "selected", "") %>><%= HtmlEncode(rsPlants("name")) %></option>
          <% rsPlants.MoveNext : Loop %>
        </select>
      </div>
      <button type="submit" class="btn btn-outline">Filtrar</button>
      <% If filterStatus <> "" Or filterPlant > 0 Then %>
      <a href="index.asp" class="btn btn-ghost">Limpiar</a>
      <% End If %>
    </form>
  </div>

  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Fecha</th>
          <th>Título</th>
          <th>Planta / Equipo</th>
          <th>Solicitante</th>
          <th>Estado</th>
          <th>OT Asignada</th>
          <th style="text-align:right">Acciones</th>
        </tr>
      </thead>
      <tbody>
        <% If rsReq.EOF Then %>
        <tr><td colspan="7" class="table-empty">No se encontraron solicitudes</td></tr>
        <% Else %>
        <% Do While Not rsReq.EOF %>
        <tr>
          <td><%= FormatDateShort(rsReq("created_at")) %></td>
          <td style="font-weight:500"><%= HtmlEncode(rsReq("title")) %></td>
          <td>
            <div style="font-size:12px"><%= HtmlEncode(rsReq("plant_name")) %></div>
            <div style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(NullStr(rsReq("asset_name"))) %></div>
          </td>
          <td><%= HtmlEncode(rsReq("requested_by")) %></td>
          <td>
            <%
            Dim rCls, rLbl
            Select Case rsReq("status")
                Case "pending": rCls = "badge-warning": rLbl = "Pendiente"
                Case "approved": rCls = "badge-success": rLbl = "Aprobada"
                Case "rejected": rCls = "badge-danger": rLbl = "Rechazada"
            End Select
            %>
            <span class="badge <%= rCls %>"><%= rLbl %></span>
          </td>
          <td>
            <% If Not IsNull(rsReq("wo_code")) Then %>
              <a href="/CMMS/modules/work_orders/detail.asp?id=<%= rsReq("work_order_id") %>" style="font-weight:600"><%= rsReq("wo_code") %></a>
            <% Else %>
              <span style="color:var(--text-muted);font-size:12px">N/A</span>
            <% End If %>
          </td>
          <td style="text-align:right">
            <% If rsReq("status") = "pending" And IsSupervisorOrAdmin() Then %>
              <a href="approve.asp?id=<%= rsReq("id") %>" class="btn btn-primary btn-sm">Evaluar</a>
            <% Else %>
              <a href="approve.asp?id=<%= rsReq("id") %>" class="btn btn-outline btn-sm">Ver Detalle</a>
            <% End If %>
          </td>
        </tr>
        <% 
        rsReq.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsReq.Close : Set rsReq = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "?status=" & filterStatus & "&plant_id=" & filterPlant) %>
</div>

<% 
rsPlants.Close : Set rsPlants = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
