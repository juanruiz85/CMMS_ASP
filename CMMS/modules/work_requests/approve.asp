<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim reqId : reqId = QSInt("id")
If reqId = 0 Then RedirectTo "index.asp"

Dim oConn : Set oConn = GetConnection()

' Get request data
Dim sql
sql = "SELECT r.*, p.name AS plant_name, a.name AS asset_name, a.code AS asset_code, u.first_name + ' ' + u.last_name AS requested_by " & _
      "FROM cmms_work_requests r " & _
      "LEFT JOIN cmms_plants p ON p.id = r.plant_id " & _
      "LEFT JOIN cmms_assets a ON a.id = r.asset_id " & _
      "LEFT JOIN cmms_users u ON u.id = r.requested_by_id " & _
      "WHERE r.id = " & reqId

Dim rsReq
Set rsReq = oConn.Execute(sql)
If rsReq.EOF Then
    rsReq.Close : Set rsReq = Nothing
    CloseConnection oConn
    SetFlashMessage "danger", "Solicitud no encontrada."
    RedirectTo "index.asp"
End If

' Handle Form Submit
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    On Error Resume Next
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token inválido."
        RedirectTo "approve.asp?id=" & reqId
    End If
    
    If Not IsSupervisorOrAdmin() Then
        SetFlashMessage "danger", "No tiene permisos para aprobar/rechazar."
        RedirectTo "index.asp"
    End If
    
    Dim actionType : actionType = FormStr("action")
    Dim adminNotes : adminNotes = Trim(Request.Form("admin_notes"))
    
    If actionType = "approve" Then
        ' Generate WO Code
        Dim woCode : woCode = GenerateWOCode()
        
        ' Create Work Order
        Dim cmdWO
        Set cmdWO = Server.CreateObject("ADODB.Command")
        cmdWO.ActiveConnection = oConn
        cmdWO.CommandText = "INSERT INTO cmms_work_orders (code, title, description, asset_id, plant_id, requester_id, type, priority, status, notes, created_at) " & _
                            "VALUES (?, ?, ?, ?, ?, ?, 'corrective', 'medium', 'open', ?, GETDATE())"
        cmdWO.Parameters.Append cmdWO.CreateParameter("@code", 200, 1, 50, woCode)
        cmdWO.Parameters.Append cmdWO.CreateParameter("@tit", 200, 1, 255, rsReq("title"))
        cmdWO.Parameters.Append cmdWO.CreateParameter("@desc", 200, 1, -1, rsReq("description"))
        cmdWO.Parameters.Append cmdWO.CreateParameter("@asset", 3, 1, , IIf(IsNull(rsReq("asset_id")), Null, rsReq("asset_id")))
        cmdWO.Parameters.Append cmdWO.CreateParameter("@plant", 3, 1, , rsReq("plant_id"))
        cmdWO.Parameters.Append cmdWO.CreateParameter("@req", 3, 1, , rsReq("requested_by_id"))
        cmdWO.Parameters.Append cmdWO.CreateParameter("@not", 200, 1, -1, adminNotes)
        cmdWO.Execute
        Set cmdWO = Nothing
        
        ' Get inserted WO ID
        Dim rsLast
        Set rsLast = oConn.Execute("SELECT IDENT_CURRENT('cmms_work_orders') AS last_id")
        Dim woId : woId = rsLast("last_id")
        rsLast.Close : Set rsLast = Nothing
        
        ' Update Request
        Dim cmdReq
        Set cmdReq = Server.CreateObject("ADODB.Command")
        cmdReq.ActiveConnection = oConn
        cmdReq.CommandText = "UPDATE cmms_work_requests SET status='approved', approved_at=GETDATE(), approved_by_id=?, work_order_id=? WHERE id=?"
        cmdReq.Parameters.Append cmdReq.CreateParameter("@appr", 3, 1, , CurrentUserId())
        cmdReq.Parameters.Append cmdReq.CreateParameter("@woid", 3, 1, , woId)
        cmdReq.Parameters.Append cmdReq.CreateParameter("@id", 3, 1, , reqId)
        cmdReq.Execute
        Set cmdReq = Nothing
        
        CheckError "Aprobar solicitud de trabajo y generar OT"
        
        CreateNotification rsReq("requested_by_id"), "Solicitud Aprobada", "Tu solicitud se convirtió en la Orden de Trabajo " & woCode, "success", "work_orders", woId
        LogActivity CurrentUserId(), "APPROVE_REQUEST", "Aprobó solicitud y generó OT " & woCode, "work_requests", reqId
        
        SetFlashMessage "success", "Solicitud aprobada y Orden de Trabajo " & woCode & " generada correctamente."
        
    ElseIf actionType = "reject" Then
        ' Update Request
        Dim cmdRej
        Set cmdRej = Server.CreateObject("ADODB.Command")
        cmdRej.ActiveConnection = oConn
        cmdRej.CommandText = "UPDATE cmms_work_requests SET status='rejected', approved_at=GETDATE(), approved_by_id=? WHERE id=?"
        cmdRej.Parameters.Append cmdRej.CreateParameter("@appr", 3, 1, , CurrentUserId())
        cmdRej.Parameters.Append cmdRej.CreateParameter("@id", 3, 1, , reqId)
        cmdRej.Execute
        Set cmdRej = Nothing
        
        CheckError "Rechazar solicitud de trabajo"
        
        CreateNotification rsReq("requested_by_id"), "Solicitud Rechazada", "Tu solicitud de trabajo fue rechazada. Razón: " & Left(adminNotes, 50), "danger", "work_requests", reqId
        LogActivity CurrentUserId(), "REJECT_REQUEST", "Rechazó solicitud", "work_requests", reqId
        
        SetFlashMessage "warning", "La solicitud ha sido rechazada."
    End If
    
    RedirectTo "index.asp"
End If

Dim PageTitle : PageTitle = "Detalle de Solicitud"
Dim PageModule : PageModule = "work_requests"

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
  </div>
  <div class="page-actions">
    <a href="index.asp" class="btn btn-outline">Volver a la lista</a>
  </div>
</div>

<div class="dashboard-grid">
  <div class="col-8">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Información de la Solicitud</h3>
        <%
        Dim rCls, rLbl
        Select Case rsReq("status")
            Case "pending": rCls = "badge-warning": rLbl = "Pendiente"
            Case "approved": rCls = "badge-success": rLbl = "Aprobada"
            Case "rejected": rCls = "badge-danger": rLbl = "Rechazada"
        End Select
        %>
        <span class="badge <%= rCls %>" style="margin-left:auto"><%= rLbl %></span>
      </div>
      <div class="card-body">
        <div class="detail-grid">
          <div class="detail-item">
            <label>Título</label>
            <div style="font-weight:600;font-size:16px"><%= HtmlEncode(rsReq("title")) %></div>
          </div>
          <div class="detail-item">
            <label>Fecha de Solicitud</label>
            <div><%= FormatDateTime(rsReq("created_at")) %></div>
          </div>
          
          <div class="detail-item">
            <label>Planta / Ubicación</label>
            <div style="font-weight:500"><%= HtmlEncode(rsReq("plant_name")) %></div>
          </div>
          <div class="detail-item">
            <label>Equipo Afectado</label>
            <% If Not IsNull(rsReq("asset_id")) Then %>
            <div><%= HtmlEncode(rsReq("asset_code")) %> - <%= HtmlEncode(rsReq("asset_name")) %></div>
            <% Else %>
            <div style="color:var(--text-muted)">General / Sin equipo específico</div>
            <% End If %>
          </div>
          
          <div class="detail-item" style="grid-column:1/-1">
            <label>Descripción del Problema</label>
            <div style="background:var(--bg-body);padding:12px;border-radius:6px;border:1px solid var(--border-subtle);min-height:80px;white-space:pre-wrap"><%= HtmlEncode(rsReq("description")) %></div>
          </div>
          
          <div class="detail-item">
            <label>Solicitado Por</label>
            <div><%= HtmlEncode(rsReq("requested_by")) %></div>
          </div>
        </div>
      </div>
    </div>
  </div>
  
  <div class="col-4">
    <% If rsReq("status") = "pending" And IsSupervisorOrAdmin() Then %>
    <div class="card">
      <div class="card-header">
        <h3 class="card-title">Acciones de Supervisor</h3>
      </div>
      <div class="card-body">
        <form method="POST" action="approve.asp?id=<%= reqId %>">
          <%= CSRFField() %>
          
          <div class="form-group">
            <label class="form-label">Notas del Supervisor (Opcional, se verán en la OT si se aprueba)</label>
            <textarea name="admin_notes" class="form-control" rows="3"></textarea>
          </div>
          
          <div style="display:flex;gap:10px;margin-top:20px">
            <button type="submit" name="action" value="approve" class="btn btn-success" style="flex:1">Aprobar y Generar OT</button>
            <button type="submit" name="action" value="reject" class="btn btn-danger" style="flex:1" onclick="return confirm('¿Está seguro de rechazar esta solicitud?');">Rechazar</button>
          </div>
        </form>
      </div>
    </div>
    <% ElseIf rsReq("status") <> "pending" Then %>
    <div class="card">
      <div class="card-header">
        <h3 class="card-title">Resolución</h3>
      </div>
      <div class="card-body">
        <div class="detail-item">
          <label>Fecha de Resolución</label>
          <div><%= FormatDateTime(rsReq("approved_at")) %></div>
        </div>
        
        <% If rsReq("status") = "approved" Then %>
        <div class="detail-item" style="margin-top:16px">
          <label>Orden de Trabajo Generada</label>
          <a href="/CMMS/modules/work_orders/detail.asp?id=<%= rsReq("work_order_id") %>" class="btn btn-outline" style="width:100%;margin-top:8px">Ver Orden de Trabajo</a>
        </div>
        <% End If %>
      </div>
    </div>
    <% End If %>
  </div>
</div>

<% 
rsReq.Close : Set rsReq = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
