<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim itemId : itemId = QSInt("id")
If itemId = 0 Then RedirectTo "/CMMS/modules/work_orders/index.asp"

Dim oConn : Set oConn = GetConnection()
Dim rsWO
Set rsWO = oConn.Execute("SELECT wo.*, a.name AS asset_name, p.name AS plant_name, u.first_name + ' ' + u.last_name AS assigned_name, req.first_name + ' ' + req.last_name AS requester_name FROM cmms_work_orders wo LEFT JOIN cmms_assets a ON a.id = wo.asset_id LEFT JOIN cmms_plants p ON p.id = wo.plant_id LEFT JOIN cmms_users u ON u.id = wo.assigned_to_id LEFT JOIN cmms_users req ON req.id = wo.requester_id WHERE wo.id = " & itemId)

If rsWO.EOF Then
    SetFlashMessage "danger", "Orden de Trabajo no encontrada."
    RedirectTo "/CMMS/modules/work_orders/index.asp"
End If

Dim PageTitle : PageTitle = "OT: " & rsWO("code")
Dim PageModule : PageModule = "work_orders"

' Post actions for tabs
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim tabAction : tabAction = Request.Form("action")
    
    If tabAction = "add_comment" Then
        Dim commentText : commentText = Trim(Request.Form("comment"))
        If commentText <> "" Then
            Dim cmdC
            Set cmdC = Server.CreateObject("ADODB.Command")
            cmdC.ActiveConnection = oConn
            cmdC.CommandText = "INSERT INTO cmms_work_order_comments (work_order_id, user_id, comment, created_at) VALUES (?, ?, ?, GETDATE())"
            cmdC.Parameters.Append cmdC.CreateParameter("@wo", 3, 1, , itemId)
            cmdC.Parameters.Append cmdC.CreateParameter("@uid", 3, 1, , CurrentUserId())
            cmdC.Parameters.Append cmdC.CreateParameter("@c", 200, 1, 4000, commentText)
            cmdC.Execute
            SetFlashMessage "success", "Comentario agregado."
        End If
    ElseIf tabAction = "add_time" Then
        Dim tHours : tHours = SafeDecimal(Request.Form("hours"))
        Dim tDesc : tDesc = Trim(Request.Form("description"))
        If tHours > 0 Then
            Dim cmdT
            Set cmdT = Server.CreateObject("ADODB.Command")
            cmdT.ActiveConnection = oConn
            cmdT.CommandText = "INSERT INTO cmms_work_order_time_logs (work_order_id, user_id, hours, description, created_at) VALUES (?, ?, " & tHours & ", ?, GETDATE())"
            cmdT.Parameters.Append cmdT.CreateParameter("@wo", 3, 1, , itemId)
            cmdT.Parameters.Append cmdT.CreateParameter("@uid", 3, 1, , CurrentUserId())
            cmdT.Parameters.Append cmdT.CreateParameter("@d", 200, 1, 500, tDesc)
            cmdT.Execute
            ' Actualizar horas reales de la OT
            oConn.Execute("UPDATE cmms_work_orders SET actual_hours = (SELECT SUM(hours) FROM cmms_work_order_time_logs WHERE work_order_id = " & itemId & ") WHERE id = " & itemId)
            SetFlashMessage "success", "Tiempo registrado."
        End If
    ElseIf tabAction = "change_status" And IsSupervisorOrAdmin() Then
        Dim newStatus : newStatus = Trim(Request.Form("status"))
        If newStatus <> "" Then
            oConn.Execute("UPDATE cmms_work_orders SET status='" & Replace(newStatus, "'", "''") & "' WHERE id=" & itemId)
            SetFlashMessage "success", "Estado actualizado."
            If newStatus = "completed" Then
                oConn.Execute("UPDATE cmms_work_orders SET completed_at=GETDATE(), closed_by_id=" & CurrentUserId() & " WHERE id=" & itemId & " AND completed_at IS NULL")
            End If
        End If
    End If
    RedirectTo "detail.asp?id=" & itemId & "&tab=" & Request.Form("current_tab")
End If

Dim activeTab : activeTab = Request.QueryString("tab")
If activeTab = "" Then activeTab = "info"

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= HtmlEncode(rsWO("code")) %>: <%= HtmlEncode(rsWO("title")) %></h1>
    <p class="page-subtitle">Planta: <%= HtmlEncode(NullStr(rsWO("plant_name"))) %> | Creado: <%= FormatDateShort(rsWO("created_at")) %></p>
  </div>
  <div class="page-actions">
    <%= WOStatusBadge(rsWO("status")) %>
    <% If IsSupervisorOrAdmin() Then %>
    <button class="btn btn-primary" onclick="openModal('modalStatus')">Cambiar Estado</button>
    <a href="/CMMS/modules/work_orders/form.asp?id=<%= itemId %>" class="btn btn-outline">Editar OT</a>
    <% End If %>
  </div>
</div>

<div class="tabs">
  <button class="tab-btn <%= IIf(activeTab="info", "active", "") %>" data-tab-group="wo" data-tab="info">Información</button>
  <button class="tab-btn <%= IIf(activeTab="comments", "active", "") %>" data-tab-group="wo" data-tab="comments">Comentarios</button>
  <button class="tab-btn <%= IIf(activeTab="time", "active", "") %>" data-tab-group="wo" data-tab="time">Tiempo y Trabajo</button>
</div>

<div class="card">
  <div class="card-body">
    
    <!-- TAB INFO -->
    <div class="tab-panel <%= IIf(activeTab="info", "active", "") %>" data-panel-group="wo" data-panel="info">
      <div class="detail-grid">
        <div class="detail-item">
          <label>Equipo</label>
          <span><%= IIf(IsNull(rsWO("asset_name")), "General Planta", HtmlEncode(rsWO("asset_name"))) %></span>
        </div>
        <div class="detail-item">
          <label>Asignado a</label>
          <span><%= IIf(IsNull(rsWO("assigned_name")), "Sin asignar", HtmlEncode(rsWO("assigned_name"))) %></span>
        </div>
        <div class="detail-item">
          <label>Tipo</label>
          <span><%= HtmlEncode(rsWO("type")) %></span>
        </div>
        <div class="detail-item">
          <label>Prioridad</label>
          <span><%= WOPriorityBadge(rsWO("priority")) %></span>
        </div>
        <div class="detail-item">
          <label>Solicitado por</label>
          <span><%= HtmlEncode(NullStr(rsWO("requester_name"))) %></span>
        </div>
        <div class="detail-item">
          <label>Fechas Programadas</label>
          <span><%= FormatDateShort(rsWO("scheduled_start")) %> al <%= FormatDateShort(rsWO("scheduled_end")) %></span>
        </div>
        <div class="detail-item">
          <label>Horas Estimadas / Reales</label>
          <span><%= NullInt(rsWO("estimated_hours")) %> / <%= NullInt(rsWO("actual_hours")) %></span>
        </div>
      </div>
      
      <div style="margin-top:24px">
        <label style="font-size:12px;color:var(--text-muted);font-weight:600;text-transform:uppercase;letter-spacing:0.05em">Descripción</label>
        <div style="background:var(--bg-elevated);border-radius:8px;padding:16px;margin-top:8px;white-space:pre-wrap;font-size:14px"><%= HtmlEncode(NullStr(rsWO("description"))) %></div>
      </div>
    </div>

    <!-- TAB COMMS -->
    <div class="tab-panel <%= IIf(activeTab="comments", "active", "") %>" data-panel-group="wo" data-panel="comments">
      <%
      Dim rsComments
      Set rsComments = oConn.Execute("SELECT c.*, u.first_name + ' ' + u.last_name AS user_name FROM cmms_work_order_comments c JOIN cmms_users u ON u.id = c.user_id WHERE c.work_order_id = " & itemId & " ORDER BY c.created_at ASC")
      If rsComments.EOF Then
      %>
      <div class="table-empty">No hay comentarios en esta orden.</div>
      <% Else %>
      <div style="margin-bottom:24px">
        <% Do While Not rsComments.EOF %>
        <div class="comment-item">
          <div class="comment-avatar"><%= UCase(Left(rsComments("user_name"),1)) %></div>
          <div>
            <div class="comment-meta"><strong><%= HtmlEncode(rsComments("user_name")) %></strong> • <%= TimeAgo(rsComments("created_at")) %></div>
            <div class="comment-text"><%= HtmlEncode(rsComments("comment")) %></div>
          </div>
        </div>
        <% rsComments.MoveNext : Loop %>
      </div>
      <% End If : rsComments.Close : Set rsComments = Nothing %>

      <form method="POST" action="detail.asp?id=<%= itemId %>">
        <input type="hidden" name="action" value="add_comment">
        <input type="hidden" name="current_tab" value="comments">
        <div class="form-group">
          <textarea name="comment" class="form-control" placeholder="Escriba un comentario..." required></textarea>
        </div>
        <button type="submit" class="btn btn-primary">Enviar Comentario</button>
      </form>
    </div>

    <!-- TAB TIME -->
    <div class="tab-panel <%= IIf(activeTab="time", "active", "") %>" data-panel-group="wo" data-panel="time">
      <div style="display:flex;gap:24px;margin-bottom:24px">
        <div style="flex:1">
          <h4 style="margin-bottom:12px">Horas Registradas</h4>
          <table class="table">
            <thead><tr><th>Fecha</th><th>Usuario</th><th>Horas</th><th>Descripción</th></tr></thead>
            <tbody>
              <%
              Dim rsTime, tTotal : tTotal = 0
              Set rsTime = oConn.Execute("SELECT t.*, u.first_name + ' ' + u.last_name AS user_name FROM cmms_work_order_time_logs t JOIN cmms_users u ON u.id = t.user_id WHERE t.work_order_id = " & itemId & " ORDER BY t.created_at DESC")
              If rsTime.EOF Then
              %>
              <tr><td colspan="4" class="table-empty">No se ha registrado tiempo</td></tr>
              <% Else
              Do While Not rsTime.EOF
                tTotal = tTotal + rsTime("hours")
              %>
              <tr>
                <td><%= FormatDateShort(rsTime("created_at")) %></td>
                <td><%= HtmlEncode(rsTime("user_name")) %></td>
                <td class="bold text-primary"><%= rsTime("hours") %> h</td>
                <td><%= HtmlEncode(NullStr(rsTime("description"))) %></td>
              </tr>
              <% rsTime.MoveNext : Loop : End If : rsTime.Close : Set rsTime = Nothing %>
            </tbody>
            <tfoot>
              <tr><th colspan="2" style="text-align:right">Total Horas:</th><th colspan="2" class="bold text-primary"><%= tTotal %> h</th></tr>
            </tfoot>
          </table>
        </div>
        <div style="flex:1;max-width:350px">
          <div style="background:var(--bg-elevated);padding:16px;border-radius:8px;border:1px solid var(--border-subtle)">
            <h4 style="margin-bottom:12px">Registrar Trabajo</h4>
            <form method="POST" action="detail.asp?id=<%= itemId %>">
              <input type="hidden" name="action" value="add_time">
              <input type="hidden" name="current_tab" value="time">
              <div class="form-group">
                <label class="form-label">Horas Empleadas</label>
                <input type="number" step="0.5" name="hours" class="form-control" required placeholder="0.5">
              </div>
              <div class="form-group">
                <label class="form-label">Descripción de las tareas realizadas</label>
                <textarea name="description" class="form-control" rows="3" required></textarea>
              </div>
              <button type="submit" class="btn btn-secondary w-full">Registrar Horas</button>
            </form>
          </div>
        </div>
      </div>
    </div>

  </div>
</div>

<!-- Modal Cambiar Estado -->
<div id="modalStatus" class="modal-overlay">
  <div class="modal modal-sm">
    <div class="modal-header">
      <h3 class="modal-title">Cambiar Estado de OT</h3>
      <button class="modal-close" onclick="closeModal('modalStatus')">×</button>
    </div>
    <div class="modal-body">
      <form method="POST" action="detail.asp?id=<%= itemId %>" id="formStatus">
        <input type="hidden" name="action" value="change_status">
        <input type="hidden" name="current_tab" value="<%= HtmlEncode(activeTab) %>">
        <div class="form-group">
          <label class="form-label">Nuevo Estado</label>
          <select name="status" class="form-control">
            <option value="open" <%= IIf(rsWO("status")="open", "selected", "") %>>Abierta</option>
            <option value="in_progress" <%= IIf(rsWO("status")="in_progress", "selected", "") %>>En Progreso</option>
            <option value="pending" <%= IIf(rsWO("status")="pending", "selected", "") %>>Pendiente (Esperando refacciones/aprobación)</option>
            <option value="completed" <%= IIf(rsWO("status")="completed", "selected", "") %>>Completada</option>
          </select>
        </div>
      </form>
    </div>
    <div class="modal-footer">
      <button class="btn btn-outline" onclick="closeModal('modalStatus')">Cancelar</button>
      <button class="btn btn-primary" onclick="document.getElementById('formStatus').submit()">Guardar</button>
    </div>
  </div>
</div>

<script>
document.addEventListener('DOMContentLoaded', function() {
    initTabs('wo');
});
</script>

<%
rsWO.Close : Set rsWO = Nothing
CloseConnection(oConn)
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
