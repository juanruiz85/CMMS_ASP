<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = "Orden de Trabajo"
Dim PageModule : PageModule = "work_orders"

Dim oConn : Set oConn = GetConnection()
Dim itemId : itemId = QSInt("id")
Dim isEdit : isEdit = (itemId > 0)

Dim woCode, woTitle, woDesc, woAssetId, woPlantId, woType, woPriority, woStatus, woAssignedTo, woEstHrs, woEstCost, woSchedStart, woSchedEnd, woNotes

If isEdit Then
    Dim rs
    Set rs = oConn.Execute("SELECT * FROM cmms_work_orders WHERE id=" & itemId)
    If Not rs.EOF Then
        woCode = rs("code")
        woTitle = rs("title")
        woDesc = NullStr(rs("description"))
        woAssetId = NullInt(rs("asset_id"))
        woPlantId = rs("plant_id")
        woType = rs("type")
        woPriority = rs("priority")
        woStatus = rs("status")
        woAssignedTo = NullInt(rs("assigned_to_id"))
        woEstHrs = NullStr(rs("estimated_hours"))
        woEstCost = NullStr(rs("estimated_cost"))
        
        woSchedStart = NullStr(rs("scheduled_start"))
        If IsDate(woSchedStart) Then woSchedStart = Year(woSchedStart) & "-" & Right("0" & Month(woSchedStart), 2) & "-" & Right("0" & Day(woSchedStart), 2)
        
        woSchedEnd = NullStr(rs("scheduled_end"))
        If IsDate(woSchedEnd) Then woSchedEnd = Year(woSchedEnd) & "-" & Right("0" & Month(woSchedEnd), 2) & "-" & Right("0" & Day(woSchedEnd), 2)
        
        woNotes = NullStr(rs("notes"))
        PageTitle = "Editar OT: " & woCode
    End If
    rs.Close : Set rs = Nothing
Else
    PageTitle = "Nueva Orden de Trabajo"
    woCode = GenerateWOCode()
    woStatus = "open"
    woPriority = "medium"
    woType = "preventive"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME") & "?id=" & itemId
    End If

    woTitle = Trim(Request.Form("title"))
    woDesc = Trim(Request.Form("description"))
    woPlantId = Request.Form("plant_id")
    woAssetId = Request.Form("asset_id")
    If woAssetId = "" Then woAssetId = "NULL"
    woType = Trim(Request.Form("type"))
    woPriority = Trim(Request.Form("priority"))
    woStatus = Trim(Request.Form("status"))
    woAssignedTo = Request.Form("assigned_to_id")
    If woAssignedTo = "" Then woAssignedTo = "NULL"
    
    woEstHrs = SafeDecimal(Request.Form("estimated_hours"))
    woEstCost = SafeDecimal(Request.Form("estimated_cost"))
    
    woSchedStart = Trim(Request.Form("scheduled_start"))
    If woSchedStart = "" Then woSchedStart = "NULL" Else woSchedStart = "'" & Replace(woSchedStart, "'", "''") & "'"
    
    woSchedEnd = Trim(Request.Form("scheduled_end"))
    If woSchedEnd = "" Then woSchedEnd = "NULL" Else woSchedEnd = "'" & Replace(woSchedEnd, "'", "''") & "'"
    
    woNotes = Trim(Request.Form("notes"))

    Dim cmd, sql
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = oConn
    
    If isEdit Then
        sql = "UPDATE cmms_work_orders SET title=?, description=?, plant_id=" & woPlantId & ", asset_id=" & woAssetId & ", type=?, priority=?, status=?, assigned_to_id=" & woAssignedTo & ", estimated_hours=" & woEstHrs & ", estimated_cost=" & woEstCost & ", scheduled_start=" & woSchedStart & ", scheduled_end=" & woSchedEnd & ", notes=?, updated_at=GETDATE() WHERE id=?"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 200, woTitle)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 4000, woDesc)
        cmd.Parameters.Append cmd.CreateParameter("@type", 200, 1, 50, woType)
        cmd.Parameters.Append cmd.CreateParameter("@pri", 200, 1, 20, woPriority)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, woStatus)
        cmd.Parameters.Append cmd.CreateParameter("@notes", 200, 1, 4000, woNotes)
        cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , itemId)
        cmd.Execute
        LogActivity CurrentUserId(), "UPDATE_WO", "Actualizó OT: " & woCode, "work_orders", itemId
        
        ' Si se cambió a completada
        If woStatus = "completed" Then
            Dim cmdComplete
            Set cmdComplete = Server.CreateObject("ADODB.Command")
            Set cmdComplete.ActiveConnection = oConn
            cmdComplete.CommandText = "UPDATE cmms_work_orders SET completed_at=GETDATE(), closed_by_id=? WHERE id=? AND completed_at IS NULL"
            cmdComplete.Parameters.Append cmdComplete.CreateParameter("@closed_by", 3, 1, , CurrentUserId())
            cmdComplete.Parameters.Append cmdComplete.CreateParameter("@id", 3, 1, , itemId)
            cmdComplete.Execute
            Set cmdComplete = Nothing
        End If
        
    Else
        sql = "INSERT INTO cmms_work_orders (code, title, description, asset_id, plant_id, requester_id, assigned_to_id, type, priority, status, estimated_hours, estimated_cost, scheduled_start, scheduled_end, notes, created_at, updated_at) " & _
              "VALUES (?, ?, ?, " & woAssetId & ", " & woPlantId & ", " & CurrentUserId() & ", " & woAssignedTo & ", ?, ?, ?, " & woEstHrs & ", " & woEstCost & ", " & woSchedStart & ", " & woSchedEnd & ", ?, GETDATE(), GETDATE())"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, woCode)
        cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 200, woTitle)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 4000, woDesc)
        cmd.Parameters.Append cmd.CreateParameter("@type", 200, 1, 50, woType)
        cmd.Parameters.Append cmd.CreateParameter("@pri", 200, 1, 20, woPriority)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, woStatus)
        cmd.Parameters.Append cmd.CreateParameter("@notes", 200, 1, 4000, woNotes)
        cmd.Execute
        LogActivity CurrentUserId(), "CREATE_WO", "Creó nueva OT: " & woCode, "work_orders", 0
        
        ' Create notification to assigned user
        If woAssignedTo <> "NULL" And CInt(woAssignedTo) <> CurrentUserId() Then
            CreateNotification CInt(woAssignedTo), "Nueva OT Asignada", "Se te ha asignado la OT: " & woCode, "info", "work_orders", 0
        End If
    End If
    
    SetFlashMessage "success", "Datos guardados correctamente."
    RedirectTo "/CMMS/modules/work_orders/index.asp"
End If

' Get Plants
Dim rsPlants
Set rsPlants = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")

' Get Assets (grouped by plant via client side logic ideally, but here all active)
Dim rsAssets
Set rsAssets = oConn.Execute("SELECT id, plant_id, code, name FROM cmms_assets WHERE status != 'retired' ORDER BY plant_id, name")

' Get Users (Technicians)
Dim rsUsers
Set rsUsers = oConn.Execute("SELECT id, first_name + ' ' + last_name AS full_name FROM cmms_users WHERE status = 'active' ORDER BY first_name")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <div class="page-subtitle">Complete los detalles de la orden de trabajo.</div>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/work_orders/index.asp" class="btn btn-outline">Cancelar</a>
  </div>
</div>

<div class="card" style="max-width:900px;margin:0 auto">
  <div class="card-body">
    <form method="POST" action="form.asp?id=<%= itemId %>" data-validate>
      <%= CSRFField() %>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Información Principal</h4>
      
      <% If Not isEdit Then %>
      <div style="margin-bottom:16px;font-size:12px;color:var(--text-muted)">
        Código autogenerado: <strong style="color:var(--text-primary)"><%= HtmlEncode(woCode) %></strong>
      </div>
      <% Else %>
      <input type="hidden" name="code" value="<%= HtmlEncode(woCode) %>">
      <% End If %>

      <div class="form-group">
        <label class="form-label">Título de la Orden <span class="required">*</span></label>
        <input type="text" name="title" class="form-control" required value="<%= HtmlEncode(woTitle) %>" placeholder="Breve descripción del problema o trabajo a realizar">
      </div>

      <div class="form-group">
        <label class="form-label">Descripción Detallada</label>
        <textarea name="description" class="form-control" rows="4"><%= HtmlEncode(woDesc) %></textarea>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Planta <span class="required">*</span></label>
          <select name="plant_id" id="plant_id" class="form-control" required onchange="filterAssets()">
            <option value="">Seleccione una planta...</option>
            <% Do While Not rsPlants.EOF %>
            <option value="<%= rsPlants("id") %>" <%= IIf(woPlantId = rsPlants("id"), "selected", "") %>><%= HtmlEncode(rsPlants("name")) %></option>
            <% rsPlants.MoveNext : Loop %>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Equipo Asociado</label>
          <select name="asset_id" id="asset_id" class="form-control">
            <option value="">Ninguno / General de la Planta</option>
            <% Do While Not rsAssets.EOF %>
            <option value="<%= rsAssets("id") %>" data-plant="<%= rsAssets("plant_id") %>" <%= IIf(woAssetId = rsAssets("id"), "selected", "") %>><%= HtmlEncode(rsAssets("code")) %> - <%= HtmlEncode(rsAssets("name")) %></option>
            <% rsAssets.MoveNext : Loop %>
          </select>
        </div>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Clasificación y Asignación</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Tipo de Trabajo <span class="required">*</span></label>
          <select name="type" class="form-control" required>
            <option value="preventive" <%= IIf(woType="preventive", "selected", "") %>>Mantenimiento Preventivo</option>
            <option value="corrective" <%= IIf(woType="corrective", "selected", "") %>>Mantenimiento Correctivo</option>
            <option value="predictive" <%= IIf(woType="predictive", "selected", "") %>>Mantenimiento Predictivo</option>
            <option value="emergency" <%= IIf(woType="emergency", "selected", "") %>>Emergencia</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Prioridad <span class="required">*</span></label>
          <select name="priority" class="form-control" required>
            <option value="low" <%= IIf(woPriority="low", "selected", "") %>>Baja</option>
            <option value="medium" <%= IIf(woPriority="medium", "selected", "") %>>Media</option>
            <option value="high" <%= IIf(woPriority="high", "selected", "") %>>Alta</option>
            <option value="urgent" <%= IIf(woPriority="urgent", "selected", "") %>>Urgente</option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Estado <span class="required">*</span></label>
          <select name="status" class="form-control" required>
            <option value="open" <%= IIf(woStatus="open", "selected", "") %>>Abierta</option>
            <option value="in_progress" <%= IIf(woStatus="in_progress", "selected", "") %>>En Progreso</option>
            <option value="pending" <%= IIf(woStatus="pending", "selected", "") %>>Pendiente</option>
            <option value="completed" <%= IIf(woStatus="completed", "selected", "") %>>Completada</option>
            <% If isEdit Then %><option value="cancelled" <%= IIf(woStatus="cancelled", "selected", "") %>>Cancelada</option><% End If %>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Asignado a (Técnico)</label>
          <select name="assigned_to_id" class="form-control">
            <option value="">Seleccione un técnico...</option>
            <% Do While Not rsUsers.EOF %>
            <option value="<%= rsUsers("id") %>" <%= IIf(woAssignedTo = rsUsers("id"), "selected", "") %>><%= HtmlEncode(rsUsers("full_name")) %></option>
            <% rsUsers.MoveNext : Loop %>
          </select>
        </div>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Planificación</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Fecha Programada (Inicio)</label>
          <input type="date" name="scheduled_start" class="form-control" value="<%= woSchedStart %>">
        </div>
        <div class="form-group">
          <label class="form-label">Fecha Programada (Fin)</label>
          <input type="date" name="scheduled_end" class="form-control" value="<%= woSchedEnd %>">
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Horas Estimadas</label>
          <input type="number" step="0.5" name="estimated_hours" class="form-control" value="<%= woEstHrs %>" placeholder="0.0">
        </div>
        <div class="form-group">
          <label class="form-label">Costo Estimado</label>
          <input type="number" step="0.01" name="estimated_cost" class="form-control" value="<%= woEstCost %>" placeholder="0.00">
        </div>
      </div>

      <div class="form-group">
        <label class="form-label">Notas Internas</label>
        <textarea name="notes" class="form-control" rows="2"><%= HtmlEncode(woNotes) %></textarea>
      </div>

      <div style="margin-top:24px;text-align:right">
        <button type="submit" class="btn btn-primary">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>
          Guardar Orden de Trabajo
        </button>
      </div>

    </form>
  </div>
</div>

<script>
function filterAssets() {
    var plantId = document.getElementById('plant_id').value;
    var assetSelect = document.getElementById('asset_id');
    var options = assetSelect.options;
    for (var i = 1; i < options.length; i++) {
        var opt = options[i];
        if (plantId === '' || opt.getAttribute('data-plant') === plantId) {
            opt.style.display = '';
        } else {
            opt.style.display = 'none';
        }
    }
    // Si la opción seleccionada se oculta, resetear a la primera
    if (assetSelect.selectedIndex > 0 && options[assetSelect.selectedIndex].style.display === 'none') {
        assetSelect.selectedIndex = 0;
    }
}
document.addEventListener('DOMContentLoaded', filterAssets);
</script>

<%
rsPlants.Close : Set rsPlants = Nothing
rsAssets.Close : Set rsAssets = Nothing
rsUsers.Close : Set rsUsers = Nothing
CloseConnection(oConn)
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
