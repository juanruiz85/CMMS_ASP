<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
Dim PageTitle : PageTitle = "Nueva Solicitud de Trabajo"
Dim PageModule : PageModule = "work_requests"

Dim oConn : Set oConn = GetConnection()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    On Error Resume Next
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME")
    End If

    Dim rTitle : rTitle = Trim(Request.Form("title"))
    Dim rDesc  : rDesc  = Trim(Request.Form("description"))
    Dim rPlant : rPlant = QSIntForm("plant_id")
    Dim rAsset : rAsset = QSIntForm("asset_id")
    If rAsset = 0 Then rAsset = "NULL"

    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = oConn
    cmd.CommandText = "INSERT INTO cmms_work_requests (title, description, plant_id, asset_id, requested_by_id, status, created_at) VALUES (?, ?, " & rPlant & ", " & rAsset & ", " & CurrentUserId() & ", 'pending', GETDATE())"
    cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 255, rTitle)
    cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 4000, rDesc)
    cmd.Execute
    
    CheckError "Al insertar la solicitud de trabajo"

    ' Notificar a los administradores/supervisores de esa planta
    Dim rsAdmins
    Set rsAdmins = oConn.Execute("SELECT id FROM cmms_users WHERE role IN ('admin', 'supervisor') AND status='active'")
    Do While Not rsAdmins.EOF
        CreateNotification rsAdmins("id"), "Nueva Solicitud", "El usuario ha solicitado mantenimiento: " & rTitle, "info", "work_requests", 0
        rsAdmins.MoveNext
    Loop
    rsAdmins.Close : Set rsAdmins = Nothing

    LogActivity CurrentUserId(), "CREATE_REQUEST", "Creó solicitud de mantenimiento: " & rTitle, "work_requests", 0
    SetFlashMessage "success", "Solicitud enviada correctamente. Será revisada por un supervisor."
    RedirectTo "index.asp"
End If

Dim rsPlants
Set rsPlants = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")
Dim rsAssets
Set rsAssets = oConn.Execute("SELECT id, plant_id, code, name FROM cmms_assets WHERE status != 'retired' ORDER BY plant_id, name")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Reporte un problema o solicite una tarea de mantenimiento</p>
  </div>
  <div class="page-actions">
    <a href="index.asp" class="btn btn-outline">Cancelar</a>
  </div>
</div>

<div class="card" style="max-width:800px;margin:0 auto">
  <div class="card-body">
    <form method="POST" action="form.asp" data-validate>
      <%= CSRFField() %>

      <div class="form-group">
        <label class="form-label">Título breve del problema <span class="required">*</span></label>
        <input type="text" name="title" class="form-control" required placeholder="Ej: Fuga de aceite en la máquina, Cambio de lámpara">
      </div>

      <div class="form-group">
        <label class="form-label">Descripción Detallada <span class="required">*</span></label>
        <textarea name="description" class="form-control" rows="4" required placeholder="Describa el problema lo más claro posible..."></textarea>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Ubicación / Planta <span class="required">*</span></label>
          <select name="plant_id" id="plant_id" class="form-control" required onchange="filterAssets()">
            <option value="">Seleccione...</option>
            <% Do While Not rsPlants.EOF %>
            <option value="<%= rsPlants("id") %>"><%= HtmlEncode(rsPlants("name")) %></option>
            <% rsPlants.MoveNext : Loop %>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Equipo Afectado (Opcional)</label>
          <select name="asset_id" id="asset_id" class="form-control">
            <option value="">No aplica / General</option>
            <% Do While Not rsAssets.EOF %>
            <option value="<%= rsAssets("id") %>" data-plant="<%= rsAssets("plant_id") %>"><%= HtmlEncode(rsAssets("code")) %> - <%= HtmlEncode(rsAssets("name")) %></option>
            <% rsAssets.MoveNext : Loop %>
          </select>
        </div>
      </div>

      <div style="margin-top:24px;text-align:right">
        <button type="submit" class="btn btn-primary">Enviar Solicitud</button>
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
    if (assetSelect.selectedIndex > 0 && options[assetSelect.selectedIndex].style.display === 'none') {
        assetSelect.selectedIndex = 0;
    }
}
document.addEventListener('DOMContentLoaded', filterAssets);
</script>

<%
rsPlants.Close : Set rsPlants = Nothing
rsAssets.Close : Set rsAssets = Nothing
CloseConnection(oConn)
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
