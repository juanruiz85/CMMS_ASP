<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
If Not IsSupervisorOrAdmin() Then
    SetFlashMessage "danger", "No tiene permisos para acceder a esta página."
    RedirectTo "/CMMS/index.asp"
End If

Dim PageTitle : PageTitle = "Equipo"
Dim PageModule : PageModule = "assets"

Dim oConn : Set oConn = GetConnection()
Dim itemId : itemId = QSInt("id")
Dim isEdit : isEdit = (itemId > 0)

Dim aCode, aName, aDesc, aPlantId, aLocation, aCategory, aManuf, aModel, aSerial, aPurchDate, aInstDate, aWarrExp, aCost, aCriticality, aStatus, aParentId

If isEdit Then
    Dim rs
    Set rs = oConn.Execute("SELECT * FROM cmms_assets WHERE id=" & itemId)
    If Not rs.EOF Then
        aCode = rs("code")
        aName = rs("name")
        aDesc = NullStr(rs("description"))
        aPlantId = rs("plant_id")
        aLocation = NullStr(rs("location"))
        aCategory = NullStr(rs("category"))
        aManuf = NullStr(rs("manufacturer"))
        aModel = NullStr(rs("model"))
        aSerial = NullStr(rs("serial_number"))
        aPurchDate = NullStr(rs("purchase_date"))
        If IsDate(aPurchDate) Then aPurchDate = Year(aPurchDate) & "-" & Right("0" & Month(aPurchDate), 2) & "-" & Right("0" & Day(aPurchDate), 2)
        aInstDate = NullStr(rs("installation_date"))
        If IsDate(aInstDate) Then aInstDate = Year(aInstDate) & "-" & Right("0" & Month(aInstDate), 2) & "-" & Right("0" & Day(aInstDate), 2)
        aWarrExp = NullStr(rs("warranty_expiry"))
        If IsDate(aWarrExp) Then aWarrExp = Year(aWarrExp) & "-" & Right("0" & Month(aWarrExp), 2) & "-" & Right("0" & Day(aWarrExp), 2)
        aCost = NullStr(rs("cost"))
        aCriticality = rs("criticality")
        aStatus = rs("status")
        aParentId = NullInt(rs("parent_id"))
        PageTitle = "Editar Equipo: " & aName
    End If
    rs.Close : Set rs = Nothing
Else
    PageTitle = "Nuevo Equipo"
    aCode = GenerateAssetCode("EQ")
    aStatus = "operational"
    aCriticality = "medium"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME") & "?id=" & itemId
    End If

    aCode = Trim(Request.Form("code"))
    aName = Trim(Request.Form("name"))
    aDesc = Trim(Request.Form("description"))
    aPlantId = Request.Form("plant_id")
    aLocation = Trim(Request.Form("location"))
    aCategory = Trim(Request.Form("category"))
    aManuf = Trim(Request.Form("manufacturer"))
    aModel = Trim(Request.Form("model"))
    aSerial = Trim(Request.Form("serial_number"))
    aPurchDate = Trim(Request.Form("purchase_date"))
    If aPurchDate = "" Then aPurchDate = "NULL" Else aPurchDate = "'" & Replace(aPurchDate, "'", "''") & "'"
    aInstDate = Trim(Request.Form("installation_date"))
    If aInstDate = "" Then aInstDate = "NULL" Else aInstDate = "'" & Replace(aInstDate, "'", "''") & "'"
    aWarrExp = Trim(Request.Form("warranty_expiry"))
    If aWarrExp = "" Then aWarrExp = "NULL" Else aWarrExp = "'" & Replace(aWarrExp, "'", "''") & "'"
    aCost = SafeDecimal(Request.Form("cost"))
    aCriticality = Trim(Request.Form("criticality"))
    aStatus = Trim(Request.Form("status"))
    aParentId = Request.Form("parent_id")
    If aParentId = "" Then aParentId = "NULL"

    Dim cmd, sql
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = oConn
    
    If isEdit Then
        sql = "UPDATE cmms_assets SET code=?, name=?, description=?, plant_id=" & aPlantId & ", location=?, category=?, manufacturer=?, model=?, serial_number=?, purchase_date=" & aPurchDate & ", installation_date=" & aInstDate & ", warranty_expiry=" & aWarrExp & ", cost=" & aCost & ", criticality=?, status=?, parent_id=" & aParentId & ", updated_at=GETDATE() WHERE id=?"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, aCode)
        cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 100, aName)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 500, aDesc)
        cmd.Parameters.Append cmd.CreateParameter("@loc", 200, 1, 100, aLocation)
        cmd.Parameters.Append cmd.CreateParameter("@cat", 200, 1, 50, aCategory)
        cmd.Parameters.Append cmd.CreateParameter("@manuf", 200, 1, 100, aManuf)
        cmd.Parameters.Append cmd.CreateParameter("@mod", 200, 1, 100, aModel)
        cmd.Parameters.Append cmd.CreateParameter("@serial", 200, 1, 100, aSerial)
        cmd.Parameters.Append cmd.CreateParameter("@crit", 200, 1, 20, aCriticality)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, aStatus)
        cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , itemId)
        cmd.Execute
        LogActivity CurrentUserId(), "UPDATE_ASSET", "Actualizó equipo: " & aName, "assets", itemId
    Else
        sql = "INSERT INTO cmms_assets (code, name, description, plant_id, location, category, manufacturer, model, serial_number, purchase_date, installation_date, warranty_expiry, cost, criticality, status, parent_id, created_at, updated_at) " & _
              "VALUES (?, ?, ?, " & aPlantId & ", ?, ?, ?, ?, ?, " & aPurchDate & ", " & aInstDate & ", " & aWarrExp & ", " & aCost & ", ?, ?, " & aParentId & ", GETDATE(), GETDATE())"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, aCode)
        cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 100, aName)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 500, aDesc)
        cmd.Parameters.Append cmd.CreateParameter("@loc", 200, 1, 100, aLocation)
        cmd.Parameters.Append cmd.CreateParameter("@cat", 200, 1, 50, aCategory)
        cmd.Parameters.Append cmd.CreateParameter("@manuf", 200, 1, 100, aManuf)
        cmd.Parameters.Append cmd.CreateParameter("@mod", 200, 1, 100, aModel)
        cmd.Parameters.Append cmd.CreateParameter("@serial", 200, 1, 100, aSerial)
        cmd.Parameters.Append cmd.CreateParameter("@crit", 200, 1, 20, aCriticality)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, aStatus)
        cmd.Execute
        LogActivity CurrentUserId(), "CREATE_ASSET", "Creó nuevo equipo: " & aName, "assets", 0
    End If
    
    SetFlashMessage "success", "Datos guardados correctamente."
    RedirectTo "/CMMS/modules/assets_module/index.asp"
End If

' Get plants
Dim rsPlants
Set rsPlants = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")

' Get parent assets (only if editing and plant is selected, or just simple list of active assets)
Dim rsParents
Set rsParents = oConn.Execute("SELECT id, code, name FROM cmms_assets WHERE status != 'retired' AND id != " & itemId & " ORDER BY name")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <div class="page-subtitle">Complete los campos para guardar la información del equipo.</div>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/assets_module/index.asp" class="btn btn-outline">Cancelar</a>
  </div>
</div>

<div class="card" style="max-width:900px;margin:0 auto">
  <div class="card-body">
    <form method="POST" action="form.asp?id=<%= itemId %>" data-validate>
      <%= CSRFField() %>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Información General</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Código <span class="required">*</span></label>
          <input type="text" name="code" class="form-control" required value="<%= HtmlEncode(aCode) %>" placeholder="Ej: EQ-123">
        </div>
        <div class="form-group">
          <label class="form-label">Nombre del Equipo <span class="required">*</span></label>
          <input type="text" name="name" class="form-control" required value="<%= HtmlEncode(aName) %>">
        </div>
      </div>

      <div class="form-group">
        <label class="form-label">Descripción</label>
        <textarea name="description" class="form-control"><%= HtmlEncode(aDesc) %></textarea>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Planta Asignada <span class="required">*</span></label>
          <select name="plant_id" class="form-control" required>
            <option value="">Seleccione una planta...</option>
            <% Do While Not rsPlants.EOF %>
            <option value="<%= rsPlants("id") %>" <%= IIf(aPlantId = rsPlants("id"), "selected", "") %>><%= HtmlEncode(rsPlants("name")) %></option>
            <% rsPlants.MoveNext : Loop %>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Ubicación Física</label>
          <input type="text" name="location" class="form-control" value="<%= HtmlEncode(aLocation) %>" placeholder="Ej: Línea 1, Cuarto de máquinas">
        </div>
      </div>
      
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Categoría</label>
          <input type="text" name="category" class="form-control" value="<%= HtmlEncode(aCategory) %>" placeholder="Ej: Motores, Bombas, HVAC">
        </div>
        <div class="form-group">
          <label class="form-label">Equipo Padre (Jerarquía)</label>
          <select name="parent_id" class="form-control">
            <option value="">-- Ninguno (Equipo Principal) --</option>
            <% Do While Not rsParents.EOF %>
            <option value="<%= rsParents("id") %>" <%= IIf(aParentId = rsParents("id"), "selected", "") %>><%= HtmlEncode(rsParents("code")) %> - <%= HtmlEncode(rsParents("name")) %></option>
            <% rsParents.MoveNext : Loop %>
          </select>
        </div>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Especificaciones Técnicas</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Fabricante</label>
          <input type="text" name="manufacturer" class="form-control" value="<%= HtmlEncode(aManuf) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Modelo</label>
          <input type="text" name="model" class="form-control" value="<%= HtmlEncode(aModel) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Número de Serie</label>
          <input type="text" name="serial_number" class="form-control" value="<%= HtmlEncode(aSerial) %>">
        </div>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Fechas y Costos</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Fecha de Compra</label>
          <input type="date" name="purchase_date" class="form-control" value="<%= aPurchDate %>">
        </div>
        <div class="form-group">
          <label class="form-label">Fecha de Instalación</label>
          <input type="date" name="installation_date" class="form-control" value="<%= aInstDate %>">
        </div>
        <div class="form-group">
          <label class="form-label">Fin de Garantía</label>
          <input type="date" name="warranty_expiry" class="form-control" value="<%= aWarrExp %>">
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Costo de Adquisición</label>
          <input type="number" step="0.01" name="cost" class="form-control" value="<%= aCost %>" placeholder="0.00">
        </div>
      </div>

      <h4 style="margin-bottom:16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Estado y Criticidad</h4>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Criticidad <span class="required">*</span></label>
          <select name="criticality" class="form-control" required>
            <option value="low" <%= IIf(aCriticality="low", "selected", "") %>>Baja</option>
            <option value="medium" <%= IIf(aCriticality="medium", "selected", "") %>>Media</option>
            <option value="high" <%= IIf(aCriticality="high", "selected", "") %>>Alta</option>
            <option value="critical" <%= IIf(aCriticality="critical", "selected", "") %>>Crítica</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Estado <span class="required">*</span></label>
          <select name="status" class="form-control" required>
            <option value="operational" <%= IIf(aStatus="operational", "selected", "") %>>Operativo</option>
            <option value="maintenance" <%= IIf(aStatus="maintenance", "selected", "") %>>En Mantenimiento</option>
            <option value="down" <%= IIf(aStatus="down", "selected", "") %>>Fuera de servicio</option>
            <option value="retired" <%= IIf(aStatus="retired", "selected", "") %>>Retirado</option>
          </select>
        </div>
      </div>

      <div style="margin-top:24px;text-align:right">
        <button type="submit" class="btn btn-primary">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>
          Guardar Datos
        </button>
      </div>

    </form>
  </div>
</div>

<%
rsPlants.Close : Set rsPlants = Nothing
rsParents.Close : Set rsParents = Nothing
CloseConnection(oConn)
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
