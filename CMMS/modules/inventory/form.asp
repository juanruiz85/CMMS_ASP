<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
If Not IsSupervisorOrAdmin() Then
    SetFlashMessage "danger", "No tiene permisos."
    RedirectTo "/CMMS/index.asp"
End If

Dim PageTitle : PageTitle = "Artículo"
Dim PageModule : PageModule = "inventory"

Dim oConn : Set oConn = GetConnection()
Dim itemId : itemId = QSInt("id")
Dim isEdit : isEdit = (itemId > 0)

Dim iCode, iName, iDesc, iCat, iManuf, iModel, iPart, iUOM, iMinStock, iMaxStock, iReorder, iCost, iPlantId, iLoc, iBin, iBarcode, iStatus

If isEdit Then
    Dim rs
    Set rs = oConn.Execute("SELECT * FROM cmms_inventory WHERE id=" & itemId)
    If Not rs.EOF Then
        iCode = rs("code")
        iName = rs("name")
        iDesc = NullStr(rs("description"))
        iCat = NullStr(rs("category"))
        iManuf = NullStr(rs("manufacturer"))
        iModel = NullStr(rs("model"))
        iPart = NullStr(rs("part_number"))
        iUOM = rs("unit_of_measure")
        iMinStock = NullStr(rs("min_stock"))
        iMaxStock = NullStr(rs("max_stock"))
        iReorder = NullStr(rs("reorder_point"))
        iCost = NullStr(rs("unit_cost"))
        iPlantId = rs("plant_id")
        iLoc = NullStr(rs("location"))
        iBin = NullStr(rs("bin_location"))
        iBarcode = NullStr(rs("barcode"))
        iStatus = rs("status")
        PageTitle = "Editar Artículo: " & iName
    End If
    rs.Close : Set rs = Nothing
Else
    PageTitle = "Nuevo Artículo"
    iCode = GenerateInventoryCode()
    iStatus = "active"
    iUOM = "unit"
    iMinStock = "5"
    iReorder = "10"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token de seguridad inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME") & "?id=" & itemId
    End If

    iCode = Trim(Request.Form("code"))
    iName = Trim(Request.Form("name"))
    iDesc = Trim(Request.Form("description"))
    iCat = Trim(Request.Form("category"))
    iManuf = Trim(Request.Form("manufacturer"))
    iModel = Trim(Request.Form("model"))
    iPart = Trim(Request.Form("part_number"))
    iUOM = Trim(Request.Form("unit_of_measure"))
    iMinStock = SafeDecimal(Request.Form("min_stock"))
    iMaxStock = SafeDecimal(Request.Form("max_stock"))
    iReorder = SafeDecimal(Request.Form("reorder_point"))
    iCost = SafeDecimal(Request.Form("unit_cost"))
    iPlantId = Request.Form("plant_id")
    iLoc = Trim(Request.Form("location"))
    iBin = Trim(Request.Form("bin_location"))
    iBarcode = Trim(Request.Form("barcode"))
    iStatus = Trim(Request.Form("status"))

    Dim cmd, sql
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = oConn
    
    If isEdit Then
        sql = "UPDATE cmms_inventory SET code=?, name=?, description=?, category=?, manufacturer=?, model=?, part_number=?, unit_of_measure=?, min_stock=" & iMinStock & ", max_stock=" & iMaxStock & ", reorder_point=" & iReorder & ", unit_cost=" & iCost & ", plant_id=" & iPlantId & ", location=?, bin_location=?, barcode=?, status=?, updated_at=GETDATE() WHERE id=?"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, iCode)
        cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 100, iName)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 500, iDesc)
        cmd.Parameters.Append cmd.CreateParameter("@cat", 200, 1, 50, iCat)
        cmd.Parameters.Append cmd.CreateParameter("@manuf", 200, 1, 100, iManuf)
        cmd.Parameters.Append cmd.CreateParameter("@mod", 200, 1, 100, iModel)
        cmd.Parameters.Append cmd.CreateParameter("@part", 200, 1, 50, iPart)
        cmd.Parameters.Append cmd.CreateParameter("@uom", 200, 1, 20, iUOM)
        cmd.Parameters.Append cmd.CreateParameter("@loc", 200, 1, 100, iLoc)
        cmd.Parameters.Append cmd.CreateParameter("@bin", 200, 1, 50, iBin)
        cmd.Parameters.Append cmd.CreateParameter("@bar", 200, 1, 100, iBarcode)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, iStatus)
        cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , itemId)
        cmd.Execute
        LogActivity CurrentUserId(), "UPDATE_INV", "Actualizó artículo: " & iName, "inventory", itemId
    Else
        sql = "INSERT INTO cmms_inventory (code, name, description, category, manufacturer, model, part_number, unit_of_measure, min_stock, max_stock, reorder_point, unit_cost, plant_id, location, bin_location, barcode, status, created_at, updated_at) " & _
              "VALUES (?, ?, ?, ?, ?, ?, ?, ?, " & iMinStock & ", " & iMaxStock & ", " & iReorder & ", " & iCost & ", " & iPlantId & ", ?, ?, ?, ?, GETDATE(), GETDATE())"
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@code", 200, 1, 50, iCode)
        cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 100, iName)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 500, iDesc)
        cmd.Parameters.Append cmd.CreateParameter("@cat", 200, 1, 50, iCat)
        cmd.Parameters.Append cmd.CreateParameter("@manuf", 200, 1, 100, iManuf)
        cmd.Parameters.Append cmd.CreateParameter("@mod", 200, 1, 100, iModel)
        cmd.Parameters.Append cmd.CreateParameter("@part", 200, 1, 50, iPart)
        cmd.Parameters.Append cmd.CreateParameter("@uom", 200, 1, 20, iUOM)
        cmd.Parameters.Append cmd.CreateParameter("@loc", 200, 1, 100, iLoc)
        cmd.Parameters.Append cmd.CreateParameter("@bin", 200, 1, 50, iBin)
        cmd.Parameters.Append cmd.CreateParameter("@bar", 200, 1, 100, iBarcode)
        cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, iStatus)
        cmd.Execute
        
        ' Obtener el ID insertado (compatible multi-BD)
        Dim newId : newId = GetLastInsertID(oConn)
        
        ' Crear registro en stock
        Dim cmdStock
        Set cmdStock = Server.CreateObject("ADODB.Command")
        Set cmdStock.ActiveConnection = oConn
        cmdStock.CommandText = "INSERT INTO cmms_inventory_stock (inventory_id, plant_id, quantity, last_updated) VALUES (?, ?, 0, GETDATE())"
        cmdStock.Parameters.Append cmdStock.CreateParameter("@inv_id", 3, 1, , newId)
        cmdStock.Parameters.Append cmdStock.CreateParameter("@plant_id", 3, 1, , iPlantId)
        cmdStock.Execute
        Set cmdStock = Nothing
        
        LogActivity CurrentUserId(), "CREATE_INV", "Creó artículo: " & iName, "inventory", newId
    End If
    
    SetFlashMessage "success", "Datos guardados."
    RedirectTo "/CMMS/modules/inventory/index.asp"
End If

Dim rsPlants
Set rsPlants = oConn.Execute("SELECT id, name FROM cmms_plants WHERE status='active' ORDER BY name")
%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
  </div>
  <div class="page-actions">
    <a href="/CMMS/modules/inventory/index.asp" class="btn btn-outline">Cancelar</a>
  </div>
</div>

<div class="card" style="max-width:800px;margin:0 auto">
  <div class="card-body">
    <form method="POST" action="form.asp?id=<%= itemId %>" data-validate>
      <%= CSRFField() %>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Código <span class="required">*</span></label>
          <input type="text" name="code" class="form-control" required value="<%= HtmlEncode(iCode) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Nombre del Artículo <span class="required">*</span></label>
          <input type="text" name="name" class="form-control" required value="<%= HtmlEncode(iName) %>">
        </div>
      </div>

      <div class="form-group">
        <label class="form-label">Descripción</label>
        <textarea name="description" class="form-control" rows="2"><%= HtmlEncode(iDesc) %></textarea>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Categoría</label>
          <input type="text" name="category" class="form-control" value="<%= HtmlEncode(iCat) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Número de Parte</label>
          <input type="text" name="part_number" class="form-control" value="<%= HtmlEncode(iPart) %>">
        </div>
      </div>

      <h4 style="margin:24px 0 16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Stock y Costos</h4>
      
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Unidad de Medida <span class="required">*</span></label>
          <select name="unit_of_measure" class="form-control" required>
            <option value="unit" <%= IIf(iUOM="unit", "selected", "") %>>Unidad (Pza)</option>
            <option value="kg" <%= IIf(iUOM="kg", "selected", "") %>>Kilogramo (Kg)</option>
            <option value="liter" <%= IIf(iUOM="liter", "selected", "") %>>Litro (L)</option>
            <option value="meter" <%= IIf(iUOM="meter", "selected", "") %>>Metro (M)</option>
            <option value="box" <%= IIf(iUOM="box", "selected", "") %>>Caja</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Costo Unitario</label>
          <input type="number" step="0.01" name="unit_cost" class="form-control" value="<%= iCost %>" placeholder="0.00">
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Stock Mínimo</label>
          <input type="number" step="1" name="min_stock" class="form-control" value="<%= iMinStock %>">
        </div>
        <div class="form-group">
          <label class="form-label">Punto de Reorden</label>
          <input type="number" step="1" name="reorder_point" class="form-control" value="<%= iReorder %>">
        </div>
        <div class="form-group">
          <label class="form-label">Stock Máximo</label>
          <input type="number" step="1" name="max_stock" class="form-control" value="<%= iMaxStock %>">
        </div>
      </div>

      <h4 style="margin:24px 0 16px;color:var(--text-primary);border-bottom:1px solid var(--border-subtle);padding-bottom:8px">Ubicación Físca</h4>

      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Planta <span class="required">*</span></label>
          <select name="plant_id" class="form-control" required>
            <% Do While Not rsPlants.EOF %>
            <option value="<%= rsPlants("id") %>" <%= IIf(iPlantId = rsPlants("id"), "selected", "") %>><%= HtmlEncode(rsPlants("name")) %></option>
            <% rsPlants.MoveNext : Loop %>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Almacén / Ubicación</label>
          <input type="text" name="location" class="form-control" value="<%= HtmlEncode(iLoc) %>">
        </div>
      </div>
      
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">Estante / Bin</label>
          <input type="text" name="bin_location" class="form-control" value="<%= HtmlEncode(iBin) %>">
        </div>
        <div class="form-group">
          <label class="form-label">Estado</label>
          <select name="status" class="form-control">
            <option value="active" <%= IIf(iStatus="active", "selected", "") %>>Activo</option>
            <option value="inactive" <%= IIf(iStatus="inactive", "selected", "") %>>Inactivo</option>
          </select>
        </div>
      </div>

      <div style="margin-top:24px;text-align:right">
        <button type="submit" class="btn btn-primary">Guardar Artículo</button>
      </div>
    </form>
  </div>
</div>

<%
rsPlants.Close : Set rsPlants = Nothing
CloseConnection(oConn)
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
