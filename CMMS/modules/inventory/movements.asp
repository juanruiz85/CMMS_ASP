<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim PageTitle : PageTitle = "Movimientos de Inventario"
Dim PageModule : PageModule = "inventory"

Dim oConn : Set oConn = GetConnection()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token inválido."
        RedirectTo "movements.asp"
    End If
    
    Dim invId : invId = QSIntForm("inventory_id")
    Dim mType : mType = Trim(Request.Form("movement_type"))
    Dim qty   : qty   = SafeDecimal(Request.Form("quantity"))
    Dim rsn   : rsn   = Trim(Request.Form("reason"))
    Dim ref   : ref   = Trim(Request.Form("reference"))
    
    If invId > 0 And qty <> 0 Then
        ' Get current stock
        Dim rsCur
        Set rsCur = oConn.Execute("SELECT ISNULL(SUM(quantity), 0) AS cur_qty FROM cmms_inventory_stock WHERE inventory_id=" & invId)
        Dim curQty : curQty = rsCur("cur_qty")
        rsCur.Close : Set rsCur = Nothing
        
        Dim newQty : newQty = curQty
        If mType = "in" Then newQty = curQty + qty
        If mType = "out" Then newQty = curQty - qty
        If mType = "adjustment" Then newQty = qty : qty = newQty - curQty ' qty is difference for log
        
        If newQty < 0 And mType = "out" Then
            SetFlashMessage "danger", "No hay stock suficiente para esta salida."
        Else
            ' Record movement
            Dim cmd
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = oConn
            cmd.CommandText = "INSERT INTO cmms_inventory_movements (inventory_id, movement_type, previous_quantity, quantity, new_quantity, reason, reference, user_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, GETDATE())"
            cmd.Parameters.Append cmd.CreateParameter("@inv", 3, 1, , invId)
            cmd.Parameters.Append cmd.CreateParameter("@mt", 200, 1, 20, mType)
            cmd.Parameters.Append cmd.CreateParameter("@pq", 14, 1, , curQty)
            cmd.Parameters.Append cmd.CreateParameter("@q", 14, 1, , Abs(qty))
            cmd.Parameters.Append cmd.CreateParameter("@nq", 14, 1, , newQty)
            cmd.Parameters.Append cmd.CreateParameter("@rsn", 200, 1, 200, rsn)
            cmd.Parameters.Append cmd.CreateParameter("@ref", 200, 1, 100, ref)
            cmd.Parameters.Append cmd.CreateParameter("@uid", 3, 1, , CurrentUserId())
            cmd.Execute
            
            ' Update stock (simplification: updating the first stock record or creating if not exists)
            Dim rsStock, cmdStock
            Set rsStock = oConn.Execute("SELECT id FROM cmms_inventory_stock WHERE inventory_id=" & invId)
            If rsStock.EOF Then
                ' Create new stock record - need to get plant_id from inventory
                Set cmdStock = Server.CreateObject("ADODB.Command")
                Set cmdStock.ActiveConnection = oConn
                cmdStock.CommandText = "INSERT INTO cmms_inventory_stock (inventory_id, plant_id, quantity, last_updated) SELECT ?, plant_id, ?, GETDATE() FROM cmms_inventory WHERE id=?"
                cmdStock.Parameters.Append cmdStock.CreateParameter("@inv_id", 3, 1, , invId)
                cmdStock.Parameters.Append cmdStock.CreateParameter("@qty", 14, 1, , newQty)
                cmdStock.Parameters.Append cmdStock.CreateParameter("@check_id", 3, 1, , invId)
                cmdStock.Execute
                Set cmdStock = Nothing
            Else
                Set cmdStock = Server.CreateObject("ADODB.Command")
                Set cmdStock.ActiveConnection = oConn
                cmdStock.CommandText = "UPDATE cmms_inventory_stock SET quantity=?, last_updated=GETDATE() WHERE id=?"
                cmdStock.Parameters.Append cmdStock.CreateParameter("@qty", 14, 1, , newQty)
                cmdStock.Parameters.Append cmdStock.CreateParameter("@id", 3, 1, , rsStock("id"))
                cmdStock.Execute
                Set cmdStock = Nothing
            End If
            rsStock.Close : Set rsStock = Nothing
            
            SetFlashMessage "success", "Movimiento registrado correctamente."
        End If
    Else
        SetFlashMessage "danger", "Datos inválidos."
    End If
    RedirectTo "movements.asp"
End If

' Pagination
Dim currentPage : currentPage = GetCurrentPage()
Dim perPage     : perPage = 25
Dim offset      : offset = (currentPage - 1) * perPage

' Count total
Dim totalRows : totalRows = 0
Dim rsTotal
Set rsTotal = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_inventory_movements")
If Not rsTotal.EOF Then totalRows = rsTotal("cnt")
rsTotal.Close : Set rsTotal = Nothing

' Get data
Dim sql
sql = "SELECT m.*, i.name AS item_name, i.code AS item_code, u.first_name + ' ' + u.last_name AS user_name " & _
      "FROM cmms_inventory_movements m " & _
      "JOIN cmms_inventory i ON i.id = m.inventory_id " & _
      "JOIN cmms_users u ON u.id = m.user_id " & _
      "ORDER BY m.created_at DESC " & _
      "OFFSET " & offset & " ROWS FETCH NEXT " & perPage & " ROWS ONLY"

Dim rsMovs
Set rsMovs = oConn.Execute(sql)

' Get items for form
Dim rsItems
Set rsItems = oConn.Execute("SELECT id, code, name FROM cmms_inventory WHERE status='active' ORDER BY name")

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Registro de entradas, salidas y ajustes de stock</p>
  </div>
  <div class="page-actions">
    <% If IsSupervisorOrAdmin() Then %>
    <button class="btn btn-primary" onclick="openModal('modalMove')">Registrar Movimiento</button>
    <% End If %>
  </div>
</div>

<div class="card">
  <div class="table-responsive">
    <table class="table">
      <thead>
        <tr>
          <th>Fecha/Hora</th>
          <th>Artículo</th>
          <th>Tipo</th>
          <th style="text-align:right">Anterior</th>
          <th style="text-align:right">Movimiento</th>
          <th style="text-align:right">Nuevo Stock</th>
          <th>Razón / Referencia</th>
          <th>Usuario</th>
        </tr>
      </thead>
      <tbody>
        <% If rsMovs.EOF Then %>
        <tr><td colspan="8" class="table-empty">No hay movimientos registrados</td></tr>
        <% Else %>
        <% Do While Not rsMovs.EOF %>
        <tr>
          <td style="font-size:12px;color:var(--text-muted)"><%= FormatDateShort(rsMovs("created_at")) %> <%= FormatDateTime(rsMovs("created_at"), 4) %></td>
          <td>
            <div style="font-weight:600"><%= HtmlEncode(rsMovs("item_name")) %></div>
            <div style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(rsMovs("item_code")) %></div>
          </td>
          <td>
            <%
            Dim mColor : mColor = "badge-muted"
            Dim mText  : mText  = rsMovs("movement_type")
            Select Case rsMovs("movement_type")
                Case "in": mColor = "badge-success": mText = "Entrada"
                Case "out": mColor = "badge-danger": mText = "Salida"
                Case "adjustment": mColor = "badge-warning": mText = "Ajuste"
            End Select
            %>
            <span class="badge <%= mColor %>"><%= mText %></span>
          </td>
          <td style="text-align:right;color:var(--text-muted)"><%= rsMovs("previous_quantity") %></td>
          <td style="text-align:right;font-weight:600;<%= IIf(rsMovs("movement_type")="in", "color:var(--success)", IIf(rsMovs("movement_type")="out", "color:var(--danger)", "")) %>">
            <%= IIf(rsMovs("movement_type")="out", "-", "+") %><%= rsMovs("quantity") %>
          </td>
          <td style="text-align:right;font-weight:700;color:var(--primary)"><%= rsMovs("new_quantity") %></td>
          <td>
            <div style="font-size:13px"><%= HtmlEncode(rsMovs("reason")) %></div>
            <% If Not IsNull(rsMovs("reference")) And rsMovs("reference") <> "" Then %>
            <div style="font-size:11px;color:var(--text-muted)">Ref: <%= HtmlEncode(rsMovs("reference")) %></div>
            <% End If %>
          </td>
          <td style="font-size:12px"><%= HtmlEncode(rsMovs("user_name")) %></td>
        </tr>
        <% 
        rsMovs.MoveNext
        Loop 
        %>
        <% End If %>
        <% rsMovs.Close : Set rsMovs = Nothing %>
      </tbody>
    </table>
  </div>
  
  <%= PaginationHTML(totalRows, currentPage, perPage, "") %>
</div>

<!-- Modal Movimiento -->
<div id="modalMove" class="modal-overlay">
  <div class="modal">
    <div class="modal-header">
      <h3 class="modal-title">Nuevo Movimiento de Inventario</h3>
      <button class="modal-close" onclick="closeModal('modalMove')">×</button>
    </div>
    <div class="modal-body">
      <form method="POST" action="movements.asp" id="formMove">
        <%= CSRFField() %>
        
        <div class="form-group">
          <label class="form-label">Artículo <span class="required">*</span></label>
          <select name="inventory_id" class="form-control" required>
            <option value="">Seleccione un artículo...</option>
            <% Do While Not rsItems.EOF %>
            <option value="<%= rsItems("id") %>"><%= HtmlEncode(rsItems("code") & " - " & rsItems("name")) %></option>
            <% rsItems.MoveNext : Loop %>
          </select>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Tipo de Movimiento <span class="required">*</span></label>
            <select name="movement_type" class="form-control" required>
              <option value="in">Entrada (Agregar Stock)</option>
              <option value="out">Salida (Descontar Stock)</option>
              <option value="adjustment">Ajuste Físico (Reemplazar Stock)</option>
            </select>
          </div>
          <div class="form-group">
            <label class="form-label">Cantidad <span class="required">*</span></label>
            <input type="number" step="0.01" name="quantity" class="form-control" required>
          </div>
        </div>
        
        <div class="form-group">
          <label class="form-label">Razón / Motivo <span class="required">*</span></label>
          <input type="text" name="reason" class="form-control" required placeholder="Ej: Compra de material, Mantenimiento Preventivo">
        </div>
        
        <div class="form-group">
          <label class="form-label">Referencia (OT, Factura, etc.)</label>
          <input type="text" name="reference" class="form-control" placeholder="Ej: OT-1024 o FAC-5501">
        </div>
      </form>
    </div>
    <div class="modal-footer">
      <button class="btn btn-outline" onclick="closeModal('modalMove')">Cancelar</button>
      <button class="btn btn-primary" onclick="document.getElementById('formMove').submit()">Guardar Movimiento</button>
    </div>
  </div>
</div>

<% 
rsItems.Close : Set rsItems = Nothing
CloseConnection(oConn) 
%>
<!--#include virtual="/CMMS/templates/footer.asp"-->
