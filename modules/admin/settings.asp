<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()
If Not IsAdmin() Then
    SetFlashMessage "danger", "Acceso denegado."
    RedirectTo "/CMMS/index.asp"
End If

Dim PageTitle : PageTitle = "Configuración del Sistema"
Dim PageModule : PageModule = "admin"

Dim oConn : Set oConn = GetConnection()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRF() Then
        SetFlashMessage "danger", "Token inválido."
        RedirectTo Request.ServerVariables("SCRIPT_NAME")
    End If
    
    Dim cName : cName = Trim(Request.Form("company_name"))
    Dim cTZ   : cTZ   = Trim(Request.Form("timezone"))
    Dim cCurr : cCurr = Trim(Request.Form("currency"))
    Dim cSess : cSess = QSIntForm("session_timeout")
    If cSess < 10 Then cSess = 60
    
    ' UPSERT for settings
    Dim rsS
    Set rsS = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_settings")
    If rsS("cnt") > 0 Then
        oConn.Execute("UPDATE cmms_settings SET company_name='" & Replace(cName, "'", "''") & "', timezone='" & Replace(cTZ, "'", "''") & "', currency='" & Replace(cCurr, "'", "''") & "', session_timeout=" & cSess)
    Else
        oConn.Execute("INSERT INTO cmms_settings (company_name, timezone, currency, session_timeout) VALUES ('" & Replace(cName, "'", "''") & "', '" & Replace(cTZ, "'", "''") & "', '" & Replace(cCurr, "'", "''") & "', " & cSess & ")")
    End If
    rsS.Close : Set rsS = Nothing
    
    LogActivity CurrentUserId(), "UPDATE_SETTINGS", "Actualizó configuración global", "system", 0
    SetFlashMessage "success", "Configuración guardada correctamente."
    RedirectTo "settings.asp"
End If

' Load settings
Dim sName, sTZ, sCurr, sSess
sName = "CMMS Company"
sTZ = "UTC"
sCurr = "USD"
sSess = 60

Dim rs
Set rs = oConn.Execute("SELECT TOP 1 * FROM cmms_settings")
If Not rs.EOF Then
    sName = NullStr(rs("company_name"))
    sTZ = NullStr(rs("timezone"))
    sCurr = NullStr(rs("currency"))
    sSess = NullInt(rs("session_timeout"))
End If
rs.Close : Set rs = Nothing

%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
  </div>
</div>

<div class="dashboard-grid">
  <div class="col-8">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Ajustes Generales</h3>
      </div>
      <div class="card-body">
        <form method="POST" action="settings.asp">
          <%= CSRFField() %>
          
          <div class="form-group">
            <label class="form-label">Nombre de la Empresa</label>
            <input type="text" name="company_name" class="form-control" value="<%= HtmlEncode(sName) %>">
          </div>
          
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Zona Horaria</label>
              <input type="text" name="timezone" class="form-control" value="<%= HtmlEncode(sTZ) %>" placeholder="Ej: America/Mexico_City">
            </div>
            <div class="form-group">
              <label class="form-label">Moneda Principal</label>
              <input type="text" name="currency" class="form-control" value="<%= HtmlEncode(sCurr) %>" placeholder="Ej: MXN, USD, EUR">
            </div>
          </div>
          
          <div class="form-group">
            <label class="form-label">Tiempo de Expiración de Sesión (Minutos)</label>
            <input type="number" name="session_timeout" class="form-control" value="<%= sSess %>" min="10">
          </div>
          
          <div style="margin-top:24px">
            <button type="submit" class="btn btn-primary">Guardar Configuración</button>
          </div>
        </form>
      </div>
    </div>
  </div>
  
  <div class="col-4">
    <div class="card h-full">
      <div class="card-header">
        <h3 class="card-title">Información del Sistema</h3>
      </div>
      <div class="card-body">
        <div class="detail-item" style="margin-bottom:12px">
          <label>Versión CMMS</label>
          <span>v1.0.0</span>
        </div>
        <div class="detail-item" style="margin-bottom:12px">
          <label>Motor de Base de Datos</label>
          <span>SQL Server</span>
        </div>
        <div class="detail-item" style="margin-bottom:12px">
          <label>Servidor Web</label>
          <span><%= Request.ServerVariables("SERVER_SOFTWARE") %></span>
        </div>
        <div class="detail-item">
          <label>Ruta Física</label>
          <span style="font-size:11px;word-break:break-all"><%= Request.ServerVariables("APPL_PHYSICAL_PATH") %></span>
        </div>
      </div>
    </div>
  </div>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
