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

Dim PageTitle : PageTitle = "Administración del Sistema"
Dim PageModule : PageModule = "admin"

Dim oConn : Set oConn = GetConnection()
%>
<!--#include virtual="/CMMS/templates/header.asp"-->
<!--#include virtual="/CMMS/templates/navigation.asp"-->

<div class="page-header">
  <div>
    <h1 class="page-title"><%= PageTitle %></h1>
    <p class="page-subtitle">Ajustes globales, roles y logs del sistema</p>
  </div>
</div>

<div class="dashboard-grid">
  <div class="col-4">
    <a href="/CMMS/modules/users/index.asp" class="card" style="display:block;text-decoration:none">
      <div class="card-body" style="text-align:center;padding:32px 16px">
        <div style="width:64px;height:64px;border-radius:16px;background:var(--primary-light);color:var(--primary);display:flex;align-items:center;justify-content:center;font-size:32px;margin:0 auto 16px">
          👥
        </div>
        <h3 style="color:var(--text-primary);font-size:18px;font-weight:600;margin-bottom:8px">Usuarios y Accesos</h3>
        <p style="color:var(--text-muted);font-size:13px">Administre los usuarios del sistema, contraseñas y asigne roles.</p>
      </div>
    </a>
  </div>
  <div class="col-4">
    <a href="#" class="card" style="display:block;text-decoration:none">
      <div class="card-body" style="text-align:center;padding:32px 16px">
        <div style="width:64px;height:64px;border-radius:16px;background:var(--secondary-light);color:var(--secondary);display:flex;align-items:center;justify-content:center;font-size:32px;margin:0 auto 16px">
          🛡️
        </div>
        <h3 style="color:var(--text-primary);font-size:18px;font-weight:600;margin-bottom:8px">Roles y Permisos</h3>
        <p style="color:var(--text-muted);font-size:13px">Gestione los roles del sistema y sus permisos específicos.</p>
      </div>
    </a>
  </div>
  <div class="col-4">
    <a href="/CMMS/modules/admin/logs.asp" class="card" style="display:block;text-decoration:none">
      <div class="card-body" style="text-align:center;padding:32px 16px">
        <div style="width:64px;height:64px;border-radius:16px;background:var(--info-light);color:var(--info);display:flex;align-items:center;justify-content:center;font-size:32px;margin:0 auto 16px">
          📋
        </div>
        <h3 style="color:var(--text-primary);font-size:18px;font-weight:600;margin-bottom:8px">Logs de Actividad</h3>
        <p style="color:var(--text-muted);font-size:13px">Auditoría completa de todas las acciones realizadas en el sistema.</p>
      </div>
    </a>
  </div>
</div>

<% CloseConnection(oConn) %>
<!--#include virtual="/CMMS/templates/footer.asp"-->
