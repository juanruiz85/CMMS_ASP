<%
' =============================================================================
' CMMS - Login Page (login.asp)
' =============================================================================
' Si ya está autenticado, ir al dashboard
If Session("user_id") <> "" And Session("user_id") <> 0 Then
    Response.Redirect "/CMMS/index.asp"
    Response.End
End If
%>
<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
Dim LoginError : LoginError = ""
Dim LoginMsg   : LoginMsg   = ""

' Mensaje de sesión expirada
If Request.QueryString("expired") = "1" Then
    LoginMsg = T("session_expired")
End If

' Procesamiento del formulario de login
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim fUser : fUser = Trim(Request.Form("username"))
    Dim fPass : fPass = Trim(Request.Form("password"))

    If fUser = "" Or fPass = "" Then
        LoginError = "Por favor ingrese usuario y contraseña."
    Else
        If DoLogin(fUser, fPass) Then
            ' Login exitoso: redirigir
            Dim redirectUrl : redirectUrl = Session("redirect_after_login")
            Session("redirect_after_login") = ""
            If redirectUrl = "" Or InStr(redirectUrl, "login") > 0 Then
                redirectUrl = "/CMMS/index.asp"
            End If
            Response.Redirect redirectUrl
            Response.End
        Else
            LoginError = T("login_error")
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="<%= IIf(Session("user_lang") = "en", "en", "es") %>">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%= T("login_title") %> — <%= T("app_name") %></title>
<link rel="stylesheet" href="/CMMS/assets/css/app.css">
<style>
.pass-toggle { position:relative; }
.pass-toggle input { padding-right: 44px; }
.pass-toggle-btn {
  position: absolute;
  right: 12px; top: 50%;
  transform: translateY(-50%);
  background: none;
  border: none;
  color: var(--text-muted);
  cursor: pointer;
  padding: 0;
  display: flex;
  align-items: center;
}
.pass-toggle-btn:hover { color: var(--text-primary); }
.login-bg-orb {
  position: absolute;
  border-radius: 50%;
  filter: blur(80px);
  opacity: 0.15;
  pointer-events: none;
}
</style>
</head>
<body>
<div id="toast-container"></div>

<div class="login-wrapper">

  <!-- ═══ Lado izquierdo: Ilustración ═══ -->
  <div class="login-left">
    <!-- Orbes decorativos -->
    <div class="login-bg-orb" style="width:400px;height:400px;background:var(--primary);top:-100px;left:-100px"></div>
    <div class="login-bg-orb" style="width:300px;height:300px;background:var(--secondary);bottom:-50px;right:50px"></div>

    <div class="login-illustration">
      <!-- Logo grande -->
      <div style="display:flex;align-items:center;justify-content:center;gap:16px;margin-bottom:32px">
        <div style="width:72px;height:72px;background:var(--gradient-primary);border-radius:20px;display:flex;align-items:center;justify-content:center;font-size:32px;font-weight:800;color:white;box-shadow:var(--shadow-glow)">CM</div>
        <div style="text-align:left">
          <div style="font-size:28px;font-weight:800;color:var(--text-primary);letter-spacing:-0.5px">CMMS</div>
          <div style="font-size:13px;color:var(--text-muted)">Sistema de Gestión de Mantenimiento</div>
        </div>
      </div>

      <h2 style="font-size:clamp(20px,3vw,32px)">Gestión inteligente<br>del mantenimiento industrial</h2>
      <p style="max-width:400px;margin:16px auto 0">Controle equipos, órdenes de trabajo e inventario con una plataforma diseñada para equipos de mantenimiento modernos.</p>

      <!-- Features -->
      <div class="login-features">
        <div class="login-feature">
          <div class="login-feature-icon" style="background:var(--primary-light);color:var(--primary)">📋</div>
          <div>
            <div style="font-weight:600;color:var(--text-primary);font-size:14px">Órdenes de Trabajo</div>
            <div style="font-size:12px;color:var(--text-muted)">Preventivo, correctivo, predictivo y emergencia</div>
          </div>
        </div>
        <div class="login-feature">
          <div class="login-feature-icon" style="background:var(--success-light);color:var(--success)">📦</div>
          <div>
            <div style="font-weight:600;color:var(--text-primary);font-size:14px">Control de Inventario</div>
            <div style="font-size:12px;color:var(--text-muted)">Alertas automáticas de stock y movimientos</div>
          </div>
        </div>
        <div class="login-feature">
          <div class="login-feature-icon" style="background:var(--warning-light);color:var(--warning)">📊</div>
          <div>
            <div style="font-weight:600;color:var(--text-primary);font-size:14px">Reportes y Analytics</div>
            <div style="font-size:12px;color:var(--text-muted)">Dashboard personalizable con gráficas en tiempo real</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ═══ Lado derecho: Formulario ═══ -->
  <div class="login-right">
    <div class="login-form-card">

      <!-- Logo pequeño para móvil -->
      <div class="login-logo" style="display:none" id="mobileLogo">
        <div class="login-logo-icon">CM</div>
        <div class="login-logo-text">
          <h1><%= T("app_short") %></h1>
          <p>Sistema de Mantenimiento</p>
        </div>
      </div>

      <h2 class="login-title"><%= T("login_title") %></h2>
      <p class="login-subtitle"><%= T("login_subtitle") %></p>

      <!-- Mensajes -->
      <% If LoginMsg <> "" Then %>
      <div class="alert alert-info" style="margin-bottom:20px">
        ℹ️ <span><%= Server.HTMLEncode(LoginMsg) %></span>
      </div>
      <% End If %>
      <% If LoginError <> "" Then %>
      <div class="alert alert-danger" style="margin-bottom:20px" id="loginErrorAlert">
        ⚠️ <span><%= Server.HTMLEncode(LoginError) %></span>
      </div>
      <% End If %>

      <!-- Formulario -->
      <form method="POST" action="login.asp" id="loginForm" novalidate>
        <input type="hidden" name="csrf_token" value="<%= GetCSRFToken() %>">

        <div class="form-group">
          <label class="form-label" for="username"><%= T("username") %></label>
          <div style="position:relative">
            <svg style="position:absolute;left:13px;top:50%;transform:translateY(-50%);color:var(--text-muted);pointer-events:none"
                 width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
              <circle cx="12" cy="7" r="4"/>
            </svg>
            <input type="text" id="username" name="username" class="form-control" required
                   style="padding-left:40px"
                   placeholder="Usuario o email"
                   value="<%= Server.HTMLEncode(Request.Form("username")) %>"
                   autocomplete="username">
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="password"><%= T("password") %></label>
          <div class="pass-toggle">
            <input type="password" id="password" name="password" class="form-control" required
                   placeholder="••••••••"
                   autocomplete="current-password">
            <button type="button" class="pass-toggle-btn" id="togglePass" aria-label="Mostrar contraseña">
              <svg id="eyeIcon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                <circle cx="12" cy="12" r="3"/>
              </svg>
            </button>
          </div>
        </div>

        <button type="submit" class="btn btn-primary w-full" id="loginBtn" style="height:44px;font-size:15px;margin-top:8px">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/>
            <polyline points="10 17 15 12 10 7"/>
            <line x1="15" y1="12" x2="3" y2="12"/>
          </svg>
          <%= T("login_btn") %>
        </button>
      </form>

      <!-- Footer del formulario -->
      <div style="margin-top:24px;text-align:center;font-size:12px;color:var(--text-muted)">
        <div style="margin-bottom:8px">Versión 1.0.0 — Classic ASP + SQL Server</div>
        <div>
          <a href="/CMMS/install.asp" style="color:var(--primary);font-size:11px">⚙️ Reinstalar sistema</a>
        </div>
      </div>

    </div>
  </div>
</div>

<script src="/CMMS/assets/js/app.js"></script>
<script>
// Toggle password visibility
document.getElementById('togglePass').addEventListener('click', function() {
    const passInput = document.getElementById('password');
    const icon = document.getElementById('eyeIcon');
    if (passInput.type === 'password') {
        passInput.type = 'text';
        icon.innerHTML = '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/>';
    } else {
        passInput.type = 'password';
        icon.innerHTML = '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>';
    }
});

// Show mobile logo on small screens
if(window.innerWidth <= 1024) {
    document.getElementById('mobileLogo').style.display = 'flex';
}

// Login button loading state
document.getElementById('loginForm').addEventListener('submit', function() {
    const btn = document.getElementById('loginBtn');
    btn.innerHTML = '<span class="spinner" style="width:18px;height:18px;border-width:2px;margin-right:8px"></span>Verificando...';
    btn.disabled = true;
});

// Auto-focus username
document.getElementById('username').focus();
</script>
</body>
</html>
