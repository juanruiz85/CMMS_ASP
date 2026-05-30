<%
' =============================================================================
' CMMS - Template: Navegación (Sidebar + Topbar)
' Incluir después del header en cada página protegida
' Variables esperadas: PageTitle, PageModule
' =============================================================================

Dim NavUserName : NavUserName = CurrentUserFullName()
Dim NavUserRole : NavUserRole = CurrentUserRole()
Dim NavUserId   : NavUserId   = CurrentUserId()

' Iniciales del usuario para el avatar
Dim NavInitials
If Session("user_firstname") <> "" Then
    NavInitials = UCase(Left(Session("user_firstname"), 1))
    If Session("user_lastname") <> "" Then
        NavInitials = NavInitials & UCase(Left(Session("user_lastname"), 1))
    End If
Else
    NavInitials = UCase(Left(Session("user_name"), 2))
End If

' Helper: determinar si la ruta actual coincide con el módulo
Function IsCurrentModule(moduleName)
    Dim currentPath
    currentPath = LCase(Request.ServerVariables("SCRIPT_NAME"))
    IsCurrentModule = (InStr(currentPath, LCase(moduleName)) > 0)
End Function
%>

<!-- ═══════════════════════════════════════════════════ SIDEBAR ═══ -->
<nav class="sidebar" id="mainSidebar" role="navigation" aria-label="Menú principal">

  <!-- Logo -->
  <div class="sidebar-logo">
    <div class="sidebar-logo-icon" aria-hidden="true">CM</div>
    <div class="sidebar-logo-text">
      <h2><%= T("app_short") %></h2>
      <span>Mantenimiento Industrial</span>
    </div>
  </div>

  <!-- Navegación -->
  <div class="sidebar-nav">

    <!-- Dashboard -->
    <a href="/CMMS/index.asp"
       class="nav-item <%= IIf(IsCurrentModule("/CMMS/index.asp") Or PageModule = "dashboard", "active", "") %>"
       data-path="index.asp"
       title="<%= T("nav_dashboard") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <rect x="3" y="3" width="7" height="7" rx="1"/>
          <rect x="14" y="3" width="7" height="7" rx="1"/>
          <rect x="3" y="14" width="7" height="7" rx="1"/>
          <rect x="14" y="14" width="7" height="7" rx="1"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_dashboard") %></span>
    </a>

    <!-- Sección: Gestión -->
    <div class="nav-section-label"><%= T("nav_maintenance") %></div>

    <!-- Plantas -->
    <a href="/CMMS/modules/plants/index.asp"
       class="nav-item <%= IIf(PageModule = "plants", "active", "") %>"
       data-path="plants"
       title="<%= T("nav_plants") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M3 21h18M5 21V7l7-4 7 4v14M9 21V11h6v10"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_plants") %></span>
    </a>

    <!-- Equipos -->
    <a href="/CMMS/modules/assets_module/index.asp"
       class="nav-item <%= IIf(PageModule = "assets", "active", "") %>"
       data-path="assets_module"
       title="<%= T("nav_assets") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <circle cx="12" cy="12" r="3"/>
          <path d="M19.07 4.93a10 10 0 0 1 0 14.14M4.93 4.93a10 10 0 0 0 0 14.14M16.24 7.76a6 6 0 0 1 0 8.49M7.76 7.76a6 6 0 0 0 0 8.49"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_assets") %></span>
    </a>

    <!-- Órdenes de Trabajo -->
    <a href="/CMMS/modules/work_orders/index.asp"
       class="nav-item <%= IIf(PageModule = "work_orders", "active", "") %>"
       data-path="work_orders"
       title="<%= T("nav_work_orders") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/>
          <rect x="9" y="3" width="6" height="4" rx="1"/>
          <line x1="9" y1="12" x2="15" y2="12"/>
          <line x1="9" y1="16" x2="13" y2="16"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_work_orders") %></span>
    </a>

    <!-- Inventario -->
    <a href="/CMMS/modules/inventory/index.asp"
       class="nav-item <%= IIf(PageModule = "inventory", "active", "") %>"
       data-path="inventory"
       title="<%= T("nav_inventory") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/>
          <polyline points="3.27 6.96 12 12.01 20.73 6.96"/>
          <line x1="12" y1="22.08" x2="12" y2="12"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_inventory") %></span>
    </a>

    <!-- Reportes -->
    <a href="/CMMS/modules/reports/index.asp"
       class="nav-item <%= IIf(PageModule = "reports", "active", "") %>"
       data-path="reports"
       title="<%= T("nav_reports") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M18 20V10M12 20V4M6 20v-6"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_reports") %></span>
    </a>

    <!-- Solicitudes de Trabajo -->
    <a href="/CMMS/modules/work_requests/index.asp"
       class="nav-item <%= IIf(PageModule = "work_requests", "active", "") %>"
       data-path="work_requests"
       title="Solicitudes">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
          <path d="M12 18v-6"/>
          <path d="M9 15h6"/>
        </svg>
      </span>
      <span class="nav-label">Solicitudes</span>
    </a>

    <!-- Sección: Admin (solo admin/supervisor) -->
    <% If IsSupervisorOrAdmin() Then %>
    <div class="nav-section-label"><%= T("nav_admin") %></div>

    <a href="/CMMS/modules/users/index.asp"
       class="nav-item <%= IIf(PageModule = "users", "active", "") %>"
       data-path="users"
       title="<%= T("nav_users") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
          <circle cx="9" cy="7" r="4"/>
          <path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_users") %></span>
    </a>

    <% If IsAdmin() Then %>
    <a href="/CMMS/modules/admin/index.asp"
       class="nav-item <%= IIf(PageModule = "admin", "active", "") %>"
       data-path="admin"
       title="<%= T("nav_settings") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <circle cx="12" cy="12" r="3"/>
          <path d="M19.07 4.93l-1.41 1.41M4.93 4.93l1.41 1.41M19.07 19.07l-1.41-1.41M4.93 19.07l1.41-1.41M20 12h-2M6 12H4M12 4V2M12 22v-2"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_settings") %></span>
    </a>

    <a href="/CMMS/modules/admin/logs.asp"
       class="nav-item <%= IIf(PageModule = "logs", "active", "") %>"
       data-path="logs"
       title="<%= T("nav_logs") %>">
      <span class="nav-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
          <line x1="16" y1="13" x2="8" y2="13"/>
          <line x1="16" y1="17" x2="8" y2="17"/>
          <polyline points="10 9 9 9 8 9"/>
        </svg>
      </span>
      <span class="nav-label"><%= T("nav_logs") %></span>
    </a>
    <% End If %>
    <% End If %>

  </div><!-- /sidebar-nav -->

  <!-- Usuario en el footer del sidebar -->
  <div class="sidebar-footer">
    <a href="/CMMS/modules/users/profile.asp" class="sidebar-user" title="<%= T("nav_profile") %>">
      <div class="user-avatar">
        <% If Session("user_avatar") <> "" Then %>
        <img src="/CMMS/uploads/<%= HtmlEncode(Session("user_avatar")) %>" alt="Avatar">
        <% Else %>
        <%= HtmlEncode(NavInitials) %>
        <% End If %>
      </div>
      <div class="user-info">
        <div class="user-name"><%= HtmlEncode(NavUserName) %></div>
        <div class="user-role"><%= HtmlEncode(NavUserRole) %></div>
      </div>
    </a>
  </div>

</nav><!-- /sidebar -->

<!-- ═════════════════════════════════════════════════════ TOPBAR ═══ -->
<header class="topbar" role="banner">

  <!-- Toggle móvil -->
  <button class="topbar-toggle" id="sidebarToggle" aria-label="Abrir menú">
    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <line x1="3" y1="6" x2="21" y2="6"/>
      <line x1="3" y1="12" x2="21" y2="12"/>
      <line x1="3" y1="18" x2="21" y2="18"/>
    </svg>
  </button>

  <!-- Breadcrumb -->
  <div class="topbar-breadcrumb">
    <span class="breadcrumb-item"><a href="/CMMS/index.asp" style="color:inherit;text-decoration:none"><%= T("app_short") %></a></span>
    <% If PageModule <> "" Then %>
    <span class="breadcrumb-sep" aria-hidden="true">/</span>
    <span class="breadcrumb-item current"><%= HtmlEncode(PageTitle) %></span>
    <% End If %>
  </div>

  <!-- Acciones del Topbar -->
  <div class="topbar-actions">

    <!-- Botón Notificaciones -->
    <div class="dropdown">
      <button class="topbar-btn" id="notifBell" data-dropdown="notifDropdown"
              aria-label="<%= T("notif_title") %>" title="<%= T("notif_title") %>">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/>
          <path d="M13.73 21a2 2 0 0 1-3.46 0"/>
        </svg>
        <% If UnreadNotifCount > 0 Then %>
        <span class="badge-dot" aria-label="<%= UnreadNotifCount %> notificaciones sin leer"></span>
        <% End If %>
      </button>

      <!-- Dropdown de notificaciones -->
      <div id="notifDropdown" class="dropdown-menu" style="min-width:320px;right:0;" role="dialog" aria-label="Notificaciones">
        <div style="display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid var(--border-subtle)">
          <span style="font-weight:600;color:var(--text-primary)"><%= T("notif_title") %></span>
          <% If UnreadNotifCount > 0 Then %>
          <button onclick="CMMS.Notifications.markAllRead()" class="btn btn-ghost btn-sm" style="font-size:11px">
            <%= T("notif_mark_all") %>
          </button>
          <% End If %>
        </div>
        <%
        Dim oConnNotif, oRSNotif
        Set oConnNotif = GetConnection()
        Set oRSNotif   = oConnNotif.Execute("SELECT TOP 8 id, title, message, type, is_read, created_at FROM cmms_notifications WHERE user_id = " & CurrentUserId() & " ORDER BY created_at DESC")
        If oRSNotif.EOF Then
        %>
        <div style="padding:24px;text-align:center;color:var(--text-muted);font-size:13px">
          <%= T("notif_none") %>
        </div>
        <% Else %>
        <% Do While Not oRSNotif.EOF %>
        <div class="dropdown-item notif-item <%= IIf(oRSNotif("is_read") = 0, "unread", "") %>"
             style="flex-direction:column;align-items:flex-start;gap:2px;<%= IIf(oRSNotif("is_read") = 0, "background:rgba(99,102,241,0.05);", "") %>">
          <span style="font-weight:500;color:var(--text-primary);font-size:13px"><%= HtmlEncode(NullStr(oRSNotif("title"))) %></span>
          <span style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(NullStr(oRSNotif("message"))) %></span>
          <span style="font-size:10px;color:var(--text-muted)"><%= TimeAgo(oRSNotif("created_at")) %></span>
        </div>
        <%   oRSNotif.MoveNext
        Loop
        oRSNotif.Close %>
        <% End If %>
        <div class="dropdown-divider"></div>
        <a href="/CMMS/modules/admin/logs.asp" class="dropdown-item" style="justify-content:center;font-size:12px">
          Ver todos los logs
        </a>
      </div>
    </div>

    <!-- Menú de usuario -->
    <div class="dropdown">
      <button class="topbar-btn" data-dropdown="userMenuDropdown"
              style="width:auto;padding:0 10px;gap:8px" aria-label="Menú de usuario">
        <div class="user-avatar" style="width:28px;height:28px;font-size:10px;flex-shrink:0">
          <% If Session("user_avatar") <> "" Then %>
          <img src="/CMMS/uploads/<%= HtmlEncode(Session("user_avatar")) %>" alt="Avatar">
          <% Else %>
          <%= HtmlEncode(NavInitials) %>
          <% End If %>
        </div>
        <span style="font-size:13px;color:var(--text-secondary);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">
          <%= HtmlEncode(NavUserName) %>
        </span>
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;color:var(--text-muted)">
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>

      <div id="userMenuDropdown" class="dropdown-menu" style="right:0;min-width:200px">
        <div style="padding:12px 16px;border-bottom:1px solid var(--border-subtle)">
          <div style="font-weight:600;color:var(--text-primary);font-size:13px"><%= HtmlEncode(NavUserName) %></div>
          <div style="font-size:11px;color:var(--text-muted)"><%= HtmlEncode(Session("user_email")) %></div>
        </div>
        <a href="/CMMS/modules/users/profile.asp" class="dropdown-item">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
            <circle cx="12" cy="7" r="4"/>
          </svg>
          <%= T("nav_profile") %>
        </a>
        <% If IsAdmin() Then %>
        <a href="/CMMS/modules/admin/settings.asp" class="dropdown-item">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="3"/>
            <path d="M19.07 4.93l-1.41 1.41M4.93 4.93l1.41 1.41M19.07 19.07l-1.41-1.41M4.93 19.07l1.41-1.41M20 12h-2M6 12H4M12 4V2M12 22v-2"/>
          </svg>
          <%= T("nav_settings") %>
        </a>
        <% End If %>
        <div class="dropdown-divider"></div>
        <a href="/CMMS/logout.asp" class="dropdown-item danger">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
            <polyline points="16 17 21 12 16 7"/>
            <line x1="21" y1="12" x2="9" y2="12"/>
          </svg>
          <%= T("logout") %>
        </a>
      </div>
    </div>

  </div><!-- /topbar-actions -->
</header><!-- /topbar -->

<!-- ═════════════════════════════════════════════════ MAIN CONTENT ═══ -->
<main class="main-content" id="mainContent" role="main">
<div class="page-wrapper">
