<%
' =============================================================================
' CMMS - Template: Header HTML
' Incluir al inicio de cada página protegida
' Variables esperadas: PageTitle, PageModule, BreadcrumbItems
' =============================================================================
' Dependencias (ya incluidas en la página que llama a este template):
'   <!--#include virtual="/CMMS/config/database.asp"-->
'   <!--#include virtual="/CMMS/core/functions.asp"-->
'   <!--#include virtual="/CMMS/core/i18n.asp"-->
'   <!--#include virtual="/CMMS/core/auth.asp"-->

' Valores por defecto para PageTitle y PageModule
' Estas variables deben ser establecidas por la página que incluye este template
' Si no se establecen, se usan los valores por defecto aquí
If Len(PageTitle & "") = 0 Then PageTitle = T("app_name")
If Len(PageModule & "") = 0 Then PageModule = ""

' Contar notificaciones no leídas
Dim UnreadNotifCount : UnreadNotifCount = CountUnreadNotifications(CurrentUserId())


%><!DOCTYPE html>
<html lang="<%= IIf(Session("user_lang") = "en", "en", "es") %>">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<meta name="theme-color" content="#6366f1">
<title><%= HtmlEncode(PageTitle) %> — <%= T("app_short") %></title>
<meta name="description" content="<%= T("app_name") %> - Sistema de Gestión de Mantenimiento">

<!-- Preconnect fonts -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>

<!-- CSS -->
<link rel="stylesheet" href="/CMMS/assets/css/app.css">

<!-- Chart.js -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
</head>
<body>
<div id="toast-container"></div>
<div class="app-layout">
