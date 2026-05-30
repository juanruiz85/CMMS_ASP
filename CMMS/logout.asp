<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/i18n.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
' =============================================================================
' CMMS - Logout (logout.asp)
' =============================================================================
DoLogout()
Response.Redirect "/CMMS/login.asp"
Response.End
%>
