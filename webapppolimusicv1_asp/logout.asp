<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: cierre de sesion
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
Call DoLogout()
Response.Redirect "login.asp"
Response.End
%>
