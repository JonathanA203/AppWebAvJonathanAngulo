<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: punto de entrada
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
If IsLoggedIn() Then
    Response.Redirect "home.asp"
Else
    Response.Redirect "login.asp"
End If
Response.End
%>
