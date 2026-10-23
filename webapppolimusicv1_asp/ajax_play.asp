<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: contador de reproducciones (AJAX)
'  Suma 1 al campo PLAYS de la cancion indicada y devuelve el total.
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
Response.ContentType = "application/json"

Dim aId, aPlays, aOk
aId = ToInt(Request.QueryString("id"), 0)
aOk = False
aPlays = 0

If IsLoggedIn() And aId > 0 Then
    On Error Resume Next
    Call DbExec("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(aId))
    If Err.Number = 0 Then
        aPlays = ToInt(DbScalar("SELECT ISNULL(PLAYS, 0) FROM TBL_SONG WHERE ID_SONG = ?", Array(aId)), 0)
        aOk = True
    End If
    Err.Clear
    On Error GoTo 0
End If

Response.Write "{""ok"":" & IIf(aOk, "true", "false") & ",""plays"":" & aPlays & "}"
Call CloseConn()
Response.End
%>
