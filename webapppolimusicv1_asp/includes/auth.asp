<%
'=====================================================================
'  PoliMusic :: control de sesion y roles
'  USER_TYPE = 0 -> usuario normal   |   USER_TYPE = 1 -> administrador
'=====================================================================
Function IsLoggedIn()
    IsLoggedIn = (Len(CStr(Session("UserID") & "")) > 0)
End Function

Function IsAdmin()
    IsAdmin = False
    If IsLoggedIn() Then
        IsAdmin = (ToInt(Session("UserType"), 0) = 1)
    End If
End Function

Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect strBase & "login.asp?e=session"
        Response.End
    End If
End Sub

Sub RequireAdmin()
    If Not IsLoggedIn() Then
        Response.Redirect strBase & "login.asp?e=session"
        Response.End
    End If
    If Not IsAdmin() Then
        Response.Redirect strBase & "home.asp?e=denied"
        Response.End
    End If
End Sub

Sub DoLogin(idUser, username, userType, email, photo)
    Session("UserID")   = idUser
    Session("Username") = username
    Session("UserType") = userType
    Session("UserEmail") = email
    Session("UserPhoto") = photo
    Session("LoginAt")  = Now()
End Sub

Sub DoLogout()
    Session.Abandon
End Sub
%>
