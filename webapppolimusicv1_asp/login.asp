<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: inicio de sesion
'  Aqui se crean las variables de sesion que consumen las demas vistas.
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
Dim lUser, lPass, lErr, lRs, lOk, lId, lName, lType, lMail, lPhoto

lErr  = ""
lOk   = False
lUser = ""
lPass = ""

If IsLoggedIn() Then
    Response.Redirect "home.asp"
    Response.End
End If

If TrimStr(Request.QueryString("e")) = "session" Then
    lErr = "Tu sesión expiró o no has iniciado sesión. Ingresa nuevamente."
End If

If IsPost() Then
    lUser = TrimStr(Request.Form("username"))
    lPass = CStr(Request.Form("password") & "")

    If Len(lUser) = 0 Or Len(lPass) = 0 Then
        lErr = "Debes completar el usuario y la contraseña."
    Else
        On Error Resume Next
        Set lRs = DbQuery("SELECT TOP 1 ID_USER, USERNAME, PASSWORD, EMAIL, USER_TYPE, USER_PHOTO " & _
                          "FROM TBL_USER WHERE USERNAME = ?", Array(lUser))
        If Err.Number <> 0 Then
            lErr = "No se pudo conectar con la base de datos. Verifica el servidor SQL Server."
            Err.Clear
        End If
        On Error GoTo 0

        If Len(lErr) = 0 And IsObject(lRs) Then
            If lRs.EOF Then
                lErr = "El usuario no existe."
            ElseIf Not PasswordMatches(lPass, lRs("PASSWORD").Value) Then
                lErr = "La contraseña no es correcta."
            Else
                lId    = lRs("ID_USER").Value
                lName  = CStr(lRs("USERNAME").Value)
                lType  = ToInt(lRs("USER_TYPE").Value, 0)
                lMail  = CStr(lRs("EMAIL").Value & "")
                lPhoto = CStr(lRs("USER_PHOTO").Value & "")

                Call DoLogin(lId, lName, lType, lMail, lPhoto)
                lOk = True
            End If
            If lRs.State <> 0 Then lRs.Close
            Set lRs = Nothing
        End If

        If lOk Then
            Response.Redirect "home.asp"
            Response.End
        End If
    End If
End If
%><!DOCTYPE html>
<html lang="es" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark light">
<title>Iniciar sesión &middot; <%=APP_NAME%></title>
<link rel="stylesheet" href="css/styles.css">
<script>
(function () {
    try {
        var t = localStorage.getItem('polimusic-theme');
        if (!t) { t = window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark'; }
        document.documentElement.setAttribute('data-theme', t);
    } catch (e) { document.documentElement.setAttribute('data-theme', 'dark'); }
})();
</script>
</head>
<body class="login-body">
<!--#include file="includes/sprite.asp"-->

<div class="login-shell">

    <section class="login-hero">
        <div class="login-hero-bg" aria-hidden="true">
            <span class="blob blob-1"></span>
            <span class="blob blob-2"></span>
            <span class="blob blob-3"></span>
            <span class="grid-lines"></span>
        </div>

        <div class="login-hero-content">
            <div class="brand brand-lg">
                <span class="brand-mark">
                    <svg viewBox="0 0 24 24" aria-hidden="true"><use href="#i-music"></use></svg>
                </span>
                <span class="brand-text"><strong>Poli</strong>Music<span class="brand-dot">.</span></span>
            </div>

            <h2 class="login-claim">Tu música suena<br><span>en azul.</span></h2>
            <p class="login-sub">
                Escucha tu catálogo, arma tu cola de reproducción y administra
                artistas, álbumes y playlists desde un solo lugar.
            </p>

            <ul class="login-features">
                <li><span class="dot"></span>Reproductor con lista consecutiva y modo bucle</li>
                <li><span class="dot"></span>Catálogo administrado por el rol <strong>admin</strong></li>
                <li><span class="dot"></span>Modo claro y oscuro con la misma identidad visual</li>
            </ul>

            <div class="login-eq" aria-hidden="true">
                <i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i>
            </div>
        </div>
    </section>

    <section class="login-panel">
        <button class="icon-btn theme-toggle login-theme" id="btnTheme" type="button"
                title="Cambiar entre modo claro y oscuro" aria-label="Cambiar tema">
            <svg class="ico-sun"><use href="#i-sun"></use></svg>
            <svg class="ico-moon"><use href="#i-moon"></use></svg>
        </button>

        <div class="login-card">
            <span class="chip"><svg><use href="#i-spark"></use></svg> Bienvenido de nuevo</span>
            <h1>Inicia sesión</h1>
            <p class="login-card-sub">Usa tu cuenta para entrar a tu biblioteca musical.</p>

            <% If Len(lErr) > 0 Then %>
            <div class="flash flash-error" role="alert">
                <svg><use href="#i-close"></use></svg>
                <span><%=H(lErr)%></span>
            </div>
            <% End If %>

            <form method="post" action="login.asp" class="login-form" autocomplete="on">
                <label class="field">
                    <span>Usuario</span>
                    <div class="input-wrap">
                        <svg><use href="#i-users"></use></svg>
                        <input type="text" name="username" value="<%=H(lUser)%>"
                               placeholder="userAdmin" required autofocus maxlength="50">
                    </div>
                </label>

                <label class="field">
                    <span>Contraseña</span>
                    <div class="input-wrap">
                        <svg><use href="#i-shield"></use></svg>
                        <input type="password" name="password" placeholder="••••••••" required maxlength="50">
                        <button type="button" class="peek" id="btnPeek" title="Mostrar contraseña">Ver</button>
                    </div>
                </label>

                <button type="submit" class="btn btn-primary btn-block">
                    Entrar
                    <svg><use href="#i-next"></use></svg>
                </button>
            </form>

            <div class="login-demo">
                <h4>Cuentas de prueba</h4>
                <div class="demo-grid">
                    <div class="demo-item">
                        <strong>Administrador</strong>
                        <code>userAdmin</code>
                        <code>12345678</code>
                    </div>
                    <div class="demo-item">
                        <strong>Usuario normal</strong>
                        <code>userNormal</code>
                        <code>12345678</code>
                    </div>
                </div>
            </div>
        </div>

        <p class="login-foot">PoliMusic · ASP clásico + SQL Server · <%=Year(Now())%></p>
    </section>
</div>

<script>
(function () {
    var btn = document.getElementById('btnTheme');
    if (btn) {
        btn.addEventListener('click', function () {
            var cur = document.documentElement.getAttribute('data-theme');
            var nxt = cur === 'dark' ? 'light' : 'dark';
            document.documentElement.setAttribute('data-theme', nxt);
            try { localStorage.setItem('polimusic-theme', nxt); } catch (e) {}
        });
    }
    var peek = document.getElementById('btnPeek');
    var pass = document.querySelector('input[name="password"]');
    if (peek && pass) {
        peek.addEventListener('click', function () {
            var show = pass.type === 'password';
            pass.type = show ? 'text' : 'password';
            peek.textContent = show ? 'Ocultar' : 'Ver';
        });
    }
})();
</script>
</body>
</html>
