<%
'=====================================================================
'  PoliMusic :: cabecera comun (barra lateral + barra superior)
'  La pagina que incluye este archivo debe definir antes:
'     strBase, pageTitle, pageId
'=====================================================================
If Len(pageTitle) = 0 Then pageTitle = APP_NAME
Dim hdrAdmin, hdrSearch, hdrPhoto, hdrName, hdrRole, hdrInitials
Dim hdrFlash, hdrFlashType, hdrQ

hdrAdmin    = IsAdmin()
hdrPhoto    = CurrentUserPhoto()
hdrName     = CurrentUserName()
hdrRole     = RoleLabel(Session("UserType"))
hdrInitials = CurrentUserInitials()
hdrSearch   = (pageId = "home")
hdrQ        = TrimStr(Request.QueryString("q"))

hdrFlash = ""
hdrFlashType = ""
If Len(CStr(Session("FlashMsg") & "")) > 0 Then
    hdrFlash = CStr(Session("FlashMsg"))
    hdrFlashType = CStr(Session("FlashType") & "")
    If Len(hdrFlashType) = 0 Then hdrFlashType = "info"
    Session("FlashMsg") = ""
    Session("FlashType") = ""
End If

Dim hdrErr
hdrErr = TrimStr(Request.QueryString("e"))
%><!DOCTYPE html>
<html lang="es" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark light">
<title><%=H(pageTitle)%> &middot; <%=APP_NAME%></title>
<link rel="stylesheet" href="<%=strBase%>css/styles.css">
<script>
(function () {
    try {
        var t = localStorage.getItem('polimusic-theme');
        if (!t) { t = window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark'; }
        document.documentElement.setAttribute('data-theme', t);
    } catch (e) {
        document.documentElement.setAttribute('data-theme', 'dark');
    }
})();
</script>
</head>
<body class="<%=strBodyClass%>">

<!--#include file="sprite.asp"-->

<div class="app">
    <aside class="sidebar" id="sidebar">
        <a class="brand" href="<%=strBase%>home.asp">
            <span class="brand-mark"><svg><use href="#i-music"></use></svg></span>
            <span class="brand-text"><strong>Poli</strong>Music<span class="brand-dot">.</span></span>
        </a>

        <nav class="nav" aria-label="Navegación principal">
            <a class="nav-item <%=IIf(pageId = "home", "is-active", "")%>" href="<%=strBase%>home.asp">
                <svg><use href="#i-home"></use></svg><span>Inicio</span>
            </a>
            <a class="nav-item" href="<%=strBase%>home.asp#songlist">
                <svg><use href="#i-list"></use></svg><span>Song List</span>
            </a>
            <a class="nav-item" href="<%=strBase%>home.asp#playlists">
                <svg><use href="#i-disc"></use></svg><span>Playlists</span>
            </a>
        </nav>

        <% If hdrAdmin Then %>
        <div class="side-title">
            <svg><use href="#i-shield"></use></svg><span>Administración</span>
        </div>
        <nav class="nav nav-admin" aria-label="Administración">
            <a class="nav-item <%=IIf(pageId = "admin", "is-active", "")%>" href="<%=strBase%>admin.asp">
                <svg><use href="#i-chart"></use></svg><span>Panel</span>
            </a>
            <a class="nav-item <%=IIf(pageId = "admin-songs", "is-active", "")%>" href="<%=strBase%>admin_songs.asp">
                <svg><use href="#i-music"></use></svg><span>Canciones</span>
            </a>
            <a class="nav-item <%=IIf(pageId = "admin-artists", "is-active", "")%>" href="<%=strBase%>admin_artists.asp">
                <svg><use href="#i-artist"></use></svg><span>Artistas</span>
            </a>
            <a class="nav-item <%=IIf(pageId = "admin-albums", "is-active", "")%>" href="<%=strBase%>admin_albums.asp">
                <svg><use href="#i-disc"></use></svg><span>Álbumes</span>
            </a>
            <a class="nav-item <%=IIf(pageId = "admin-playlists", "is-active", "")%>" href="<%=strBase%>admin_playlists.asp">
                <svg><use href="#i-list"></use></svg><span>Playlists</span>
            </a>
            <a class="nav-item <%=IIf(pageId = "admin-users", "is-active", "")%>" href="<%=strBase%>admin_users.asp">
                <svg><use href="#i-users"></use></svg><span>Usuarios</span>
            </a>
        </nav>
        <% End If %>

        <div class="side-user">
            <img class="avatar" src="<%=hdrPhoto%>" alt="Foto de <%=H(hdrName)%>">
            <div class="side-user-info">
                <strong title="<%=H(hdrName)%>"><%=H(hdrName)%></strong>
                <span class="role-pill <%=IIf(hdrAdmin, "role-admin", "role-user")%>"><%=H(hdrRole)%></span>
            </div>
            <form method="post" action="<%=strBase%>logout.asp" class="logout-form">
                <button type="submit" class="icon-btn" title="Cerrar sesión" aria-label="Cerrar sesión">
                    <svg><use href="#i-logout"></use></svg>
                </button>
            </form>
        </div>
    </aside>

    <div class="main">
        <header class="topbar">
            <button class="icon-btn only-mobile" id="btnMenu" type="button" aria-label="Abrir menú">
                <svg><use href="#i-menu"></use></svg>
            </button>

            <div class="topbar-title">
                <span class="topbar-eyebrow">PoliMusic</span>
                <h1><%=H(pageTitle)%></h1>
            </div>

            <% If hdrSearch Then %>
            <form class="search" method="get" action="<%=strBase%>home.asp" role="search">
                <svg><use href="#i-search"></use></svg>
                <input type="search" name="q" id="songSearch" value="<%=H(hdrQ)%>"
                       placeholder="Buscar canciones, artistas o álbumes…" autocomplete="off">
                <button type="submit" class="search-submit">Buscar</button>
            </form>
            <% End If %>

            <div class="topbar-actions">
                <button class="icon-btn theme-toggle" id="btnTheme" type="button"
                        title="Cambiar entre modo claro y oscuro" aria-label="Cambiar tema">
                    <svg class="ico-sun"><use href="#i-sun"></use></svg>
                    <svg class="ico-moon"><use href="#i-moon"></use></svg>
                </button>
                <div class="topbar-user" title="<%=H(hdrName)%> · <%=H(hdrRole)%>">
                    <img class="avatar sm" src="<%=hdrPhoto%>" alt="">
                    <div class="topbar-user-text">
                        <strong><%=H(hdrName)%></strong>
                        <span><%=H(hdrRole)%></span>
                    </div>
                </div>
            </div>
        </header>

        <main class="content" id="content">
            <% If Len(hdrFlash) > 0 Then %>
            <div class="flash flash-<%=H(hdrFlashType)%>" role="status">
                <svg><use href="#i-check"></use></svg>
                <span><%=H(hdrFlash)%></span>
            </div>
            <% End If %>
            <% If hdrErr = "denied" Then %>
            <div class="flash flash-error" role="alert">
                <svg><use href="#i-close"></use></svg>
                <span>Acceso restringido: solo el usuario administrador puede gestionar el catálogo.</span>
            </div>
            <% End If %>
