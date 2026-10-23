<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: panel de administracion
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase   = "./"
pageTitle = "Panel de administración"
pageId    = "admin"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim cSongs, cArtists, cAlbums, cUsers, cPlaylists, cPlays, cAdmins
Dim rsTop, rsRecent, i

cSongs     = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG", Array()), 0)
cArtists   = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", Array()), 0)
cAlbums    = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", Array()), 0)
cUsers     = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER", Array()), 0)
cPlaylists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST", Array()), 0)
cPlays     = ToInt(DbScalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", Array()), 0)
cAdmins    = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1", Array()), 0)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración</span>
        <h2>Resumen general</h2>
        <p>
            Sesión iniciada como <strong><%=H(Session("Username"))%></strong> con rol
            <strong>Administrador</strong>. Desde aquí puedes gestionar el catálogo completo:
            canciones, artistas, álbumes, playlists y usuarios.
        </p>
    </div>
    <div class="page-head-actions">
        <a class="btn btn-primary" href="<%=strBase%>admin_song_form.asp">
            <svg><use href="#i-plus"></use></svg> Nueva canción
        </a>
        <a class="btn btn-ghost" href="<%=strBase%>home.asp">
            <svg><use href="#i-play"></use></svg> Ir al reproductor
        </a>
    </div>
</section>

<div class="stat-grid">
    <article class="stat-card">
        <span class="stat-icon"><svg><use href="#i-music"></use></svg></span>
        <div><strong><%=FmtNum(cSongs)%></strong><span>Canciones</span></div>
        <a href="<%=strBase%>admin_songs.asp" class="stat-link">Gestionar</a>
    </article>
    <article class="stat-card">
        <span class="stat-icon"><svg><use href="#i-artist"></use></svg></span>
        <div><strong><%=FmtNum(cArtists)%></strong><span>Artistas</span></div>
        <a href="<%=strBase%>admin_artists.asp" class="stat-link">Gestionar</a>
    </article>
    <article class="stat-card">
        <span class="stat-icon"><svg><use href="#i-disc"></use></svg></span>
        <div><strong><%=FmtNum(cAlbums)%></strong><span>Álbumes</span></div>
        <a href="<%=strBase%>admin_albums.asp" class="stat-link">Gestionar</a>
    </article>
    <article class="stat-card">
        <span class="stat-icon"><svg><use href="#i-list"></use></svg></span>
        <div><strong><%=FmtNum(cPlaylists)%></strong><span>Playlists</span></div>
        <a href="<%=strBase%>admin_playlists.asp" class="stat-link">Gestionar</a>
    </article>
    <article class="stat-card">
        <span class="stat-icon"><svg><use href="#i-users"></use></svg></span>
        <div><strong><%=FmtNum(cUsers)%></strong><span>Usuarios</span></div>
        <a href="<%=strBase%>admin_users.asp" class="stat-link">Gestionar</a>
    </article>
    <article class="stat-card stat-accent">
        <span class="stat-icon"><svg><use href="#i-chart"></use></svg></span>
        <div><strong><%=FmtNum(cPlays)%></strong><span>Reproducciones</span></div>
        <span class="stat-note"><%=FmtNum(cAdmins)%> admin(es)</span>
    </article>
</div>

<div class="split">
    <section class="panel">
        <header class="panel-head">
            <h3>Últimas canciones registradas</h3>
            <a class="link-more" href="<%=strBase%>admin_songs.asp">Ver todas</a>
        </header>
        <div class="table-wrap">
            <table class="data-table compact">
                <thead>
                    <tr><th>#</th><th>Canción</th><th>Artistas</th><th class="num">Plays</th><th></th></tr>
                </thead>
                <tbody>
                <%
                Set rsRecent = DbQuery("SELECT TOP 6 s.ID_SONG, s.SONG_NAME, ISNULL(s.PLAYS,0) AS PLAYS, " & _
                    " ISNULL((SELECT STUFF((SELECT DISTINCT ', ' + a.ARTIST_NAME FROM TBL_PLAYER p " & _
                    "   INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG " & _
                    "   FOR XML PATH('')), 1, 2, '')), 'Sin artista') AS ARTISTS " & _
                    "FROM TBL_SONG s ORDER BY s.ID_SONG DESC", Array())
                i = 0
                Do While Not rsRecent.EOF
                    i = i + 1
                %>
                    <tr>
                        <td class="muted"><%=Right("0" & i, 2)%></td>
                        <td class="strong"><%=H(rsRecent("SONG_NAME"))%></td>
                        <td class="muted"><%=H(rsRecent("ARTISTS"))%></td>
                        <td class="num"><%=FmtNum(rsRecent("PLAYS"))%></td>
                        <td class="row-actions">
                            <a class="mini-btn" href="<%=strBase%>admin_song_form.asp?id=<%=rsRecent("ID_SONG")%>"
                               title="Editar"><svg><use href="#i-edit"></use></svg></a>
                        </td>
                    </tr>
                <%
                    rsRecent.MoveNext
                Loop
                If rsRecent.State <> 0 Then rsRecent.Close
                Set rsRecent = Nothing
                If i = 0 Then
                %>
                    <tr><td colspan="5" class="muted center">Sin registros.</td></tr>
                <% End If %>
                </tbody>
            </table>
        </div>
    </section>

    <section class="panel">
        <header class="panel-head">
            <h3>Artistas con más canciones</h3>
            <a class="link-more" href="<%=strBase%>admin_artists.asp">Ver todos</a>
        </header>
        <ul class="rank-list">
        <%
        Set rsTop = DbQuery("SELECT TOP 6 a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
            " (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS TOTAL " & _
            "FROM TBL_ARTIST a ORDER BY TOTAL DESC, a.ARTIST_NAME ASC", Array())
        i = 0
        Do While Not rsTop.EOF
            i = i + 1
        %>
            <li>
                <span class="rank-num"><%=i%></span>
                <img class="rank-img" src="<%=ArtistPhotoUrl(rsTop("ARTIST_IMAGE").Value, rsTop("ARTIST_NAME").Value)%>"
                     alt="<%=H(rsTop("ARTIST_NAME"))%>">
                <div class="rank-info">
                    <strong><%=H(rsTop("ARTIST_NAME") & " " & Nz(rsTop("ARTIST_LASTNAME"), ""))%></strong>
                    <span><%=FmtNum(rsTop("TOTAL").Value)%> canciones</span>
                </div>
                <a class="mini-btn" href="<%=strBase%>admin_artist_form.asp?id=<%=rsTop("ID_ARTIST")%>"
                   title="Editar"><svg><use href="#i-edit"></use></svg></a>
            </li>
        <%
            rsTop.MoveNext
        Loop
        If rsTop.State <> 0 Then rsTop.Close
        Set rsTop = Nothing
        If i = 0 Then
        %>
            <li class="muted center">Sin artistas registrados.</li>
        <% End If %>
        </ul>
    </section>
</div>

<section class="panel">
    <header class="panel-head">
        <h3>Accesos rápidos</h3>
        <span class="muted small">Solo el rol administrador puede ejecutar estas acciones</span>
    </header>
    <div class="quick-grid">
        <a class="quick-card" href="<%=strBase%>admin_songs.asp">
            <span class="quick-icon"><svg><use href="#i-music"></use></svg></span>
            <strong>Canciones</strong>
            <span>Nombre, ruta del archivo y reproducciones</span>
        </a>
        <a class="quick-card" href="<%=strBase%>admin_artists.asp">
            <span class="quick-icon"><svg><use href="#i-artist"></use></svg></span>
            <strong>Artistas</strong>
            <span>Datos e imagen del artista</span>
        </a>
        <a class="quick-card" href="<%=strBase%>admin_albums.asp">
            <span class="quick-icon"><svg><use href="#i-disc"></use></svg></span>
            <strong>Álbumes</strong>
            <span>Título, año y portada</span>
        </a>
        <a class="quick-card" href="<%=strBase%>admin_playlists.asp">
            <span class="quick-icon"><svg><use href="#i-list"></use></svg></span>
            <strong>Playlists</strong>
            <span>Propietario, tipo y portada</span>
        </a>
        <a class="quick-card" href="<%=strBase%>admin_users.asp">
            <span class="quick-icon"><svg><use href="#i-users"></use></svg></span>
            <strong>Usuarios</strong>
            <span>Credenciales, foto y rol</span>
        </a>
    </div>
</section>

<!--#include file="includes/footer.asp"-->
