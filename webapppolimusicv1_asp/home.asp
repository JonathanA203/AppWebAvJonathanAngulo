<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: vista principal (Song List + reproductor)
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase   = "./"
pageTitle = "Song List"
pageId    = "home"
strBodyClass = "page-home"

Call RequireLogin()

Dim q, sqlSongs, rsSongs, rsStats
Dim totSongs, totArtists, totPlaylists, totPlays
Dim idx, idSong, sName, sPath, sPlays, sAlbum, sYear, sCover, sArtists, sAudio, sSearch
Dim jsonItem, jsonList
Dim rsPl, plCount

q = TrimStr(Request.QueryString("q"))

' ---- estadisticas ---------------------------------------------------
totSongs     = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG", Array()), 0)
totArtists   = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", Array()), 0)
totPlaylists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST", Array()), 0)
totPlays     = ToInt(DbScalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", Array()), 0)

' ---- listado de canciones -------------------------------------------
sqlSongs = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
           " ISNULL((SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t " & _
           "         INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
           "         WHERE t.ID_SONG = s.ID_SONG), '') AS ALBUM_NAME, " & _
           " ISNULL((SELECT TOP 1 al.RELEASE_YEAR FROM TBL_TRACKLIST t " & _
           "         INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
           "         WHERE t.ID_SONG = s.ID_SONG), 0) AS ALBUM_YEAR, " & _
           " ISNULL((SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t " & _
           "         INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
           "         WHERE t.ID_SONG = s.ID_SONG), '') AS ALBUM_IMAGE, " & _
           " ISNULL((SELECT STUFF((SELECT DISTINCT ', ' + a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,'') " & _
           "         FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
           "         WHERE p.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, '')), '') AS ARTISTS " & _
           "FROM TBL_SONG s "

If Len(q) > 0 Then
    sqlSongs = sqlSongs & "WHERE (s.SONG_NAME LIKE ? OR EXISTS (SELECT 1 FROM TBL_PLAYER px " & _
               "INNER JOIN TBL_ARTIST ax ON ax.ID_ARTIST = px.ID_ARTIST WHERE px.ID_SONG = s.ID_SONG " & _
               "AND (ax.ARTIST_NAME LIKE ? OR ISNULL(ax.ARTIST_LASTNAME,'') LIKE ?))) "
End If

sqlSongs = sqlSongs & "ORDER BY s.SONG_NAME"

If Len(q) > 0 Then
    Set rsSongs = DbQuery(sqlSongs, Array("%" & q & "%", "%" & q & "%", "%" & q & "%"))
Else
    Set rsSongs = DbQuery(sqlSongs, Array())
End If

idx = 0
jsonList = ""
%>
<!--#include file="includes/header.asp"-->

<section class="hero">
    <div class="hero-glow" aria-hidden="true"></div>
    <div class="hero-text">
        <span class="chip"><svg><use href="#i-spark"></use></svg> Tu biblioteca musical</span>
        <h2 class="hero-title">Bienvenido, <span class="grad"><%=H(Session("Username"))%></span></h2>
        <p class="hero-sub">
            Nos alegra verte de nuevo. Tu sesión está activa como
            <strong><%=H(RoleLabel(Session("UserType")))%></strong> y tienes
            <strong><%=FmtNum(totSongs)%> canciones</strong> listas para sonar.
            <% If IsAdmin() Then %>Como administrador puedes crear, editar y eliminar el catálogo.<% End If %>
        </p>

        <div class="hero-actions">
            <button type="button" class="btn btn-primary" id="btnPlayAll">
                <svg><use href="#i-play"></use></svg> Reproducir todo
            </button>
            <a class="btn btn-ghost" href="#songlist">
                <svg><use href="#i-list"></use></svg> Ir al listado
            </a>
        </div>

        <dl class="hero-stats">
            <div><dt>Canciones</dt><dd><%=FmtNum(totSongs)%></dd></div>
            <div><dt>Artistas</dt><dd><%=FmtNum(totArtists)%></dd></div>
            <div><dt>Playlists</dt><dd><%=FmtNum(totPlaylists)%></dd></div>
            <div><dt>Reproducciones</dt><dd><%=FmtNum(totPlays)%></dd></div>
        </dl>
    </div>

    <div class="hero-art" aria-hidden="true">
        <div class="hero-disc">
            <img src="<%=strBase%>SongFilesAndImages_resources/PlaylistImages/play-back-in-black.jpg" alt="">
            <span class="hero-disc-hole"></span>
        </div>
        <div class="hero-float">
            <span class="dot-live"></span>
            <div>
                <strong>Modo bucle</strong>
                <span>La lista no se detiene</span>
            </div>
        </div>
        <div class="hero-float hero-float-2">
            <svg><use href="#i-music"></use></svg>
            <div>
                <strong><%=FmtNum(totPlays)%> plays</strong>
                <span>en toda la biblioteca</span>
            </div>
        </div>
    </div>
</section>

<section class="section" id="songlist">
    <header class="section-head">
        <div class="section-copy">
            <span class="section-kicker">Song List</span>
            <h2>Listado de canciones</h2>
            <p>
                A continuación se presenta el listado de canciones disponibles en PoliMusic.
                Elige una portada para reproducirla en la barra inferior, o usa los controles
                para avanzar, retroceder y cambiar el modo de reproducción.
            </p>
        </div>
        <div class="section-tools">
            <span class="count-pill" id="songCount"><%=FmtNum(IIf(Len(q) > 0, 0, totSongs))%> canciones</span>
            <button type="button" class="btn btn-soft" id="btnPlayAllTop">
                <svg><use href="#i-play"></use></svg> Reproducir lista
            </button>
        </div>
    </header>

    <% If Len(q) > 0 Then %>
    <div class="filter-bar">
        <span>Resultados para <strong>&ldquo;<%=H(q)%>&rdquo;</strong></span>
        <a class="link-clear" href="<%=strBase%>home.asp">Limpiar búsqueda</a>
    </div>
    <% End If %>

    <% If rsSongs.EOF Then %>
    <div class="empty-state">
        <svg><use href="#i-music"></use></svg>
        <h3>No hay canciones para mostrar</h3>
        <p>
            <% If Len(q) > 0 Then %>
            No encontramos coincidencias para tu búsqueda. Prueba con otro nombre.
            <% Else %>
            Todavía no existen canciones registradas en el catálogo.
            <% End If %>
        </p>
        <% If IsAdmin() Then %>
        <a class="btn btn-primary" href="<%=strBase%>admin_song_form.asp">
            <svg><use href="#i-plus"></use></svg> Registrar canción
        </a>
        <% End If %>
    </div>
    <% Else %>
    <div class="song-grid" id="songGrid">
        <%
        Do While Not rsSongs.EOF
            idSong  = ToInt(rsSongs("ID_SONG"), 0)
            sName   = CStr(rsSongs("SONG_NAME") & "")
            sPath   = CStr(rsSongs("SONG_PATH") & "")
            sPlays  = ToInt(rsSongs("PLAYS"), 0)
            sAlbum  = TrimStr(rsSongs("ALBUM_NAME"))
            sYear   = ToInt(rsSongs("ALBUM_YEAR"), 0)
            sCover  = SongCoverUrl(idSong, sName, rsSongs("ALBUM_IMAGE").Value)
            sArtists = TrimStr(rsSongs("ARTISTS"))
            sAudio  = SongAudioUrl(idSong, sPath)
            If Len(sArtists) = 0 Then sArtists = "Artista desconocido"
            If Len(sAlbum) = 0 Then sAlbum = "Sencillo"
            sSearch = LCase(sName & " " & sArtists & " " & sAlbum)

            jsonItem = "{""id"":" & idSong & _
                       ",""title"":" & JsStr(sName) & _
                       ",""artist"":" & JsStr(sArtists) & _
                       ",""album"":" & JsStr(sAlbum) & _
                       ",""src"":" & JsStr(sAudio) & _
                       ",""cover"":" & JsStr(sCover) & _
                       ",""plays"":" & sPlays & "}"
            If Len(jsonList) > 0 Then jsonList = jsonList & ","
            jsonList = jsonList & jsonItem
        %>
        <article class="song-card" data-index="<%=idx%>" data-id="<%=idSong%>"
                 data-search="<%=H(sSearch)%>" tabindex="0" role="button"
                 aria-label="Reproducir <%=H(sName)%>">
            <div class="song-cover">
                <img src="<%=sCover%>" alt="Portada de <%=H(sName)%>" loading="lazy">
                <span class="song-cover-shade" aria-hidden="true"></span>
                <button type="button" class="song-play" data-play="<%=idx%>"
                        aria-label="Reproducir <%=H(sName)%>">
                    <svg class="ico-play"><use href="#i-play"></use></svg>
                    <svg class="ico-pause"><use href="#i-pause"></use></svg>
                </button>
                <span class="plays-pill">
                    <svg><use href="#i-play"></use></svg><%=FmtShortNum(sPlays)%>
                </span>
                <span class="song-number"><%=Right("0" & (idx + 1), 2)%></span>
            </div>

            <div class="song-body">
                <h3 title="<%=H(sName)%>"><%=H(sName)%></h3>
                <p class="song-artist" title="<%=H(sArtists)%>"><%=H(sArtists)%></p>
                <p class="song-album">
                    <%=H(sAlbum)%><% If sYear > 0 Then %> <span class="sep">·</span> <%=sYear%><% End If %>
                </p>
            </div>

            <% If IsAdmin() Then %>
            <div class="song-admin">
                <a class="mini-btn" href="<%=strBase%>admin_song_form.asp?id=<%=idSong%>" title="Editar canción">
                    <svg><use href="#i-edit"></use></svg>
                </a>
                <form method="post" action="<%=strBase%>admin_songs.asp" class="inline-form"
                      data-confirm="¿Eliminar la canción &quot;<%=H(sName)%>&quot;? Esta acción no se puede deshacer.">
                    <input type="hidden" name="action" value="delete">
                    <input type="hidden" name="id" value="<%=idSong%>">
                    <button type="submit" class="mini-btn danger" title="Eliminar canción">
                        <svg><use href="#i-trash"></use></svg>
                    </button>
                </form>
            </div>
            <% End If %>
        </article>
        <%
            idx = idx + 1
            rsSongs.MoveNext
        Loop
        %>
    </div>
    <% End If %>
    <%
    ' Se cierra el cursor de canciones antes de abrir la consulta de playlists:
    ' la conexion SQLOLEDB no admite dos resultsets activos al mismo tiempo.
    If rsSongs.State <> 0 Then rsSongs.Close
    Set rsSongs = Nothing
    %>

    <div class="grid-note" id="gridNote" hidden>
        <p>No hay coincidencias con el filtro aplicado.</p>
    </div>
</section>

<section class="section" id="playlists">
    <header class="section-head">
        <div class="section-copy">
            <span class="section-kicker">Playlists</span>
            <h2>Playlists de la comunidad</h2>
            <p>Colecciones creadas por los usuarios, con su portada y cantidad de canciones.</p>
        </div>
        <% If IsAdmin() Then %>
        <div class="section-tools">
            <a class="btn btn-soft" href="<%=strBase%>admin_playlists.asp">
                <svg><use href="#i-edit"></use></svg> Gestionar
            </a>
        </div>
        <% End If %>
    </header>

    <div class="playlist-grid">
        <%
        Set rsPl = DbQuery("SELECT p.ID_PLAYLIST, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, " & _
                           " (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TOTAL " & _
                           "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
                           "WHERE p.TYPE = 0 OR p.ID_USER = ? " & _
                           "ORDER BY p.CREATION_DATE DESC", Array(ToInt(Session("UserID"), 0)))
        plCount = 0
        Do While Not rsPl.EOF
            plCount = plCount + 1
        %>
        <article class="playlist-card">
            <div class="playlist-cover">
                <img src="<%=PlaylistCoverUrl(rsPl("ID_PLAYLIST").Value, rsPl("PHOTO").Value)%>"
                     alt="Portada de <%=H(rsPl("TITLE"))%>" loading="lazy">
                <span class="pl-type <%=IIf(ToInt(rsPl("TYPE"), 0) = 1, "is-private", "is-public")%>">
                    <%=IIf(ToInt(rsPl("TYPE"), 0) = 1, "Privada", "Pública")%>
                </span>
            </div>
            <div class="playlist-body">
                <h3><%=H(rsPl("TITLE"))%></h3>
                <p><%=H(rsPl("USERNAME"))%> · <%=FmtDate(rsPl("CREATION_DATE").Value)%></p>
                <span class="pl-count"><%=FmtNum(rsPl("TOTAL").Value)%> canciones</span>
            </div>
        </article>
        <%
            rsPl.MoveNext
        Loop
        If rsPl.State <> 0 Then rsPl.Close
        Set rsPl = Nothing

        If plCount = 0 Then
        %>
        <div class="empty-state compact">
            <svg><use href="#i-disc"></use></svg>
            <h3>Sin playlists todavía</h3>
            <p>Cuando existan playlists públicas o propias aparecerán en esta sección.</p>
        </div>
        <% End If %>
    </div>
</section>

<%
strPlaylistJson = "[" & jsonList & "]"
%>
<!--#include file="includes/footer.asp"-->
