<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: administracion de canciones (listado + eliminar)
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Canciones"
pageId       = "admin-songs"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim aAction, aId, aQ, aSql, rsList, i, aTotal, aParams, aCover

aAction = TrimStr(Request.Form("action"))

If IsPost() And aAction = "delete" Then
    aId = ToInt(Request.Form("id"), 0)
    If aId > 0 Then
        If DbTry("DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(aId)) Then
            Call SetFlash("Canción eliminada correctamente.", "success")
        Else
            Call SetFlash("No se pudo eliminar la canción: tiene registros asociados en artistas, álbumes o playlists.", "error")
        End If
    End If
    Response.Redirect "admin_songs.asp"
    Response.End
End If

aQ = TrimStr(Request.QueryString("q"))
aTotal = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG", Array()), 0)

aSql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
       " ISNULL((SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al " & _
       "   ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG), '') AS ALBUM_NAME, " & _
       " ISNULL((SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al " & _
       "   ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG), '') AS ALBUM_IMAGE, " & _
       " ISNULL((SELECT STUFF((SELECT DISTINCT ', ' + a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,'') " & _
       "   FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
       "   WHERE p.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, '')), '') AS ARTISTS " & _
       "FROM TBL_SONG s "

If Len(aQ) > 0 Then
    aSql = aSql & "WHERE s.SONG_NAME LIKE ? "
    aParams = Array("%" & aQ & "%")
Else
    aParams = Array()
End If
aSql = aSql & "ORDER BY s.SONG_NAME"

Set rsList = DbQuery(aSql, aParams)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Canciones</span>
        <h2>Gestión de canciones</h2>
        <p>
            Crea, edita y elimina las canciones del catálogo. La ruta corresponde al archivo de audio
            dentro de <code>SongFilesAndImages_resources/SongFiles</code>.
        </p>
    </div>
    <div class="page-head-actions">
        <span class="count-pill"><%=FmtNum(aTotal)%> registros</span>
        <a class="btn btn-primary" href="<%=strBase%>admin_song_form.asp">
            <svg><use href="#i-plus"></use></svg> Nueva canción
        </a>
    </div>
</section>

<div class="table-card">
    <div class="table-toolbar">
        <form class="table-search" method="get" action="<%=strBase%>admin_songs.asp">
            <svg><use href="#i-search"></use></svg>
            <input type="search" name="q" value="<%=H(aQ)%>" placeholder="Buscar por nombre de canción…">
            <button type="submit" class="btn btn-soft sm">Filtrar</button>
            <% If Len(aQ) > 0 Then %>
            <a class="link-clear" href="<%=strBase%>admin_songs.asp">Limpiar</a>
            <% End If %>
        </form>
        <span class="muted small">Catálogo con <%=FmtNum(aTotal)%> canciones registradas</span>
    </div>

    <div class="table-wrap">
        <table class="data-table">
            <thead>
                <tr>
                    <th class="col-cover">Portada</th>
                    <th>Canción</th>
                    <th>Artistas</th>
                    <th>Álbum</th>
                    <th class="num">Plays</th>
                    <th class="col-actions">Acciones</th>
                </tr>
            </thead>
            <tbody>
            <%
            i = 0
            Do While Not rsList.EOF
                i = i + 1
                aCover = SongCoverUrl(rsList("ID_SONG").Value, rsList("SONG_NAME").Value, rsList("ALBUM_IMAGE").Value)
            %>
                <tr>
                    <td class="col-cover">
                        <img class="thumb" src="<%=aCover%>" alt="Portada de <%=H(rsList("SONG_NAME"))%>">
                    </td>
                    <td>
                        <strong class="cell-title"><%=H(rsList("SONG_NAME"))%></strong>
                        <span class="cell-sub mono"><%=H(rsList("SONG_PATH"))%></span>
                    </td>
                    <td class="muted"><%=H(IIf(Len(TrimStr(rsList("ARTISTS"))) = 0, "Sin artista", rsList("ARTISTS")))%></td>
                    <td class="muted"><%=H(IIf(Len(TrimStr(rsList("ALBUM_NAME"))) = 0, "Sencillo", rsList("ALBUM_NAME")))%></td>
                    <td class="num"><span class="plays-cell"><%=FmtNum(rsList("PLAYS"))%></span></td>
                    <td class="col-actions">
                        <div class="row-actions">
                            <a class="mini-btn" href="<%=strBase%>admin_song_form.asp?id=<%=rsList("ID_SONG")%>"
                               title="Editar"><svg><use href="#i-edit"></use></svg></a>
                            <form method="post" action="<%=strBase%>admin_songs.asp" class="inline-form"
                                  data-confirm="¿Eliminar la canción &quot;<%=H(rsList("SONG_NAME"))%>&quot;?">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=rsList("ID_SONG")%>">
                                <button type="submit" class="mini-btn danger" title="Eliminar">
                                    <svg><use href="#i-trash"></use></svg>
                                </button>
                            </form>
                        </div>
                    </td>
                </tr>
            <%
                rsList.MoveNext
            Loop
            If rsList.State <> 0 Then rsList.Close
            Set rsList = Nothing
            If i = 0 Then
            %>
                <tr>
                    <td colspan="6">
                        <div class="empty-inline">
                            <svg><use href="#i-music"></use></svg>
                            <span>No se encontraron canciones<%=IIf(Len(aQ) > 0, " con ese filtro", "")%>.</span>
                        </div>
                    </td>
                </tr>
            <% End If %>
            </tbody>
        </table>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
