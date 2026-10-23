<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: administracion de playlists
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Playlists"
pageId       = "admin-playlists"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim aAction, aId, aQ, aSql, aParams, rsList, i, aTotal

aAction = TrimStr(Request.Form("action"))

If IsPost() And aAction = "delete" Then
    aId = ToInt(Request.Form("id"), 0)
    If aId > 0 Then
        If DbTry("DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(aId)) Then
            Call SetFlash("Playlist eliminada correctamente.", "success")
        Else
            Call SetFlash("No se pudo eliminar la playlist: tiene canciones asociadas.", "error")
        End If
    End If
    Response.Redirect "admin_playlists.asp"
    Response.End
End If

aQ = TrimStr(Request.QueryString("q"))
aTotal = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST", Array()), 0)

aSql = "SELECT p.ID_PLAYLIST, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, " & _
       " (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TOTAL " & _
       "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER "
If Len(aQ) > 0 Then
    aSql = aSql & "WHERE (p.TITLE LIKE ? OR u.USERNAME LIKE ?) "
    aParams = Array("%" & aQ & "%", "%" & aQ & "%")
Else
    aParams = Array()
End If
aSql = aSql & "ORDER BY p.CREATION_DATE DESC"

Set rsList = DbQuery(aSql, aParams)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Playlists</span>
        <h2>Gestión de playlists</h2>
        <p>Cada playlist pertenece a un usuario y puede ser pública o privada según <code>TBL_CATALOG</code>.</p>
    </div>
    <div class="page-head-actions">
        <span class="count-pill"><%=FmtNum(aTotal)%> registros</span>
        <a class="btn btn-primary" href="<%=strBase%>admin_playlist_form.asp">
            <svg><use href="#i-plus"></use></svg> Nueva playlist
        </a>
    </div>
</section>

<div class="table-card">
    <div class="table-toolbar">
        <form class="table-search" method="get" action="<%=strBase%>admin_playlists.asp">
            <svg><use href="#i-search"></use></svg>
            <input type="search" name="q" value="<%=H(aQ)%>" placeholder="Buscar por título o usuario…">
            <button type="submit" class="btn btn-soft sm">Filtrar</button>
            <% If Len(aQ) > 0 Then %>
            <a class="link-clear" href="<%=strBase%>admin_playlists.asp">Limpiar</a>
            <% End If %>
        </form>
    </div>

    <div class="table-wrap">
        <table class="data-table">
            <thead>
                <tr>
                    <th class="col-cover">Portada</th>
                    <th>Playlist</th>
                    <th>Propietario</th>
                    <th>Creada</th>
                    <th>Tipo</th>
                    <th class="num">Canciones</th>
                    <th class="col-actions">Acciones</th>
                </tr>
            </thead>
            <tbody>
            <%
            i = 0
            Do While Not rsList.EOF
                i = i + 1
            %>
                <tr>
                    <td class="col-cover">
                        <img class="thumb" src="<%=PlaylistCoverUrl(rsList("ID_PLAYLIST").Value, rsList("PHOTO").Value)%>"
                             alt="Portada de <%=H(rsList("TITLE"))%>">
                    </td>
                    <td>
                        <strong class="cell-title"><%=H(rsList("TITLE"))%></strong>
                        <span class="cell-sub">ID <%=rsList("ID_PLAYLIST")%></span>
                    </td>
                    <td class="muted"><%=H(rsList("USERNAME"))%></td>
                    <td class="muted"><%=FmtDate(rsList("CREATION_DATE").Value)%></td>
                    <td>
                        <span class="pl-type <%=IIf(ToInt(rsList("TYPE"), 0) = 1, "is-private", "is-public")%>">
                            <%=IIf(ToInt(rsList("TYPE"), 0) = 1, "Privada", "Pública")%>
                        </span>
                    </td>
                    <td class="num"><span class="plays-cell"><%=FmtNum(rsList("TOTAL").Value)%></span></td>
                    <td class="col-actions">
                        <div class="row-actions">
                            <a class="mini-btn" href="<%=strBase%>admin_playlist_form.asp?id=<%=rsList("ID_PLAYLIST")%>" title="Editar">
                                <svg><use href="#i-edit"></use></svg>
                            </a>
                            <form method="post" action="<%=strBase%>admin_playlists.asp" class="inline-form"
                                  data-confirm="¿Eliminar la playlist &quot;<%=H(rsList("TITLE"))%>&quot;?">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=rsList("ID_PLAYLIST")%>">
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
                <tr><td colspan="7">
                    <div class="empty-inline">
                        <svg><use href="#i-list"></use></svg>
                        <span>No se encontraron playlists.</span>
                    </div>
                </td></tr>
            <% End If %>
            </tbody>
        </table>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
