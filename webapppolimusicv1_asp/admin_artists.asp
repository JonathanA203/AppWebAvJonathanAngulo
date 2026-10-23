<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: administracion de artistas
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Artistas"
pageId       = "admin-artists"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim aAction, aId, aQ, aSql, aParams, rsList, i, aTotal

aAction = TrimStr(Request.Form("action"))

If IsPost() And aAction = "delete" Then
    aId = ToInt(Request.Form("id"), 0)
    If aId > 0 Then
        If DbTry("DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(aId)) Then
            Call SetFlash("Artista eliminado correctamente.", "success")
        Else
            Call SetFlash("No se pudo eliminar el artista: tiene canciones asociadas.", "error")
        End If
    End If
    Response.Redirect "admin_artists.asp"
    Response.End
End If

aQ = TrimStr(Request.QueryString("q"))
aTotal = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", Array()), 0)

aSql = "SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
       " (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS TOTAL " & _
       "FROM TBL_ARTIST a "
If Len(aQ) > 0 Then
    aSql = aSql & "WHERE (a.ARTIST_NAME LIKE ? OR ISNULL(a.ARTIST_LASTNAME,'') LIKE ?) "
    aParams = Array("%" & aQ & "%", "%" & aQ & "%")
Else
    aParams = Array()
End If
aSql = aSql & "ORDER BY a.ARTIST_NAME"

Set rsList = DbQuery(aSql, aParams)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Artistas</span>
        <h2>Gestión de artistas</h2>
        <p>Administra los intérpretes del catálogo. Cada artista se relaciona con sus canciones mediante <code>TBL_PLAYER</code>.</p>
    </div>
    <div class="page-head-actions">
        <span class="count-pill"><%=FmtNum(aTotal)%> registros</span>
        <a class="btn btn-primary" href="<%=strBase%>admin_artist_form.asp">
            <svg><use href="#i-plus"></use></svg> Nuevo artista
        </a>
    </div>
</section>

<div class="table-card">
    <div class="table-toolbar">
        <form class="table-search" method="get" action="<%=strBase%>admin_artists.asp">
            <svg><use href="#i-search"></use></svg>
            <input type="search" name="q" value="<%=H(aQ)%>" placeholder="Buscar por nombre o apellido…">
            <button type="submit" class="btn btn-soft sm">Filtrar</button>
            <% If Len(aQ) > 0 Then %>
            <a class="link-clear" href="<%=strBase%>admin_artists.asp">Limpiar</a>
            <% End If %>
        </form>
    </div>

    <div class="table-wrap">
        <table class="data-table">
            <thead>
                <tr>
                    <th class="col-cover">Foto</th>
                    <th>Artista</th>
                    <th>Imagen registrada</th>
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
                        <img class="thumb round" src="<%=ArtistPhotoUrl(rsList("ARTIST_IMAGE").Value, rsList("ARTIST_NAME").Value)%>"
                             alt="<%=H(rsList("ARTIST_NAME"))%>">
                    </td>
                    <td>
                        <strong class="cell-title"><%=H(rsList("ARTIST_NAME") & " " & Nz(rsList("ARTIST_LASTNAME"), ""))%></strong>
                        <span class="cell-sub">ID <%=rsList("ID_ARTIST")%></span>
                    </td>
                    <td class="muted mono small"><%=H(Nz(rsList("ARTIST_IMAGE"), "—"))%></td>
                    <td class="num"><span class="plays-cell"><%=FmtNum(rsList("TOTAL").Value)%></span></td>
                    <td class="col-actions">
                        <div class="row-actions">
                            <a class="mini-btn" href="<%=strBase%>admin_artist_form.asp?id=<%=rsList("ID_ARTIST")%>" title="Editar">
                                <svg><use href="#i-edit"></use></svg>
                            </a>
                            <form method="post" action="<%=strBase%>admin_artists.asp" class="inline-form"
                                  data-confirm="¿Eliminar al artista &quot;<%=H(rsList("ARTIST_NAME"))%>&quot;?">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=rsList("ID_ARTIST")%>">
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
                <tr><td colspan="5">
                    <div class="empty-inline">
                        <svg><use href="#i-artist"></use></svg>
                        <span>No se encontraron artistas.</span>
                    </div>
                </td></tr>
            <% End If %>
            </tbody>
        </table>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
