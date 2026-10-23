<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: administracion de albumes
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Álbumes"
pageId       = "admin-albums"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim aAction, aId, aQ, aSql, aParams, rsList, i, aTotal

aAction = TrimStr(Request.Form("action"))

If IsPost() And aAction = "delete" Then
    aId = ToInt(Request.Form("id"), 0)
    If aId > 0 Then
        If DbTry("DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(aId)) Then
            Call SetFlash("Álbum eliminado correctamente.", "success")
        Else
            Call SetFlash("No se pudo eliminar el álbum: tiene canciones asociadas.", "error")
        End If
    End If
    Response.Redirect "admin_albums.asp"
    Response.End
End If

aQ = TrimStr(Request.QueryString("q"))
aTotal = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", Array()), 0)

aSql = "SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
       " (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS TOTAL " & _
       "FROM TBL_ALBUM al "
If Len(aQ) > 0 Then
    aSql = aSql & "WHERE al.TITLE_ALBUM LIKE ? "
    aParams = Array("%" & aQ & "%")
Else
    aParams = Array()
End If
aSql = aSql & "ORDER BY al.RELEASE_YEAR DESC, al.TITLE_ALBUM"

Set rsList = DbQuery(aSql, aParams)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Álbumes</span>
        <h2>Gestión de álbumes</h2>
        <p>Agrupa canciones en álbumes. La relación se almacena en <code>TBL_TRACKLIST</code>.</p>
    </div>
    <div class="page-head-actions">
        <span class="count-pill"><%=FmtNum(aTotal)%> registros</span>
        <a class="btn btn-primary" href="<%=strBase%>admin_album_form.asp">
            <svg><use href="#i-plus"></use></svg> Nuevo álbum
        </a>
    </div>
</section>

<div class="table-card">
    <div class="table-toolbar">
        <form class="table-search" method="get" action="<%=strBase%>admin_albums.asp">
            <svg><use href="#i-search"></use></svg>
            <input type="search" name="q" value="<%=H(aQ)%>" placeholder="Buscar por título del álbum…">
            <button type="submit" class="btn btn-soft sm">Filtrar</button>
            <% If Len(aQ) > 0 Then %>
            <a class="link-clear" href="<%=strBase%>admin_albums.asp">Limpiar</a>
            <% End If %>
        </form>
    </div>

    <div class="table-wrap">
        <table class="data-table">
            <thead>
                <tr>
                    <th class="col-cover">Portada</th>
                    <th>Álbum</th>
                    <th class="num">Año</th>
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
                        <img class="thumb" src="<%=AlbumCoverUrl(rsList("ID_ALBUM").Value, rsList("ALBUM_IMAGE_PATH").Value)%>"
                             alt="Portada de <%=H(rsList("TITLE_ALBUM"))%>">
                    </td>
                    <td>
                        <strong class="cell-title"><%=H(rsList("TITLE_ALBUM"))%></strong>
                        <span class="cell-sub mono"><%=H(Nz(rsList("ALBUM_IMAGE_PATH"), "—"))%></span>
                    </td>
                    <td class="num"><%=IIf(IsNull(rsList("RELEASE_YEAR")), "—", rsList("RELEASE_YEAR"))%></td>
                    <td class="num"><span class="plays-cell"><%=FmtNum(rsList("TOTAL").Value)%></span></td>
                    <td class="col-actions">
                        <div class="row-actions">
                            <a class="mini-btn" href="<%=strBase%>admin_album_form.asp?id=<%=rsList("ID_ALBUM")%>" title="Editar">
                                <svg><use href="#i-edit"></use></svg>
                            </a>
                            <form method="post" action="<%=strBase%>admin_albums.asp" class="inline-form"
                                  data-confirm="¿Eliminar el álbum &quot;<%=H(rsList("TITLE_ALBUM"))%>&quot;?">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=rsList("ID_ALBUM")%>">
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
                        <svg><use href="#i-disc"></use></svg>
                        <span>No se encontraron álbumes.</span>
                    </div>
                </td></tr>
            <% End If %>
            </tbody>
        </table>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
