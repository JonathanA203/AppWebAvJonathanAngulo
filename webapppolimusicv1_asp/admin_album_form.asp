<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: alta / edicion de albumes
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Nuevo álbum"
pageId       = "admin-albums"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim fId, fTitle, fYear, fImage, fErr, isEdit, rsOne, saved, rsTracks, imgPreview, yearTxt

fId = 0
fTitle = ""
fYear = Year(Now())
fImage = ""
fErr = ""
saved = False

If IsPost() Then
    fId    = ToInt(Request.Form("id"), 0)
    fTitle = TrimStr(Request.Form("title_album"))
    fYear  = ToInt(Request.Form("release_year"), 0)
    fImage = TrimStr(Request.Form("album_image"))

    If Len(fTitle) = 0 Then
        fErr = "El título del álbum es obligatorio."
    ElseIf Len(fTitle) > 50 Then
        fErr = "El título no puede superar los 50 caracteres."
    ElseIf fYear < 0 Or fYear > Year(Now()) + 1 Then
        fErr = "El año de lanzamiento no es válido."
    ElseIf Len(fImage) > 255 Then
        fErr = "La ruta de la imagen no puede superar los 255 caracteres."
    End If

    If Len(fErr) = 0 Then
        On Error Resume Next
        If fId = 0 Then
            Call DbExec("INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?)", _
                        Array(fTitle, fYear, fImage))
        Else
            Call DbExec("UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", _
                        Array(fTitle, fYear, fImage, fId))
        End If
        If Err.Number <> 0 Then
            Err.Clear
            fErr = "No se pudo guardar el álbum."
        Else
            saved = True
        End If
        On Error GoTo 0
    End If

    If saved Then
        Call SetFlash("Álbum guardado correctamente.", "success")
        Response.Redirect "admin_albums.asp"
        Response.End
    End If
Else
    fId = ToInt(Request.QueryString("id"), 0)
    If fId > 0 Then
        Set rsOne = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(fId))
        If Not rsOne.EOF Then
            fTitle = CStr(rsOne("TITLE_ALBUM") & "")
            fYear  = ToInt(rsOne("RELEASE_YEAR"), 0)
            fImage = CStr(Nz(rsOne("ALBUM_IMAGE_PATH"), ""))
        Else
            fId = 0
        End If
        If rsOne.State <> 0 Then rsOne.Close
        Set rsOne = Nothing
    End If
End If

isEdit = (fId > 0)
If isEdit Then pageTitle = "Editar álbum"
imgPreview = AlbumCoverUrl(fId, fImage)
yearTxt = CStr(fYear)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Álbumes</span>
        <h2><%=IIf(isEdit, "Editar álbum", "Registrar nuevo álbum")%></h2>
        <p>Las canciones se agregan a un álbum desde el formulario de cada canción.</p>
    </div>
    <div class="page-head-actions">
        <a class="btn btn-ghost" href="<%=strBase%>admin_albums.asp">
            <svg><use href="#i-close"></use></svg> Volver al listado
        </a>
    </div>
</section>

<% If Len(fErr) > 0 Then %>
<div class="flash flash-error" role="alert">
    <svg><use href="#i-close"></use></svg><span><%=H(fErr)%></span>
</div>
<% End If %>

<div class="form-layout">
    <form class="form-card" method="post" action="<%=strBase%>admin_album_form.asp" autocomplete="off">
        <input type="hidden" name="id" value="<%=fId%>">

        <div class="form-section">
            <h3><svg><use href="#i-disc"></use></svg> Datos del álbum</h3>

            <label class="field">
                <span>Título <em>*</em></span>
                <input type="text" name="title_album" value="<%=H(fTitle)%>" maxlength="50" required
                       placeholder="Ej. Rock Album">
            </label>

            <div class="form-row">
                <label class="field">
                    <span>Año de lanzamiento</span>
                    <input type="number" name="release_year" value="<%=H(yearTxt)%>" min="1900" max="<%=Year(Now()) + 1%>" step="1">
                </label>
                <label class="field">
                    <span>Ruta de la portada</span>
                    <input type="text" name="album_image" value="<%=H(fImage)%>" maxlength="255"
                           placeholder="~/albumImages/album1.png">
                </label>
            </div>
            <small class="hint">Si la portada no existe en los recursos se usará una imagen por defecto.</small>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-primary">
                <svg><use href="#i-check"></use></svg> <%=IIf(isEdit, "Guardar cambios", "Registrar álbum")%>
            </button>
            <a class="btn btn-ghost" href="<%=strBase%>admin_albums.asp">Cancelar</a>
        </div>
    </form>

    <aside class="form-aside">
        <div class="aside-card">
            <h4>Portada</h4>
            <img class="aside-preview" src="<%=imgPreview%>" alt="Vista previa de la portada">
        </div>

        <% If isEdit Then %>
        <div class="aside-card">
            <h4>Canciones del álbum</h4>
            <ul class="tips plain">
                <%
                Set rsTracks = DbQuery("SELECT s.ID_SONG, s.SONG_NAME FROM TBL_TRACKLIST t " & _
                                       "INNER JOIN TBL_SONG s ON s.ID_SONG = t.ID_SONG " & _
                                       "WHERE t.ID_ALBUM = ? ORDER BY s.SONG_NAME", Array(fId))
                i = 0
                Do While Not rsTracks.EOF
                    i = i + 1
                %>
                <li><a href="<%=strBase%>admin_song_form.asp?id=<%=rsTracks("ID_SONG")%>"><%=H(rsTracks("SONG_NAME"))%></a></li>
                <%
                    rsTracks.MoveNext
                Loop
                If rsTracks.State <> 0 Then rsTracks.Close
                Set rsTracks = Nothing
                If i = 0 Then
                %>
                <li class="muted">Este álbum todavía no tiene canciones.</li>
                <% End If %>
            </ul>
        </div>

        <div class="aside-card">
            <h4>Zona sensible</h4>
            <p class="muted small">Solo se puede eliminar un álbum sin canciones asociadas.</p>
            <form method="post" action="<%=strBase%>admin_albums.asp" class="inline-form"
                  data-confirm="¿Eliminar el álbum &quot;<%=H(fTitle)%>&quot;?">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="id" value="<%=fId%>">
                <button type="submit" class="btn btn-danger btn-block">
                    <svg><use href="#i-trash"></use></svg> Eliminar álbum
                </button>
            </form>
        </div>
        <% End If %>
    </aside>
</div>

<!--#include file="includes/footer.asp"-->
