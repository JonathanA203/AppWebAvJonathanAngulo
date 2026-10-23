<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: alta / edicion de playlists
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Nueva playlist"
pageId       = "admin-playlists"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim fId, fTitle, fUser, fDate, fType, fPhoto, fErr, isEdit
Dim rsOne, rsUsers, rsTypes, rsSongs, saved, imgPreview

fId = 0
fTitle = ""
fUser = ToInt(Session("UserID"), 0)
fDate = IsoDate(Now())
fType = 0
fPhoto = ""
fErr = ""
saved = False

If IsPost() Then
    fId    = ToInt(Request.Form("id"), 0)
    fTitle = TrimStr(Request.Form("title"))
    fUser  = ToInt(Request.Form("id_user"), 0)
    fDate  = TrimStr(Request.Form("creation_date"))
    fType  = ToInt(Request.Form("type"), 0)
    fPhoto = TrimStr(Request.Form("photo"))

    If Len(fTitle) = 0 Then
        fErr = "El título de la playlist es obligatorio."
    ElseIf Len(fTitle) > 50 Then
        fErr = "El título no puede superar los 50 caracteres."
    ElseIf fUser <= 0 Then
        fErr = "Debes seleccionar un propietario."
    ElseIf Not IsDate(fDate) Then
        fErr = "La fecha de creación no es válida."
    ElseIf Len(fPhoto) > 255 Then
        fErr = "La ruta de la imagen no puede superar los 255 caracteres."
    End If

    If Len(fErr) = 0 Then
        On Error Resume Next
        If fId = 0 Then
            Call DbExec("INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (?, ?, ?, ?, ?)", _
                        Array(fUser, fTitle, CDate(fDate), fType, fPhoto))
        Else
            Call DbExec("UPDATE TBL_PLAYLIST SET ID_USER = ?, TITLE = ?, CREATION_DATE = ?, TYPE = ?, PHOTO = ? " & _
                        "WHERE ID_PLAYLIST = ?", _
                        Array(fUser, fTitle, CDate(fDate), fType, fPhoto, fId))
        End If
        If Err.Number <> 0 Then
            Err.Clear
            fErr = "No se pudo guardar la playlist."
        Else
            saved = True
        End If
        On Error GoTo 0
    End If

    If saved Then
        Call SetFlash("Playlist guardada correctamente.", "success")
        Response.Redirect "admin_playlists.asp"
        Response.End
    End If
Else
    fId = ToInt(Request.QueryString("id"), 0)
    If fId > 0 Then
        Set rsOne = DbQuery("SELECT ID_PLAYLIST, ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(fId))
        If Not rsOne.EOF Then
            fTitle = CStr(rsOne("TITLE") & "")
            fUser  = ToInt(rsOne("ID_USER"), 0)
            fDate  = IsoDate(rsOne("CREATION_DATE").Value)
            fType  = ToInt(rsOne("TYPE"), 0)
            fPhoto = CStr(Nz(rsOne("PHOTO"), ""))
        Else
            fId = 0
        End If
        If rsOne.State <> 0 Then rsOne.Close
        Set rsOne = Nothing
    End If
End If

isEdit = (fId > 0)
If isEdit Then pageTitle = "Editar playlist"
imgPreview = PlaylistCoverUrl(fId, fPhoto)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Playlists</span>
        <h2><%=IIf(isEdit, "Editar playlist", "Registrar nueva playlist")%></h2>
        <p>Define el propietario, la visibilidad y la portada de la colección.</p>
    </div>
    <div class="page-head-actions">
        <a class="btn btn-ghost" href="<%=strBase%>admin_playlists.asp">
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
    <form class="form-card" method="post" action="<%=strBase%>admin_playlist_form.asp" autocomplete="off">
        <input type="hidden" name="id" value="<%=fId%>">

        <div class="form-section">
            <h3><svg><use href="#i-list"></use></svg> Datos de la playlist</h3>

            <label class="field">
                <span>Título <em>*</em></span>
                <input type="text" name="title" value="<%=H(fTitle)%>" maxlength="50" required
                       placeholder="Ej. Playlist 1">
            </label>

            <div class="form-row">
                <label class="field">
                    <span>Propietario <em>*</em></span>
                    <select name="id_user" required>
                        <option value="0">— Selecciona un usuario —</option>
                        <%
                        Set rsUsers = DbQuery("SELECT ID_USER, USERNAME FROM TBL_USER ORDER BY USERNAME", Array())
                        Do While Not rsUsers.EOF
                        %>
                        <option value="<%=rsUsers("ID_USER")%>"
                            <%=IIf(ToInt(rsUsers("ID_USER"), 0) = fUser, "selected", "")%>>
                            <%=H(rsUsers("USERNAME"))%>
                        </option>
                        <%
                            rsUsers.MoveNext
                        Loop
                        If rsUsers.State <> 0 Then rsUsers.Close
                        Set rsUsers = Nothing
                        %>
                    </select>
                </label>

                <label class="field">
                    <span>Fecha de creación <em>*</em></span>
                    <input type="date" name="creation_date" value="<%=H(fDate)%>" required>
                </label>
            </div>

            <div class="form-row">
                <label class="field">
                    <span>Tipo de playlist</span>
                    <select name="type">
                        <%
                        Set rsTypes = DbQuery("SELECT VALUE, COD_CATALOG FROM TBL_CATALOG " & _
                                              "WHERE COD_CATALOG_PARENT = '2' ORDER BY COD_CATALOG", Array())
                        Do While Not rsTypes.EOF
                        %>
                        <option value="<%=ToInt(rsTypes("COD_CATALOG"), 0)%>"
                            <%=IIf(ToInt(rsTypes("COD_CATALOG"), 0) = fType, "selected", "")%>>
                            <%=H(rsTypes("VALUE"))%>
                        </option>
                        <%
                            rsTypes.MoveNext
                        Loop
                        If rsTypes.State <> 0 Then rsTypes.Close
                        Set rsTypes = Nothing
                        %>
                    </select>
                </label>

                <label class="field">
                    <span>Ruta de la portada</span>
                    <input type="text" name="photo" value="<%=H(fPhoto)%>" maxlength="255"
                           placeholder="~/playlistImages/playlist1.png">
                </label>
            </div>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-primary">
                <svg><use href="#i-check"></use></svg> <%=IIf(isEdit, "Guardar cambios", "Registrar playlist")%>
            </button>
            <a class="btn btn-ghost" href="<%=strBase%>admin_playlists.asp">Cancelar</a>
        </div>
    </form>

    <aside class="form-aside">
        <div class="aside-card">
            <h4>Portada</h4>
            <img class="aside-preview" src="<%=imgPreview%>" alt="Vista previa de la portada">
        </div>

        <% If isEdit Then %>
        <div class="aside-card">
            <h4>Canciones incluidas</h4>
            <ul class="tips plain">
                <%
                Set rsSongs = DbQuery("SELECT s.SONG_NAME, s.ID_SONG FROM TBL_PLAYLIST_SONG ps " & _
                                      "INNER JOIN TBL_SONG s ON s.ID_SONG = ps.ID_SONG " & _
                                      "WHERE ps.ID_PLAYLIST = ? ORDER BY s.SONG_NAME", Array(fId))
                Dim nSongs
                nSongs = 0
                Do While Not rsSongs.EOF
                    nSongs = nSongs + 1
                %>
                <li><a href="<%=strBase%>admin_song_form.asp?id=<%=rsSongs("ID_SONG")%>"><%=H(rsSongs("SONG_NAME"))%></a></li>
                <%
                    rsSongs.MoveNext
                Loop
                If rsSongs.State <> 0 Then rsSongs.Close
                Set rsSongs = Nothing
                If nSongs = 0 Then
                %>
                <li class="muted">Esta playlist todavía no tiene canciones.</li>
                <% End If %>
            </ul>
        </div>

        <div class="aside-card">
            <h4>Zona sensible</h4>
            <p class="muted small">Solo se puede eliminar una playlist sin canciones asociadas.</p>
            <form method="post" action="<%=strBase%>admin_playlists.asp" class="inline-form"
                  data-confirm="¿Eliminar la playlist &quot;<%=H(fTitle)%>&quot;?">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="id" value="<%=fId%>">
                <button type="submit" class="btn btn-danger btn-block">
                    <svg><use href="#i-trash"></use></svg> Eliminar playlist
                </button>
            </form>
        </div>
        <% End If %>
    </aside>
</div>

<!--#include file="includes/footer.asp"-->
