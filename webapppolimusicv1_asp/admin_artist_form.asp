<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: alta / edicion de artistas
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Nuevo artista"
pageId       = "admin-artists"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim fId, fName, fLast, fImage, fErr, isEdit, rsOne, saved, imgPreview

fId = 0
fName = ""
fLast = ""
fImage = ""
fErr = ""
saved = False

If IsPost() Then
    fId    = ToInt(Request.Form("id"), 0)
    fName  = TrimStr(Request.Form("artist_name"))
    fLast  = TrimStr(Request.Form("artist_lastname"))
    fImage = TrimStr(Request.Form("artist_image"))

    If Len(fName) = 0 Then
        fErr = "El nombre del artista es obligatorio."
    ElseIf Len(fName) > 100 Then
        fErr = "El nombre no puede superar los 100 caracteres."
    ElseIf Len(fLast) > 100 Then
        fErr = "El apellido no puede superar los 100 caracteres."
    ElseIf Len(fImage) > 255 Then
        fErr = "La ruta de la imagen no puede superar los 255 caracteres."
    End If

    If Len(fErr) = 0 Then
        On Error Resume Next
        If fId = 0 Then
            Call DbExec("INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)", _
                        Array(fName, fLast, fImage))
        Else
            Call DbExec("UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", _
                        Array(fName, fLast, fImage, fId))
        End If
        If Err.Number <> 0 Then
            Err.Clear
            fErr = "No se pudo guardar el artista."
        Else
            saved = True
        End If
        On Error GoTo 0
    End If

    If saved Then
        Call SetFlash("Artista guardado correctamente.", "success")
        Response.Redirect "admin_artists.asp"
        Response.End
    End If
Else
    fId = ToInt(Request.QueryString("id"), 0)
    If fId > 0 Then
        Set rsOne = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(fId))
        If Not rsOne.EOF Then
            fName  = CStr(rsOne("ARTIST_NAME") & "")
            fLast  = CStr(Nz(rsOne("ARTIST_LASTNAME"), ""))
            fImage = CStr(Nz(rsOne("ARTIST_IMAGE"), ""))
        Else
            fId = 0
        End If
        If rsOne.State <> 0 Then rsOne.Close
        Set rsOne = Nothing
    End If
End If

isEdit = (fId > 0)
If isEdit Then pageTitle = "Editar artista"
imgPreview = ArtistPhotoUrl(fImage, fName)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Artistas</span>
        <h2><%=IIf(isEdit, "Editar artista", "Registrar nuevo artista")%></h2>
        <p>Los artistas se asocian a las canciones desde el formulario de cada canción.</p>
    </div>
    <div class="page-head-actions">
        <a class="btn btn-ghost" href="<%=strBase%>admin_artists.asp">
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
    <form class="form-card" method="post" action="<%=strBase%>admin_artist_form.asp" autocomplete="off">
        <input type="hidden" name="id" value="<%=fId%>">

        <div class="form-section">
            <h3><svg><use href="#i-artist"></use></svg> Datos del artista</h3>

            <div class="form-row">
                <label class="field">
                    <span>Nombre <em>*</em></span>
                    <input type="text" name="artist_name" value="<%=H(fName)%>" maxlength="100" required
                           placeholder="Ej. Ana">
                </label>
                <label class="field">
                    <span>Apellido</span>
                    <input type="text" name="artist_lastname" value="<%=H(fLast)%>" maxlength="100"
                           placeholder="Ej. Gabriel">
                </label>
            </div>

            <label class="field">
                <span>Ruta de la imagen</span>
                <input type="text" name="artist_image" value="<%=H(fImage)%>" maxlength="255"
                       placeholder="~/ArtistImages/artist1.jpg">
                <small class="hint">
                    Si la imagen no existe en los recursos, se usará una portada por defecto.
                </small>
            </label>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-primary">
                <svg><use href="#i-check"></use></svg> <%=IIf(isEdit, "Guardar cambios", "Registrar artista")%>
            </button>
            <a class="btn btn-ghost" href="<%=strBase%>admin_artists.asp">Cancelar</a>
        </div>
    </form>

    <aside class="form-aside">
        <div class="aside-card">
            <h4>Vista previa</h4>
            <img class="aside-preview round" src="<%=imgPreview%>" alt="Vista previa del artista">
            <p class="muted small">La imagen se resuelve automáticamente desde la carpeta de recursos.</p>
        </div>
        <div class="aside-card">
            <h4>Relaciones</h4>
            <p class="muted small">
                Un artista puede interpretar varias canciones. La relación se guarda en
                <code>TBL_PLAYER (ID_SONG, ID_ARTIST)</code>.
            </p>
        </div>
        <% If isEdit Then %>
        <div class="aside-card">
            <h4>Zona sensible</h4>
            <p class="muted small">Solo se puede eliminar un artista sin canciones asociadas.</p>
            <form method="post" action="<%=strBase%>admin_artists.asp" class="inline-form"
                  data-confirm="¿Eliminar al artista &quot;<%=H(fName & " " & fLast)%>&quot;?">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="id" value="<%=fId%>">
                <button type="submit" class="btn btn-danger btn-block">
                    <svg><use href="#i-trash"></use></svg> Eliminar artista
                </button>
            </form>
        </div>
        <% End If %>
    </aside>
</div>

<!--#include file="includes/footer.asp"-->
