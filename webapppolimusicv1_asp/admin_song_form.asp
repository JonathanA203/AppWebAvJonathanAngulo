<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: alta / edicion de canciones
'  Tambien administra las relaciones TBL_PLAYER (artistas) y
'  TBL_TRACKLIST (album) de la cancion.
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Nueva canción"
pageId       = "admin-songs"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim fId, fName, fPath, fPlays, fAlbum, fArtistsCsv, fErr, isEdit
Dim rsOne, rsRel, rsArtists, rsAlbums, rsFiles, fso, fld, f, optPath
Dim saved, previewUrl, previewOk, i, parts, pid, tmp

fId = 0
fName = ""
fPath = ""
fPlays = 0
fAlbum = 0
fArtistsCsv = ""
fErr = ""
saved = False

Function CsvHas(csv, val)
    Dim arr, k
    CsvHas = False
    If Len(Trim(csv)) = 0 Then Exit Function
    arr = Split(csv, ",")
    For k = 0 To UBound(arr)
        If Trim(arr(k)) = CStr(val) Then CsvHas = True
    Next
End Function

If IsPost() Then
    fId         = ToInt(Request.Form("id"), 0)
    fName       = TrimStr(Request.Form("song_name"))
    fPath       = TrimStr(Request.Form("song_path"))
    fPlays      = ToInt(Request.Form("plays"), 0)
    fAlbum      = ToInt(Request.Form("album"), 0)
    fArtistsCsv = TrimStr(Request.Form("artists"))

    If Len(fName) = 0 Then
        fErr = "El nombre de la canción es obligatorio."
    ElseIf Len(fName) > 50 Then
        fErr = "El nombre no puede superar los 50 caracteres."
    ElseIf Len(fPath) = 0 Then
        fErr = "La ruta del archivo de audio es obligatoria."
    ElseIf Len(fPath) > 255 Then
        fErr = "La ruta no puede superar los 255 caracteres."
    ElseIf fPlays < 0 Then
        fErr = "Las reproducciones no pueden ser negativas."
    End If

    If Len(fErr) = 0 Then
        Dim cn
        Set cn = GetConn()
        cn.BeginTrans
        On Error Resume Next

        If fId = 0 Then
            Call DbExec("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?)", _
                        Array(fName, fPath, fPlays))
            If Err.Number = 0 Then
                fId = ToInt(DbScalar("SELECT CAST(@@IDENTITY AS int) AS NEWID", Array()), 0)
            End If
        Else
            Call DbExec("UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", _
                        Array(fName, fPath, fPlays, fId))
        End If

        If Err.Number = 0 And fId > 0 Then
            Call DbExec("DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(fId))
        End If

        If Err.Number = 0 And fId > 0 Then
            Call DbExec("DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(fId))
        End If

        If Err.Number = 0 And fId > 0 And fAlbum > 0 Then
            Call DbExec("INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(fAlbum, fId))
        End If

        If Err.Number = 0 And fId > 0 And Len(fArtistsCsv) > 0 Then
            parts = Split(fArtistsCsv, ",")
            For i = 0 To UBound(parts)
                If Err.Number = 0 Then
                    pid = ToInt(Trim(parts(i)), 0)
                    If pid > 0 Then
                        Call DbExec("INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", Array(fId, pid))
                    End If
                End If
            Next
        End If

        If Err.Number <> 0 Then
            Err.Clear
            cn.RollbackTrans
            fErr = "No se pudo guardar la canción. Revisa los datos e inténtalo otra vez."
        Else
            cn.CommitTrans
            saved = True
        End If
        On Error GoTo 0
    End If

    If saved Then
        Call SetFlash("Canción guardada correctamente.", "success")
        Response.Redirect "admin_songs.asp"
        Response.End
    End If
Else
    fId = ToInt(Request.QueryString("id"), 0)
    If fId > 0 Then
        Set rsOne = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS,0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(fId))
        If Not rsOne.EOF Then
            fName  = CStr(rsOne("SONG_NAME") & "")
            fPath  = CStr(rsOne("SONG_PATH") & "")
            fPlays = ToInt(rsOne("PLAYS"), 0)
        Else
            fId = 0
        End If
        If rsOne.State <> 0 Then rsOne.Close
        Set rsOne = Nothing

        If fId > 0 Then
            fAlbum = ToInt(DbScalar("SELECT TOP 1 ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(fId)), 0)
            Set rsRel = DbQuery("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", Array(fId))
            Do While Not rsRel.EOF
                If Len(fArtistsCsv) > 0 Then fArtistsCsv = fArtistsCsv & ","
                fArtistsCsv = fArtistsCsv & CStr(rsRel("ID_ARTIST").Value)
                rsRel.MoveNext
            Loop
            If rsRel.State <> 0 Then rsRel.Close
            Set rsRel = Nothing
        End If
    End If
End If

isEdit = (fId > 0)
If isEdit Then pageTitle = "Editar canción"

previewUrl = ""
previewOk = False
If Len(fPath) > 0 Then
    previewUrl = SongAudioUrl(fId, fPath)
    previewOk = True
End If

' Archivos disponibles para el datalist de rutas
Dim fileList, fileCount
Set fso = GetFso()
fileList = ""
fileCount = 0
On Error Resume Next
Set fld = fso.GetFolder(Server.MapPath(strBase & RES_ROOT & "SongFiles"))
If Err.Number = 0 Then
    For Each f In fld.Files
        If LCase(Right(f.Name, 4)) = ".mp3" Then
            fileList = fileList & "<option value=""~/songFiles/" & H(f.Name) & """></option>"
            fileCount = fileCount + 1
        End If
    Next
End If
Err.Clear
On Error GoTo 0
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Canciones</span>
        <h2><%=IIf(isEdit, "Editar canción", "Registrar nueva canción")%></h2>
        <p>
            <%=IIf(isEdit, "Actualiza los datos de la canción y sus relaciones.", "Completa los datos para agregar una canción al catálogo.")%>
        </p>
    </div>
    <div class="page-head-actions">
        <a class="btn btn-ghost" href="<%=strBase%>admin_songs.asp">
            <svg><use href="#i-close"></use></svg> Volver al listado
        </a>
    </div>
</section>

<% If Len(fErr) > 0 Then %>
<div class="flash flash-error" role="alert">
    <svg><use href="#i-close"></use></svg>
    <span><%=H(fErr)%></span>
</div>
<% End If %>

<div class="form-layout">
    <form class="form-card" method="post" action="<%=strBase%>admin_song_form.asp" autocomplete="off">
        <input type="hidden" name="id" value="<%=fId%>">

        <div class="form-section">
            <h3><svg><use href="#i-music"></use></svg> Datos de la canción</h3>

            <label class="field">
                <span>Nombre de la canción <em>*</em></span>
                <input type="text" name="song_name" value="<%=H(fName)%>" maxlength="50" required
                       placeholder="Ej. Back in black">
            </label>

            <label class="field">
                <span>Ruta del archivo de audio <em>*</em></span>
                <input type="text" name="song_path" value="<%=H(fPath)%>" maxlength="255" required
                       list="songFiles" placeholder="~/songFiles/nombre.mp3">
                <datalist id="songFiles"><%=fileList%></datalist>
                <small class="hint">
                    Archivos disponibles: <strong><%=fileCount%></strong>.
                    Usa el formato <code>~/songFiles/archivo.mp3</code>.
                </small>
            </label>

            <label class="field">
                <span>Reproducciones</span>
                <input type="number" name="plays" value="<%=fPlays%>" min="0" step="1">
                <small class="hint">Se incrementa automáticamente cada vez que la canción suena.</small>
            </label>
        </div>

        <div class="form-section">
            <h3><svg><use href="#i-disc"></use></svg> Relaciones</h3>

            <label class="field">
                <span>Álbum</span>
                <select name="album">
                    <option value="0">— Sin álbum (sencillo) —</option>
                    <%
                    Set rsAlbums = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Array())
                    Do While Not rsAlbums.EOF
                    %>
                    <option value="<%=rsAlbums("ID_ALBUM")%>" <%=IIf(ToInt(rsAlbums("ID_ALBUM"),0) = fAlbum, "selected", "")%>>
                        <%=H(rsAlbums("TITLE_ALBUM"))%><% If Not IsNull(rsAlbums("RELEASE_YEAR")) Then %> (<%=rsAlbums("RELEASE_YEAR")%>)<% End If %>
                    </option>
                    <%
                        rsAlbums.MoveNext
                    Loop
                    If rsAlbums.State <> 0 Then rsAlbums.Close
                    Set rsAlbums = Nothing
                    %>
                </select>
            </label>

            <div class="field">
                <span>Artistas <small class="muted">(Ctrl para seleccionar varios)</small></span>
                <select name="artists" multiple size="6" class="multi">
                    <%
                    Set rsArtists = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME", Array())
                    Do While Not rsArtists.EOF
                    %>
                    <option value="<%=rsArtists("ID_ARTIST")%>" <%=IIf(CsvHas(fArtistsCsv, rsArtists("ID_ARTIST").Value), "selected", "")%>>
                        <%=H(rsArtists("ARTIST_NAME") & " " & Nz(rsArtists("ARTIST_LASTNAME"), ""))%>
                    </option>
                    <%
                        rsArtists.MoveNext
                    Loop
                    If rsArtists.State <> 0 Then rsArtists.Close
                    Set rsArtists = Nothing
                    %>
                </select>
                <small class="hint">Si no seleccionas artistas, la canción aparecerá como “Artista desconocido”.</small>
            </div>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-primary">
                <svg><use href="#i-check"></use></svg> <%=IIf(isEdit, "Guardar cambios", "Registrar canción")%>
            </button>
            <a class="btn btn-ghost" href="<%=strBase%>admin_songs.asp">Cancelar</a>
        </div>
    </form>

    <aside class="form-aside">
        <div class="aside-card">
            <h4>Vista previa del audio</h4>
            <% If previewOk Then %>
            <p class="muted small">Archivo resuelto desde la carpeta de recursos:</p>
            <audio class="preview-audio" controls preload="none" src="<%=previewUrl%>"></audio>
            <p class="mono small break"><%=H(previewUrl)%></p>
            <% Else %>
            <p class="muted small">Escribe una ruta para escuchar el archivo antes de guardar.</p>
            <% End If %>
        </div>

        <div class="aside-card">
            <h4>Buenas prácticas</h4>
            <ul class="tips">
                <li>El nombre se muestra en el listado y en el reproductor.</li>
                <li>La ruta debe apuntar a un archivo <code>.mp3</code> existente.</li>
                <li>Un álbum agrupa varias canciones mediante <code>TBL_TRACKLIST</code>.</li>
                <li>Los artistas se relacionan con la canción en <code>TBL_PLAYER</code>.</li>
            </ul>
        </div>

        <% If isEdit Then %>
        <div class="aside-card">
            <h4>Zona sensible</h4>
            <p class="muted small">Eliminar la canción también borra sus relaciones con álbumes, artistas y playlists.</p>
            <form method="post" action="<%=strBase%>admin_songs.asp" class="inline-form"
                  data-confirm="¿Eliminar la canción &quot;<%=H(fName)%>&quot;?">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="id" value="<%=fId%>">
                <button type="submit" class="btn btn-danger btn-block">
                    <svg><use href="#i-trash"></use></svg> Eliminar canción
                </button>
            </form>
        </div>
        <% End If %>
    </aside>
</div>

<!--#include file="includes/footer.asp"-->
