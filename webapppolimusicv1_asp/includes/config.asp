<%
'=====================================================================
'  PoliMusic :: Configuracion general y utilidades compartidas
'  ASP clasico + VBScript  |  Base de datos: BDD_PoliMusic (SQL Server)
'=====================================================================
Response.Buffer   = True
Response.CodePage = 65001
Response.CharSet  = "utf-8"
Session.CodePage  = 65001

Const APP_NAME   = "PoliMusic"
Const APP_SLOGAN = "Tu música, en azul"

' ---- Conexion a SQL Server -----------------------------------------
Const DB_SERVER = "localhost"
Const DB_NAME   = "BDD_PoliMusic"
Const DB_USER   = "usr_polimusic"
Const DB_PASS   = "Politecnica1"

' ---- Carpeta de recursos multimedia --------------------------------
Const RES_ROOT  = "SongFilesAndImages_resources/"

' ---- Variables globales de pagina ----------------------------------
Dim strBase, pageTitle, pageId, strFlash, strFlashType
Dim strPlaylistJson, strBodyClass, strSearchTerm

strBase         = "./"
pageTitle       = APP_NAME
pageId          = ""
strFlash        = ""
strFlashType    = ""
strPlaylistJson = "[]"
strBodyClass    = ""
strSearchTerm   = ""

Function ConnString()
    ConnString = "Provider=SQLOLEDB;Data Source=" & DB_SERVER & ";" & _
                 "Initial Catalog=" & DB_NAME & ";" & _
                 "User ID=" & DB_USER & ";Password=" & DB_PASS & ";"
End Function

'---------------------------------------------------------------------
'  Utilidades de texto / tipos
'---------------------------------------------------------------------
Function H(v)
    If IsNull(v) Or IsEmpty(v) Then
        H = ""
    Else
        H = Server.HTMLEncode(CStr(v))
    End If
End Function

Function Nz(v, d)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = d
    Else
        Nz = v
    End If
End Function

Function TrimStr(v)
    TrimStr = Trim(CStr(Nz(v, "")))
End Function

' VBScript no trae IIf, se agrega como utilidad
Function IIf(cond, valTrue, valFalse)
    If cond Then
        IIf = valTrue
    Else
        IIf = valFalse
    End If
End Function

Function ToInt(v, d)
    If IsNumeric(v) Then
        ToInt = CLng(v)
    Else
        ToInt = d
    End If
End Function

Function IsPost()
    IsPost = (UCase(TrimStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
End Function

Sub SetFlash(msg, kind)
    Session("FlashMsg")  = msg
    Session("FlashType") = kind
End Sub

Function FmtDate(v)
    Dim d
    If IsNull(v) Or Len(TrimStr(v)) = 0 Then
        FmtDate = "&mdash;"
        Exit Function
    End If
    On Error Resume Next
    d = CDate(v)
    If Err.Number <> 0 Then
        Err.Clear
        FmtDate = H(v)
    Else
        FmtDate = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
    End If
    On Error GoTo 0
End Function

' Fecha en formato ISO (yyyy-mm-dd) para inputs type="date"
Function IsoDate(v)
    Dim d
    IsoDate = ""
    If IsNull(v) Or Len(TrimStr(v)) = 0 Then Exit Function
    On Error Resume Next
    d = CDate(v)
    If Err.Number = 0 Then
        IsoDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function FmtNum(v)
    If IsNull(v) Or Not IsNumeric(v) Then
        FmtNum = "0"
    Else
        FmtNum = FormatNumber(CLng(v), 0, -1, 0, -1)
    End If
End Function

Function FmtShortNum(v)
    Dim n
    If IsNull(v) Or Not IsNumeric(v) Then
        FmtShortNum = "0"
        Exit Function
    End If
    n = CLng(v)
    If n >= 1000000 Then
        FmtShortNum = Replace(FormatNumber(n / 1000000, 1, -1, 0, -1), ",", ".") & "M"
    ElseIf n >= 1000 Then
        FmtShortNum = Replace(FormatNumber(n / 1000, 1, -1, 0, -1), ",", ".") & "K"
    Else
        FmtShortNum = CStr(n)
    End If
End Function

'---------------------------------------------------------------------
'  Utilidades de archivos / recursos multimedia
'---------------------------------------------------------------------
Dim g_fso
Function GetFso()
    If Not IsObject(g_fso) Then
        Set g_fso = Server.CreateObject("Scripting.FileSystemObject")
    End If
    Set GetFso = g_fso
End Function

Function FileNameOf(p)
    Dim s, i, j
    s = TrimStr(p)
    i = InStrRev(s, "/")
    j = InStrRev(s, "\")
    If j > i Then i = j
    If i > 0 Then s = Mid(s, i + 1)
    FileNameOf = s
End Function

Function NormalizeKey(s)
    Dim t
    t = LCase(TrimStr(s))
    t = Replace(t, " ", "")
    t = Replace(t, "-", "")
    t = Replace(t, "_", "")
    NormalizeKey = t
End Function

Function UrlFile(n)
    Dim t
    t = CStr(n)
    t = Replace(t, " ", "%20")
    t = Replace(t, "#", "%23")
    t = Replace(t, "&", "%26")
    t = Replace(t, "+", "%2B")
    UrlFile = t
End Function

' Devuelve True si el archivo existe dentro de la carpeta de recursos
Function ResFileExists(subFolder, fileName)
    Dim ph, fso
    ResFileExists = False
    If Len(TrimStr(fileName)) = 0 Then Exit Function
    ph = Server.MapPath(strBase & RES_ROOT & subFolder & "/" & fileName)
    On Error Resume Next
    Set fso = GetFso()
    If IsObject(fso) Then ResFileExists = fso.FileExists(ph)
    If Err.Number <> 0 Then ResFileExists = False
    Err.Clear
    On Error GoTo 0
End Function

Function ResUrl(subFolder, fileName)
    ResUrl = strBase & RES_ROOT & subFolder & "/" & UrlFile(fileName)
End Function

' Busca un archivo dentro de una carpeta de recursos comparando nombres
' normalizados (sin espacios, guiones ni guion bajo).
Function FindResFile(subFolder, fileName)
    Dim fso, folder, f, key
    FindResFile = ""
    key = NormalizeKey(fileName)
    If Len(key) = 0 Then Exit Function
    On Error Resume Next
    Set fso = GetFso()
    folder = Server.MapPath(strBase & RES_ROOT & subFolder)
    If fso.FolderExists(folder) Then
        For Each f In fso.GetFolder(folder).Files
            If NormalizeKey(f.Name) = key Then
                FindResFile = f.Name
                Exit For
            End If
        Next
    End If
    If Err.Number <> 0 Then
        Err.Clear
        FindResFile = ""
    End If
    On Error GoTo 0
End Function

'---------------------------------------------------------------------
'  Resolucion de archivos multimedia de cada cancion
'  La base de datos guarda rutas "~/songFiles/xxx.mp3"; el proyecto
'  guarda los archivos reales en SongFilesAndImages_resources/SongFiles.
'---------------------------------------------------------------------
Function SongAudioUrl(idSong, songPath)
    Dim fn, realName, arr, i

    fn = FileNameOf(songPath)

    ' 1) coincidencia exacta del nombre
    If ResFileExists("SongFiles", fn) Then
        SongAudioUrl = ResUrl("SongFiles", fn)
        Exit Function
    End If

    ' 2) coincidencia por nombre normalizado (espacios / guiones)
    realName = FindResFile("SongFiles", fn)
    If Len(realName) > 0 Then
        SongAudioUrl = ResUrl("SongFiles", realName)
        Exit Function
    End If

    ' 3) repertorio disponible: se asigna de forma ciclica
    arr = Array("bensound-adventure.mp3", "BackinBlack-ACDC.mp3", "Loser-Tame Impala.mp3")
    i = (ToInt(idSong, 1) - 1) Mod 3
    If i < 0 Then i = 0
    SongAudioUrl = ResUrl("SongFiles", arr(i))
End Function

Function SongCoverUrl(idSong, songName, albumImage)
    Dim fn, realName, arr, i, nm

    fn = FileNameOf(albumImage)
    If ResFileExists("PlaylistImages", fn) Then
        SongCoverUrl = ResUrl("PlaylistImages", fn)
        Exit Function
    End If
    realName = FindResFile("PlaylistImages", fn)
    If Len(realName) > 0 Then
        SongCoverUrl = ResUrl("PlaylistImages", realName)
        Exit Function
    End If

    nm = LCase(TrimStr(songName))
    If InStr(nm, "black") > 0 And ResFileExists("PlaylistImages", "play-back-in-black.jpg") Then
        SongCoverUrl = ResUrl("PlaylistImages", "play-back-in-black.jpg")
        Exit Function
    End If
    If InStr(nm, "loser") > 0 And ResFileExists("PlaylistImages", "play-loser.jpg") Then
        SongCoverUrl = ResUrl("PlaylistImages", "play-loser.jpg")
        Exit Function
    End If

    arr = Array("playlist1.png", "playlist2.png", "playlist3.png", "playlist4.png", "playlist5.png")
    i = (ToInt(idSong, 1) - 1) Mod 5
    If i < 0 Then i = 0
    SongCoverUrl = ResUrl("PlaylistImages", arr(i))
End Function

Function PlaylistCoverUrl(idPlaylist, photo)
    Dim fn, realName, arr, i
    fn = FileNameOf(photo)
    If ResFileExists("PlaylistImages", fn) Then
        PlaylistCoverUrl = ResUrl("PlaylistImages", fn)
        Exit Function
    End If
    realName = FindResFile("PlaylistImages", fn)
    If Len(realName) > 0 Then
        PlaylistCoverUrl = ResUrl("PlaylistImages", realName)
        Exit Function
    End If
    arr = Array("playlist1.png", "playlist2.png", "playlist3.png", "playlist4.png", "playlist5.png")
    i = (ToInt(idPlaylist, 1) - 1) Mod 5
    If i < 0 Then i = 0
    PlaylistCoverUrl = ResUrl("PlaylistImages", arr(i))
End Function

' Portada de album: usa la imagen registrada si existe en los recursos
Function AlbumCoverUrl(idAlbum, image)
    AlbumCoverUrl = PlaylistCoverUrl(idAlbum, image)
End Function

Function UserPhotoUrl(photo, userType)
    Dim fn, realName
    fn = FileNameOf(photo)
    If ResFileExists("UserImages", fn) Then
        UserPhotoUrl = ResUrl("UserImages", fn)
        Exit Function
    End If
    realName = FindResFile("UserImages", fn)
    If Len(realName) > 0 Then
        UserPhotoUrl = ResUrl("UserImages", realName)
        Exit Function
    End If
    If ToInt(userType, 0) = 1 And ResFileExists("UserImages", "root.png") Then
        UserPhotoUrl = ResUrl("UserImages", "root.png")
        Exit Function
    End If
    If ResFileExists("UserImages", "default.jpg") Then
        UserPhotoUrl = ResUrl("UserImages", "default.jpg")
    Else
        UserPhotoUrl = strBase & "css/avatar.svg"
    End If
End Function

Function ArtistPhotoUrl(image, artistName)
    Dim fn, realName
    fn = FileNameOf(image)
    If ResFileExists("ArtistImages", fn) Then
        ArtistPhotoUrl = ResUrl("ArtistImages", fn)
        Exit Function
    End If
    realName = FindResFile("ArtistImages", fn)
    If Len(realName) > 0 Then
        ArtistPhotoUrl = ResUrl("ArtistImages", realName)
        Exit Function
    End If
    If ResFileExists("UserImages", "default.jpg") Then
        ArtistPhotoUrl = ResUrl("UserImages", "default.jpg")
    Else
        ArtistPhotoUrl = strBase & "css/avatar.svg"
    End If
End Function

'---------------------------------------------------------------------
'  Serializacion JSON minima (para enviar la lista al reproductor)
'---------------------------------------------------------------------
Function JsStr(v)
    Dim t
    t = CStr(Nz(v, ""))
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, "</", "<\/")
    JsStr = """" & t & """"
End Function

'---------------------------------------------------------------------
'  Sesion
'---------------------------------------------------------------------
Function SessionStr(name)
    SessionStr = CStr(Session(name) & "")
End Function

Function CurrentUserName()
    CurrentUserName = CStr(Session("Username") & "")
End Function

Function CurrentUserType()
    CurrentUserType = ToInt(Session("UserType"), 0)
End Function

Function CurrentUserPhoto()
    CurrentUserPhoto = UserPhotoUrl(Session("UserPhoto"), Session("UserType"))
End Function

Function CurrentUserInitials()
    Dim n
    n = TrimStr(Session("Username"))
    If Len(n) = 0 Then n = "?"
    If Len(n) > 2 Then n = Left(n, 2)
    CurrentUserInitials = UCase(n)
End Function

Function RoleLabel(t)
    If ToInt(t, 0) = 1 Then
        RoleLabel = "Administrador"
    Else
        RoleLabel = "Usuario"
    End If
End Function
%>
<!--#include file="db.asp"-->
<!--#include file="sha1.asp"-->
<!--#include file="auth.asp"-->
