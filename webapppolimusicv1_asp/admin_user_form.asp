<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: alta / edicion de usuarios
'  Las contraseñas nuevas se guardan con SHA-1 (el login acepta
'  tambien los valores en texto plano que ya existen en la tabla).
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Nuevo usuario"
pageId       = "admin-users"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim fId, fUser, fPass, fMail, fBirth, fPhoto, fType, fErr, isEdit
Dim rsOne, saved, imgPreview, myId, dup

fId = 0
fUser = ""
fPass = ""
fMail = ""
fBirth = ""
fPhoto = ""
fType = 0
fErr = ""
saved = False
myId = ToInt(Session("UserID"), 0)

If IsPost() Then
    fId    = ToInt(Request.Form("id"), 0)
    fUser  = TrimStr(Request.Form("username"))
    fPass  = CStr(Request.Form("password") & "")
    fMail  = TrimStr(Request.Form("email"))
    fBirth = TrimStr(Request.Form("birthday"))
    fPhoto = TrimStr(Request.Form("user_photo"))
    fType  = ToInt(Request.Form("user_type"), 0)

    If Len(fUser) = 0 Then
        fErr = "El nombre de usuario es obligatorio."
    ElseIf Len(fUser) > 50 Then
        fErr = "El usuario no puede superar los 50 caracteres."
    ElseIf Len(fMail) = 0 Or InStr(fMail, "@") = 0 Then
        fErr = "Ingresa un correo electrónico válido."
    ElseIf Len(fMail) > 50 Then
        fErr = "El correo no puede superar los 50 caracteres."
    ElseIf fId = 0 And Len(fPass) < 4 Then
        fErr = "La contraseña debe tener al menos 4 caracteres."
    ElseIf Len(fPass) > 0 And Len(fPass) < 4 Then
        fErr = "La contraseña debe tener al menos 4 caracteres."
    ElseIf Len(fPass) > 50 Then
        fErr = "La contraseña no puede superar los 50 caracteres."
    ElseIf Len(fBirth) > 0 And Not IsDate(fBirth) Then
        fErr = "La fecha de cumpleaños no es válida."
    ElseIf fId = myId And fType <> 1 Then
        fErr = "No puedes quitarte el rol de administrador a ti mismo."
    End If

    If Len(fErr) = 0 Then
        dup = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", _
                             Array(fUser, fId)), 0)
        If dup > 0 Then fErr = "Ya existe un usuario con ese nombre."
    End If

    If Len(fErr) = 0 Then
        Dim bVal
        If Len(fBirth) = 0 Then
            bVal = Null
        Else
            bVal = CDate(fBirth)
        End If

        On Error Resume Next
        If fId = 0 Then
            Call DbExec("INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                        "VALUES (?, ?, ?, ?, ?, ?)", _
                        Array(fUser, Sha1Hex(fPass), fMail, bVal, fPhoto, fType))
        Else
            If Len(fPass) > 0 Then
                Call DbExec("UPDATE TBL_USER SET USERNAME = ?, PASSWORD = ?, EMAIL = ?, BIRTHDAY = ?, " & _
                            "USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?", _
                            Array(fUser, Sha1Hex(fPass), fMail, bVal, fPhoto, fType, fId))
            Else
                Call DbExec("UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, " & _
                            "USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?", _
                            Array(fUser, fMail, bVal, fPhoto, fType, fId))
            End If
        End If

        If Err.Number <> 0 Then
            Err.Clear
            fErr = "No se pudo guardar el usuario. Verifica los datos ingresados."
        Else
            saved = True
        End If
        On Error GoTo 0
    End If

    If saved Then
        ' Si el admin edita su propio nombre, se refresca la sesión
        If fId = myId Then
            Session("Username") = fUser
            Session("UserType") = fType
            Session("UserPhoto") = fPhoto
            Session("UserEmail") = fMail
        End If
        Call SetFlash("Usuario guardado correctamente.", "success")
        Response.Redirect "admin_users.asp"
        Response.End
    End If
Else
    fId = ToInt(Request.QueryString("id"), 0)
    If fId > 0 Then
        Set rsOne = DbQuery("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(fId))
        If Not rsOne.EOF Then
            fUser  = CStr(rsOne("USERNAME") & "")
            fMail  = CStr(rsOne("EMAIL") & "")
            fBirth = IsoDate(rsOne("BIRTHDAY").Value)
            fPhoto = CStr(Nz(rsOne("USER_PHOTO"), ""))
            fType  = ToInt(rsOne("USER_TYPE"), 0)
        Else
            fId = 0
        End If
        If rsOne.State <> 0 Then rsOne.Close
        Set rsOne = Nothing
    End If
End If

isEdit = (fId > 0)
If isEdit Then pageTitle = "Editar usuario"
imgPreview = UserPhotoUrl(fPhoto, fType)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Usuarios</span>
        <h2><%=IIf(isEdit, "Editar usuario", "Registrar nuevo usuario")%></h2>
        <p>El tipo de usuario define los permisos dentro de la aplicación.</p>
    </div>
    <div class="page-head-actions">
        <a class="btn btn-ghost" href="<%=strBase%>admin_users.asp">
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
    <form class="form-card" method="post" action="<%=strBase%>admin_user_form.asp" autocomplete="off">
        <input type="hidden" name="id" value="<%=fId%>">

        <div class="form-section">
            <h3><svg><use href="#i-users"></use></svg> Cuenta</h3>

            <div class="form-row">
                <label class="field">
                    <span>Usuario <em>*</em></span>
                    <input type="text" name="username" value="<%=H(fUser)%>" maxlength="50" required
                           placeholder="Ej. userNormal">
                </label>
                <label class="field">
                    <span>Contraseña <%=IIf(isEdit, "", "*")%></span>
                    <input type="password" name="password" maxlength="50" <%=IIf(isEdit, "", "required")%>
                           placeholder="<%=IIf(isEdit, "Dejar vacío para conservarla", "Mínimo 4 caracteres")%>">
                    <small class="hint">
                        <%=IIf(isEdit, "Si escribes una nueva contraseña se guardará cifrada con SHA-1.", "La contraseña se almacena cifrada con SHA-1.")%>
                    </small>
                </label>
            </div>

            <label class="field">
                <span>Correo electrónico <em>*</em></span>
                <input type="email" name="email" value="<%=H(fMail)%>" maxlength="50" required
                       placeholder="usuario@epn.edu.ec">
            </label>

            <div class="form-row">
                <label class="field">
                    <span>Cumpleaños</span>
                    <input type="date" name="birthday" value="<%=H(fBirth)%>">
                </label>
                <label class="field">
                    <span>Tipo de usuario</span>
                    <select name="user_type">
                        <%
                        Dim rsTypes
                        Set rsTypes = DbQuery("SELECT VALUE, COD_CATALOG FROM TBL_CATALOG " & _
                                              "WHERE COD_CATALOG_PARENT = '1' ORDER BY COD_CATALOG", Array())
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
                    <small class="hint">Solo <strong>ADMIN</strong> puede crear, editar y eliminar registros.</small>
                </label>
            </div>

            <label class="field">
                <span>Ruta de la foto</span>
                <input type="text" name="user_photo" value="<%=H(fPhoto)%>" maxlength="255"
                       placeholder="~/userImages/default.jpg">
            </label>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-primary">
                <svg><use href="#i-check"></use></svg> <%=IIf(isEdit, "Guardar cambios", "Registrar usuario")%>
            </button>
            <a class="btn btn-ghost" href="<%=strBase%>admin_users.asp">Cancelar</a>
        </div>
    </form>

    <aside class="form-aside">
        <div class="aside-card">
            <h4>Vista previa</h4>
            <img class="aside-preview round" src="<%=imgPreview%>" alt="Vista previa del usuario">
            <p class="muted small">Si la imagen no existe, se usa la foto por defecto del rol.</p>
        </div>

        <div class="aside-card">
            <h4>Roles del catálogo</h4>
            <ul class="tips">
                <li><strong>USER</strong> (0): solo puede escuchar el catálogo.</li>
                <li><strong>ADMIN</strong> (1): acceso total a los CRUD.</li>
                <li>El tipo se guarda en <code>TBL_USER.USER_TYPE</code> y se valida contra <code>TBL_CATALOG</code>.</li>
            </ul>
        </div>

        <% If isEdit And fId <> myId Then %>
        <div class="aside-card">
            <h4>Zona sensible</h4>
            <p class="muted small">Solo se puede eliminar un usuario sin playlists asociadas.</p>
            <form method="post" action="<%=strBase%>admin_users.asp" class="inline-form"
                  data-confirm="¿Eliminar al usuario &quot;<%=H(fUser)%>&quot;?">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="id" value="<%=fId%>">
                <button type="submit" class="btn btn-danger btn-block">
                    <svg><use href="#i-trash"></use></svg> Eliminar usuario
                </button>
            </form>
        </div>
        <% End If %>
    </aside>
</div>

<!--#include file="includes/footer.asp"-->
