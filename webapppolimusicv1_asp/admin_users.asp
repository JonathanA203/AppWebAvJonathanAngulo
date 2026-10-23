<%@ Language="VBScript" CodePage="65001" %>
<%
'=====================================================================
'  PoliMusic :: administracion de usuarios
'=====================================================================
Response.Buffer = True
%>
<!--#include file="includes/config.asp"-->
<%
strBase      = "./"
pageTitle    = "Usuarios"
pageId       = "admin-users"
strBodyClass = "page-admin"

Call RequireAdmin()

Dim aAction, aId, aQ, aSql, aParams, rsList, i, aTotal, myId

aAction = TrimStr(Request.Form("action"))
myId = ToInt(Session("UserID"), 0)

If IsPost() And aAction = "delete" Then
    aId = ToInt(Request.Form("id"), 0)
    If aId = myId Then
        Call SetFlash("No puedes eliminar tu propia cuenta mientras la sesión está activa.", "error")
    ElseIf aId > 0 Then
        If DbTry("DELETE FROM TBL_USER WHERE ID_USER = ?", Array(aId)) Then
            Call SetFlash("Usuario eliminado correctamente.", "success")
        Else
            Call SetFlash("No se pudo eliminar el usuario: tiene playlists asociadas.", "error")
        End If
    End If
    Response.Redirect "admin_users.asp"
    Response.End
End If

aQ = TrimStr(Request.QueryString("q"))
aTotal = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER", Array()), 0)

aSql = "SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, " & _
       " (SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS TOTAL " & _
       "FROM TBL_USER u "
If Len(aQ) > 0 Then
    aSql = aSql & "WHERE (u.USERNAME LIKE ? OR u.EMAIL LIKE ?) "
    aParams = Array("%" & aQ & "%", "%" & aQ & "%")
Else
    aParams = Array()
End If
aSql = aSql & "ORDER BY u.USERNAME"

Set rsList = DbQuery(aSql, aParams)
%>
<!--#include file="includes/header.asp"-->

<section class="page-head">
    <div>
        <span class="section-kicker">Administración · Usuarios</span>
        <h2>Gestión de usuarios</h2>
        <p>
            Controla las cuentas y el rol de cada usuario. El rol <strong>Administrador</strong>
            (<code>USER_TYPE = 1</code>) es el único que puede ejecutar los CRUD del catálogo.
        </p>
    </div>
    <div class="page-head-actions">
        <span class="count-pill"><%=FmtNum(aTotal)%> registros</span>
        <a class="btn btn-primary" href="<%=strBase%>admin_user_form.asp">
            <svg><use href="#i-plus"></use></svg> Nuevo usuario
        </a>
    </div>
</section>

<div class="table-card">
    <div class="table-toolbar">
        <form class="table-search" method="get" action="<%=strBase%>admin_users.asp">
            <svg><use href="#i-search"></use></svg>
            <input type="search" name="q" value="<%=H(aQ)%>" placeholder="Buscar por usuario o correo…">
            <button type="submit" class="btn btn-soft sm">Filtrar</button>
            <% If Len(aQ) > 0 Then %>
            <a class="link-clear" href="<%=strBase%>admin_users.asp">Limpiar</a>
            <% End If %>
        </form>
    </div>

    <div class="table-wrap">
        <table class="data-table">
            <thead>
                <tr>
                    <th class="col-cover">Foto</th>
                    <th>Usuario</th>
                    <th>Correo</th>
                    <th>Cumpleaños</th>
                    <th>Rol</th>
                    <th class="num">Playlists</th>
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
                        <img class="thumb round" src="<%=UserPhotoUrl(rsList("USER_PHOTO").Value, rsList("USER_TYPE").Value)%>"
                             alt="Foto de <%=H(rsList("USERNAME"))%>">
                    </td>
                    <td>
                        <strong class="cell-title"><%=H(rsList("USERNAME"))%></strong>
                        <span class="cell-sub">ID <%=rsList("ID_USER")%><% If ToInt(rsList("ID_USER"),0) = myId Then %> · tu cuenta<% End If %></span>
                    </td>
                    <td class="muted"><%=H(rsList("EMAIL"))%></td>
                    <td class="muted"><%=FmtDate(rsList("BIRTHDAY").Value)%></td>
                    <td>
                        <span class="role-pill <%=IIf(ToInt(rsList("USER_TYPE"),0) = 1, "role-admin", "role-user")%>">
                            <%=H(RoleLabel(rsList("USER_TYPE").Value))%>
                        </span>
                    </td>
                    <td class="num"><%=FmtNum(rsList("TOTAL").Value)%></td>
                    <td class="col-actions">
                        <div class="row-actions">
                            <a class="mini-btn" href="<%=strBase%>admin_user_form.asp?id=<%=rsList("ID_USER")%>" title="Editar">
                                <svg><use href="#i-edit"></use></svg>
                            </a>
                            <% If ToInt(rsList("ID_USER"), 0) <> myId Then %>
                            <form method="post" action="<%=strBase%>admin_users.asp" class="inline-form"
                                  data-confirm="¿Eliminar al usuario &quot;<%=H(rsList("USERNAME"))%>&quot;?">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=rsList("ID_USER")%>">
                                <button type="submit" class="mini-btn danger" title="Eliminar">
                                    <svg><use href="#i-trash"></use></svg>
                                </button>
                            </form>
                            <% Else %>
                            <span class="mini-btn is-disabled" title="No puedes eliminar tu propia cuenta">
                                <svg><use href="#i-trash"></use></svg>
                            </span>
                            <% End If %>
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
                        <svg><use href="#i-users"></use></svg>
                        <span>No se encontraron usuarios.</span>
                    </div>
                </td></tr>
            <% End If %>
            </tbody>
        </table>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
