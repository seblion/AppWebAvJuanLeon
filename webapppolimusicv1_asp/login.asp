<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - login.asp
'  Inicio de sesion contra TBL_USER (SQL Server localhost)
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
If SesionIniciada() Then Response.Redirect RUTA_BASE & "index.asp"

Dim mensaje, sesionCerrada
mensaje = ""
sesionCerrada = (Request.QueryString("logout") = "1")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim usuario, clave
    usuario = Trim(Request.Form("txtUsuario") & "")
    clave   = Request.Form("txtClave") & ""

    If usuario = "" Or clave = "" Then
        mensaje = "Debes ingresar tu usuario y tu contrasena."
    Else
        Dim errLogin
        errLogin = ""
        On Error Resume Next
        AbrirConexion()
        If Err.Number <> 0 Then
            errLogin = "No se pudo conectar con la base de datos. Verifica que SQL Server este ejecutandose."
            Err.Clear
            On Error GoTo 0
        Else
            On Error GoTo 0
            Dim cmdL, rsL
            ' Valida contrasena en texto plano o hash SHA1 (HASHBYTES en el servidor)
            Set cmdL = NuevoComando("SELECT ID_USER, USERNAME, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE USERNAME = ? AND (PASSWORD = ? OR PASSWORD = CONVERT(VARCHAR(64), HASHBYTES('SHA1', CAST(? AS VARCHAR(50))), 2))")
ParTexto cmdL, "@u",  usuario, 50
ParTexto cmdL, "@p1", clave, 50
ParTexto cmdL, "@p2", clave, 50
            On Error Resume Next
            Set rsL = cmdL.Execute
            If Err.Number <> 0 Then
                errLogin = "Error al validar las credenciales. Intentalo de nuevo."
                Err.Clear
            ElseIf Not rsL.EOF Then
                ' Evita fijacion de sesion: se abandona la sesion anterior
                Session.Contents.RemoveAll
                Session("ID_USER")    = rsL("ID_USER")
                Session("USERNAME")   = CStr(rsL("USERNAME") & "")
                Session("USER_PHOTO") = CStr(rsL("USER_PHOTO") & "")
                Session("USER_TYPE")  = rsL("USER_TYPE")
                CerrarConexion()
                Response.Redirect RUTA_BASE & "index.asp"
            Else
                mensaje = "Usuario o contrasena incorrectos. Intentalo de nuevo."
            End If
            On Error GoTo 0
        End If
        If errLogin <> "" Then mensaje = errLogin
    End If
End If

TituloPagina = "Iniciar sesion"
ClaseBody    = "pagina-login sin-reproductor"
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%=TituloPagina%> &mdash; PoliMusic</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect x='3' y='4' width='4' height='16' rx='1.5' fill='%232f7bff'/%3E%3Crect x='10' y='4' width='4' height='16' rx='1.5' fill='%2338bdf8'/%3E%3Crect x='17' y='4' width='4' height='16' rx='1.5' fill='%232f7bff'/%3E%3C/svg%3E">
<link rel="stylesheet" href="<%=RUTA_BASE%>css/styles.css">
<script>window.PM_BASE = "<%=RUTA_BASE%>";</script>
</head>
<body class="<%=ClaseBody%>">
<div class="fondo-degradado" aria-hidden="true"></div>

<div class="pagina-login">
  <!-- Panel artistico -->
  <aside class="login-arte">
    <img src="<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/playlist1.png" alt="" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
    <div class="login-arte-contenido">
      <div class="login-arte-marca">
        <span class="logo-ecualizador" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
        <span class="logo-texto">Poli<em>Music</em></span>
      </div>
      <h2 class="login-eslogan">Tu musica, <em>sin limites.</em></h2>
      <ul class="login-lista">
        <li><span><%=Icono("nota")%></span> Escucha tus canciones favoritas al instante</li>
        <li><span><%=Icono("disco")%></span> Explora albumes y artistas</li>
        <li><span><%=Icono("lista")%></span> Organiza tus listas de reproduccion</li>
        <li><span><%=Icono("escudo")%></span> Panel de administracion para el rol admin</li>
      </ul>
    </div>
  </aside>

  <!-- Formulario -->
  <section class="login-panel">
    <div class="login-caja">
      <% If sesionCerrada Then %>
        <div class="aviso aviso-exito" style="margin-bottom:20px"><span><%=Icono("check")%></span><span>Sesion cerrada correctamente. Vuelve pronto.</span></div>
      <% End If %>

      <% If mensaje <> "" Then %>
        <div class="aviso aviso-error"><span><%=Icono("alerta")%></span><span><%=EscaparHTML(mensaje)%></span></div>
      <% End If %>

      <h1>Bienvenido a PoliMusic</h1>
      <p>Inicia sesion para acceder a tu musica.</p>

      <form method="post" action="<%=RUTA_BASE%>login.asp" autocomplete="on">
        <div class="campo">
          <label for="txtUsuario">Usuario</label>
          <input type="text" id="txtUsuario" name="txtUsuario" maxlength="50" placeholder="Tu nombre de usuario" required value="<%=EscaparHTML(Request.Form("txtUsuario"))%>">
        </div>
        <div class="campo">
          <label for="txtClave">Contrasena</label>
          <input type="password" id="txtClave" name="txtClave" maxlength="50" placeholder="Tu contrasena" required>
        </div>
        <button type="submit" class="boton boton-primario boton-grande">Iniciar sesion</button>
      </form>

      <div class="cuenta-demo">
        <b>Cuentas de prueba</b> (incluidas en BDD_PoliMusic.sql):<br>
        <code>userNormal</code> / contrasena <code>12345678</code> &middot; rol USER<br>
        <code>userAdmin</code> / contrasena <code>12345678</code> &middot; rol ADMIN
      </div>
    </div>
  </section>
</div>

<script src="<%=RUTA_BASE%>js/app.js"></script>
</body>
</html>
<%
CerrarConexion()
%>
