<%
'=====================================================================
'  PoliMusic - includes/header.asp
'  Apertura del documento y barra superior (consciente del rol)
'  Requiere: conexion.asp incluido antes.
'=====================================================================
If Len(TituloPagina) = 0 Then TituloPagina = "PoliMusic"
Dim nomPagina
nomPagina = PaginaActual()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%=EscaparHTML(TituloPagina)%> &mdash; PoliMusic</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect x='3' y='4' width='4' height='16' rx='1.5' fill='%232f7bff'/%3E%3Crect x='10' y='4' width='4' height='16' rx='1.5' fill='%2338bdf8'/%3E%3Crect x='17' y='4' width='4' height='16' rx='1.5' fill='%232f7bff'/%3E%3C/svg%3E">
<link rel="stylesheet" href="<%=RUTA_BASE%>css/styles.css">
<script>window.PM_BASE = "<%=RUTA_BASE%>";</script>
</head>
<body class="<%=ClaseBody%>">
<div class="fondo-degradado" aria-hidden="true"></div>

<header class="barra-superior">
  <a class="logo" href="<%=RUTA_BASE%>index.asp">
    <span class="logo-ecualizador" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
    <span class="logo-texto">Poli<em>Music</em></span>
  </a>

  <nav class="nav-principal" aria-label="Navegacion principal">
    <a href="<%=RUTA_BASE%>index.asp"<% If nomPagina = "index.asp" Then Response.Write " class=""activo""" %>><span class="nav-icono"><%=Icono("inicio")%></span><span class="nav-texto">Inicio</span></a>
    <a href="<%=RUTA_BASE%>index.asp#song-list"><span class="nav-icono"><%=Icono("nota")%></span><span class="nav-texto">Canciones</span></a>
    <a href="<%=RUTA_BASE%>index.asp#albumes"><span class="nav-icono"><%=Icono("disco")%></span><span class="nav-texto">Albumes</span></a>
    <a href="<%=RUTA_BASE%>index.asp#artistas"><span class="nav-icono"><%=Icono("micro")%></span><span class="nav-texto">Artistas</span></a>
    <% If EsAdmin() Then %>
    <a href="<%=RUTA_BASE%>admin/canciones.asp" class="nav-admin"><span class="nav-icono"><%=Icono("escudo")%></span><span class="nav-texto">Administracion</span></a>
    <% End If %>
  </nav>

  <div class="barra-derecha">
    <% If SesionIniciada() Then %>
      <a class="chip-usuario" href="<%=RUTA_BASE%>index.asp" title="Mi perfil">
        <img class="chip-foto" src="<%=RutaRecurso(Session("USER_PHOTO"), "imagen")%>" alt="" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/UserImages/default.jpg';">
        <span class="chip-datos">
          <span class="chip-nombre"><%=EscaparHTML(Session("USERNAME"))%></span>
          <span class="chip-rol"><%=EscaparHTML(EtiquetaCatalogo("1", CStr(Session("USER_TYPE") & "")))%></span>
        </span>
      </a>
      <a class="boton-icono boton-salir" href="<%=RUTA_BASE%>logout.asp" title="Cerrar sesion" aria-label="Cerrar sesion"><%=Icono("salir")%></a>
    <% Else %>
      <a class="boton boton-primario" href="<%=RUTA_BASE%>login.asp">Iniciar sesion</a>
    <% End If %>
  </div>
</header>

<main class="contenido">
