<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - index.asp
'  Pagina principal: bienvenida, listado de canciones, albumes,
'  artistas y listas de reproduccion. Requiere sesion iniciada.
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
RequiereLogin()
AbrirConexion()

TituloPagina = "Inicio"
UsarPlayer = True

Dim avisoPermiso
avisoPermiso = (Request.QueryString("error") = "permiso")

'---------------- Canciones (con artistas, albumes y portada) ----------------
Dim canciones(), nCanciones
nCanciones = 0
Dim cmdC, rsC
Set cmdC = NuevoComando("SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG ORDER BY ID_SONG")
Set rsC = cmdC.Execute
Do While Not rsC.EOF
    ReDim Preserve canciones(nCanciones)
    Set canciones(nCanciones) = Server.CreateObject("Scripting.Dictionary")
    canciones(nCanciones)("id")       = rsC("ID_SONG")
    canciones(nCanciones)("nombre")   = CStr(rsC("SONG_NAME") & "")
    canciones(nCanciones)("src")      = RutaRecurso(rsC("SONG_PATH"), "cancion")
    canciones(nCanciones)("portada")  = PortadaDeCancion(rsC("ID_SONG"))
    canciones(nCanciones)("artistas") = ArtistasDeCancion(rsC("ID_SONG"))
    canciones(nCanciones)("albumes")  = AlbumesDeCancion(rsC("ID_SONG"))
    If IsNull(rsC("PLAYS")) Or IsEmpty(rsC("PLAYS")) Then
        canciones(nCanciones)("plays") = 0
    Else
        canciones(nCanciones)("plays") = rsC("PLAYS")
    End If
    nCanciones = nCanciones + 1
    rsC.MoveNext
Loop
rsC.Close
Set rsC = Nothing
Set cmdC = Nothing

'---------------- Albumes ----------------
Dim albumes(), nAlbumes
nAlbumes = 0
Set cmdC = NuevoComando("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY TITLE_ALBUM")
Set rsC = cmdC.Execute
Do While Not rsC.EOF
    ReDim Preserve albumes(nAlbumes)
    Set albumes(nAlbumes) = Server.CreateObject("Scripting.Dictionary")
    albumes(nAlbumes)("id")      = rsC("ID_ALBUM")
    albumes(nAlbumes)("titulo")  = CStr(rsC("TITLE_ALBUM") & "")
    If IsNull(rsC("RELEASE_YEAR")) Or IsEmpty(rsC("RELEASE_YEAR")) Then
        albumes(nAlbumes)("anio") = ""
    Else
        albumes(nAlbumes)("anio") = CStr(rsC("RELEASE_YEAR"))
    End If
    albumes(nAlbumes)("imagen") = RutaRecurso(rsC("ALBUM_IMAGE_PATH"), "imagen")
    nAlbumes = nAlbumes + 1
    rsC.MoveNext
Loop
rsC.Close
Set rsC = Nothing
Set cmdC = Nothing

'---------------- Artistas ----------------
Dim artistas(), nArtistas
nArtistas = 0
Set cmdC = NuevoComando("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ARTIST_NAME")
Set rsC = cmdC.Execute
Do While Not rsC.EOF
    ReDim Preserve artistas(nArtistas)
    Set artistas(nArtistas) = Server.CreateObject("Scripting.Dictionary")
    artistas(nArtistas)("id")      = rsC("ID_ARTIST")
    artistas(nArtistas)("nombre")  = Trim(CStr(rsC("ARTIST_NAME") & "") & " " & CStr(rsC("ARTIST_LASTNAME") & ""))
    artistas(nArtistas)("imagen")  = RutaRecurso(rsC("ARTIST_IMAGE"), "imagen")
    nArtistas = nArtistas + 1
    rsC.MoveNext
Loop
rsC.Close
Set rsC = Nothing
Set cmdC = Nothing

'---------------- Listas de reproduccion del usuario ----------------
Dim listas(), nListas
nListas = 0
Set cmdC = NuevoComando("SELECT ID_PLAYLIST, TITLE, CREATION_DATE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY CREATION_DATE DESC, ID_PLAYLIST DESC")
ParEntero cmdC, "@u", CLng(Session("ID_USER"))
Set rsC = cmdC.Execute
Do While Not rsC.EOF
    ReDim Preserve listas(nListas)
    Set listas(nListas) = Server.CreateObject("Scripting.Dictionary")
    listas(nListas)("id")     = rsC("ID_PLAYLIST")
    listas(nListas)("titulo") = CStr(rsC("TITLE") & "")
    listas(nListas)("tipo")   = EtiquetaCatalogo("2", CStr(rsC("TYPE") & ""))
    listas(nListas)("foto")   = RutaRecurso(rsC("PHOTO"), "imagen")
    nListas = nListas + 1
    rsC.MoveNext
Loop
rsC.Close
Set rsC = Nothing
Set cmdC = Nothing
%>
<!--#include file="includes/header.asp"-->
<% If avisoPermiso Then %>
<div class="aviso aviso-error"><span><%=Icono("alerta")%></span><span>No tienes permisos de administrador para acceder a esa seccion.</span></div>
<% End If %>

<!-- ============ Hero de bienvenida ============ -->
<section class="hero">
  <div class="hero-texto">
    <p class="hero-saludo">PoliMusic &middot; Escucha sin limites</p>
    <h1 class="hero-titulo">Bienvenido, <%=EscaparHTML(Session("USERNAME"))%></h1>
    <p class="hero-sub">Tu musica favorita reunida en un solo lugar. Explora el listado de canciones, reproducelas al instante y descubre nuevos albumes y artistas.</p>
    <div class="hero-chips">
      <span class="chip-estadistica"><%=Icono("nota")%> <%=nCanciones%> canciones</span>
      <span class="chip-estadistica"><%=Icono("disco")%> <%=nAlbumes%> albumes</span>
      <span class="chip-estadistica"><%=Icono("micro")%> <%=nArtistas%> artistas</span>
      <span class="chip-estadistica"><%=Icono("lista")%> <%=nListas%> listas</span>
    </div>
    <div class="hero-acciones">
      <a class="boton boton-primario boton-grande" href="#song-list"><%=Icono("play")%> Escuchar ahora</a>
      <% If EsAdmin() Then %>
      <a class="boton boton-fantasma boton-grande" href="<%=RUTA_BASE%>admin/canciones.asp"><%=Icono("escudo")%> Panel de administracion</a>
      <% End If %>
    </div>
  </div>
  <div class="hero-arte" aria-hidden="true">
    <img src="<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/playlist1.png" alt="" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
    <img src="<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/playlist3.png" alt="" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
  </div>
</section>

<!-- ============ Listado de canciones ============ -->
<section class="seccion" id="song-list">
  <div class="seccion-cabecera">
    <h2 class="seccion-titulo">Song List</h2>
    <span class="seccion-info">Listado de canciones &middot; <%=nCanciones%> en total</span>
    <div class="seccion-extra filtro-caja">
      <%=Icono("buscar")%>
      <input type="search" id="filtro-canciones" placeholder="Filtrar canciones..." aria-label="Filtrar canciones">
    </div>
  </div>

  <div class="lista-canciones">
    <div class="lista-cabecera">
      <span>#</span>
      <span>Titulo / Artista</span>
      <span>Album</span>
      <span>Reproducciones</span>
    </div>

    <% If nCanciones = 0 Then %>
      <div class="estado-vacio">
        <%=Icono("nota")%>
        <strong>Todavia no hay canciones registradas</strong>
        <span>Cuando el administrador agregue canciones apareceran aqui para reproducirlas.</span>
      </div>
    <% Else %>
      <% Dim i %>
      <% For i = 0 To nCanciones - 1 %>
        <div class="fila-cancion" data-id="<%=canciones(i)("id")%>">
          <span class="fila-indice">
            <span class="num"><%=i + 1%></span>
            <button class="boton-reproducir" type="button" aria-label="Reproducir <%=EscaparHTML(canciones(i)("nombre"))%>">
              <span class="ico-play"><%=Icono("play")%></span>
              <span class="ico-pausa"><%=Icono("pausa")%></span>
            </button>
          </span>
          <span class="fila-info">
            <img class="fila-portada" src="<%=canciones(i)("portada")%>" alt="" loading="lazy" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
            <span class="fila-textos">
              <span class="fila-titulo"><%=EscaparHTML(canciones(i)("nombre"))%></span>
              <span class="fila-artista"><%=EscaparHTML(canciones(i)("artistas"))%></span>
            </span>
          </span>
          <span class="fila-album"><%=EscaparHTML(canciones(i)("albumes"))%></span>
          <span class="fila-plays"><%=canciones(i)("plays")%> reproducciones</span>
        </div>
      <% Next %>
    <% End If %>
  </div>
</section>

<!-- ============ Albumes ============ -->
<section class="seccion" id="albumes">
  <div class="seccion-cabecera">
    <h2 class="seccion-titulo">Albumes</h2>
    <span class="seccion-info"><%=nAlbumes%> albumes en la coleccion</span>
  </div>
  <% If nAlbumes = 0 Then %>
    <div class="estado-vacio">
      <%=Icono("disco")%>
      <strong>No hay albumes registrados</strong>
      <span>Los albumes creados por el administrador se mostraran en esta seccion.</span>
    </div>
  <% Else %>
    <div class="rejilla-tarjetas">
      <% For i = 0 To nAlbumes - 1 %>
        <article class="tarjeta">
          <div class="tarjeta-portada">
            <img src="<%=albumes(i)("imagen")%>" alt="Portada de <%=EscaparHTML(albumes(i)("titulo"))%>" loading="lazy" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
          </div>
          <h3 class="tarjeta-nombre"><%=EscaparHTML(albumes(i)("titulo"))%></h3>
          <p class="tarjeta-meta">Album &middot; <%=EscaparHTML(albumes(i)("anio"))%></p>
        </article>
      <% Next %>
    </div>
  <% End If %>
</section>

<!-- ============ Artistas ============ -->
<section class="seccion" id="artistas">
  <div class="seccion-cabecera">
    <h2 class="seccion-titulo">Artistas</h2>
    <span class="seccion-info"><%=nArtistas%> artistas</span>
  </div>
  <% If nArtistas = 0 Then %>
    <div class="estado-vacio">
      <%=Icono("micro")%>
      <strong>No hay artistas registrados</strong>
      <span>Los artistas creados por el administrador se mostraran en esta seccion.</span>
    </div>
  <% Else %>
    <div class="rejilla-tarjetas">
      <% For i = 0 To nArtistas - 1 %>
        <article class="tarjeta">
          <div class="tarjeta-portada">
            <img src="<%=artistas(i)("imagen")%>" alt="Imagen de <%=EscaparHTML(artistas(i)("nombre"))%>" loading="lazy" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
          </div>
          <h3 class="tarjeta-nombre"><%=EscaparHTML(artistas(i)("nombre"))%></h3>
          <p class="tarjeta-meta">Artista</p>
        </article>
      <% Next %>
    </div>
  <% End If %>
</section>

<!-- ============ Listas de reproduccion ============ -->
<section class="seccion" id="listas">
  <div class="seccion-cabecera">
    <h2 class="seccion-titulo">Listas de reproduccion</h2>
    <span class="seccion-info">Tus listas &middot; <%=nListas%></span>
  </div>
  <% If nListas = 0 Then %>
    <div class="estado-vacio">
      <%=Icono("lista")%>
      <strong>No tienes listas de reproduccion</strong>
      <span>Crea tus primeras listas para organizar tus canciones favoritas.</span>
    </div>
  <% Else %>
    <div class="rejilla-tarjetas">
      <% For i = 0 To nListas - 1 %>
        <article class="tarjeta">
          <div class="tarjeta-portada">
            <img src="<%=listas(i)("foto")%>" alt="Portada de <%=EscaparHTML(listas(i)("titulo"))%>" loading="lazy" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
          </div>
          <h3 class="tarjeta-nombre"><%=EscaparHTML(listas(i)("titulo"))%></h3>
          <p class="tarjeta-meta">Lista &middot; <b><%=EscaparHTML(listas(i)("tipo"))%></b></p>
        </article>
      <% Next %>
    </div>
  <% End If %>
</section>

<!-- ============ Reproductor ============ -->
<!--#include file="includes/reproductor.asp"-->

<script>
window.PM_BASE = "<%=RUTA_BASE%>";
window.PM_COLA = [
  <% Dim j, coma %>
  <% For j = 0 To nCanciones - 1 %>
    <% If j < nCanciones - 1 Then coma = "," Else coma = "" %>
    { id: <%=canciones(j)("id")%>, titulo: "<%=EscaparJS(canciones(j)("nombre"))%>", artista: "<%=EscaparJS(canciones(j)("artistas"))%>", portada: "<%=EscaparJS(canciones(j)("portada"))%>", src: "<%=EscaparJS(canciones(j)("src"))%>" }<%=coma%>
  <% Next %>
];

/* Filtro del listado de canciones */
(function () {
  var filtro = document.getElementById("filtro-canciones");
  if (!filtro) return;
  filtro.addEventListener("input", function () {
    var q = filtro.value.toLowerCase();
    var filas = document.querySelectorAll(".fila-cancion");
    for (var i = 0; i < filas.length; i++) {
      filas[i].style.display = filas[i].textContent.toLowerCase().indexOf(q) >= 0 ? "" : "none";
    }
  });
})();
</script>
<!--#include file="includes/footer.asp"-->
