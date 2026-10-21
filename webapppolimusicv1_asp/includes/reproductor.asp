<%
'=====================================================================
'  PoliMusic - includes/reproductor.asp
'  Barra del reproductor fija en la parte inferior (solo index.asp)
'=====================================================================
%>
<div class="reproductor" id="reproductor" aria-label="Reproductor de musica">
  <audio id="audio" preload="metadata"></audio>

  <div class="rep-izq">
    <img class="rep-portada" id="rep-portada" src="<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg" alt="" onerror="this.onerror=null;this.src='<%=RUTA_BASE%>SongFilesAndImages_resources/PlaylistImages/default.jpg';">
    <div class="rep-info">
      <span class="rep-titulo" id="rep-titulo">PoliMusic</span>
      <span class="rep-artista" id="rep-artista">Selecciona una cancion</span>
    </div>
  </div>

  <div class="rep-centro">
    <div class="rep-controles">
      <button class="boton-control" id="btn-prev" type="button" title="Anterior" aria-label="Cancion anterior"><%=Icono("anterior")%></button>
      <button class="boton-play" id="btn-play" type="button" title="Reproducir" aria-label="Reproducir o pausar"><%=Icono("play")%></button>
      <button class="boton-control" id="btn-next" type="button" title="Siguiente" aria-label="Siguiente cancion"><%=Icono("siguiente")%></button>
    </div>
    <div class="rep-progreso">
      <span class="rep-tiempo" id="rep-tiempo">0:00</span>
      <input type="range" id="rep-barra" min="0" max="100" value="0" step="0.1" aria-label="Progreso de la cancion">
      <span class="rep-tiempo" id="rep-duracion">0:00</span>
    </div>
  </div>

  <div class="rep-der">
    <span class="rep-estado" id="rep-estado"></span>
    <span class="rep-volumen-icono" aria-hidden="true"><%=Icono("volumen")%></span>
    <input type="range" id="rep-volumen" min="0" max="100" value="80" aria-label="Volumen">
  </div>
</div>
