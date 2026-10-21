/* ============================================================
   PoliMusic - Reproductor (JavaScript puro)
   Lee window.PM_COLA (generado por index.asp desde SQL Server)
   ============================================================ */
(function () {
  "use strict";

  var cola = Array.isArray(window.PM_COLA) ? window.PM_COLA : [];
  var base = window.PM_BASE || "/";

  var audio = document.getElementById("audio");
  var contenedor = document.getElementById("reproductor");
  if (!audio || !contenedor) return;

  function el(id) { return document.getElementById(id); }
  var portada = el("rep-portada"),
      tituloTxt = el("rep-titulo"),
      artistaTxt = el("rep-artista"),
      btnPlay = el("btn-play"),
      btnPrev = el("btn-prev"),
      btnNext = el("btn-next"),
      barra = el("rep-barra"),
      tiempoTxt = el("rep-tiempo"),
      duracionTxt = el("rep-duracion"),
      volumen = el("rep-volumen"),
      estadoTxt = el("rep-estado");

  var SVG_PLAY = '<svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor" aria-hidden="true"><path d="M8 5.2v13.6c0 .8.9 1.3 1.6.9l10.4-6.8c.6-.4.6-1.4 0-1.8L9.6 4.3C8.9 3.9 8 4.4 8 5.2z"/></svg>';
  var SVG_PAUSA = '<svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor" aria-hidden="true"><rect x="6" y="5" width="4" height="14" rx="1.2"/><rect x="14" y="5" width="4" height="14" rx="1.2"/></svg>';

  var indice = -1;
  var buscando = false; /* el usuario esta arrastrando la barra de progreso */

  function formatear(seg) {
    if (!isFinite(seg) || seg < 0) seg = 0;
    var m = Math.floor(seg / 60);
    var s = Math.floor(seg % 60);
    return m + ":" + (s < 10 ? "0" : "") + s;
  }

  function buscarIndice(id) {
    for (var i = 0; i < cola.length; i++) {
      if (String(cola[i].id) === String(id)) return i;
    }
    return -1;
  }

  function actualizarBoton() {
    btnPlay.innerHTML = audio.paused ? SVG_PLAY : SVG_PAUSA;
    var etiqueta = audio.paused ? "Reproducir" : "Pausar";
    btnPlay.setAttribute("aria-label", etiqueta);
    btnPlay.setAttribute("title", etiqueta);
    contenedor.classList.toggle("reproduciendo", !audio.paused);
    var activa = document.querySelector(".fila-cancion.activa");
    if (activa) activa.classList.toggle("sonando", !audio.paused);
  }

  function marcarActiva() {
    var filas = document.querySelectorAll(".fila-cancion");
    for (var i = 0; i < filas.length; i++) {
      var coincide = cola[indice] && String(cola[indice].id) === filas[i].getAttribute("data-id");
      if (coincide) filas[i].classList.add("activa");
      else filas[i].classList.remove("activa", "sonando");
    }
    actualizarBoton();
  }

  /* Incrementa el contador PLAYS en el servidor */
  function registrarReproduccion(cancion) {
    if (!cancion || !window.fetch) return;
    fetch(base + "api_reproducir.asp?id=" + encodeURIComponent(cancion.id))
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d || !d.ok) return;
        var celda = document.querySelector('.fila-cancion[data-id="' + d.id + '"] .fila-plays');
        if (celda) celda.textContent = d.plays + " reproducciones";
      })
      .catch(function () { /* silencioso: el contador es informativo */ });
  }

  function cargar(i, reproducir) {
    if (cola.length === 0) return;
    if (i < 0) i = cola.length - 1;
    if (i >= cola.length) i = 0;
    indice = i;
    var c = cola[i];
    audio.src = c.src;
    portada.src = c.portada;
    portada.alt = "Portada de " + c.titulo;
    tituloTxt.textContent = c.titulo;
    artistaTxt.textContent = c.artista;
    marcarActiva();
    registrarReproduccion(c);
    if (reproducir !== false) {
      var p = audio.play();
      if (p && p.catch) p.catch(function () { actualizarBoton(); });
    }
  }

  function alternar() {
    if (indice < 0) { cargar(0, true); return; }
    if (audio.paused) {
      var p = audio.play();
      if (p && p.catch) p.catch(function () { actualizarBoton(); });
    } else {
      audio.pause();
    }
  }

  function siguiente() {
    if (cola.length === 0) return;
    cargar(indice + 1 > cola.length - 1 ? 0 : indice + 1, true);
  }

  function anterior() {
    if (cola.length === 0) return;
    if (audio.currentTime > 3) { audio.currentTime = 0; return; }
    cargar(indice - 1 < 0 ? cola.length - 1 : indice - 1, true);
  }

  /* ---- Sin canciones: reproductor deshabilitado ---- */
  if (cola.length === 0) {
    contenedor.classList.add("vacio");
    if (estadoTxt) estadoTxt.textContent = "Sin canciones";
    btnPlay.disabled = true;
    btnNext.disabled = true;
    btnPrev.disabled = true;
    duracionTxt.textContent = "0:00";
    return;
  }

  /* ---- Estado inicial ---- */
  var vol = 0.8;
  try {
    vol = parseFloat(localStorage.getItem("pm_volumen"));
    if (isNaN(vol)) vol = 0.8;
  } catch (e) { vol = 0.8; }
  audio.volume = vol;
  volumen.value = Math.round(vol * 100);
  if (estadoTxt) estadoTxt.textContent = cola.length + " canciones en cola";
  duracionTxt.textContent = "0:00";

  btnPlay.addEventListener("click", alternar);
  btnNext.addEventListener("click", siguiente);
  btnPrev.addEventListener("click", anterior);

  audio.addEventListener("play", actualizarBoton);
  audio.addEventListener("pause", actualizarBoton);
  audio.addEventListener("ended", siguiente);

  audio.addEventListener("loadedmetadata", function () {
    duracionTxt.textContent = formatear(audio.duration);
    barra.max = isFinite(audio.duration) && audio.duration > 0 ? audio.duration : 0;
  });

  audio.addEventListener("timeupdate", function () {
    if (buscando) return;
    barra.value = audio.currentTime;
    tiempoTxt.textContent = formatear(audio.currentTime);
  });

  audio.addEventListener("error", function () {
    if (estadoTxt && indice >= 0) estadoTxt.textContent = "No se pudo cargar el audio";
  });

  barra.addEventListener("input", function () {
    buscando = true;
    tiempoTxt.textContent = formatear(parseFloat(barra.value));
  });
  barra.addEventListener("change", function () {
    if (isFinite(audio.duration)) audio.currentTime = parseFloat(barra.value);
    buscando = false;
  });

  volumen.addEventListener("input", function () {
    audio.volume = parseFloat(volumen.value) / 100;
    try { localStorage.setItem("pm_volumen", String(audio.volume)); } catch (e) { /* ignorar */ }
  });

  /* Clic sobre una fila del listado: reproducir o pausar */
  var lista = document.querySelector(".lista-canciones");
  if (lista) {
    lista.addEventListener("click", function (ev) {
      var fila = ev.target && ev.target.closest ? ev.target.closest(".fila-cancion") : null;
      if (!fila) return;
      var i = buscarIndice(fila.getAttribute("data-id"));
      if (i < 0) return;
      if (i === indice) alternar();
      else cargar(i, true);
    });
  }

  /* Clic en el boton de una tarjeta de album/artista/lista */
  var tarjetas = document.querySelectorAll(".tarjeta-play");
  for (var t = 0; t < tarjetas.length; t++) {
    tarjetas[t].addEventListener("click", function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      var i = buscarIndice(this.getAttribute("data-id"));
      if (i >= 0) cargar(i, true);
    });
  }

  /* Atajos de teclado: espacio reproduce/pausa, flechas buscan */
  document.addEventListener("keydown", function (ev) {
    var t = ev.target;
    if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT")) return;
    if (ev.code === "Space") { ev.preventDefault(); alternar(); }
    else if (ev.code === "ArrowRight" && indice >= 0 && isFinite(audio.duration)) {
      audio.currentTime = Math.min(audio.currentTime + 5, audio.duration);
    } else if (ev.code === "ArrowLeft" && indice >= 0) {
      audio.currentTime = Math.max(audio.currentTime - 5, 0);
    }
  });

  actualizarBoton();
})();
