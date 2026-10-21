/* Reemplazo de imagenes que no cargan por una imagen por defecto. */
(function () {
  "use strict";
  var base = window.PM_BASE || "/";
  document.addEventListener("error", function (ev) {
    var el = ev.target;
    if (el && el.tagName === "IMG" && !el.getAttribute("data-fb")) {
      el.setAttribute("data-fb", "1");
      el.src = base + "SongFilesAndImages_resources/PlaylistImages/default.jpg";
    }
  }, true);
})();
