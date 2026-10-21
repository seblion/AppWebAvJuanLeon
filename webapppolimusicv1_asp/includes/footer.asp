</main>

<footer class="pie">
  <p>PoliMusic &mdash; Clon tipo Spotify desarrollado con ASP clasico, HTML5, CSS3 y JavaScript puro.</p>
  <p>Base de datos BDD_PoliMusic en SQL Server (localhost) &middot; Escuela Politecnica Nacional</p>
</footer>

<script src="<%=RUTA_BASE%>js/app.js"></script>
<% If UsarPlayer Then %>
<script src="<%=RUTA_BASE%>js/player.js"></script>
<% End If %>
</body>
</html>
<%
' Cierra la conexion abierta en AbrirConexion()
CerrarConexion()
%>
