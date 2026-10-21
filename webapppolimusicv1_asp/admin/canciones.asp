<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - admin/canciones.asp
'  CRUD de canciones (TBL_SONG). SOLO rol admin (USER_TYPE = 1).
'=====================================================================
Option Explicit
%>
<!--#include virtual="/webapppolimusicv1_asp/includes/conexion.asp"-->
<%
'--------- Restriccion estricta: solo admin ----------
RequiereAdmin()

Dim mensaje, mensajeTipo
mensaje = ""
mensajeTipo = "exito"

AbrirConexion()
TituloPagina = "Administrar canciones"
ClaseBody    = "pagina-admin"

'---------------------------------------------------------------------
'  Eliminar cancion: se quita primero de las tablas dependientes
'  (TBL_PLAYLIST_SONG, TBL_TRACKLIST, TBL_PLAYER) y luego de TBL_SONG.
'  Todo dentro de una transaccion.
'---------------------------------------------------------------------
If Request.Form("accion") = "eliminar" Then
    Dim idElim, errDesc
    idElim = CLng(Request.Form("hdnIDEliminar"))
    On Error Resume Next
    CONN.BeginTrans
    Dim d1, d2, d3, d4
    Set d1 = NuevoComando("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?")
    ParEntero d1, "@id", idElim
    d1.Execute
    Set d2 = NuevoComando("DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?")
    ParEntero d2, "@id", idElim
    d2.Execute
    Set d3 = NuevoComando("DELETE FROM TBL_PLAYER WHERE ID_SONG = ?")
    ParEntero d3, "@id", idElim
    d3.Execute
    Set d4 = NuevoComando("DELETE FROM TBL_SONG WHERE ID_SONG = ?")
    ParEntero d4, "@id", idElim
    d4.Execute
    If Err.Number <> 0 Then
        errDesc = Err.Description
        Err.Clear
        CONN.RollbackTrans
        On Error GoTo 0
        Response.Redirect "canciones.asp?msg=" & Server.URLEncode("No se pudo eliminar la cancion. " & errDesc) & "&tipo=error"
    Else
        CONN.CommitTrans
        On Error GoTo 0
        Response.Redirect "canciones.asp?msg=" & Server.URLEncode("Cancion eliminada correctamente.") & "&tipo=exito"
    End If
End If

'---------------------------------------------------------------------
'  Guardar (insertar o actualizar)
'---------------------------------------------------------------------
Dim vID, vNombre, vRuta, vPlays
vID = "" : vNombre = "" : vRuta = "" : vPlays = "0"
Dim errorForm
errorForm = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    vID     = Trim(Request.Form("hdnID") & "")
    vNombre = Trim(Request.Form("txtNombre") & "")
    vRuta   = Trim(Request.Form("txtRuta") & "")
    vPlays  = Trim(Request.Form("txtPlays") & "")

    If Len(vNombre) = 0 Or Len(vNombre) > 50 Then
        errorForm = "El nombre de la cancion es obligatorio (maximo 50 caracteres)."
    ElseIf Len(vRuta) = 0 Or Len(vRuta) > 255 Then
        errorForm = "La ruta del archivo de audio es obligatoria (maximo 255 caracteres)."
    ElseIf Not IsNumeric(vPlays) Then
        errorForm = "Las reproducciones deben ser un numero entero mayor o igual a cero."
    ElseIf CLng(vPlays) < 0 Then
        errorForm = "Las reproducciones deben ser un numero entero mayor o igual a cero."
    End If

    If errorForm = "" Then
        On Error Resume Next
        If vID = "" Then
            Dim cmdI
            Set cmdI = NuevoComando("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?)")
            ParTexto cmdI, "@nombre", vNombre, 50
            ParTexto cmdI, "@ruta", vRuta, 255
            ParEntero cmdI, "@plays", CLng(vPlays)
            cmdI.Execute
        Else
            Dim cmdU
            Set cmdU = NuevoComando("UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?")
            ParTexto cmdU, "@nombre", vNombre, 50
            ParTexto cmdU, "@ruta", vRuta, 255
            ParEntero cmdU, "@plays", CLng(vPlays)
            ParEntero cmdU, "@id", CLng(vID)
            cmdU.Execute
        End If
        If Err.Number <> 0 Then
            errDesc = Err.Description
            Err.Clear
            On Error GoTo 0
            errorForm = "Error al guardar en la base de datos: " & errDesc
        Else
            On Error GoTo 0
            Response.Redirect "canciones.asp?msg=" & Server.URLEncode("Cancion guardada correctamente.") & "&tipo=exito"
        End If
    End If
'--------- Cargar datos para edicion ---------
ElseIf Request.QueryString("editar") <> "" And IsNumeric(Request.QueryString("editar")) Then
    vID = CLng(Request.QueryString("editar"))
    Dim cmdE, rsE
    Set cmdE = NuevoComando("SELECT SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = ?")
    ParEntero cmdE, "@id", vID
    Set rsE = cmdE.Execute
    If Not rsE.EOF Then
        vNombre = CStr(rsE("SONG_NAME") & "")
        vRuta   = CStr(rsE("SONG_PATH") & "")
        If IsNull(rsE("PLAYS")) Or IsEmpty(rsE("PLAYS")) Then vPlays = "0" Else vPlays = CStr(rsE("PLAYS"))
    Else
        vID = ""
    End If
    rsE.Close
    Set rsE = Nothing
    Set cmdE = Nothing
End If

'--------- Mensajes desde redireccion (PRG) ---------
If Request.QueryString("msg") <> "" Then
    mensaje = Request.QueryString("msg")
    If Request.QueryString("tipo") = "error" Then mensajeTipo = "error" Else mensajeTipo = "exito"
End If

'---------------------------------------------------------------------
'  Listado de canciones con sus artistas
'---------------------------------------------------------------------
Dim cmdL, rsL
Set cmdL = NuevoComando("SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG ORDER BY ID_SONG")
Set rsL = cmdL.Execute
%>
<!--#include virtual="/webapppolimusicv1_asp/includes/header.asp"-->
<div class="pagina-cabecera">
  <div>
    <h1 class="pagina-titulo">Administrar canciones</h1>
    <p class="pagina-sub">Alta, baja y modificacion de canciones (TBL_SONG). Seccion exclusiva del rol admin.</p>
  </div>
</div>
<nav class="admin-subnav" aria-label="Secciones de administracion">
  <a href="canciones.asp" class="activo">Canciones</a>
  <a href="artistas.asp">Artistas</a>
  <a href="albumes.asp">Albumes</a>
</nav>

<% If mensaje <> "" Then %>
<div class="aviso aviso-<%=mensajeTipo%>"><span><% If mensajeTipo = "exito" Then %><%=Icono("check")%><% Else %><%=Icono("alerta")%><% End If %></span><span><%=EscaparHTML(mensaje)%></span></div>
<% End If %>
<% If errorForm <> "" Then %>
<div class="aviso aviso-error"><span><%=Icono("alerta")%></span><span><%=EscaparHTML(errorForm)%></span></div>
<% End If %>

<div class="admin-grid">
  <!-- ======== Formulario ======== -->
  <section class="panel">
    <h2 class="panel-titulo"><%=Icono("mas")%> <% If vID = "" Then %>Nueva cancion<% Else %>Editar cancion<% End If %></h2>
    <form method="post" action="canciones.asp">
      <input type="hidden" name="hdnID" value="<%=vID%>">
      <div class="campo">
        <label for="txtNombre">Nombre de la cancion</label>
        <input type="text" id="txtNombre" name="txtNombre" maxlength="50" required value="<%=EscaparHTML(vNombre)%>" placeholder="Ej: Adventure">
      </div>
      <div class="campo">
        <label for="txtRuta">Ruta del audio</label>
        <input type="text" id="txtRuta" name="txtRuta" maxlength="255" required value="<%=EscaparHTML(vRuta)%>" placeholder="~/songFiles/bensound-adventure.mp3">
        <span class="campo-ayuda">Los archivos deben existir en <b>SongFilesAndImages_resources/SongFiles</b>. Si no existen se reproducira bensound-adventure.mp3.</span>
      </div>
      <div class="campo">
        <label for="txtPlays">Reproducciones</label>
        <input type="number" id="txtPlays" name="txtPlays" min="0" step="1" value="<%=vPlays%>">
      </div>
      <div class="form-acciones">
        <button type="submit" class="boton boton-primario">Guardar</button>
        <a class="boton boton-fantasma" href="canciones.asp">Cancelar</a>
      </div>
    </form>
  </section>

  <!-- ======== Tabla ======== -->
  <section class="panel">
    <h2 class="panel-titulo"><%=Icono("nota")%> Canciones registradas</h2>
    <div class="tabla-envoltura">
      <table class="tabla">
        <thead>
          <tr>
            <th>ID</th>
            <th>Nombre</th>
            <th>Ruta</th>
            <th>Artistas</th>
            <th>Reprod.</th>
            <th>Acciones</th>
          </tr>
        </thead>
        <tbody>
          <% If rsL.EOF Then %>
            <tr><td colspan="6">
              <div class="estado-vacio">
                <%=Icono("nota")%>
                <strong>No hay canciones registradas</strong>
                <span>Usa el formulario para crear la primera cancion.</span>
              </div>
            </td></tr>
          <% End If %>
          <% Do While Not rsL.EOF %>
            <tr>
              <td><%=rsL("ID_SONG")%></td>
              <td><b><%=EscaparHTML(rsL("SONG_NAME"))%></b></td>
              <td class="celda-ruta"><%=EscaparHTML(rsL("SONG_PATH"))%></td>
              <td><%=EscaparHTML(ArtistasDeCancion(rsL("ID_SONG")))%></td>
              <td><% If IsNull(rsL("PLAYS")) Or IsEmpty(rsL("PLAYS")) Then %>0<% Else %><%=rsL("PLAYS")%><% End If %></td>
              <td>
                <div class="celda-acciones">
                  <a class="boton-icono" href="canciones.asp?editar=<%=rsL("ID_SONG")%>" title="Editar" aria-label="Editar"><%=Icono("editar")%></a>
                  <form method="post" action="canciones.asp" class="form-eliminar" onsubmit="return confirm('Se eliminara la cancion y sus referencias en listas, tracklists y artistas. Continuar?');">
                    <input type="hidden" name="accion" value="eliminar">
                    <input type="hidden" name="hdnIDEliminar" value="<%=rsL("ID_SONG")%>">
                    <button type="submit" class="boton-icono boton-peligro" title="Eliminar" aria-label="Eliminar"><%=Icono("basura")%></button>
                  </form>
                </div>
              </td>
            </tr>
            <% rsL.MoveNext %>
          <% Loop %>
        </tbody>
      </table>
    </div>
  </section>
</div>
<!--#include virtual="/webapppolimusicv1_asp/includes/footer.asp"-->