<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - test_diagnostico.asp
'  Pagina de diagnostico (NO requiere sesion). Verifica ASP, sesion,
'  FSO, proveedores ADO y la conexion a BDD_PoliMusic.
'  Abrir como: http://localhost/webapppolimusicv1_asp/test_diagnostico.asp
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
Dim numPrueba, numOk, numFail
numPrueba = 0 : numOk = 0 : numFail = 0

Sub Prueba(ByVal titulo, ByVal ok, ByVal detalle)
    numPrueba = numPrueba + 1
    If ok Then numOk = numOk + 1 Else numFail = numFail + 1
    %>
    <tr>
      <td><% If ok Then %><span style="color:var(--exito)"><%=Icono("check")%></span><% Else %><span style="color:var(--peligro)"><%=Icono("alerta")%></span><% End If %></td>
      <td><b><%=EscaparHTML(titulo)%></b></td>
      <td><%=detalle%></td>
    </tr>
    <%
End Sub
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Diagnostico - PoliMusic</title>
<link rel="stylesheet" href="<%=RUTA_BASE%>css/styles.css">
</head>
<body class="sin-reproductor">
<div class="fondo-degradado" aria-hidden="true"></div>
<main class="contenido" style="max-width:920px">

  <div class="pagina-cabecera" style="margin-bottom:22px">
    <div>
      <h1 class="pagina-titulo">Diagnostico de PoliMusic</h1>
      <p class="pagina-sub">Verifica que ASP, la sesion, las rutas de archivos y la conexion a SQL Server funcionen.</p>
    </div>
  </div>

  <div class="aviso aviso-info">
    <span><%=Icono("buscar")%></span>
    <span>Esta pagina no requiere inicio de sesion. Sirve para detectar que falla en el despliegue.</span>
  </div>

  <section class="panel">
    <h2 class="panel-titulo"><%=Icono("check")%> Resultados</h2>
    <div class="tabla-envoltura">
      <table class="tabla">
        <thead>
          <tr><th>Estado</th><th>Prueba</th><th>Detalle</th></tr>
        </thead>
        <tbody>
          <%
          '--------- 1. ASP funcionando ----------
          Prueba "ASP clasico en ejecucion", True, "Servidor: " & EscaparHTML(Request.ServerVariables("SERVER_SOFTWARE") & "") & " &middot; " & EscaparHTML(Request.ServerVariables("SCRIPT_NAME") & "")

          '--------- 2. Sesion ----------
          Dim sesionOk, sesionDetalle
          sesionOk = False : sesionDetalle = ""
          On Error Resume Next
          Session("PM_TEST") = "ok"
          If Err.Number <> 0 Then
            sesionDetalle = "Error de sesion: " & EscaparHTML(Err.Description)
            Err.Clear
          ElseIf CStr(Session("PM_TEST") & "") = "ok" Then
            sesionOk = True
            sesionDetalle = "Lectura/escritura de Session correcta (login funcional)"
          Else
            sesionDetalle = "La sesion no conserva valores"
          End If
          On Error GoTo 0
          Prueba "Estado de sesion (Session)", sesionOk, sesionDetalle

          '--------- 3. FSO y carpeta de recursos ----------
          Dim fsoT, rutaFisica, mp3Ok
          Set fsoT = Server.CreateObject("Scripting.FileSystemObject")
          rutaFisica = Server.MapPath("/SongFilesAndImages_resources")
          Prueba "Carpeta SongFilesAndImages_resources", fsoT.FolderExists(rutaFisica), EscaparHTML(rutaFisica)
          mp3Ok = fsoT.FileExists(Server.MapPath("/SongFilesAndImages_resources/SongFiles/bensound-adventure.mp3"))
          Prueba "MP3 de prueba (bensound-adventure.mp3)", mp3Ok, IIf(mp3Ok, "Archivo encontrado", "No se encuentra el MP3")
          Set fsoT = Nothing

          '--------- 4. Proveedores ADO ----------
          Dim provList, i, provOk, provDetalle, connT, cadenaT
          provList = Array("SQLNCLI11", "MSOLEDBSQL", "SQLOLEDB")
          provDetalle = ""
          provOk = False
          For i = 0 To UBound(provList)
            Set connT = Server.CreateObject("ADODB.Connection")
            cadenaT = "Provider=" & provList(i) & ";Server=" & DB_SERVIDOR & ";Database=" & DB_BASE & ";Uid=" & DB_USUARIO & ";Pwd=" & DB_CLAVE & ";"
            On Error Resume Next
            Err.Clear
            connT.Open cadenaT
            If Err.Number = 0 Then
              provOk = True
              provDetalle = "Funciona con el proveedor <b>" & provList(i) & "</b>"
              On Error GoTo 0
              connT.Close
              Exit For
            End If
            Dim errProv
            errProv = Err.Description
            On Error GoTo 0
            If provDetalle <> "" Then provDetalle = provDetalle & "<br>"
            provDetalle = provDetalle & provList(i) & ": " & EscaparHTML(errProv)
            Set connT = Nothing
          Next
          Prueba "Proveedor OLE DB para SQL Server", provOk, provDetalle

          '--------- 5. Consulta de prueba ----------
          Dim consultaOk, consultaDetalle, cmdT, rsT
          consultaOk = False : consultaDetalle = "Requiere un proveedor ADO funcionando (prueba anterior)"
          If provOk Then
            On Error Resume Next
            AbrirConexion()
            If Err.Number <> 0 Then
              Err.Clear
            ElseIf IsObject(CONN) Then
              If CONN.State = adStateOpen Then
                Set cmdT = NuevoComando("SELECT (SELECT COUNT(*) FROM TBL_SONG) AS N_CANCIONES, (SELECT COUNT(*) FROM TBL_USER) AS N_USUARIOS, (SELECT COUNT(*) FROM TBL_ALBUM) AS N_ALBUMES")
                Set rsT = cmdT.Execute
                If Err.Number <> 0 Then
                  consultaDetalle = "No se pudo consultar. Se ejecuto el script BDD_PoliMusic.sql? Detalle: " & EscaparHTML(Err.Description)
                  Err.Clear
                Else
                  consultaOk = True
                  consultaDetalle = "BDD_PoliMusic OK &middot; canciones: " & (rsT("N_CANCIONES") & "") & " &middot; usuarios: " & (rsT("N_USUARIOS") & "") & " &middot; albumes: " & (rsT("N_ALBUMES") & "")
                End If
                rsT.Close
                Set rsT = Nothing
                Set cmdT = Nothing
              End If
            End If
            On Error GoTo 0
            If IsObject(CONN) Then
              If CONN.State = adStateOpen Then CONN.Close
              Set CONN = Nothing
            End If
          End If
          Prueba "Conexion y consulta a BDD_PoliMusic", consultaOk, consultaDetalle

          '--------- 6. Ruta base ----------
          Prueba "Ruta base de la aplicacion (RUTA_BASE)", (RUTA_BASE = "/"), "Valor: <b>" & RUTA_BASE & "</b> &middot; Si el sitio NO esta en la raiz de IIS, cambiar en includes/conexion.asp"
          %>
        </tbody>
      </table>
    </div>

    <div style="margin-top:18px">
      <% If numFail = 0 Then %>
        <div class="aviso aviso-exito"><span><%=Icono("check")%></span><span>Todo correcto (<%=numOk%>/<%=numPrueba%>). Abre <b><%=RUTA_BASE%>login.asp</b> e ingresa con <b>userAdmin</b> / <b>12345678</b>.</span></div>
      <% Else %>
        <div class="aviso aviso-error"><span><%=Icono("alerta")%></span><span><b><%=numFail%> de <%=numPrueba%></b> pruebas fallaron. Revisa el detalle arriba y la seccion 7 del README. Si todo aparece correcto pero el error continua, abre el <b>Visor de sucesos</b> (eventvwr.msc) → Registros de Windows → Aplicacion → fuente <b>Active Server Pages</b>: ahi aparece el codigo exacto del error y la linea.</span></div>
      <% End If %>
    </div>
  </section>

  <p style="margin-top:18px"><a class="boton boton-fantasma" href="<%=RUTA_BASE%>login.asp">Ir al inicio de sesion</a></p>
</main>
<script src="<%=RUTA_BASE%>js/app.js"></script>
</body>
</html>
<%
CerrarConexion()
%>
