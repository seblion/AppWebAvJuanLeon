<%
'=====================================================================
'  PoliMusic - includes/conexion.asp
'  Configuracion, conexion a SQL Server y funciones utilitarias.
'  Base de datos: BDD_PoliMusic (localhost) - usuario usr_polimusic
'
'  IMPORTANTE: este archivo es un INCLUDE. NO debe contener directivas @
'  ni "Option Explicit": al fusionarse con la pagina principal quedarian
'  duplicados y el motor de ASP falla al compilar (error HTTP 500
'  generico en todas las paginas que lo incluyen).
'  Option Explicit se declara una sola vez en cada pagina principal.
'  TAMPOCO escribir los delimitadores de script de ASP en comentarios
'  ni en cadenas: el preprocesador cerraria el bloque de codigo.
'=====================================================================

'---------------- Configuracion de la base de datos ----------------
' AbrirConexion() prueba en orden: SQLNCLI11, MSOLEDBSQL, SQLOLEDB.
' Si se necesita forzar uno, editar el Array dentro de AbrirConexion().
Const DB_PROVEEDOR = "SQLNCLI11"          ' Proveedor preferido (informativo)
Const DB_SERVIDOR  = "localhost\SQLSERVER2026"
Const DB_BASE      = "BDD_PoliMusic"
Const DB_USUARIO   = "usr_polimusic"
Const DB_CLAVE     = "Politecnica1"

'---------------- Constantes ADO ----------------
Const adCmdText    = 1
Const adVarChar    = 200
Const adInteger    = 3
Const adParamInput = 1
Const adStateOpen  = 1

'---------------- Variables globales ----------------
Dim CONN
Dim CATALOGO
Dim TituloPagina
Dim ClaseBody
Dim UsarPlayer
Dim RUTA_BASE

TituloPagina = "PoliMusic"
ClaseBody    = ""
UsarPlayer   = False

' Ruta raiz de la aplicacion (para enlaces, CSS, imagenes y includes).
' Se detecta automaticamente si el sitio esta en la raiz de IIS ("/") o
' un nivel debajo (ej: "/webapppolimusicv1_asp/"). Si se publica mas
' profundo, escribir la ruta manualmente aqui.
Dim scriptName, segundaBarra
scriptName = Request.ServerVariables("SCRIPT_NAME")
segundaBarra = InStr(2, scriptName, "/")
If segundaBarra > 0 Then
    RUTA_BASE = Left(scriptName, segundaBarra)
Else
    RUTA_BASE = "/"
End If

'---------------- Ajustes de respuesta ----------------
Response.Buffer = True
On Error Resume Next
Session.CodePage = 65001
Err.Clear
On Error GoTo 0
Response.CodePage = 65001
Response.Charset = "UTF-8"

'---------------------------------------------------------------------
'  Abre la conexion con SQL Server (una por pagina).
'  Si falla muestra una pagina de diagnostico.
'---------------------------------------------------------------------
Sub AbrirConexion()
    Dim proveedores, i, cadena, intentos
    ' Se prueban varios proveedores en orden: SQLNCLI11 (SQL Server 2012-2017),
    ' MSOLEDBSQL (cliente actual) y SQLOLEDB (MDAC, siempre presente en Windows).
    proveedores = Array("SQLNCLI11", "MSOLEDBSQL", "SQLOLEDB")
    intentos = ""
    On Error Resume Next
    Set CONN = Server.CreateObject("ADODB.Connection")
    For i = 0 To UBound(proveedores)
        Err.Clear
        cadena = "Provider=" & proveedores(i) & ";Server=" & DB_SERVIDOR & _
                 ";Database=" & DB_BASE & ";Uid=" & DB_USUARIO & ";Pwd=" & DB_CLAVE & ";"
        CONN.Open cadena
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        If intentos <> "" Then intentos = intentos & "<br>"
        intentos = intentos & "<b>" & proveedores(i) & "</b>: " & EscaparHTML(Err.Description)
    Next
    On Error GoTo 0
    Response.Clear
    %>
        <!DOCTYPE html>
        <html lang="es">
        <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Error de conexion - PoliMusic</title>
        <link rel="stylesheet" href="<%=RUTA_BASE%>css/styles.css">
        </head>
        <body class="sin-reproductor">
        <div class="fondo-degradado" aria-hidden="true"></div>
        <main class="contenido" style="max-width:760px">
          <div class="aviso aviso-error" style="margin-top:40px">
            <span>No se pudo conectar con la base de datos <b>BDD_PoliMusic</b>.</span>
          </div>
          <div class="panel">
            <div class="panel-titulo">Verifica lo siguiente</div>
            <ul style="margin-left:20px;color:var(--texto-2);font-size:14px;line-height:2">
              <li>Que el servicio <b>SQL Server</b> este ejecutandose en <b>localhost</b>.</li>
              <li>Que el protocolo <b>TCP/IP</b> este habilitado (SQL Server Configuration Manager).</li>
              <li>Que SQL Server este en modo de <b>autenticacion mixta</b> (SQL Server + Windows).</li>
              <li>Que se haya ejecutado el script <b>BDD_PoliMusic.sql</b> (crea login, base y tablas).</li>
              <li>Que el proveedor <b>SQLNCLI11</b> (SQL Server Native Client 11.0), <b>MSOLEDBSQL</b> o <b>SQLOLEDB</b> este instalado.</li>
            </ul>
            <p style="margin-top:16px;color:var(--texto-3);font-size:13px;line-height:1.8">Detalle tecnico de cada intento:<br><%=intentos%></p>
          </div>
        </main>
        </body>
        </html>
        <%
        Response.End
End Sub

'---------------------------------------------------------------------
'  Cierra la conexion abierta
'---------------------------------------------------------------------
Sub CerrarConexion()
    On Error Resume Next
    If IsObject(CONN) Then
        If CONN.State = adStateOpen Then CONN.Close
    End If
    Set CONN = Nothing
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------
'  Sesion
'---------------------------------------------------------------------
Function SesionIniciada()
    SesionIniciada = False
    If IsEmpty(Session("ID_USER")) Then Exit Function
    If IsNull(Session("ID_USER")) Then Exit Function
    If CStr(Session("ID_USER")) = "" Then Exit Function
    SesionIniciada = True
End Function

Function EsAdmin()
    EsAdmin = False
    If Not SesionIniciada() Then Exit Function
    If IsEmpty(Session("USER_TYPE")) Or IsNull(Session("USER_TYPE")) Then Exit Function
    EsAdmin = (CInt(Session("USER_TYPE")) = 1)
End Function

Sub RequiereLogin()
    If Not SesionIniciada() Then
        Response.Redirect RUTA_BASE & "login.asp"
    End If
End Sub

Sub RequiereAdmin()
    RequiereLogin()
    If Not EsAdmin() Then
        Response.Redirect RUTA_BASE & "index.asp?error=permiso"
    End If
End Sub

'---------------------------------------------------------------------
'  Utilidades de texto
'---------------------------------------------------------------------
Function EscaparHTML(ByVal texto)
    If IsNull(texto) Then EscaparHTML = "" : Exit Function
    If IsEmpty(texto) Then EscaparHTML = "" : Exit Function
    texto = CStr(texto)
    texto = Replace(texto, "&", "&amp;")
    texto = Replace(texto, "<", "&lt;")
    texto = Replace(texto, ">", "&gt;")
    texto = Replace(texto, """", "&quot;")
    texto = Replace(texto, "'", "&#39;")
    EscaparHTML = texto
End Function

Function EscaparJS(ByVal texto)
    If IsNull(texto) Or IsEmpty(texto) Then EscaparJS = "" : Exit Function
    texto = CStr(texto)
    texto = Replace(texto, "\", "\\")
    texto = Replace(texto, """", "\""")
    texto = Replace(texto, "'", "\'")
    texto = Replace(texto, vbCr, "")
    texto = Replace(texto, vbLf, "\n")
    texto = Replace(texto, vbTab, "")
    EscaparJS = texto
End Function

Function PaginaActual()
    Dim s
    s = LCase(Request.ServerVariables("SCRIPT_NAME"))
    PaginaActual = Mid(s, InStrRev(s, "/") + 1)
End Function
'---------------------------------------------------------------------
'  Archivos del sistema (para verificar existencia de recursos)
'---------------------------------------------------------------------
Function ExisteArchivo(ByVal rutaRelativa)
    Dim fso, ok
    ok = False
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    ok = fso.FileExists(Server.MapPath("/" & rutaRelativa))
    Set fso = Nothing
    On Error GoTo 0
    ExisteArchivo = ok
End Function

Function ExtraerNumero(ByVal texto)
    Dim i, c, acum
    acum = ""
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
        If c >= "0" And c <= "9" Then acum = acum & c
    Next
    If acum = "" Then ExtraerNumero = 0 Else ExtraerNumero = CInt(acum)
End Function

'---------------------------------------------------------------------
'  Resuelve la ruta de un archivo de recursos.
'  La BD guarda rutas como ~/songFiles/x.mp3, pero los archivos reales
'  estan en la carpeta SongFilesAndImages_resources. Aqui se mapean y,
'  si el archivo no existe, se usa un reemplazo de la misma carpeta.
'    tipo: "cancion" (audio) o "imagen"
'---------------------------------------------------------------------
Function RutaRecurso(ByVal rutaDb, ByVal tipo)
    Dim base, predet, ruta, pos, carpetaDb, nombreDb, carpeta, rel, num
    base = "SongFilesAndImages_resources"

    If LCase(tipo) = "cancion" Then
        predet = base & "/SongFiles/bensound-adventure.mp3"
    Else
        predet = base & "/PlaylistImages/default.jpg"
    End If

    If IsNull(rutaDb) Or IsEmpty(rutaDb) Then
        RutaRecurso = RUTA_BASE & predet
        Exit Function
    End If
    If Trim(CStr(rutaDb)) = "" Then
        RutaRecurso = RUTA_BASE & predet
        Exit Function
    End If

    ruta = Replace(Replace(Trim(CStr(rutaDb)), "\", "/"), "~/", "")
    pos = InStrRev(ruta, "/")
    If pos > 0 Then
        carpetaDb = LCase(Left(ruta, pos - 1))
        nombreDb = Mid(ruta, pos + 1)
    Else
        carpetaDb = ""
        nombreDb = ruta
    End If

    Select Case carpetaDb
        Case "songfiles"
            carpeta = base & "/SongFiles"
        Case "userimages"
            carpeta = base & "/UserImages"
        Case "playlistimages"
            carpeta = base & "/PlaylistImages"
        Case "albumimages", "artistimages"
            ' El proyecto solo incluye las imagenes de SongFilesAndImages_resources:
            ' se reutilizan las portadas playlist1..5 segun el numero del album/artista.
            num = ExtraerNumero(nombreDb)
            If num >= 1 And num <= 5 Then
                carpeta = base & "/PlaylistImages"
                nombreDb = "playlist" & num & ".png"
            Else
                carpeta = base & "/PlaylistImages"
                nombreDb = "default.jpg"
            End If
        Case Else
            carpeta = base & "/PlaylistImages"
            nombreDb = "default.jpg"
    End Select

    rel = carpeta & "/" & nombreDb
    If ExisteArchivo(rel) Then
        RutaRecurso = RUTA_BASE & rel
    Else
        RutaRecurso = RUTA_BASE & predet
    End If
End Function

'---------------------------------------------------------------------
'  Etiqueta <img> con reemplazo si la imagen no existe
'---------------------------------------------------------------------
Function EtiquetaImg(ByVal rutaDb, ByVal clase, ByVal alt)
    Dim url, fb
    url = RutaRecurso(rutaDb, "imagen")
    fb  = RUTA_BASE & "SongFilesAndImages_resources/PlaylistImages/default.jpg"
    EtiquetaImg = "<img class=""" & clase & """ src=""" & url & """ alt=""" & EscaparHTML(alt) & """ loading=""lazy"" onerror=""this.onerror=null;this.src='" & fb & "';"">"
End Function

'---------------------------------------------------------------------
'  Catalogo (TBL_CATALOG) - se carga una vez por pagina
'---------------------------------------------------------------------
Sub CargarCatalogo()
    Dim cmdC, rsC
    Set CATALOGO = Server.CreateObject("Scripting.Dictionary")
    Set cmdC = Server.CreateObject("ADODB.Command")
    Set cmdC.ActiveConnection = CONN
    cmdC.CommandText = "SELECT COD_CATALOG_PARENT, COD_CATALOG, VALUE FROM TBL_CATALOG"
    cmdC.CommandType = adCmdText
    Set rsC = cmdC.Execute
    Do While Not rsC.EOF
        CATALOGO(Trim(CStr(rsC("COD_CATALOG_PARENT") & "")) & "|" & Trim(CStr(rsC("COD_CATALOG") & ""))) = CStr(rsC("VALUE") & "")
        rsC.MoveNext
    Loop
    rsC.Close
    Set rsC = Nothing
    Set cmdC = Nothing
End Sub

Function EtiquetaCatalogo(ByVal padre, ByVal codigo)
    Dim clave
    clave = CStr(padre) & "|" & CStr(codigo)
    On Error Resume Next
    If CATALOGO Is Nothing Then
        If IsObject(CONN) Then
            If CONN.State = adStateOpen Then CargarCatalogo()
        End If
    End If
    If Not CATALOGO Is Nothing Then
        If CATALOGO.Exists(clave) Then
            EtiquetaCatalogo = CATALOGO(clave)
            On Error GoTo 0
            Exit Function
        End If
    End If
    On Error GoTo 0
    EtiquetaCatalogo = CStr(codigo)
End Function

'---------------------------------------------------------------------
'  Comandos SQL parametrizados
'---------------------------------------------------------------------
Function NuevoComando(ByVal sql)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = CONN
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    Set NuevoComando = cmd
End Function

Sub ParTexto(ByVal cmd, ByVal nombre, ByVal valor, ByVal tam)
    cmd.Parameters.Append cmd.CreateParameter(nombre, adVarChar, adParamInput, tam, valor)
End Sub

Sub ParEntero(ByVal cmd, ByVal nombre, ByVal valor)
    cmd.Parameters.Append cmd.CreateParameter(nombre, adInteger, adParamInput, Empty, valor)
End Sub
'---------------------------------------------------------------------
'  Consultas de apoyo usadas por la pagina principal
'---------------------------------------------------------------------
Function ArtistasDeCancion(ByVal idCancion)
    Dim cmdA, rsA, nombres
    nombres = ""
    Set cmdA = NuevoComando("SELECT a.ARTIST_NAME, a.ARTIST_LASTNAME FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = ? ORDER BY a.ID_ARTIST")
    ParEntero cmdA, "@id", CLng(idCancion)
    Set rsA = cmdA.Execute
    Do While Not rsA.EOF
        If nombres <> "" Then nombres = nombres & ", "
        nombres = nombres & Trim(CStr(rsA("ARTIST_NAME") & "") & " " & CStr(rsA("ARTIST_LASTNAME") & ""))
        rsA.MoveNext
    Loop
    rsA.Close
    Set rsA = Nothing
    Set cmdA = Nothing
    If nombres = "" Then nombres = "Artista desconocido"
    ArtistasDeCancion = nombres
End Function

Function AlbumesDeCancion(ByVal idCancion)
    Dim cmdA, rsA, nombres
    nombres = ""
    Set cmdA = NuevoComando("SELECT al.TITLE_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = ? ORDER BY al.ID_ALBUM")
    ParEntero cmdA, "@id", CLng(idCancion)
    Set rsA = cmdA.Execute
    Do While Not rsA.EOF
        If nombres <> "" Then nombres = nombres & ", "
        nombres = nombres & CStr(rsA("TITLE_ALBUM") & "")
        rsA.MoveNext
    Loop
    rsA.Close
    Set rsA = Nothing
    Set cmdA = Nothing
    If nombres = "" Then nombres = "Sin album"
    AlbumesDeCancion = nombres
End Function

Function PortadaDeCancion(ByVal idCancion)
    Dim cmdA, rsA, img
    img = ""
    Set cmdA = NuevoComando("SELECT TOP 1 al.ALBUM_IMAGE_PATH AS IMG FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = ? ORDER BY al.ID_ALBUM")
    ParEntero cmdA, "@id", CLng(idCancion)
    Set rsA = cmdA.Execute
    If Not rsA.EOF Then
        If Not IsNull(rsA("IMG")) And Not IsEmpty(rsA("IMG")) Then img = RutaRecurso(rsA("IMG"), "imagen")
    End If
    rsA.Close
    Set rsA = Nothing
    Set cmdA = Nothing

    If img = "" Then
        Set cmdA = NuevoComando("SELECT TOP 1 a.ARTIST_IMAGE AS IMG FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = ? ORDER BY a.ID_ARTIST")
        ParEntero cmdA, "@id", CLng(idCancion)
        Set rsA = cmdA.Execute
        If Not rsA.EOF Then
            If Not IsNull(rsA("IMG")) And Not IsEmpty(rsA("IMG")) Then img = RutaRecurso(rsA("IMG"), "imagen")
        End If
        rsA.Close
        Set rsA = Nothing
        Set cmdA = Nothing
    End If

    If img = "" Then img = RUTA_BASE & "SongFilesAndImages_resources/PlaylistImages/default.jpg"
    PortadaDeCancion = img
End Function

'---------------------------------------------------------------------
'  Iconos SVG en linea (sin dependencias externas)
'---------------------------------------------------------------------
Function Icono(ByVal nombre)
    Dim inner, tam
    tam = 20
    Select Case LCase(Trim(nombre))
        Case "play"
            inner = "<path d='M8 5.2v13.6c0 .8.9 1.3 1.6.9l10.4-6.8c.6-.4.6-1.4 0-1.8L9.6 4.3C8.9 3.9 8 4.4 8 5.2z'/>"
        Case "pausa"
            inner = "<rect x='6' y='5' width='4' height='14' rx='1.2'/><rect x='14' y='5' width='4' height='14' rx='1.2'/>"
        Case "siguiente"
            inner = "<path d='M5 5.6v12.8c0 .8.9 1.3 1.6.9l8.6-6.4c.6-.4.6-1.4 0-1.8L6.6 4.7C5.9 4.3 5 4.8 5 5.6z'/><rect x='16.6' y='5' width='2.6' height='14' rx='1.2'/>"
        Case "anterior"
            inner = "<path d='M19 5.6v12.8c0 .8-.9 1.3-1.6.9L8.8 12.9c-.6-.4-.6-1.4 0-1.8l8.6-6.4c.7-.4 1.6.1 1.6.9z'/><rect x='4.8' y='5' width='2.6' height='14' rx='1.2'/>"
        Case "volumen"
            inner = "<path d='M4 9.2v5.6c0 .6.5 1.2 1.1 1.2h2.4l4.2 3.5c.7.5 1.7 0 1.7-.9V5.4c0-.9-1-1.4-1.7-.9L7.5 8H5.1C4.5 8 4 8.6 4 9.2z'/><path d='M16.2 8.6a4.8 4.8 0 0 1 0 6.8' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linecap='round'/><path d='M18.8 6.2a8.2 8.2 0 0 1 0 11.6' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linecap='round'/>"
        Case "inicio"
            inner = "<path d='M4 10.8 12 4.2l8 6.6V19a1 1 0 0 1-1 1h-4.5v-5.5h-5V20H5a1 1 0 0 1-1-1z'/>"
        Case "nota"
            inner = "<path d='M9 17.2V6.3l10-2.1v10.6' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linejoin='round'/><circle cx='6.8' cy='17.4' r='2.4'/><circle cx='16.8' cy='15.3' r='2.4'/>"
        Case "disco"
            inner = "<path d='M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zm0 5.2a3.8 3.8 0 1 1 0 7.6 3.8 3.8 0 0 1 0-7.6z'/><circle cx='12' cy='12' r='1.4'/>"
        Case "micro"
            inner = "<rect x='9.2' y='3.4' width='5.6' height='10.4' rx='2.8'/><path d='M6 11.4a6 6 0 0 0 12 0' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/><path d='M12 17.4V21M9 21h6' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/>"
        Case "lista"
            inner = "<path d='M4 5.5h2.4M9.2 5.5H20M4 12h2.4M9.2 12H20M4 18.5h2.4M9.2 18.5H20' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/><path d='M6.4 3.6v16.8' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/>"
        Case "escudo"
            inner = "<path d='M12 3.2 18.6 6v5.1c0 4.4-2.8 7.7-6.6 9.7-3.8-2-6.6-5.3-6.6-9.7V6z'/><path d='m9 11.8 2.1 2.1 4.1-4.1' fill='none' stroke='#eaf1ff' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round'/>"
        Case "usuario"
            inner = "<circle cx='12' cy='8' r='3.6'/><path d='M5 20c.8-3.6 3.6-5.4 7-5.4s6.2 1.8 7 5.4z'/>"
        Case "salir"
            inner = "<path d='M14.5 4h3.4A1.6 1.6 0 0 1 19.5 5.6v12.8a1.6 1.6 0 0 1-1.6 1.6h-3.4' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/><path d='M9.5 8.2 5.7 12l3.8 3.8M6.4 12h8.4' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'/>"
        Case "editar"
            inner = "<path d='M4.5 19.5 5.4 15 16.3 4.1a2.1 2.1 0 0 1 3 3L8.4 18z' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linejoin='round'/><path d='M14.5 6.5 17.5 9.5' fill='none' stroke='currentColor' stroke-width='1.7'/>"
        Case "basura"
            inner = "<path d='M5.5 7.5h13M9.5 7.5V5.8A1.3 1.3 0 0 1 10.8 4.5h2.4a1.3 1.3 0 0 1 1.3 1.3v1.7' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linecap='round'/><path d='M7 7.5 7.9 19a1.4 1.4 0 0 0 1.4 1.3h5.4A1.4 1.4 0 0 0 16.1 19L17 7.5' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linejoin='round'/><path d='M10.5 11v6M13.5 11v6' fill='none' stroke='currentColor' stroke-width='1.5' stroke-linecap='round'/>"
        Case "mas"
            inner = "<path d='M12 5.5v13M5.5 12h13' fill='none' stroke='currentColor' stroke-width='1.9' stroke-linecap='round'/>"
        Case "check"
            inner = "<path d='m5 12.8 4.2 4.2L19 7.2' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'/>"
        Case "alerta"
            inner = "<path d='M12 4.2 21 19.4H3z' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linejoin='round'/><path d='M12 9.6v4.2' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/><circle cx='12' cy='16.6' r='1'/>"
        Case "buscar"
            inner = "<circle cx='11' cy='11' r='6.4' fill='none' stroke='currentColor' stroke-width='1.8'/><path d='m15.8 15.8 4.2 4.2' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round'/>"
        Case Else
            inner = "<circle cx='12' cy='12' r='8' fill='none' stroke='currentColor' stroke-width='1.7'/>"
    End Select
    Icono = "<svg class='icono-svg' viewBox='0 0 24 24' width='" & tam & "' height='" & tam & "' fill='currentColor' aria-hidden='true'>" & inner & "</svg>"
End Function
%>
