<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - hola.asp  (pagina minima de diagnostico)
'  No usa includes, ni sesion, ni base de datos: solo imprime por
'  etapas SIN buffer, de modo que la ULTIMA linea visible indica
'  exactamente donde falla el motor de ASP.
'  Abrir: http://localhost/webapppolimusicv1_asp/hola.asp
'=====================================================================
Response.Buffer = False
Response.ContentType = "text/html"
Response.Charset = "UTF-8"
%>
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>hola - diagnostico</title></head>
<body style="font-family:Consolas,monospace;background:#04060c;color:#e9eefb;padding:30px;line-height:2">
<%
Response.Write "1. ASP clasico: FUNCIONA<br>" & vbCrLf

Dim fso, rutaFisica
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Response.Write "2. FileSystemObject: FUNCIONA<br>" & vbCrLf

rutaFisica = Server.MapPath("/")
Response.Write "3. Server.MapPath raiz del sitio: " & rutaFisica & "<br>" & vbCrLf

If fso.FolderExists(Server.MapPath("SongFilesAndImages_resources")) Then
    Response.Write "4. Carpeta SongFilesAndImages_resources: ENCONTRADA<br>" & vbCrLf
Else
    Response.Write "4. Carpeta SongFilesAndImages_resources: NO ENCONTRADA (revisar ruta fisica)<br>" & vbCrLf
End If
Set fso = Nothing

Session("PM_HOLA") = "ok"
If CStr(Session("PM_HOLA") & "") = "ok" Then
    Response.Write "5. Estado de sesion: FUNCIONA<br>" & vbCrLf
Else
    Response.Write "5. Estado de sesion: NO DISPONIBLE (revisar web.config / IIS)<br>" & vbCrLf
End If

Dim conn, errDB
errDB = ""
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open "Provider=SQLNCLI11;Server=localhost\SQLSERVER2026;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;"
If Err.Number <> 0 Then
    errDB = Err.Description
    Err.Clear
    On Error GoTo 0
    Response.Write "6. Conexion SQLNCLI11: FALLO - " & errDB & "<br>" & vbCrLf
    ' Segundo intento con SQLOLEDB (MDAC, siempre presente)
    Set conn = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open "Provider=SQLOLEDB;Server=localhost;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;"
    If Err.Number <> 0 Then
        errDB = Err.Description
        Err.Clear
        On Error GoTo 0
        Response.Write "7. Conexion SQLOLEDB: FALLO - " & errDB & "<br>" & vbCrLf
    Else
        On Error GoTo 0
        conn.Close
        Response.Write "7. Conexion SQLOLEDB: FUNCIONA (usar SQLOLEDB)<br>" & vbCrLf
    End If
Else
    On Error GoTo 0
    conn.Close
    Response.Write "6. Conexion SQLNCLI11: FUNCIONA<br>" & vbCrLf
End If
Set conn = Nothing

Response.Write "8. FIN - todas las etapas completadas<br>" & vbCrLf
%>
<p style="color:#93a4c7;font-size:13px">Si una linea no aparece, el error esta en esa etapa. Revisa tambien el Visor de sucesos → Aplicacion → "Active Server Pages" para ver el codigo ASP exacto.</p>
</body>
</html>
