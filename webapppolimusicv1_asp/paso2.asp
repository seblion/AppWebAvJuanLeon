<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - paso2.asp  (diagnostico nivel 2)
'  Carga el include includes/conexion.asp y prueba la conexion
'  directa a SQL Server, imprimiendo por etapas SIN buffer.
'  Abrir: http://localhost/webapppolimusicv1_asp/paso2.asp
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
Response.Buffer = False
Response.Charset = "UTF-8"
Response.Write "A. Include conexion.asp: CARGADO OK. RUTA_BASE=" & RUTA_BASE & "<br>" & vbCrLf

Dim connT, errD
errD = ""
Set connT = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connT.Open "Provider=SQLNCLI11;Server=localhost;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;"
If Err.Number <> 0 Then
    errD = Err.Description
    Err.Clear
    On Error GoTo 0
    Response.Write "B. Conexion SQLNCLI11: FALLO -> " & errD & "<br>" & vbCrLf
    Set connT = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    connT.Open "Provider=SQLOLEDB;Server=localhost;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;"
    If Err.Number <> 0 Then
        errD = Err.Description
        Err.Clear
        On Error GoTo 0
        Response.Write "C. Conexion SQLOLEDB: FALLO -> " & errD & "<br>" & vbCrLf
    Else
        On Error GoTo 0
        connT.Close
        Response.Write "C. Conexion SQLOLEDB: OK (revisar SQLNCLI11)<br>" & vbCrLf
    End If
Else
    On Error GoTo 0
    connT.Close
    Response.Write "B. Conexion SQLNCLI11: OK<br>" & vbCrLf
End If
Set connT = Nothing

Response.Write "D. FIN - include y conexion verificados<br>" & vbCrLf
%>
