<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - api_reproducir.asp
'  Incrementa el contador PLAYS de una cancion (llamada AJAX
'  desde js/player.js) y devuelve JSON.
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
Response.ContentType = "application/json"

Dim resultado, idCancion, plays
resultado = "{""ok"":false}"
idCancion = 0
plays = 0

If IsNumeric(Request.QueryString("id")) Then
    idCancion = CLng(Request.QueryString("id"))
    If idCancion > 0 Then
        AbrirConexion()
        On Error Resume Next
        Dim cmdU, cmdS, rsS
        Set cmdU = NuevoComando("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = @id")
        ParEntero cmdU, "@id", idCancion
        cmdU.Execute
        If Err.Number = 0 Then
            Set cmdS = NuevoComando("SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = @id")
            ParEntero cmdS, "@id", idCancion
            Set rsS = cmdS.Execute
            If Err.Number = 0 Then
                If Not rsS.EOF Then
                    If IsNull(rsS("PLAYS")) Or IsEmpty(rsS("PLAYS")) Then plays = 0 Else plays = CStr(rsS("PLAYS"))
                    resultado = "{""ok"":true,""id"":" & idCancion & ",""plays"":" & plays & "}"
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        CerrarConexion()
    End If
End If

Response.Write resultado
%>
