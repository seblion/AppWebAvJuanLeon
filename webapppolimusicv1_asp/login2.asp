<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - login.asp
'  Inicio de sesion contra TBL_USER (SQL Server localhost)
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
If SesionIniciada() Then Response.Redirect RUTA_BASE & "index.asp"

Dim mensaje, sesionCerrada
mensaje = ""
sesionCerrada = (Request.QueryString("logout") = "1")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim usuario, clave
    usuario = Trim(Request.Form("txtUsuario") & "")
    clave   = Request.Form("txtClave") & ""

    If usuario = "" Or clave = "" Then
        mensaje = "Debes ingresar tu usuario y tu contrasena."
    Else
        Dim errLogin
        errLogin = ""
        On Error Resume Next
        AbrirConexion()
        If Err.Number <> 0 Then
            errLogin = "No se pudo conectar con la base de datos. Verifica que SQL Server este ejecutandose."
            Err.Clear
            On Error GoTo 0
        Else
            On Error GoTo 0
            Dim cmdL, rsL
            ' Valida contrasena en texto plano o hash SHA1 (HASHBYTES en el servidor)
            Set cmdL = NuevoComando("SELECT ID_USER, USERNAME, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE USERNAME = @u AND (PASSWORD = @p OR PASSWORD = CONVERT(VARCHAR(64), HASHBYTES('SHA1', @p), 2))")
            ParTexto cmdL, "@u", usuario, 50
            ParTexto cmdL, "@p", clave, 50
            On Error Resume Next
            Set rsL = cmdL.Execute
            If Err.Number <> 0 Then
                errLogin = "Error al validar las credenciales. Intentalo de nuevo."
                Err.Clear
            ElseIf Not rsL.EOF Then
                ' Evita fijacion de sesion: se abandona la sesion anterior
                Session.Abandon
                Session("ID_USER")    = rsL("ID_USER")
                Session("USERNAME")   = CStr(rsL("USERNAME") & "")
                Session("USER_PHOTO") = CStr(rsL("USER_PHOTO") & "")
                Session("USER_TYPE")  = rsL("USER_TYPE")
                CerrarConexion()
                Response.Redirect RUTA_BASE & "index.asp"
            Else
                mensaje = "Usuario o contrasena incorrectos. Intentalo de nuevo."
            End If
            On Error GoTo 0
        End If
        If errLogin <> "" Then mensaje = errLogin
    End If
End If

TituloPagina = "Iniciar sesion"
ClaseBody    = "pagina-login sin-reproductor"
%>
