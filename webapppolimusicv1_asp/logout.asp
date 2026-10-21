<%@ Language=VBScript CodePage=65001 %>
<%
'=====================================================================
'  PoliMusic - logout.asp
'  Cierra la sesion y regresa al login
'=====================================================================
Option Explicit
%>
<!--#include file="includes/conexion.asp"-->
<%
Session.Contents.RemoveAll()
Session.Abandon()
Response.Redirect RUTA_BASE & "login.asp?logout=1"
%>
