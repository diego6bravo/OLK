<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C"
	user = Session("UserName")
	MainDoc = "clientes" %><!--#include file="clientTop.asp"-->
<% 
If (Session("UserName") = "-Anon-") Then Response.Redirect "default.asp"
Case "V"
	user = Session("vendid")
	MainDoc = "ventas" %><!--#include file="agentTop.asp"-->
<%
End Select
%>
<!--#include file="messages/messagePost.asp" -->
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>