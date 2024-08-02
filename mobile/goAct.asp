<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<%
Session("ActRetVal") = Request("LogNum")
Session("UserName") = Request("CardCode")
Session("ActReadOnly") = False

Response.Redirect "operaciones.asp?cmd=activity"
%>