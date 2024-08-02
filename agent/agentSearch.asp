<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myAut.HasAuthorization(1) Then Response.Redirect "unauthorized.asp" %>
<!--#include file="ventas/search_arte.asp" -->
<!--#include file="agentBottom.asp"-->