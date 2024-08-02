<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myAut.HasBPAccess Then Response.Redirect "unauthorized.asp" %>
<!--#include file="ventas/search_inte.asp" -->
<!--#include file="agentBottom.asp"-->