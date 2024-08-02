<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOITM Then Response.Redirect "unauthorized.asp" %>
<!--#include file="ventas/itemsX.asp" -->
<!--#include file="agentBottom.asp"-->