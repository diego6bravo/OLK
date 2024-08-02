<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (comDocsMenu or myApp.EnableORCT) Then Response.Redirect "unauthorized.asp" %>
<!--#include file="ventas/ventasX.asp" -->
<!--#include file="agentBottom.asp"-->