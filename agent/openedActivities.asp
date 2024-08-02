<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOCLG Then Response.Redirect "unauthorized.asp" %>
<!--#include file="ventas/search_activityX.asp" -->
<!--#include file="agentBottom.asp"-->