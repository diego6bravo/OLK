<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOITM Then Response.Redirect "unauthorized.asp" %>
<!--#include file="addItem/addItem.asp" -->
<!--#include file="agentBottom.asp"-->