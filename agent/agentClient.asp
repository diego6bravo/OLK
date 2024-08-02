<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOCRD Then Response.Redirect "unauthorized.asp" %>
<!--#include file="addCard/addCard.asp" -->
<!--#include file="agentBottom.asp"-->