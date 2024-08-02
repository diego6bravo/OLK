<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOOPR Then Response.Redirect "unauthorized.asp" %>
<!--#include file="addSO/searchOpenedSO.asp"-->
<!--#include file="agentBottom.asp"-->