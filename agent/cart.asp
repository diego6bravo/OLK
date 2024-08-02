<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% 
If (not comDocsMenu or strScriptName = "submitcart.asp" or strScriptName = "cartcancel.asp") Then Response.Redirect "unauthorized.asp" %>
<% If userType = "V" and (session("RetVal") = "" or strScriptName = "submitcart.asp" or strScriptName = "cartcancel.asp") Then Response.Redirect "default.asp" %>
<!--#include file="cartApp.asp" -->
<!--#include file="agentBottom.asp"-->