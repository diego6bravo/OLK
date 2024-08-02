<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (optOfert and myAut.HasAuthorization(7)) Then Response.Redirect "unauthorized.asp" %>
<!--#include file="ventas/search_OfertX.asp" -->
<!--#include file="agentBottom.asp"-->