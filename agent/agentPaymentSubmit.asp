<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (Session("UserName") <> "" and CardType = "C" and myApp.EnableORCT) Then Response.Redirect "unauthorized.asp" %>
<!--#include file="payments/submitPayment.asp" -->
<!--#include file="agentBottom.asp"-->