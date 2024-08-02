<!--#include file="clientInc.asp"-->
<!--#include file="expire.inc"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% 
If Session("UserName") = "-Anon-" or not optCXC and userType = "C" Then Response.Redirect "default.asp"
Case "V" %><!--#include file="agentTop.asp"-->
<% 
If Not (myAut.HasAuthorization(24) and IsBPAssigned) Then Response.Redirect "unauthorized.asp"
End Select %>
<!--#include file="cxcData.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>