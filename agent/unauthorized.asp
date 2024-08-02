<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% End Select %>
<% addLngPathStr = "" %>
<!---|P:LangLink|-->
<center><h4>Unauthorized access!</h4></center>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>