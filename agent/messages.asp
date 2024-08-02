<!--#include file="clientInc.asp"-->
<% If userType = "C" Then %><!--#include file="clientTop.asp"--><% End If %>
<% If Not optMsg Then Response.Redirect "default.asp" %>
<!--#include file="messages/messages.asp"-->
<% If userType = "C" Then %><!--#include file="clientBottom.asp"--><% End If %>