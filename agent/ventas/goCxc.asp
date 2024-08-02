<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<% Session("UserName") = saveHTMLDecode(Request("c1"), True)
Response.redirect "../cxc.asp" %>