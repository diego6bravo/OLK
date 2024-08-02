<% Response.Expires = -1 %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="chkLogin.asp" -->
<!--#include file="cxcDocDetail.asp" -->
