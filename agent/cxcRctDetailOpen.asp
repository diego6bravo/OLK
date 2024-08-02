<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="authorizationClass.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = New clsAuthorization

%>
<!--#include file="chkLogin.asp" -->
<!--#include file="cxcRctDetail.asp" -->
