<%@ Language=VBScript %>
<% If session("ID") = "" Then response.redirect "lock.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
imgAddPath = "" %>
<!--#include file="sectionInc.inc"-->
<% set rs = nothing
conn.close %>