<%@ Language=VBScript %>
<!-- #include file="chkLogin.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
sql = "select SelDes, DirectRate from OLKCommon cross join oadm"
set rs = conn.execute(sql)
If userType = "C" Then SelDes = rs("SelDes") Else SelDes = 0
imgAddPath = "" %>
<!--#include file="design/section.inc"-->
<% set rs = nothing
conn.close %>