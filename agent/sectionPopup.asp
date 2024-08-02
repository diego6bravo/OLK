<%@ Language=VBScript %>
<html>
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

%>

<!--|P:LangLink|-->
<%
'Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
sql = "select SelDes, DirectRate from OLKCommon cross join oadm"
set rs = conn.execute(sql)
If userType = "C" Then SelDes = rs("SelDes") Else SelDes = 0
imgAddPath = "" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
<!--#include file="licid.inc"-->
<% If Request("Excell") <> "Y" and Request("itemSmallRep") <> "Y" Then %>
<link rel="stylesheet" href="design/<%=SelDes%>/style/stylenuevo.css"><% End If %>

</head>
<body topmargin="0">
<!--#include file="design/section.inc"-->
</body>
<% set rs = nothing
conn.close %></html>