<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->

<!--#include file="lang/crdConfDetailOpen.asp" -->
<!--#include file="../authorizationClass.asp"-->
<!--#include file="../loadAlterNames.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
Dim varx
varx = "0"
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%

Dim myAut
set myAut = New clsAuthorization

Dim DocName
set rs = Server.CreateObject("ADODB.recordset")
CardRetVal = Request("DocEntry")

If Request("CardCode") <> "" Then 
	ClientShowBalance = CheckAgentClientFilter(Request("CardCode"), 2)
End If
%>
<title><% If Request("CardCode") = "" Then %><%=getcrdConfDetailOpenLngStr("LttlClientConfDetails")%><% Else %><%=getcrdConfDetailOpenLngStr("LttlClientDetails")%><% End If %></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylenuevo.css">
</head>
<body>
<!--#include file="crdConfDetail.asp"-->
</body>
</html>
