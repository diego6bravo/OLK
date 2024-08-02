<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

          
set rs = Server.CreateObject("ADODB.RecordSet")
%>
<html xmlns="http://www.w3.org/1999/xhtml" <% If Session("rtl") <> "" Then %> dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<title>Untitled 1</title>
<link rel="stylesheet" type="text/css" href="design/0/style/stylenuevo.css"/>
</head>

<body style="background-color: #EDF5FE; ">
<!--#include file="myTitle.asp"-->
</body>

</html>
