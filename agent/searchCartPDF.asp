<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<head>
<link rel="stylesheet" href="design/0/style/stylenuevo.css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body topmargin="0" leftmargin="0" link="#4783C5" vlink="#4783C5" bgcolor="#0166CB">
<!--#include file="conn.asp" -->
<!--#include file="lang.asp"-->
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

If Request("pdf") = "Y" Then 
	myApp.LoadDBConfigData CInt(Request("dbID"))
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "OLKCheckPDFAccess"
	cmd.Parameters.Refresh()
	cmd("@ID") = Session("id")
	cmd("@Rnd") = Request("myRnd")
	cmd.execute()
	If cmd.Parameters.Item(0).value = 0 Then Response.Redirect "accessDenied.asp"
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAgentReloadLogin" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@SlpCode") =  Request("vendid")
	cmd("@branch") = Request("branch")
	set rd = Server.CreateObject("ADODB.RecordSet")
	set rd = cmd.execute()
	
	Session("branch") = Request("branch")
	Session("vendid") = Request("vendid")
	Session("useraccess") = rd("Access")
	Session("BranchWhs") = rd("BranchWhs")
	Session("AgentWhs") = rd("AgentWhs")
	Session("AgentLastUpdate") = rd("LastUpdate")
	
	myAut.LoadAuthorization Session("vendid"), Session("ID")
	mySession.LoginAgent


End If

searchCmd = "searchCatalog"
%>
<!--#include file="searchCart.asp"-->
</body>
</html>
