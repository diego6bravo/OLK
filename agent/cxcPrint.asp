<%@ Language=VBScript %>
<!--#include file="authorizationClass.asp"-->
<html <% If Request("newLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Request("pdf") = "Y" Then
	If Request("newLng") <> "" Then %>
	<!--#include file="lang.asp" -->
	<% End If %>
	<!--#include file="conn.asp" -->
	<% 
	myApp.LoadDBConfigData CInt(Request("dbID"))
	userType = Request("UserType")
	Session("UserName") = Request("CardCode")
ElseIf Request("LinkRep") = "Y" Then
	Session("UserName") = Request("c1")
	Session("Cart") = ""
	Session("RetVal") = ""
	Session("PayRetVal") = ""
End If

set rs = Server.CreateObject("ADODB.RecordSet")

If Request("pdf") = "Y" Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "OLKCheckPDFAccess"
	cmd.Parameters.Refresh()
	cmd("@ID") = Session("id")
	cmd("@Rnd") = Request("myRnd")
	cmd.execute()
	If cmd.Parameters.Item(0).value = 0 Then Response.Redirect "accessDenied.asp"
End If

CmpName = mySession.GetCompanyName

If userType = "C" Then
	sql = "select SelDes from OLKCommon"
	set rs = conn.execute(sql)
	SelDes = rs("SelDes")
Else
	SelDes = 0
End If


Dim myAut
set myAut = New clsAuthorization
 %>

<head>
<!--#include file="loadAlterNames.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>OLK</title>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
<% If Request("LinkRep") = "Y" Then %><link rel="stylesheet" type="text/css" media="all" href="design/0/style/style_cal.css" title="winter" /><% End If %>
<!--#include file="licid.inc"-->
</head>
<body topmargin="0">
<!--#include file="cxcData.asp"-->
</body>
</html>
