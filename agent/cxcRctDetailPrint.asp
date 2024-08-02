<% If Request("newLng") <> "" Then %>
<!--#include file="lang.asp" -->
<% End If %>
<% 
Session("ConfRetVal") = Request("ConfRetVal")
userType = Request("UserType")
Session("vendid") = Request("vendid")

Session("cart") = "cart"
Session("useraccess") = Request("useraccess")
 %>
<html <% If Request("newLng") = "he" Then %>dir="rtl"<% End If %>>
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

myApp.LoadDBConfigData CInt(Request("dbID"))

set rs = Server.CreateObject("ADODB.recordset")

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

If userType = "C" Then 
	sql = "select SelDes from OLKCommon"
	set rs = conn.execute(sql)
	SelDes = rs("SelDes") 
Else 
	SelDes = 0
	myAut.LoadAuthorization Session("vendid"), Session("id")
End If
CmpName = mySession.GetCompanyName %>
<head>
<!--#include file="loadAlterNames.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>OLK</title>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
<!--#include file="licid.inc"-->
<!--#include file="clearItem.asp"-->
</head>
<body topmargin="0">
<!--#include file="cxcRctDetail.asp" -->
</body>
</html>