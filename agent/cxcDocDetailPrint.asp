<% If Request("newLng") <> "" Then %>
<!--#include file="lang.asp" -->
<% End If %>

<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

myApp.LoadDBConfigData CInt(Request("dbID"))

Session("ConfRetVal") = Request("ConfRetVal")
userType = Request("UserType")
Session("vendid") = Request("vendid")
Session("UserName") = Request("CardCode")
Session("cart") = "cart"
Session("useraccess") = Request("useraccess")
 %>
<html <% If Request("newLng") = "he" Then %>dir="rtl"<% End If %>>
<%
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
<!--#include file="cxcDocDetail.asp" -->
</body>
</html>