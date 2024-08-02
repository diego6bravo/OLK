<%@ Language=VBScript %>
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
%>
<!--#include file="chkLogin.asp" -->

<% 

Session("ConfRetVal") = Request("ConfRetVal")
userType = Request("UserType")
Session("vendid") = Request("vendid")

cartPDFAddStr = ""

If Request("document") = "" Then
	If Request.Cookies("catMethod") = "" Then Response.Cookies("catMethod") = "T"
Else
	Response.Cookies("catMethod") = Request("document")
End If

Session("cart") = "cart"
 %>
<html <% If Request("newLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="lcidReturn.inc"-->
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

sql = "select SelDes, " & _
"(select CardCode from R3_ObsCommon..TDOC where LogNum = " & Request("ConfRetVal") & ") CardCode, " & _
"(select PriceList from R3_ObsCommon..TLOGControl where LogNum = " & Session("ConfRetVal") & ") PriceList " & _
"from OLKCommon"
set rs = conn.execute(sql)
If userType = "C" Then SelDes = rs("SelDes") Else SelDes = 0
CmpName = mySession.GetCompanyName
Session("UserName") = rs("CardCode")
Session("PriceList") = rs("PriceList")
If userType = "V" Then
   ShowClientRef = True
End If %>
<head>
<!--#include file="loadAlterNames.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Olk - Interfase <%=txtClient%></title>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
<!--#include file="licid.inc"-->
<!--#include file="clearItem.asp"-->
</head>
<body topmargin="0">
<!--#include file="cartSubmitConfirm.asp" -->
</body>
<% set rs = nothing
conn.close %></html>