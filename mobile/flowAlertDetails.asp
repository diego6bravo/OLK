<!--#include file="lang/flowAlertDetails.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="myHTMLEncode.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><title>New 
Page 1</title>
<link rel="stylesheet" type="text/css" href="design/0/style/stylePopUp.css">
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="lcidReturn.inc"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

ExecAt = Request("ExecAt")
set rs = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(Request("FlowID"), "-", "_") & "_line"
LoadCmdParams
set rs = cmd.execute()
%>
<body topmargin="0" leftmargin="0" bgcolor="#9BC4FF">

<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
	<% For each item in rs.Fields %>
		<td>
		<p align="center"><%=item.Name%>&nbsp;</td>
	<% Next %>
	</tr>
	<% do while not rs.eof %>
	<tr class="GeneralTbl">
	<% For each item in rs.Fields %>
		<td><% If Not IsNull(item) Then %><%=item%><% End If %>&nbsp;</td>
	<% next %>
	</tr>
	<% rs.movenext
	loop %>
</table>
</body>

</html>
<%
Sub LoadCmdParams
	cmd.Parameters.Refresh()

	cmd("@LanID") = Session("LanID")
	cmd("@SlpCode") = Session("vendid")
	cmd("@dbName") = Session("olkdb")
	cmd("@branch") = Session("branch")
	cmd("@UserType") = "V"
	
	Select Case ExecAt
		'Case "O0", "O1", "O7" ' Aprove Sales Order, Convert Quotation to Sales Order, Convert Sales Order to Invoice
		'	cmd("@Entry") = arrVars(0)
		'Case "O2", "O3", "O4" ' Close  Object, Cancel Object, Remove Object
		'	cmd("@ObjectCode") = arrVars(0) 
		'	cmd("@Entry") = arrVars(1)
		Case "D2" ' Add Item
			cmd("@LogNum") = Session("RetVal")
			cmd("@CardCode") = Session("UserName")
			cmd("@ItemCode") = Request("Item")
			If Request("Quantity") <> "" Then cmd("@Quantity") = CDbl(getNumericOut(Request("Quantity")))
			If Request("SaleType") <> "" Then cmd("@Unit") = Request("SaleType")
			If Request("Price") <> "" Then cmd("@Price") = CDbl(getNumericOut(Request("Price")))
			If Request("WhsCode") <> "" Then cmd("@WhsCode") = Request("WhsCode")
			If Request("SellAll") = "Y" Then cmd("@All") = "Y"
		Case "D3" ' LtxtDocConf
			cmd("@LogNum") = Session("RetVal")
		Case "R1" ' LtxtCreation	******clean*******
		Case "R2" ' LtxtRcpConf
			cmd("@LogNum") = Session("PayRetVal")
		Case "A1" ' LtxtItmConf
			cmd("@LogNum") = Session("ItmRetVal")
		Case "C1" ' LtxtClientConf
			cmd("@LogNum") = Session("CrdRetVal")
		Case "C2" ' LtxtActivityConf
			cmd("@LogNum") = Session("ActRetVal")
		Case "C3" ' LtxtActivityConf
			cmd("@LogNum") = Session("SORetVal")
	End Select	
	
	Select Case ExecAt
		Case "C2", "C3", "R1", "R2", "D3", "D1"
			cmd("@CardCode") = Session("UserName")
	End Select
End Sub
%>