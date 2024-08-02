<!--#include file="lang/flowAlertDetails.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lcidReturn.inc"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title>New 
Page 1</title>
<link rel="stylesheet" type="text/css" href="design/<%=Request("SelDes")%>/style/stylePopUp.css">
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select LineQuery from OLKUAF where FlowID = " & Request("FlowID")
set rs = conn.execute(sql)

	ExecAt = Request("ExecAt")
	If Request("LogNum") <> "" Then
		LogNum = Request("LogNum")
	ElseIf Left(ExecAt,1) = "D" Then 
		LogNum = Session("RetVal")
	ElseIf Left(ExecAt,1) = "R" Then
		LogNum = Session("PayRetVal")
	ElseIf ExecAt = "C2" Then
		LogNum = Session("ActRetVal")
	End If
	If ExecAt <> "D1" and ExecAt <> "R1" Then sqlBase = "declare @LogNum int set @LogNum = " & LogNum & " "
	
	sqlBase = 	sqlBase & "declare @LanID int set @LanID = " & Session("LanID") & " " & _
				"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
				"declare @dbName nvarchar(100) set @dbName = db_name() " & _
				"declare @branch int set @branch = " & Session("branch") & " "
	
	If Left(ExecAt,1) = "D" or Left(ExecAt,1) = "R" or ExecAt = "C2" Then sqlBase = sqlBase & "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
		
	If ExecAt = "D2" Then 
		If Request("SaleType") <> "" Then SaleType = Request("SaleType") Else SaleType = "NULL"
		If Request("WhsCode") <> "" Then WhsCode = "N'" & Request("WhsCode") & "'" Else WhsCode = "OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode)"
		If Request("precio") <> "" Then precio = getNumeric(Request("precio")) Else precio = "NULL"
		If Request("chkAddAll") <> "Y" Then 
			If Request("addQty") <> "" Then addQty = getNumeric(Request("addQty")) Else addQty = "1"
		Else
			addQty = "OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "Val(@ItemCode, @WhsCode, @dbName, @LogNum, -1)"
		End If


		sqlBase = sqlBase & "declare @ItemCode nvarchar(20) set @ItemCode = N'" & Request("Item") & "' " & _
								"declare @WhsCode nvarchar(8) set @WhsCode = " & WhsCode & " " & _
								"declare @Unit smallint set @Unit = " & SaleType & " " & _
								"If @Unit is null Begin " & _
								"set @Unit = Case '" & userType & "'  " & _
								"		When 'C' Then (select ClientSaleUnit from olkcommon)  " & _
								"		When 'V' Then (select AgentSaleUnit from olkcommon) End End " & _
								"declare @Quantity numeric(19,6) set @Quantity = " & addQty & " " & _
								"declare @Price numeric(19,6) set @Price = " & precio & " " & _
								"If @Price is null begin " & _
								"	EXEC OLKCommon..DBOLKGetItemPrice" & Session("ID") & " @ItemCode = @ItemCode, @CardCode = @CardCode, @PriceList = " & Session("PriceList") & ", @UserType = '" & userType & "', @ItemPrice = @Price out " & _
								"End If @Price is null begin set @Price = 0 End "					


	End If

sql = sqlBase & rs("LineQuery")
set rs = conn.execute(sql)
%>
<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" width="100%" id="table1">
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