<% response.expires = 0 %>
<% If Session("VendId") = "" Then response.redirect "default.asp" %>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/topGetDocLink.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
set rd = server.createobject("ADODB.RecordSet") %>
<!--#include file="loadAlterNames.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) CardName from OCRD where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
set rs = conn.execute(sql)
CardName = rs(0)
rs.close

sql = "select DocEntry, DocNum, DocDate, Comments, DocDueDate from "
Select Case Request("DocType")
	Case 23
		sql = sql & "OQUT"
		docNames = txtQuotes
	Case 17
		sql = sql & "ORDR"
		docNames = txtOrdrs
	Case 15
		sql = sql & "ODLN"
		docNames = txtOdlns
	Case 16
		sql = sql & "ORDN"
		docNames = txtOrnds
	Case 13
		sql = sql & "OINV"
		docNames = txtInvs
	Case 14
		sql = sql & "ORIN"
		docNames = txtOrins
	Case 203
		sql = sql & "ODPI"
		docNames = gettopGetDocLinkLngStr("LtxtInvDownPay")
	Case 24
		sql = sql & "ORCT"
		docNames = txtRcts
	Case 46
		sql = sql & "OVPM"
		docNames = txtOvpms
	Case 67
		sql = sql & "OWTR"
		docNames = gettopGetDocLinkLngStr("LtxtInvTrans")
End Select
sql = sql & " where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DocNum like '" & Replace(Request("DocNum"), "*", "%") & "' " & _
"order by Convert(int,DocDate) desc "
set rd = conn.execute(sql) %>

<head>
<title><%=gettopGetDocLinkLngStr("LtxtDocSel")%> - <%=docNames%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="design/0/style/stylePopUp.css">
<script language="javascript" src="general.js"></script>
</head>

<body topmargin="0" leftmargin="00" rightmargin="0" bottommargin="0">
<form name="frm">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="CSpecialTlt">
		<td><%=gettopGetDocLinkLngStr("LtxtDocSel")%> - <%=docNames%>&nbsp;</td>
	</tr>
	<tr class="CSpecialTlt">
		<td><%=gettopGetDocLinkLngStr("DtxtClient")%> - <%=CardName%>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="0" id="table2" cellpadding="0" style="width: 100%">
			<tr class="CSpecialTlt2">
				<td style="width: 80px"><%=gettopGetDocLinkLngStr("LtxtDocNum")%>&nbsp;</td>
				<td style="width: 80px"><%=gettopGetDocLinkLngStr("DtxtDate")%>&nbsp;</td>
				<td><%=gettopGetDocLinkLngStr("LtxtComments")%>&nbsp;</td>
				<td style="width: 80px"><%=gettopGetDocLinkLngStr("DtxtDueDate")%>&nbsp;</td>
			</tr>
			<% If Not rd.Eof Then
			do while not rd.eof
			myVal = rd("DocEntry") & "," & rd("DocNum") %>
			<tr class="CSpecialTbl">
				<td style="width: 80px"><a href="#" class="LinkCSpecial" onclick="window.returnValue = '<%=myHTMLEncode(myVal)%>'; window.close()"><%=rd("DocNum")%></a>&nbsp;</td>
				<td style="width: 80px"><a href="#" class="LinkCSpecial" onclick="window.returnValue = '<%=myHTMLEncode(myVal)%>'; window.close()"><%=FormatDate(rd("DocDate"), True)%></a>&nbsp;</td>
				<td><a href="#" class="LinkCSpecial" onclick="window.returnValue = '<%=myHTMLEncode(myVal)%>'; window.close()"><%=rd("Comments")%></a>&nbsp;</td>
				<td style="width: 80px"><a href="#" class="LinkCSpecial" onclick="window.returnValue = '<%=myHTMLEncode(myVal)%>'; window.close()"><%=FormatDate(rd("DocDueDate"), True)%></a>&nbsp;</td>
			</tr>
			<% rd.movenext
			loop
			else %>
			<tr class="CSpecialTbl">
				<td colspan="5">
				<p align="center"><%=gettopGetDocLinkLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
			<tr class="CSpecialTbl">
				<td colspan="5" align="center">
				<input type="button" name="btnCancel" value="<%=gettopGetDocLinkLngStr("DtxtCancel")%>" onclick="javascript:window.close();"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
</body>

</html>

<% conn.close
set rd = nothing %>