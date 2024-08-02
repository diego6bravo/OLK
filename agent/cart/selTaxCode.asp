<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/selTaxCode.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../loadAlterNames.asp" -->
<title><%=Replace(LttlSelTax, "{0}", LCase(txtTax))%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<script language="javascript" src="../general.js"></script>
</head>

<body topmargin="0" leftmargin="0">
<%      set rs = Server.CreateObject("ADODB.recordset")
      
      
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			set rs = cmd.execute()
      %>
      
<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<tr class="CSpecialTlt2">
		<td><b><%=getselTaxCodeLngStr("DtxtCode")%></b></td>
		<td><b><%=getselTaxCodeLngStr("DtxtNote")%></b></td>
		<td><b><%=getselTaxCodeLngStr("DtxtRate")%></b></td>
	</tr>
	<% do while not rs.eof %>
	<tr class="GeneralTbl" style="cursor: hand" onmouseover="javascript:this.className='GeneralTblHigh'" onmouseout="javascript:this.className='GeneralTbl'" onclick="javascript:opener.setTaxCode('<%=myHTMLEncode(rs("Code"))%>');window.close()">
		<td><%=rs("Code")%>&nbsp;</td>
		<td><%=rs("Name")%>&nbsp;</td>
		<td>
		<p align="right"><%=FormatNumber(rs("Rate"),myApp.PercentDec)%>&nbsp;</td>
	</tr>
	<% rs.movenext
	loop %>
</table>
      
</body>

</html>