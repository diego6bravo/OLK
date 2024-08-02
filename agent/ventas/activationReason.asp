<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="lang/activationReason.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<%
varx = 0
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select ReasonIndex, ReasonName from OLKAcctRejectNotes order by 2 asc"
set rs = conn.execute(sql)
           %>
<title><%=getactivationReasonLngStr("LtxtAcctActive")%> - <%=getactivationReasonLngStr("LtxtRejReson")%></title>
</head>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="CSpecialTlt">
		<td>&nbsp;<%=getactivationReasonLngStr("LtxtRejReson")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="CSpecialTlt2">
				<td width="120">
				<p align="left"><%=getactivationReasonLngStr("LtxtPredNotes")%>:&nbsp;</td>
				<td>
				<select name="cmbReason" size="1" onchange="javascript:getValue(this.value);" style="width: 267; height:16">
				<option></option>				
				<% do while not rs.eof %>
				<option value="<%=rs(0)%>"><%=myHTMLEncode(rs(1))%></option>
				<% rs.movenext
				loop %>
				</select>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="CSpecialTbl">
		<td colspan="2">
		<textarea cols="40" rows="25" name="Reason" style="width: 100%"><%=myHTMLEncode(Request("note"))%></textarea></td>
	</tr>
	<tr class="CSpecialTlt2">
		<td height="24">
		<p align="center"><input type="button" name="btnAccept" value="<%=getactivationReasonLngStr("DtxtAccept")%>" onclick="javascript:opener.setReason(Reason.value);window.close();"></td>
	</tr>
	</table>
	<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="169" width="167" src=""></iframe>
	<form method="post" target="ifGetValue" name="frmGetValue" action="../topGetValue.asp">
	<input type="hidden" name="Type" value="AcctRejReason">
	<input type="hidden" name="searchStr" value="">
	</form>

<script language="javascript">
<!--
function getValue(ReasonIndex)
{
	if (ReasonIndex != '')
	{
		document.frmGetValue.searchStr.value = ReasonIndex;
		document.frmGetValue.submit();
	}
}
function setValue(src, value, myType){
	Reason.value = value;
	/*if (value != '') 
	{ updFld.value = value; setTargetVal(value); }
	else { if(src == 0)launchSelect(myType, updFld.value); }*/
}
//-->
</script>
</body>
<% conn.close
set rs = nothing %>
</html>