<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminAlertsBranchs.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getadminAlertsBranchsLngStr("LttlAlertBranches")%></title>
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="general.js"></script>
<script language="javascript">
function setTblSet()
{
	if (browserDetect() == 'msie')
	{
		tblSave.style.top = document.body.offsetHeight-31+document.body.scrollTop;
	}
	else if (browserDetect() == 'opera')
	{
		tblSave.style.top = document.body.offsetHeight-27+document.body.scrollTop;
	}
	else //firefox & others
	{
		tblSave.style.top = window.innerHeight-27+document.body.scrollTop;
	}
}
</script>
</head>

<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();" onload="setTblSet();" onscroll="setTblSet();">
<% If Request.Form.Count = 0 Then %>
<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
	<form method="POST" name="frmBranchs" action="adminAlertsBranchs.asp">
	<tr>
		<td colspan="2" class="popupTtl"><%=getadminAlertsBranchsLngStr("DtxtBranches")%></td>
	</tr>
	<%
	sql = "select T0.branchIndex, T0.branchName, "
	
	If Request("branchIndex") <> "" Then
		sql = sql & "Case When T0.branchIndex in (" & Request("branchIndex") & ") Then 'Y' Else 'N' End "
	Else
		sql = sql & " 'N' "
	End If
	
	sql = sql & " Verfy from OLKBranchs T0 "
	
	set rs = conn.execute(sql)
	do while not rs.eof 
	%>
	<tr class="popupOptValue">
		<td width="20"><input <% If rs("Verfy") = "Y" Then %>checked<% End If %> class="noborder" type="checkbox" name="branchIndex" id="branchIndex<%=rs("branchIndex")%>" value="<%=rs("branchIndex")%>"></td>
		<td><p align="left">
		<font size="1" face="Verdana" color="#4783C5"><label for="branchIndex<%=rs("branchIndex")%>"><%=rs("branchName")%></label>&nbsp;</font></td>
	</tr>
	<% rs.movenext
	loop %>
	<tr height="27">
		<td width="20">&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	</form>
</table>
<table cellpadding="0" border="0" width="100%" style="position: absolute; z-index: 1;" id="tblSave" bgcolor="#FFFFFF">
	<tr>
		<td style="width: 75px"><input type="button" value="<%=getadminAlertsBranchsLngStr("DtxtAccept")%>" name="cmdGuardar" class="OlkBtn" onclick="document.frmBranchs.submit();"></td>
		<td>
			<hr color="#0D85C6" size="1"></td>
		<td style="width: 75px"><input type="button" value="<%=getadminAlertsBranchsLngStr("DtxtClose")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
	</tr>
</table>
<% Else %>
<script language="javascript">
opener.acceptBranchs('<%=Request("branchIndex")%>');
window.close();
</script>
<% End If %>
</body>

</html>
<% conn.close
set rs = nothing %>