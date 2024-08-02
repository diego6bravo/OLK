<!--#include file="chkLogin.asp" -->
<!--#include file="lang/adminRepEditSaveAs.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getadminRepEditSaveAsLngStr("LttlRepSaveAs")%></title>
<link rel="stylesheet" type="text/css" href="style/style_pop.css">
<script type="text/javascript">
<!--
function saveCopy()
{
	if (document.frmSaveAs.rsName.value == '')
	{
		alert('<%=getadminRepEditSaveAsLngStr("LtxtValRepNam")%>');
		document.frmSaveAs.rsName.focus();
		return;
	}
	
	opener.saveCopy(document.frmSaveAs.rsName.value, document.frmSaveAs.rgIndex.value);
	window.close();
}
//-->
</script>
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="repVars.inc" -->
<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();">
<form method="post" action="adminRepEditSaveAs.asp" name="frmSaveAs" onsubmit="javascript:saveCopy();return false;" webbot-action="--WEBBOT-SELF--">
<table border="0" width="100%" id="table1" cellpadding="0">
	<tr class="TblGreenTlt">
		<td colspan="2">
			<%=getadminRepEditSaveAsLngStr("LttlRepSaveAs")%>
		</td>
	</tr>
	<tr>
		<td style="width: 100px" class="TblGreenTlt">
				<%=getadminRepEditSaveAsLngStr("DtxtName")%></td>
		<td class="TblGreenNrm">
		<input type="text" name="rsName" size="20" style="width: 100%" maxlength="60" value="<%=Request("rsName")%> (<%=getadminRepEditSaveAsLngStr("DtxtCopy")%>)"></td>
	</tr>
	<tr>
		<td style="width: 100px" class="TblGreenTlt">
				<%=getadminRepEditSaveAsLngStr("DtxtGroup")%></td>
		<td class="TblGreenNrm">
		<% set rd = server.createobject("ADODB.RecordSet")
		sql = "select rgIndex, rgName from " & repTbl & "RG where UserType = '" & Request("UserType") & "' and rgIndex >= 0 order by 2"
		set rd = conn.execute(sql)%>
		<select size="1" name="rgIndex" style="width: 100%;">
		<% do while not rd.eof %>
		<option <% If CInt(Request("rgIndex")) = CInt(rd("rgIndex")) Then %>selected<% End If %> value="<%=rd("rgIndex")%>"><%=myHTMLEncode(rd("rgName"))%></option>
		<% rd.movenext
		loop
		set rd = nothing %>
		</select></td>
	</tr>
	<tr>
		<td colspan="2">
		<table style="width: 100%" cellpadding="0">
			<tr>
				<td style="width: 75px">
				<font color="#4783C5" face="Verdana" size="1"><input type="button" value="<%=getadminRepEditSaveAsLngStr("DtxtSave")%>" name="btnSave" onclick="javascript:saveCopy();" class="OlkBtn"></font></td>
				<td><hr size="1"></td>
				<td style="width: 75px">
				<input type="button" value="<%=getadminRepEditSaveAsLngStr("DtxtCancel")%>" name="btnCancel" onclick="window.close();" class="OlkBtn"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
</body>

</html>
