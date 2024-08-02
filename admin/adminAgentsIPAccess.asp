<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminAgentsIPAccess.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select IsNull(SlpName, '') SlpName, Case When Exists(select 'A' from OLKAgentsAccess where SlpCode = T0.SlpCode) Then 'Y' Else 'N' End VerfyUser from OSLP T0 where SlpCode = " & Request("SlpCode")
set rs = conn.execute(sql)
SlpName = rs("SlpName")
VerfyUser = rs("VerfyUser")

If Request.Form.Count <> 0 Then
	sql = "declare @SlpCode int set @SlpCode = " & Request("SlpCode") & " "
	
	If Request("delIndex") <> "" Then sql = sql & "delete OLKAgentsAccessIPS where SlpCode = @SlpCode and IPIndex in (" & Request("delIndex") & ") "
	
	If Request("IPFrom0") <> "" Then
		sql = sql & "declare @IPIndex int set @IPIndex = IsNull((select Max(IPIndex)+1 from OLKAgentsAccessIPS where SlpCode = @SlpCode), 0) " & _
		"insert OLKAgentsAccessIPS(SlpCode, IPIndex, IPFrom, IPTo) " & _
		"values(@SlpCode, @IPIndex, N'" & Request("IPFrom0") & "." & Request("IPFrom1") & "." & Request("IPFrom2") & "." & Request("IPFrom3") & "', " & _
		"N'" & Request("IPTo0") & "." & Request("IPTo1") & "." & Request("IPTo2") & "." & Request("IPTo3") & "') "
	End If
	
	If sql <> "" Then conn.execute(sql)
End If

sql = "select IPIndex, IPFrom, IPTo from OLKAgentsAccessIPS where SlpCode = " & Request("SlpCode") & " "
set rs = conn.execute(sql)
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getadminAgentsIPAccessLngStr("LttlAccessIP")%> - <%=SlpName%></title>
<script language="javascript" src="general.js"></script>
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<script language="javascript">
<% If Request("btnSave") <> "" Then %>window.close();<% End If %>
function valKeyDown(o, next, e)
{
	if (e.keyCode == 96 || (e.keyCode >= 48 && e.keyCode <= 57) || (e.keyCode >= 97 && e.keyCode <= 105) || e.keyCode == 110 || 
		e.keyCode == 8 || e.keyCode == 46 || (e.keyCode >= 37 && e.keyCode <= 40) || e.keyCode == 9 || e.keyCode == 16)
	{
		if (e.keyCode == 110)
		{
			if (next != null) next.focus();
			return false;
		}
		else
		{
			return true;
		}
	}
	else
		return false;
}
function doChange(o, e, next)
{
	if (o.value.length == 3 && e.keyCode != 9 && e.keyCode != 16)
	{
		if (parseInt(o.value) > 255)
		{
			alert('<%=getadminAgentsIPAccessLngStr("LtxtVal255")%>');
			o.value = '';
			o.focus();
		}
		else
			if (next != null) next.focus();
	}
}
function valFrm()
{
	<% If Not rs.Eof Then %>
	if (document.frmIP.delIndex.length)
	{
		var delCount = 0;
		for (var i = 0;i<document.frmIP.delIndex.length;i++)
		{
			if (document.frmIP.delIndex[i].checked) delCount++
		}
		if (delCount > 0)
			if (!confirm('<%=getadminAgentsIPAccessLngStr("LtxtConfDelIP")%>'.replace('{0}', delCount))) return false;
	}
	else
	{
		if (document.frmIP.delIndex.checked)
			if (!confirm('<%=getadminAgentsIPAccessLngStr("LtxtConfDelIP")%>'.replace('{0}', '1'))) return false;
	}
	<% End If %>
	
	var arrFrom = new Array(document.frmIP.IPFrom0.value, document.frmIP.IPFrom1.value, document.frmIP.IPFrom2.value, document.frmIP.IPFrom3.value);
	var arrTo = new Array(document.frmIP.IPTo0.value, document.frmIP.IPTo1.value, document.frmIP.IPTo2.value, document.frmIP.IPTo3.value);
	
	if (!valArrIP(arrFrom))
	{
		alert('<%=getadminAgentsIPAccessLngStr("LtxtValFromIP")%>');
		return false;
	}
	else if (!valArrIP(arrTo))
	{
		alert('<%=getadminAgentsIPAccessLngStr("LtxtValToIP")%>');
		return false;
	}
	
	if (document.frmIP.IPFrom0.value != '' && document.frmIP.IPTo0.value == '' ||
		document.frmIP.IPFrom0.value == '' && document.frmIP.IPTo0.value != '')
	{
		alert('<%=getadminAgentsIPAccessLngStr("LtxtValIPFromTo")%>');
		return false;
	}
	
	return true;
}
function valArrIP(arr)
{
	var vCount = 0;
	for (var i = 0;i<arr.length;i++)
	{
		if (arr[i] != '') vCount++;
	}
	return (vCount == 0 || vCount == 4);
}

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
function doSubmit(btn)
{
	switch (btn)
	{
		case 'btnApply':
			document.frmIP.btnApply.value='Y';
			document.frmIP.btnSave.value='';
			break;
		case 'btnSave':
			document.frmIP.btnApply.value='';
			document.frmIP.btnSave.value='Y';
			break;
	}
	document.frmIP.submit();
}
</script>
<style type="text/css">
.style1 {
	text-align: center;
}
.style2 {
	background-color: #E1F3FD;
	font-family: Verdana;
	font-weight: bold;
	font-size: 10px;
	color: #31659C;
	text-align: center;
}
</style>
</head>

<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();" onload="setTblSet();" onscroll="setTblSet();">

<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<form method="POST" action="adminAgentsIPAccess.asp" name="frmIP" onsubmit="return valFrm();" webbot-action="--WEBBOT-SELF--">
	<tr>
		<td class="popupTtl" colspan="4"><%=getadminAgentsIPAccessLngStr("LttlAccessIP")%> - <%=SlpName%></td>
	</tr>
	<tr>
		<td class="style2"><%=getadminAgentsIPAccessLngStr("DtxtFrom")%></td>
		<td class="style2" colspan="2"><%=getadminAgentsIPAccessLngStr("DtxtTo")%></td>
		<td class="style2"><%=getadminAgentsIPAccessLngStr("LtxtDelete")%></td>
	</tr>
	<% do while not rs.eof %>
	<tr class="popupOptValue">
		<td class="style1"><%=rs("IPFrom")%></td>
		<td colspan="2" class="style1"><%=rs("IPTo")%></td>
		<td>
		<p align="center">
		<input type="checkbox" class="noborder" name="delIndex" id="delIndex" value="<%=rs("IPIndex")%>"></td>
	</tr>
	<% rs.movenext
	loop %>
	<tr class="popupOptValue">
		<td class="style1">
		<input type="text" name="IPFrom0" id="IPFrom0" size="3" class="input" onkeydown="return valKeyDown(this, IPFrom1, event);" onkeyup="doChange(this, event, IPFrom1);" maxlength="3">.<input type="text" name="IPFrom1" id="IPFrom1" size="3" class="input" onkeydown="return valKeyDown(this, IPFrom2, event);" onkeyup="doChange(this, event, IPFrom2);" maxlength="3">.<input type="text" name="IPFrom2" id="IPFrom2" size="3" class="input" onkeydown="return valKeyDown(this, IPFrom3, event);" onkeyup="doChange(this, event, IPFrom3);" maxlength="3">.<input type="text" name="IPFrom3" id="IPFrom3" size="3" class="input" onkeydown="return valKeyDown(this, IPTo0, event);" onkeyup="doChange(this, event, IPTo0);" maxlength="3"></td>
		<td colspan="2" class="style1">
		<input type="text" name="IPTo0" id="IPTo0" size="3" class="input" onkeydown="return valKeyDown(this, IPTo1, event);" onkeyup="doChange(this, event, IPTo1);" maxlength="3">.<input type="text" name="IPTo1" id="IPTo1" size="3" class="input" onkeydown="return valKeyDown(this, IPTo2, event);" onkeyup="doChange(this, event, IPTo2);" maxlength="3">.<input type="text" name="IPTo2" id="IPTo2" size="3" class="input" onkeydown="return valKeyDown(this, IPTo3, event);" onkeyup="doChange(this, event, IPTo3);" maxlength="3">.<input type="text" name="IPTo3" id="IPTo3" size="3" class="input" onkeydown="return valKeyDown(this, null, event);" onkeyup="doChange(this, event, null);" maxlength="3"></td>
		<td>
		&nbsp;</td>
	</tr>
	<input type="hidden" name="SlpCode" value="<%=Request("SlpCode")%>">
	<input type="hidden" name="VerfyUser" value="<%=VerfyUser%>">
	<input type="hidden" name="btnApply" value="">
	<input type="hidden" name="btnSave" value="">
	</form>
</table>
<table cellpadding="0" border="0" width="100%" id="tblSave" style="position: absolute; ">
	<tr>
		<td width="75">
		<input type="button" name="btnApply" value="<%=getadminAgentsIPAccessLngStr("DtxtApply")%>" class="OlkBtn" onclick="javascript:doSubmit('btnApply');">
		</td>
		<td width="75">
		<input type="button" name="btnSave" value="<%=getadminAgentsIPAccessLngStr("DtxtSave")%>" class="OlkBtn" onclick="javascript:doSubmit('btnSave');"></td>
		<td><hr size="1"></td>
		<td width="75">
		<input type="button" name="btnCancel" value="<%=getadminAgentsIPAccessLngStr("DtxtCancel")%>" class="OlkBtn" onclick="window.close();"></td>
	</tr>	
</table>
</body>

</html>
<% conn.close
set rs = nothing %>