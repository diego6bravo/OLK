<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminRepExport.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="repVars.inc"--><% 

set rd = Server.CreateObject("ADODB.RecordSet")
sql = 	"select rgIndex, rgName, SuperUser, (select count('A') from " & repTbl & "RS where rgIndex = " & repTbl & "RG.rgIndex) rsCount " & _
		"from " & repTbl & "RG where UserType = '" & Request("UserType") & "' order by rgName asc"
rd.open sql, conn, 3, 1

set rs = Server.CreateObject("ADODB.RecordSet")
sql = 	"select rsIndex, rsName + Case LinkOnly When 'N' Then N'' Else N' (" & getadminRepExportLngStr("DtxtLink") & ")' end rsName, rgName, Case SuperUser When 'Y' Then 'Super Usuario' When 'N' Then 'Todos' End Access, T0.Active, T0.rgIndex " & _
		"from " & repTbl & "rs T0 " & _
		"inner join " & repTbl & "rg T1 on T1.rgIndex = T0.rgIndex " & _
		"where T1.UserType = '" & Request("UserType") & "' " & _
		" order by rgName, rsName"
rs.open sql, conn, 3, 1
If Request("rgIndex") <> "" Then rs.Filter = "rgIndex = " & Request("rgIndex")
%>
<html <% if session("rtl") <> "" then %>dir="rtl" <% end if %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getadminRepExportLngStr("LttlRepExp")%></title>

<link rel="stylesheet" type="text/css" title="admin" href="style/style_admin_<%=Session("style")%>.css">
<script language="javascript" src="<% If repTbl = "TMRP" Then %><% End If %>general.js"></script>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onbeforeunload="opener.clearWin();">
<script language="javascript">
function valFrm()
{
	if (document.frmExp.rsIndex.length)
	{
		for (var i = 0;i<document.frmExp.rsIndex.length;i++)
		{
			if (document.frmExp.rsIndex[i].checked) 
			{
				doSaveChanged();
				return true;
			}
		}
	}
	else
	{
		if (document.frmExp.rsIndex.checked) 
		{
			doSaveChanged();
			return true;
		}
	}	
	alert('<%=getadminRepExportLngStr("LtxtValSelRep")%>');
	return false;
}

function doSaveChanged()
{
	trOK.style.display = '';
	tdSave.style.display = 'none';
	trAll.style.display = 'none';
	document.frmExp.btnCancel.value = '<%=getadminRepExportLngStr("DtxtClose")%>';
	if (document.frmExp.rsIndex.length)
	{
		for (var i = 0;i<document.frmExp.rsIndex.length;i++)
		{
			if (!document.frmExp.rsIndex[i].checked) 
			{
				document.getElementById('tr' + document.frmExp.rsIndex[i].value).style.display = 'none';
			}
		}
	}
	else
	{
		if (!document.frmExp.rsIndex.checked) 
		{
			document.getElementById('tr' + document.frmExp.rsIndex.value).style.display = 'none';
		}
	}	
}

function chkAllRS(chk)
{
	if (document.frmExp.rsIndex.length)
	{
		for (var i = 0;i<document.frmExp.rsIndex.length;i++)
		{
			document.frmExp.rsIndex[i].checked = chk;
		}
	}
	else
	{
		document.frmExp.rsIndex.checked = chk;
	}
}

function chkRS()
{
	retVal = true;
	if (document.frmExp.rsIndex.length)
	{
		for (var i = 0;i<document.frmExp.rsIndex.length;i++)
		{
			if (!document.frmExp.rsIndex[i].checked)
			{
				retVal = false;
				break;
			}
		}
	}
	else
	{
		retVal = document.frmExp.rsIndex.checked;
	}
	document.frmExp.chkAll.checked = retVal;
}
</script>
<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<form method="POST" action="adminRepExportSave.asp" name="frmExp" onsubmit="return valFrm();">
		<tr class="TblRepTlt">
			<td colspan="3">
			<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
				<tr class="TblRepTlt">
					<td><%=getadminRepExportLngStr("LttlRepExp")%></td>
					<td>
					<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<select name="rgIndex" size="1" onchange="window.location.href='?UserType=<%=Request("UserType")%>&amp;rgIndex='+this.value">
					<option value><%=getadminRepExportLngStr("DtxtAll")%></option>
					<% If Not rd.Eof Then 
					do while not rd.eof %>
					<option <% if request("rgindex") = cstr(rd("rgindex")) then %>selected<% end if %> value="<%=rd("rgIndex")%>">
					<%=myHTMLEncode(rd("rgName"))%></option>
					<% rd.movenext
					loop
					rd.movefirst
					End If %></select></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr class="TblRepTlt">
			<td width="10">
			&nbsp;</td>
			<% If Request("rgIndex") = "" Then %>
			<td class="style1">
			<%=getadminRepExportLngStr("DtxtGroup")%></td><% End If %>
			<td class="style1">
			<%=getadminRepExportLngStr("DtxtName")%></td>
		</tr>
		<% 
		If Not rs.Eof Then
		do while not rs.eof
		rsIndex = rs("rsIndex") %>
		<tr id="tr<%=rsIndex%>" class="TblRepTbl">
			<td width="10">
			<input type="checkbox" name="rsIndex" value="<%=rsIndex%>" id="rsIndex<%=rsIndex%>" class="noborder" onclick="javascript:chkRS();"></td>
			<% If Request("rgIndex") = "" Then %>
			<td>
			<label for="rsIndex<%=rsIndex%>"><%=rs("rgName")%></label>&nbsp;</td><% End If %>
			<td>
			<label for="rsIndex<%=rsIndex%>"><%=rs("rsName")%></label></td>
		</tr>
		<% rs.movenext
		loop %>
		<tr id="trAll" class="TblRep<% If Alter Then %>A<% End If %>Tbl">
			<td colspan="2">
			<input type="checkbox" name="chkAll" value="Y" onclick="javascript:chkAllRS(this.checked);" class="noborder" id="chkAll"><label for="chkAll"><%=getadminRepExportLngStr("DtxtAll")%></label></td>
			<td>
			&nbsp;</td>
		</tr>
		<tr id="trOK" style="display: none;" class="TblRepTbl">
			<td colspan="3" align="center">
			<%=getadminRepExportLngStr("LtxtDataExpOK")%></td>
		</tr>
		<tr>
			<td bordercolor="#C8E7E8" colspan="3">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table3">
				<tr>
					<td width="75" id="tdSave">
			<input type="submit" name="btnSave" value="<%=getadminRepExportLngStr("DtxtSave")%>" class="BtnRep"></td>
					<td>
			<hr size="1">
					</td>
					<td width="75">
					<p align="right">
			<input type="button" name="btnCancel" value="<%=getadminRepExportLngStr("DtxtCancel")%>" class="BtnRep" onclick="window.close();"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="UserType" value="<%=Request("UserType")%>">
		<% Else %>
		<tr>
			<td bordercolor="#C8E7E8" colspan="3">
			<p align="center">
			<%=getadminRepExportLngStr("DtxtNoData")%></p>
			</td>
		</tr>
		<% End If %>
		<input type="hidden" name="pop" value="Y">
	</form>
</table>

</body>

</html>