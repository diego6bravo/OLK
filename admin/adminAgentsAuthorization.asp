<% response.Expires=-1 %>
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="lang/adminAgentsAuthorization.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl" <% End If %>="">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" href="style/style_admin_ie.css">
<title>Untitled 2</title>
<style type="text/css">
.clsA {
	font-family: Verdana;
	font-size: 10px;
	background-color: #F5FBFE; 
	color: #4783C5;
}
.clsT {
	font-family: Verdana;
	font-size: 10px;
	background-color: #E1F3FD; 
	color: #4783C5;
}
li{
	height: 20px;
	list-style-type: none;
}
.style3 {
	background-color: #FFFFFF;
}
.style4 {
	font-size: xx-small;
	color: #4783C5;
}
.clsButton
{
	color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size: 10px; height: 23px; font-weight: bold;
}
</style>
</head>

<%

If Request("dbID") = "" Then 
	dbID = Session("ID") 
Else 
	dbID = CInt(Request("dbID"))
	UserName = Request("UserName")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "OLKSetDBAgentData"
	cmd.Parameters.Refresh
	cmd("@dbID") = dbID
	cmd("@UserName") = UserName
	cmd("@SlpCode") = CInt(Request("SlpCode"))
	cmd("@Access") = Request("Access")
	cmd("@WhsCode") = "##"
	cmd.execute()
End If
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim startSub
set rd = Server.CreateObject("ADODB.RecordSet")

If Request("Access") = "U" Then 
	set rDep = Server.CreateObject("ADODB.RecordSet")
	sql = "select T0.AutID, T0.DepID, dbo.OLKGetAutDepDesc(T0.DepID, " & Session("LanID") & ") DepDesc from OLKAuthorizationDependance T0"
	rDep.open sql, connCommon, 3, 1
End If
         
set cmd = Server.CreateObject("ADODB.Command")
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetAuthorization" & dbID
cmd.ActiveConnection = connCommon
cmd.Parameters.Refresh
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = CInt(Request("SlpCode"))
%>
<script type="text/javascript" src="general.js"></script>
<script type="text/javascript">
<!--

function showSubNode(AutID, Show)
{
	document.getElementById('sub' + AutID).style.display = Show ? '' : 'none';
	document.getElementById('sign' + AutID).innerHTML = Show ? '[-]' : '[+]';
}
function subNode(AutID)
{
	showSubNode(AutID, document.getElementById('sub' + AutID).style.display == 'none');
}
function setTblSet()
{
	if (browserDetect() == 'msie')
	{
		tblSet.style.top = document.body.offsetHeight-33+document.body.scrollTop;
	}
	else if (browserDetect() == 'opera')
	{
		tblSet.style.top = document.body.offsetHeight-46+document.body.scrollTop;
	}
	else if (browserDetect() == 'safari')
	{
		tblSet.style.top = window.innerHeight-30+document.body.scrollTop;
	}
	else //firefox & others
	{
		tblSet.style.top = window.innerHeight-46+document.body.scrollTop;
	}
}
function expandAll()
{
	aut = document.frmUpdAut.TitleIndex;
	for (var i = 0;i<aut.length;i++)
	{
		showSubNode(aut[i].value, true);
	}
	scroll(0, 0);
}
function collapseAll()
{
	aut = document.frmUpdAut.TitleIndex;
	for (var i = 0;i<aut.length;i++)
	{
		showSubNode(aut[i].value, false);
	}
	scroll(0, 0);
}
function fullAut()
{
	var chk = document.frmUpdAut.AutID;
	for (var i = 0;i<chk.length;i++)
	{
		if (!document.getElementById('chkAut' + chk[i].value).disabled)
			document.getElementById('chkAut' + chk[i].value).checked = true;
	}
	var CheckID = document.frmUpdAut.CheckID;
	for (var i = 0;i<CheckID.length;i++)
	{
		document.getElementById(CheckID[i].value).checked = true;
	}
	expandAll();
}
function noAut()
{
	var chk = document.frmUpdAut.AutID;
	for (var i = 0;i<chk.length;i++)
	{
		document.getElementById('chkAut' + chk[i].value).checked = false;
	}
	var CheckID = document.frmUpdAut.CheckID;
	for (var i = 0;i<CheckID.length;i++)
	{
		document.getElementById(CheckID[i].value).checked = false;
	}
	expandAll();
}
function chkMaxDisc(fld, old)
{
	if (!IsNumeric(fld.value))
	{
		alert('<%=getadminAgentsAuthorizationLngStr("DtxtValNumVal")%>');
		fld.value = old.value;
	}
	else if (parseFloat(fld.value) > 100)
	{
		alert('<%=getadminAgentsAuthorizationLngStr("DtxtValNumMaxVal")%>'.replace('{0}', 100));
		fld.value = 100;
	}
	else if (parseFloat(fld.value) < 0)
	{
		alert('<%=getadminAgentsAuthorizationLngStr("DtxtValNumMinVal")%>'.replace('{0}', 0));
		fld.value = 0;
	}
	fld.value = formatNumber(fld.value, <% If Request("dbID") = "" Then Response.Write myApp.PercentDec Else Response.Write 6%>);
	old.value = fld.value;
}
function chkDep(chk, AutID)
{
	if (chk.checked && document.getElementById('Dep' + AutID) != null)
	{
		var depStr = document.getElementById('Dep' + AutID).value;
		if (depStr != '')
		{
			arrDep = depStr.split('|');
			var found = false;
			for (var i=0;i<arrDep.length;i++)
			{
				varID = arrDep[i].split('{S}')[0];
				if (document.getElementById('chkAutS' + varID).checked)
				{
					found = true;
					break;
				}
			}
			if (!found)
			{
				var optDesc = '';
				for (var i=0;i<arrDep.length;i++)
				{
					if (optDesc != '') optDesc += '\n';
					
					varDesc = arrDep[i].split('{S}')[1];
					optDesc += varDesc;
				}
				alert('<%=getadminAgentsAuthorizationLngStr("LtxtAlertDep")%>'.replace('{0}', optDesc));
			}
		}
	}
}
function enableSave()
{
	document.getElementById('btnSave').disabled = false;
	document.getElementById('btnCopy').disabled = true;
}
function enableCopy()
{
	document.getElementById('btnSave').disabled = true;
	document.getElementById('btnCopy').disabled = false;
}

//-->
</script>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onload="setTblSet();" onscroll="setTblSet();">

<table border="0" cellpadding="0" width="100%" id="tblAut">
	<form method="POST" name="frmUpdAut" action="autSubmit.asp" target="iFrameSave">
	<% If Request("dbID") <> "" Then %><input type="hidden" name="UserName" value="<%=Server.HTMLEncode(UserName)%>">
	<input type="hidden" name="dbID" value="<%=Request("dbID")%>"><% End If %>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td colspan="2">
					<table cellpadding="0" border="0" width="100%">
						<tr>
							<td>
							<ul style="margin-left: 0px; margin-right: 0px; ">
							<% If Request("Access") = "U" Then %>
							<% GenSubAut "S-1", "U" %>
							<% Else %>
							<% GenSubAut "S8", "P" %>
							<% End If %>
							</ul>
							</td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td></td>
		</tr>
		<input type="hidden" name="SlpCode" value="<%=Request("SlpCode")%>">
		<input type="hidden" name="submitCmd" value="UserAut">
		<input type="hidden" name="Access" value="<%=Request("Access")%>">
		<input type="hidden" name="parent" value="Y">
	</form>
</table>
<table cellpadding="0" border="0" cellspacing="0" width="100%" id="tblSet" style="position: absolute; z-index: 1;" class="style3">
	<tr>
		<td>
		<table width="100%">
			<tr>
				<td style="width: 75px">
				<input type="button" value="<%=getadminAgentsAuthorizationLngStr("DtxtSave")%>" name="btnSave" id="btnSave" disabled style="width: 75px; " class="clsButton" onclick="document.frmUpdAut.submit();"></td>
				<td style="width: 75px">
				<input type="button" value="<%=getadminAgentsAuthorizationLngStr("LtxtCopyTo")%>..." name="btnCopy" id="btnCopy" style="width: 75px; " class="clsButton" onclick="javascript:parent.copyAut(<% If Request("dbID") = "" Then Response.Write Request("SlpCode") Else Response.Write Request("dbID") %>, '<%=Request("Access")%>');"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td style="width: 125px">
				<input type="button" value="<%=getadminAgentsAuthorizationLngStr("DtxtExpandAll")%>" style="width: 125px; " class="clsButton" onclick="expandAll();"></td>
				<td style="width: 125px">
				<input type="button" value="<%=getadminAgentsAuthorizationLngStr("DtxtCollapseAll")%>" style="width: 125px; " class="clsButton" onclick="collapseAll();"></td>
				<% If Request("Access") = "U" Then %>
				<td style="width: 125px">
				<input type="button" value="<%=getadminAgentsAuthorizationLngStr("LtxtFullAut")%>" style="width: 125px; " class="clsButton" onclick="fullAut();enableSave();"></td>
				<td style="width: 125px">
				<input type="button" value="<%=getadminAgentsAuthorizationLngStr("LtxtNoAut")%>" style="width: 125px; " class="clsButton" onclick="noAut();enableSave();"></td>
				<% End If %>
			</tr>
		</table>
		</td>
	</tr>
</table>
<% Sub GenSubAut(ByVal ParentID, ByVal Access)
cmd("@ParentID") = ParentID
cmd("@Access") = Access
set rAut = Server.CreateObject("ADODB.RecordSet")
set rAut = cmd.execute

If rAut.State = adStateOpen Then
do while not rAut.eof
Select Case rAut("AutID")
	Case "F_1"
		Name = getadminAgentsAuthorizationLngStr("DtxtForms")
	Case "F_2"
		Name = getadminAgentsAuthorizationLngStr("DtxtAgent")
	Case "F_3"
		Name = getadminAgentsAuthorizationLngStr("DtxtPocket")
	Case Else
		Name = rAut("AutName")
End Select %>
<li>
<% If rAut("Type") = "A" Then %><input type="hidden" name="AutID" value="<%=rAut("AutID")%>">
<% If Request("Access") = "U" Then
depStr = ""
If Left(rAut("AutID"), 1) = "S" Then
	rDep.Filter = "AutID = " & Right(rAut("AutID"), Len(rAut("AutID"))-1)
	do while not rDep.eof
		If depStr <> "" Then depStr = depStr & "|"
		depStr = depStr & rDep("DepID") & "{S}" & rDep("DepDesc")
	rDep.movenext
	loop
End If %><input type="hidden" name="Dep<%=rAut("AutID")%>" id="Dep<%=rAut("AutID")%>" value="<%=depStr%>"><% End If %><% End If %>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<% If startSub Then %><tr><td style="font-size: 1pt; ">&nbsp;</td></tr><% End If %>
<tr class="cls<%=rAut("Type")%>" height="20" onmouseover="window.status='<%=rAut("AutID")%>';" onmouseout="window.status='';">
	<% If Request("Access") = "U" or rAut("Type") = "T" Then %><td <% If rAut("Type") = "T" Then %> onclick="subNode('<%=rAut("AutID")%>');" style="cursor: hand; " <% End If %> width="20"><% If rAut("Type") = "A" Then %><input type="checkbox" <% If rAut("ViewOnly") = "Y" Then %>disabled<% End If %> name="chkAut<%=rAut("AutID")%>" <% If rAut("Verfy") = "Y" Then %>checked<% End If %> id="chkAut<%=rAut("AutID")%>" class="noborder" value="Y" onclick="chkDep(this, '<%=rAut("AutID")%>');enableSave();"><% Else %><span id="sign<%=rAut("AutID")%>">[+]</span><% End If %></td><% End If %>
	<td <% If rAut("Type") = "T" Then %> onclick="subNode('<%=rAut("AutID")%>');" style="cursor: hand; " <% End If %> width="260"><% If rAut("Type") = "A" Then %><label for='chkAut<%=rAut("AutID")%>'><% End If %><nobr><%=Name%></nobr><% If rAut("Type") = "A" Then %></label><% End If %></td>
	<% If rAut("Series") = "Y" Then %>
	<td width="200"><% GetSeries rAut("AutID"), rAut("ObjectCode"), rAut("SeriesValue"), ""  %></td>
	<% End If
	If rAut("Series") = "Y" and rAut("AutID") = "S35" Then %>
	<td width="10" align="center">&nbsp;/&nbsp;</td>
	<td width="200"><% GetSeries rAut("AutID"), 24, rAut("SeriesValue2"), "2"  %></td>
	<% End If %>
	<% If Request("Access") = "U" Then %>
	<% If rAut("ObjConf") = "Y" Then %>
	<td width="20"><input class="noborder" type="checkbox" name="chkConf<%=rAut("AutID")%>" value="Y" id="chkConf<%=rAut("AutID")%>" <% If rAut("Confirm") = "Y" Then %>checked<% End If %> onclick="javascript:enableSave();"></td>
	<td width="60"><label for="chkConf<%=rAut("AutID")%>"><%=getadminAgentsAuthorizationLngStr("DtxtConfirm")%></label></td>
	<input type="hidden" name="CheckID" value="chkConf<%=rAut("AutID")%>">
	<% End If
	If rAut("ObjView") = "Y" Then %>
	<td width="20"><input class="noborder" type="checkbox" name="chkAutView<%=rAut("AutID")%>" value="Y" id="chkAutView<%=rAut("AutID")%>" <% If rAut("AutView") = "Y" Then %>checked<% End If %> onclick="javascript:enableSave();"></td>
	<td width="120"><nobr><label for="chkAutView<%=rAut("AutID")%>"><%=getadminAgentsAuthorizationLngStr("LtxtAutView")%></label></nobr></td>
	<input type="hidden" name="CheckID" value="chkAutView<%=rAut("AutID")%>">
	<% End If %>
	<% If rAut("AutID") = "S68" or rAut("AutID") = "S91" Then %>
	<td width="120"><nobr><%=getadminAgentsAuthorizationLngStr("LtxtMaxDiscount")%></nobr></td>
	<td width="100"><input type="text" name="Max<% If rAut("AutID") = "S91" Then Response.Write "Doc" %>Discount" size="6" class="input" style="text-align: right; " value='<%=FormatNumber(CDbl(rAut("MaxDiscount")), myApp.PercentDec)%>' onfocus="this.select()" onchange="chkMaxDisc(this, oldMaxDisc<% If rAut("AutID") = "S91" Then Response.Write "Doc" %>);enableSave();">
	<input type="hidden" name="oldMaxDisc<% If rAut("AutID") = "S91" Then Response.Write "Doc" %>" value="<%=FormatNumber(CDbl(rAut("MaxDiscount")), myApp.PercentDec)%>"></td>
	<% End If %>
	<% End If %>
	<td>&nbsp;</td>
</tr>
</table>
</li>
<% startSub = False
If rAut("Type") = "T" Then
startSub = True %>
<input type="hidden" id="TitleIndex" value="<%=rAut("AutID")%>">
<ul id='sub<%=rAut("AutID")%>' style="display: none; ">
<% GenSubAut rAut("AutID"), Access %>
</ul>
<% End If  %>
<% rAut.movenext
loop
End If
End Sub
Sub GetSeries(AutID, ObjCode, Series, AddID)
set rd = Server.CreateObject("ADODB.RecordSet")
GetQuery rd, 4, ObjCode, null %>
<select size="1" name="Series<%=AutID%><%=AddID%>" class="input" style="width: 100%; " onchange="enableSave();">
<option value="-1"><%=getadminAgentsAuthorizationLngStr("DtxtDefault")%></option>
	<%  do While NOT RD.EOF %>
		<option <% If rd("Series") = Series Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
	<%  RD.MoveNext
	loop    %>
</select>
<% End Sub %>
<iframe name="iFrameSave" id="iFrameSave" style="display: none;"></iframe>
</body>

</html>
