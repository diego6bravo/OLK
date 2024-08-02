<!--#include file="top.asp" -->
<!--#include file="lang/adminAgentsAccess.asp" -->

<head>
<%
conn.execute("use [" & Session("OLKDB") & "]")
SQL = 	"select T0.SlpCode, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) slpname,  " & _
		"UserName, EMail, EMailInbox, Password, Access, WhsCode, AsignedSLP, T1.Admin " & _
		"from OSLP T0 " & _
		"inner join OLKAgentsAccess T1 on T1.SlpCode = T0.SlpCode where T0.slpcode <> -1 " & _
		"order by 2"
rs.open sql, conn, 3, 1

set rw = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
rw.open cmd, , 3, 1
If Not rw.eof then EnableWhs = True Else EnableWhs = False
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
function Pic(name, page, w, h, s, r) {
var winleft = (screen.width - w) / 2;
var winUp = (screen.height - h) / 2;
OpenWin = this.open(page, name, "toolbar=no,menubar=no,location=no,left="+winleft+",top="+winUp+",scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
//OpenWin.focus()
}

<% If Request("err") = "-2147217873" Then %>
alert("<%=getadminAgentsAccessLngStr("LtxtValUsr")%>")
<% End If %>
function valFrmUpdAccess()
{
	<% If rs.recordcount > 1 Then %>
	uName = document.frmUpdAccess.UserName;
	uMail = document.frmUpdAccess.EMail;
	uAdmin = document.frmUpdAccess.Admin;
	for (var i = 0;i<uName.length;i++)
	{
		if (uName(i).value == '') {
			alert('<%=getadminAgentsAccessLngStr("LtxtValUName")%>');
			uName(i).focus();
			return false;
		}
		else if (!chkName(uName(i).value))
		{
			alert('<%=getadminAgentsAccessLngStr("LtxtVarUNamAsign")%>'.replace('{0}', uName(i).value));
			uName(i).focus();
			return false;
		}
		else if (uMail(i).value != '' && !emailCheck(uMail(i).value))
		{
			alert('<%=getadminAgentsAccessLngStr("LtxtValEMail")%>');
			uMail(i).focus();
			return false;
		}
		else if (uAdmin(i).checked && uMail(i).value == '')
		{
			alert('<%=getadminAgentsAccessLngStr("LtxtAdminMail")%>');
			uMail(i).focus();
			return false;
		}
	}
	<% ElseIf rs.recordcount = 1 Then %>
	if (document.frmUpdAccess.UserName.value == '') {
		alert('<%=getadminAgentsAccessLngStr("LtxtValUName")%>');
		document.frmUpdAccess.UserName.focus();
		return false;
	}
	else if (document.frmUpdAccess.EMail.value != '' && !emailCheck(document.frmUpdAccess.EMail.value))
	{
		alert('<%=getadminAgentsAccessLngStr("LtxtValEMail")%>');
		document.frmUpdAccess.EMail.focus();
		return false;
	}
	else if (document.frmUpdAccess.Admin.checked && document.frmUpdAccess.EMail.value == '')
	{
		alert('<%=getadminAgentsAccessLngStr("LtxtAdminMail")%>');
		document.frmUpdAccess.EMail.focus();
		return false;
	}
	<% End If %>
	return true;
}

function chkName(n)
{
	var c = 0;
	uName = document.frmUpdAccess.UserName;
	for (var i = 0;i<uName.length;i++)
	{
		if (uName(i).value == n)c++;
	}
	if (c<=1) { return true; } else { return false; }
}
function changeAccess(value, SlpCode)
{
	//document.getElementById('btnEditGroups' + SlpCode).style.display = value == 'D' ? 'none' : '';
	document.getElementById('lnkAut' + SlpCode).style.display = value == 'D' ? 'none' : '';
	if (document.getElementById('Aut' + SlpCode).style.display != 'none') goUserAut(SlpCode);
}
function editRGAccess(SlpCode)
{
	Pic('RGAccess', 'adminAgentsRGAccess.asp?SlpCode=' + SlpCode+'&pop=Y', 400, 300, 'Y', 'N');
}
function editIPAccess(SlpCode)
{
	Pic('IPAccess', 'adminAgentsIPAccess.asp?SlpCode=' + SlpCode+'&pop=Y', 400, 300, 'Y', 'N');
}

function doDelUsr(UserName)
{
	if(confirm('<%=getadminAgentsAccessLngStr("LtxtConfDelUsr")%>'.replace('{0}', UserName)))
		window.location.href='adminSubmit.asp?submitCmd=vUserRem&UName=' + UserName;
}
</script>
<style type="text/css">
.style1 {
	text-align: center;
	font-weight: bold;
	background-color: #E1F3FD;
}
.style2 {
	background-color: #F5FBFE;
}
.style3 {
	background-color: #E1F3FD;
}
.style4 {
	font-weight: bold;
	background-color: #E1F3FD;
}
.style5 {
	text-align: center;
	background-color: #F5FBFE;
}
.style6 {
	text-align: center;
	background-color: #E1F3FD;
}
.style7 {
	text-align: center;
	background-color: #F3FBFE;
}
.style8 {
	background-color: #F3FBFE;
}
</style>
<script language="javascript">
function expandUserAut(SlpCode)
{
	if (document.getElementById('Aut' + SlpCode).style.display == 'none' && document.getElementById('Access' + SlpCode).value != 'D')
	{
		showUserAut(SlpCode, true);
		if (document.getElementById('iFrame' + SlpCode).src == '')
		{
			goUserAut(SlpCode);
		}
	}
	else
	{
		showUserAut(SlpCode, false);
	}
}

function goUserAut(SlpCode)
{
	document.getElementById('iFrame' + SlpCode).src = 'adminAgentsAuthorization.asp?SlpCode=' + SlpCode + '&parent=Y&Access=' + document.getElementById('Access' + SlpCode).value;
}

function showUserAut(SlpCode, Show)
{
	document.getElementById('Aut' + SlpCode).style.display = Show ? '' : 'none';
	document.getElementById('lnkAut' + SlpCode).innerHTML = Show ? '[-]' : '[+]';	
}
function copyAut(SlpCode, Access)
{
	var winleft = (screen.width - 400) / 2;
	var winUp = (screen.height - 300) / 2;
	OpenWin = this.open('adminAgentsAutCopy.asp?SlpCode=' + SlpCode+ '&pop=Y&Access=' + Access, 'CopyAut', "toolbar=no,menubar=no,location=no,left="+winleft+",top="+winUp+",scrollbars=1,resizable=0, width=400,height=300");
}
</script>
</head>

<table border="0" cellpadding="0" width="100%" id="table6">
	<tr>
		<td height="15"></td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmUpdAccess" onsubmit="return valFrmUpdAccess();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("LtxtAgAccDef")%> </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminAgentsAccessLngStr("LtxtAgAccNote")%> </font></td>
	</tr>
	
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td colspan="2">
					<table cellpadding="0" border="0" width="100%">
						<tr>
							<td class="style1" style="width: 20px">
							&nbsp;</td>
							<td class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAgent")%></font></td>
							<td width="130" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtUser")%></font></td>
							<td width="220" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtEMail")%></font></td>
							<td width="110" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("LtxtMsgAlr")%></font></td>
							<td width="110" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAccess")%></font></td>
							<td class="style1" style="width: 200px">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtWarehouse")%></font></td>
							<td class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAdministrator")%></font></td>
							<% If 1 = 2 Then %>
							<td width="120" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("LtxtAsignClient")%></font></td>
							<td align="center" class="style4">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtGroups")%></font></td>
							<% End If %>
							<td align="center" class="style4">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtIP")%></font></td>
							<td style="width: 16px" class="style3">
							&nbsp;</td>
						</tr>
						<% do While NOT RS.EOF %>
						<tr bgcolor="#F3FBFE">
							<td class="style2" style="width: 20px">
							<font face="Verdana" size="1" color="#3366CC">
							<a href="#" id="lnkAut<%=rs("SlpCode")%>" onclick="javascript:expandUserAut(<%=rs("SlpCode")%>);" style="text-decoration: none;<% If rs("Access") = "D" Then %>display: none;<% End If %> ">[+]</a></font></td>
							<td class="style2">
							<a href="#" onclick="javascript:expandUserAut(<%=rs("SlpCode")%>);" style="text-decoration: none; ">
							<font face="Verdana" size="1" color="#4783C5"><%=Server.HTMLEncode(RS("SLPName"))%></font></a></td>
							<td width="130" class="style2">
							<table><tr><td><a href="#" onclick="javascript:Pic('userPwd', 'adminAgentsAccessPwd.asp?UName=<%=Replace(rs("Username"), "'", "\'")%>', 260, 140, 'no', 'no')"><img border="0" src="images/img_lock.gif" width="15" height="15"></a></td><td><font face="Verdana" size="1" color="#3366CC">
								<input type="text" value="<%=Server.HTMLEncode(RS("UserName"))%>" id="UserName" name="UserName<%=rs("SlpCode")%>" class="input" onkeydown="return chkMax(event, this, 20);" size="16"></font></td></tr></table></td>
							<td width="220" class="style2">
							<input type="text" value="<% If Not IsNull(RS("EMail")) Then %><%=Server.HTMLEncode(RS("EMail"))%><% End If %>" id="EMail" name="EMail<%=rs("SlpCode")%>" class="input" onkeydown="return chkMax(event, this, 100);" size="34"></td>
							<td width="110" class="style2" align="center">
							<input type="checkbox" value="Y" name="chkEMailInbox<%=rs("SlpCode")%>" <% If rs("EMailInbox") = "Y" Then %>checked<% End If %> class="noborder"></td>
							<td width="110" class="style2">
							<select name="Access<%=RS("SLPCode")%>" id="Access<%=RS("SLPCode")%>" size="1" class="input" onchange="changeAccess(this.value,<%=rs("SlpCode")%>);">
							<option value="D" <% If Rs("Access") = "D" Then %>selected<%end if %>><%=getadminAgentsAccessLngStr("DtxtDisabled")%></option>
							<option value="U" <% If Rs("Access") = "U" Then %>selected<%end if %>><%=getadminAgentsAccessLngStr("DtxtUser")%></option>
							<option value="P" <% If Rs("Access") = "P" Then %>selected<%end if %>><%=getadminAgentsAccessLngStr("LtxtSuperUser")%></option>
							</select></td>
							<td class="style2" style="width: 200px">
							<select name="WhsCode<%=RS("SLPCode")%>" size="1" class="input">
							<option value="##"><%=getadminAgentsAccessLngStr("DtxtDefault")%></option>
							<% If EnableWhs Then
							do while not rw.eof %>
							<option <% If rs("WhsCode") = rw("WhsCode") Then %>selected<% End If %> value="<%=rw("WhsCode")%>"><%=myHTMLEncode(rw("WhsName"))%></option>
							<% rw.movenext
							loop
							rw.movefirst
							end if %>
							</select></td>
							<td width="120" class="style2">
							<p align="center">
							<input type="checkbox" id="Admin" name="Admin<%=RS("SLPCode")%>" <% If rs("Admin") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<% If 1 = 2 Then %>
							<td width="120" class="style2">
							<p align="center">
							<input type="checkbox" name="AsignedSLP<%=RS("SLPCode")%>" <% If rs("AsignedSLP") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<td class="style5">
							<input type="button" value="<%=getadminAgentsAccessLngStr("DtxtEdit")%>" name="btnEditGroups<%=RS("SLPCode")%>" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; font-weight:bold<% If rs("Access") <> "U" Then %>;display: none<% End If %>" onclick="editRGAccess(<%=rs("SlpCode")%>);"></td>
							<% End If %>
							<td class="style5">
							<input type="button" value="<%=getadminAgentsAccessLngStr("DtxtEdit")%>" name="btnEditIP<%=RS("SLPCode")%>" onclick="editIPAccess(<%=rs("SlpCode")%>);" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; font-weight:bold"></td>
							<td style="width: 16px" class="style2">
							<p align="right">
							<a href="javascript:doDelUsr('<%=Replace(rs("UserName"), "'", "\'")%>');"><img border="0" src="images/remove.gif" width="16" height="16"></a></td>
						</tr>
						<tr bgcolor="#F3FBFE" id="Aut<%=rs("SlpCode")%>" style="display: none;">
							<td class="style2" colspan="10">
							<iframe id="iFrame<%=rs("SlpCode")%>" width="100%" height="350" scrolling="yes"></iframe></td>
						</tr>
						<% RS.MoveNext
						loop %>
						</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="0" border="0" width="100%">
				<td width="75"><% If rs.recordcount > 0 then %><input type="submit" value="<%=getadminAgentsAccessLngStr("DtxtSave")%>" name="B1" class="OlkBtn"><% End If %></td>
				<td><hr color="#0D85C6" size="1"></td>
			</table>
		</td>
	</tr>
		<input type="hidden" name="submitCmd" value="vUserUpd">
	</form>
</table>
<table cellpadding="0" border="0" width="100%">
	<form method="POST" action="adminSubmit.asp" onsubmit="return valFrm()" name="frmAddUser">
	<tr>
		<td bgcolor="#E1F3FD" colspan="2"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAdd")%>&nbsp;<%=getadminAgentsAccessLngStr("DtxtAgent")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" colspan="2">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"> <%=getadminAgentsAccessLngStr("LtxtAddUserNote")%></font></td>
	</tr>
	<tr>
		<td colspan="2">
			<table cellpadding="0" border="0" width="100%">
				<tr>
					<td width="200" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAgent")%></font></td>
					<td width="130" class="style6">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtUser")%></font></b>&nbsp;</td>
					<td width="220" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtEMail")%></font></td>
					<td width="110" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("LtxtMsgAlr")%></font></td>
					<td width="110" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAccess")%></font></td>
					<td width="180" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtWarehouse")%></font></td>
					<td width="180" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtAdministrator")%></font></td>
					<% If 1 = 2 Then %>
					<td width="120" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("LtxtAsignClient")%></font></td>
					<td align="center" class="style4">
							<font face="Verdana" size="1" color="#31659C"><%=getadminAgentsAccessLngStr("DtxtGroups")%></font></td>
					<% End If %>							
					<td class="style3">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="200" height="23" class="style8">
					<select name="SlpCode" size="1" class="input" style="width: 100%;"><%
					sql = 	"select slpcode, " & _
							"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) slpname " & _
							"from oslp T0 " & _
							"where not exists(select slpcode from olkagentsaccess where slpcode = T0.slpcode) and slpcode <> '-1' " & _
							"order by 2"
					set rs = conn.execute(sql)
					do while not rs.eof
					slp = True
					%><option value="<%=rs("slpcode")%>"><%=myHTMLEncode(rs("slpname"))%></option><% rs.movenext
					loop %>
					</select></td>
					<td width="130" height="23" class="style8">
					<table><tr><td><a href="#" onclick="javascript:Pic('userPwd', 'adminAgentsAccessPwd.asp?NewName=' + document.frmAddUser.UserName.value, 260, 140, 'no', 'no')"><img border="0" src="images/img_lock.gif" width="15" height="15"></a><input type="hidden" name="Password" value=""><input type="hidden" name="ChangePwd" value="N"></td><td><font face="Verdana" size="1" color="#3366CC">
						<input type="text" value="" name="UserName" class="input" onkeydown="return chkMax(event, this, 20);" size="16"></font></td></tr></table>
					</td>
					<td width="220" height="23" class="style8">
					<font face="Verdana" size="1" color="#3366CC">
					<input type="text" name="EMail" class="input" onkeydown="return chkMax(event, this, 100);" size="34"></font></td>
					<td width="120" height="23" class="style8">
					<p align="center">
					<input type="checkbox" name="EMailInbox" value="Y" class="noborder"></td>
					<td width="110" height="23" class="style8">
					<select name="Access" size="1" class="input" onchange="//document.frmAddUser.btnEditGroups.style.display=this.value!='U' ? 'none' : '';">
					<option value="D"><%=getadminAgentsAccessLngStr("DtxtDisabled")%></option>
					<option value="U"><%=getadminAgentsAccessLngStr("DtxtUser")%></option>
					<option value="P"><%=getadminAgentsAccessLngStr("LtxtSuperUser")%></option>
					</select></td>
					<td width="180" height="23" class="style8">
					<select name="WhsCode" size="1" class="input">
					<option value="##"><%=getadminAgentsAccessLngStr("DtxtDefault")%></option>
					<% If EnableWhs Then
					do while not rw.eof %>
					<option value="<%=rw("WhsCode")%>"><%=myHTMLEncode(rw("WhsName"))%></option>
					<% rw.movenext
					loop
					rw.movefirst 
					end if %>
					</select></td>
					<td width="180" height="23" class="style7">
					<input type="checkbox" id="Admin" name='Admin' value="Y" class="noborder"></td>
					<% If 1 = 2 Then %>
					<td width="120" height="23" class="style8">
					<p align="center">
					<input type="checkbox" name="AsignedSLP" value="Y" class="noborder"></td>
					<td height="23" class="style7">
					<input type="button" value="<%=getadminAgentsAccessLngStr("DtxtEdit")%>" name="btnEditGroups" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; font-weight:bold;display: none" onclick="editRGAccess(document.frmAddUser.SlpCode.value);"></td>
					<% End If %>
					<td bgcolor="#F5FBFE" height="23">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="29">
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminAgentsAccessLngStr("DtxtAdd")%>" name="B2" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="vUserAdd">
	<input type="hidden" name="ipIndex" value="">
	<input type="hidden" name="rgIndex" value="">
	</form>
</table>
<script language="javascript">
<!--
function setNewRGAccess(rgIndex)
{
	document.frmAddUser.rgIndex.value = rgIndex;
}
function setNewPwd(newPwd, changePwd)
{
	document.frmAddUser.Password.value = newPwd;
	document.frmAddUser.ChangePwd.value = changePwd;
}

function valFrm() {
Validation = true
	if (document.frmAddUser.UserName.value == "") 
	{
		alert("<%=getadminAgentsAccessLngStr("LtxtValUName")%>");
		Validation = false;
		document.frmAddUser.UserName.focus() ;
	}
	else if (document.frmAddUser.Password.value == "")
	{
		alert("<%=getadminAgentsAccessLngStr("LtxtValPwd")%>");
		Validation = false;
		Pic('userPwd', 'adminAgentsAccessPwd.asp?NewName=' + document.frmAddUser.UserName.value, 200, 140, 'no', 'no');
	}
	else if (document.frmAddUser.EMail.value != "" && !emailCheck(document.frmAddUser.EMail.value))
	{
		alert("<%=getadminAgentsAccessLngStr("LtxtValEMail")%>");
		Validation = false;
		document.frmAddUser.EMail.focus();
	}
	else if (document.frmAddUser.Admin.checked && document.frmAddUser.EMail.value == '')
	{
		alert('<%=getadminAgentsAccessLngStr("LtxtAdminMail")%>');
		Validation = false;
		document.frmAddUser.EMail.focus();
	}
return Validation 
}
//-->
</script>
<% If Not Slp Then %>
<font face="Verdana" size="1" color="#FF5050"><center><%=getadminAgentsAccessLngStr("LtxtErrAgents")%></center></font>
<% End If %><!--#include file="bottom.asp" -->