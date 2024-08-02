<!--#include file="top.asp" -->
<!--#include file="lang/adminSingleAccess.asp" -->

<head>
<%

SQL = 	"select UserName, EMail, EMailInbox, T1.Admin, Status " & _
		"from OLKAgentsAccess T1 " & _
		"order by 1"
rs.open sql, connCommon, 3, 1


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
alert("<%=getadminSingleAccessLngStr("LtxtValUsr")%>")
<% End If %>
function valFrmUpdAccess()
{
	<% If rs.recordcount > 1 Then %>
	uName = document.frmUpdAccess.UserName;
	uMail = document.frmUpdAccess.EMail;
	uAdmin = document.frmUpdAccess.Admin;
	for (var i = 0;i<uName.length;i++)
	{
		if (uMail(i).value != '' && !emailCheck(uMail(i).value))
		{
			alert('<%=getadminSingleAccessLngStr("LtxtValEMail")%>');
			uMail(i).focus();
			return false;
		}
		
		if (uAdmin(i).checked && uMail(i).value == '')
		{
			alert('<%=getadminSingleAccessLngStr("LtxtAdminMail")%>');
			uMail(i).focus();
			return false;
		}
	}
	<% ElseIf rs.recordcount = 1 Then %>
	if (document.frmUpdAccess.EMail.value != '' && !emailCheck(document.frmUpdAccess.EMail.value))
	{
		alert('<%=getadminSingleAccessLngStr("LtxtValEMail")%>');
		document.frmUpdAccess.EMail.focus();
		return false;
	}
	
	if (document.frmUpdAccess.Admin.checked && document.frmUpdAccess.EMail.value == '')
	{
		alert('<%=getadminSingleAccessLngStr("LtxtAdminMail")%>');
		document.frmUpdAccess.EMail.focus();
		return false;
	}
	<% End If %>
	return true;
}

function changeAccess(value, SlpCode)
{
	//document.getElementById('btnEditGroups' + SlpCode).style.display = value == 'D' ? 'none' : '';
	document.getElementById('lnkAut' + SlpCode).style.display = value == 'D' ? 'none' : '';
	if (document.getElementById('Aut' + SlpCode).style.display != 'none') goUserAut(SlpCode);
}
function editIPAccess(UserName)
{
	Pic('IPAccess', 'adminSingleIPAccess.asp?UserName=' + escape(UserName) +'&pop=Y', 400, 300, 'Y', 'N');
}

function doDelUsr(UserName)
{
	if(confirm('<%=getadminSingleAccessLngStr("LtxtConfDelUsr")%>'.replace('{0}', UserName)))
		window.location.href='adminSubmit.asp?submitCmd=SingleUserRem&UName=' + UserName;
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
</head>

<table border="0" cellpadding="0" width="100%">
	<tr>
		<td height="15"></td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmUpdAccess" onsubmit="return valFrmUpdAccess();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("LtxtAgAccDef")%> </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminSingleAccessLngStr("LtxtAgAccNote")%> </font></td>
	</tr>
	
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td colspan="2">
					<table cellpadding="0" border="0" width="100%">
						<tr>
							<td class="style1" style="width: 20px">
							&nbsp;</td>
							<td width="130" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtUser")%></font></td>
							<td width="220" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtEMail")%></font></td>
							<td width="110" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("LtxtMsgAlr")%></font></td>
							<td width="110" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtActive")%></font></td>
							<td class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtAdministrator")%></font></td>
							<td align="center" class="style4">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtIP")%></font></td>
							<td style="width: 16px" class="style3">
							&nbsp;</td>
						</tr>
						<% do While NOT RS.EOF %>
						<tr bgcolor="#F3FBFE">
							<td class="style2" style="width: 20px"><input type="hidden" name="UserName" value="<%=Server.HTMLEncode(RS("UserName"))%>">
							<font face="Verdana" size="1" color="#3366CC">
							<a href="adminSingleAccessDB.asp?UserName=<%=Server.URLEncode(RS("UserName"))%>" id="lnkAut<%=rs.bookmark%>" style="text-decoration: none;"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></font></td>
							<td width="130" class="style2">
							<table><tr><td><a href="#" onclick="javascript:Pic('userPwd', 'adminSingleAccessPwd.asp?UName=<%=Replace(rs("Username"), "'", "\'")%>', 260, 140, 'no', 'no')"><img border="0" src="images/img_lock.gif" width="15" height="15"></a></td>
							<td><font face="Verdana" size="1" color="#3366CC"><%=Server.HTMLEncode(RS("UserName"))%></font></td></tr></table></td>
							<td width="220" class="style2">
							<input type="text" value="<% If Not IsNull(RS("EMail")) Then %><%=Server.HTMLEncode(RS("EMail"))%><% End If %>" id="EMail" name="EMail<%=rs.bookmark%>" class="input" onkeydown="return chkMax(event, this, 100);" size="34"></td>
							<td width="110" class="style2" align="center">
							<input type="checkbox" value="Y" name="chkEMailInbox<%=rs.bookmark%>" <% If rs("EMailInbox") = "Y" Then %>checked<% End If %> class="noborder"></td>
							<td width="110" class="style2" align="center">
							<input type="checkbox" value="Y" name="chkStatus<%=rs.bookmark%>" <% If rs("Status") = "Y" Then %>checked<% End If %> class="noborder"></td>
							<td width="120" class="style2">
							<p align="center">
							<input type="checkbox" id="Admin" name="Admin<%=rs.bookmark%>" <% If rs("Admin") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<td class="style5">
							<input type="button" value="<%=getadminSingleAccessLngStr("DtxtEdit")%>" name="btnEditIP<%=rs.bookmark%>" onclick="editIPAccess('<%=Replace(rs("UserName"), "'", "\'")%>');" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; font-weight:bold"></td>
							<td style="width: 16px" class="style2">
							<p align="right">
							<a href="javascript:doDelUsr('<%=Replace(rs("UserName"), "'", "\'")%>');"><img border="0" src="images/remove.gif" width="16" height="16"></a></td>
						</tr>
						<tr bgcolor="#F3FBFE" id="Aut<%=rs.bookmark%>" style="display: none;">
							<td class="style2" colspan="8">
							<div id="db<%=rs.bookmark%>"></div>
							</td>
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
				<td width="75"><% If rs.recordcount > 0 then %><input type="submit" value="<%=getadminSingleAccessLngStr("DtxtSave")%>" name="B1" class="OlkBtn"><% End If %></td>
				<td><hr color="#0D85C6" size="1"></td>
			</table>
		</td>
	</tr>
		<input type="hidden" name="submitCmd" value="SingleUserUpd">
	</form>
</table>
<table cellpadding="0" border="0" width="100%">
	<form method="POST" action="adminSubmit.asp" onsubmit="return valFrm()" name="frmAddUser">
	<tr>
		<td bgcolor="#E1F3FD" colspan="2"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtAdd")%>&nbsp;<%=getadminSingleAccessLngStr("DtxtAgent")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" colspan="2">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"> <%=getadminSingleAccessLngStr("LtxtAddUserNote")%></font></td>
	</tr>
	<tr>
		<td colspan="2">
			<table cellpadding="0" border="0" width="100%">
				<tr>
					<td width="130" class="style6">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtUser")%></font></b>&nbsp;</td>
					<td width="220" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtEMail")%></font></td>
					<td width="110" class="style1">
							<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("LtxtMsgAlr")%></font></td>
					<td width="110" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtActive")%></font></td>
					<td width="180" class="style1">
					<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessLngStr("DtxtAdministrator")%></font></td>						
					<td class="style3">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="130" height="23" class="style8">
					<table><tr><td><a href="#" onclick="javascript:Pic('userPwd', 'adminAgentsAccessPwd.asp?NewName=' + document.frmAddUser.UserName.value, 260, 140, 'no', 'no')"><img border="0" src="images/img_lock.gif" width="15" height="15"></a><input type="hidden" name="Password" value=""><input type="hidden" name="ChangePwd" value="N"></td><td><font face="Verdana" size="1" color="#3366CC">
					<input type="text" value="" name="UserName" class="input" onkeydown="return chkMax(event, this, 20);" size="16"></font></td></tr></table>
					</td>
					<td width="220" height="23" class="style8">
					<font face="Verdana" size="1" color="#3366CC">
					<input type="text" name="EMail" class="input" onkeydown="return chkMax(event, this, 100);" size="34"></font></td>
					<td width="110" height="23" class="style8" align="center">
							<input type="checkbox" value="Y" name="EMailInbox" class="noborder"></td>
					<td width="110" height="23" class="style8" align="center">
					<input type="checkbox" name="chkStatus" value="Y" class="noborder"></td>
					<td width="180" height="23" class="style7">
					<input type="checkbox" id="Admin" name='Admin' value="Y" class="noborder"></td>
					<td bgcolor="#F5FBFE" height="23">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="29">
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSingleAccessLngStr("DtxtAdd")%>" name="B2" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="SingleUserAdd">
	<input type="hidden" name="ipIndex" value="">
	</form>
</table>
<script language="javascript">
<!--
function setNewPwd(newPwd, changePwd)
{
	document.frmAddUser.Password.value = newPwd;
	document.frmAddUser.ChangePwd.value = changePwd;
}

function valFrm() {

	var chkUser = document.frmAddUser.UserName.value;
	if (chkUser == "") 
	{
		alert("<%=getadminSingleAccessLngStr("LtxtValUName")%>");
		document.frmAddUser.UserName.focus() ;
		return false;
	}
	
	if (document.frmUpdAccess.UserName)
	{
		if (document.frmUpdAccess.UserName.length)
		{
			for (var i = 0;i<document.frmUpdAccess.UserName.length;i++)
			{
				if (chkUser == document.frmUpdAccess.UserName[i].value)
				{
					alert('<%=getadminSingleAccessLngStr("LtxtVarUNamAsign")%>'.replace('{0}', chkUser));
					document.frmAddUser.UserName.focus();
					return false
				}	
			}
		}
		else
		{
			if (chkUser == document.frmUpdAccess.UserName.value)
			{
				alert('<%=getadminSingleAccessLngStr("LtxtVarUNamAsign")%>'.replace('{0}', chkUser));
				document.frmAddUser.UserName.focus();
				return false
			}
		}
	}
	
	if (document.frmAddUser.Password.value == "")
	{
		alert("<%=getadminSingleAccessLngStr("LtxtValPwd")%>");
		Pic('userPwd', 'adminSingleAccessPwd.asp?NewName=' + document.frmAddUser.UserName.value, 200, 140, 'no', 'no');
		return false;
	}
	
	if (document.frmAddUser.EMail.value != "" && !emailCheck(document.frmAddUser.EMail.value))
	{
		alert("<%=getadminSingleAccessLngStr("LtxtValEMail")%>");
		document.frmAddUser.EMail.focus();
		return false;
	}
	
	if (document.frmAddUser.Admin.checked && document.frmAddUser.EMail.value == '')
	{
		alert('<%=getadminSingleAccessLngStr("LtxtAdminMail")%>');
		document.frmAddUser.EMail.focus();
		return false;
	}
	
	return true; 
}
//-->
</script><!--#include file="bottom.asp" -->