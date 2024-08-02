<!--#include file="top.asp" -->
<!--#include file="lang/adminPwd.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	color: #31659C;
}
.style2 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<SCRIPT LANGUAGE="JavaScript">
function checkPw() {
	if (document.form1.OldPwd.value == '')
	{
		alert('<%=getadminPwdLngStr("LtxtValOldPwd")%>');
		document.form1.OldPwd.focus();
		return false;
	}
	else if (document.form1.NewPwd.value == '')
	{
		alert('<%=getadminPwdLngStr("LtxtValNewPwd")%>');
		document.form1.NewPwd.focus();
		return false;
	}
	else if (document.form1.NewPwd.value != document.form1.ConfPwd.value)
	{
		alert ("\n<%=getadminPwdLngStr("LtxtValConfPwd")%>")
		document.form1.NewPwd.value = '';
		document.form1.ConfPwd.value = '';
		document.form1.NewPwd.focus();
		return false;
	}
	
	return true;
}
</script>
<% 
         
If Request("ErrMsg") = "True" Then
	ErrMsg = True
Else
	ErrMsg = False
End If %>
<form method="post" action="adminSubmit.asp" name="form1" onsubmit="return checkPw()">

<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">&nbsp;<%=getadminPwdLngStr("LttlPwdChange")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminPwdLngStr("LttlPwdChangeNote")%></font></td>
	</tr>
	<tr>
		<td>
		<div>
			<table border="0" cellpadding="0" width="300" id="table6">
				<tr>
					<td bgcolor="#E1F3FD" class="style1">
					<font size="1" face="Verdana"><strong>&nbsp;<%=getadminPwdLngStr("LtxtCurPwd")%></strong></font></td>
					<td bgcolor="#E1F3FD" width="1">
					<input type="password" name="OldPwd" size="20" class="input" onkeydown="return chkMax(event, this, 20);"></td>
				</tr>
				<tr>
					<td bgcolor="#E1F3FD" class="style2">
					<font face="Verdana" size="1"><strong>&nbsp;<%=getadminPwdLngStr("LtxtNewPwd")%></strong></font></td>
					<td bgcolor="#E1F3FD" width="1">
					<input type="password" name="NewPwd" size="20" class="input" onkeydown="return chkMax(event, this, 20);"></td>
				</tr>
				<tr>
					<td bgcolor="#E1F3FD" class="style2">
					<font face="Verdana" size="1"><strong>&nbsp;<%=getadminPwdLngStr("LtxtPwdConf")%></strong></font></td>
					<td bgcolor="#E1F3FD" width="1">
					<input type="password" name="ConfPwd" size="20" class="input" onkeydown="return chkMax(event, this, 20);"></td>
				</tr>
				<% If Request("ErrMsg") <> "" and ErrMsg Then %>
				<tr>
					<td bgcolor="#F5FBFE" colspan="2">
					<p align="center">
					<font face="Verdana" size="1" color="#FF0000">&nbsp;<%=getadminPwdLngStr("LtxtValCurPwd")%></font></td>
				</tr>
				<% ElseIf Request("ErrMsg") <> "" and Not ErrMsg Then %>
				<tr>
					<td bgcolor="#F5FBFE" colspan="2">
					<p align="center">
					<font face="Verdana" size="1" color="#4783C5">&nbsp;<%=getadminPwdLngStr("LtxtPwdChanged")%></font></td>
				</tr>
				<% end if %>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminPwdLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="submitCmd" value="adminPwd">
</form>
<!--#include file="bottom.asp" -->