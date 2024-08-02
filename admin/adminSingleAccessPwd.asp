<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<!--#include file="lang/adminSingleAccessPwd.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getadminSingleAccessPwdLngStr("LttlChangePwd")%></title>
<style>
<!--
input        { font-family: Verdana; font-size: 10px; border: 1px solid #73B9B9; background-color: #D8EDEE }
.noborder {
	border-style: solid;
	border-width: 0;
	background: background-image;
}

-->
</style>
</head>
<% If Request.Form.Count = 0 Then %>
<script language="javascript" src="general.js"></script>
<script language="javascript">
function ValidateForm(frm)
{
	if (frm.pwd1.value != frm.pwd2.value)
	{
		alert("<%=getadminSingleAccessPwdLngStr("LtxtValConfPwd")%>");
		frm.pwd1.value = "";
		frm.pwd2.value = "";
		frm.pwd1.focus();
		return false;
	}
	else if (frm.pwd1.value.length < 5)
	{
		alert("<%=getadminSingleAccessPwdLngStr("LtxtValMinChar")%>");
		frm.pwd1.value = "";
		frm.pwd2.value = "";
		frm.pwd1.focus();
		return false;
	}
}
</script>
<body bgcolor="#F9FDFF" onload="javascript:document.frmPwd.pwd1.focus()" marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onbeforeunload="javascript:opener.clearWin();">

<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<form method="POST" name="frmPwd" action="<% If Request("UName") <> "" Then %>adminSubmit.asp<% Else %>adminSingleAccessPwd.asp<% End If %>" onsubmit="javascript:return ValidateForm(this)">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">
		&nbsp;<%=getadminSingleAccessPwdLngStr("LttlChangePwd")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD" align="center">
		<font size="1" face="Verdana" color="#3366CC"><%=getadminSingleAccessPwdLngStr("DtxtUser")%>: <%=Request("UName")%><%=Request("NewName")%></font></td>
	</tr>
	<tr>
		<td bordercolor="#C8E7E8" align="center">
		<input type="password" name="pwd1" size="20" style="border:1px solid #68A6C0; color: #3F7B96; font-family: Verdana; font-size: 10px; width: 276; padding: 0; background-color: #D9F0FD; width: 150" onkeydown="javascript:return chkMax(event, this, 20);"></td>
	</tr>
	<tr>
		<td bordercolor="#C8E7E8" align="center">
		<input type="password" name="pwd2" size="20" style="border:1px solid #68A6C0; color: #3F7B96; font-family: Verdana; font-size: 10px; width: 276; padding: 0; background-color: #D9F0FD; width: 150" onkeydown="javascript:return chkMax(event, this, 20);"></td>
	</tr>
		<tr>
		<td bordercolor="#C8E7E8" align="center">
		<font size="1" face="Verdana" color="#3366CC"><%=getadminSingleAccessPwdLngStr("LtxtMinPwd")%></font></td>
		</tr>
	<tr>
		<td bordercolor="#C8E7E8" align="center">
		<table cellpadding="0" cellspacing="0" border="0">
			<tr>
				<td><input class="noborder" type="checkbox" name="ChangePwd" id="ChangePwd" value="Y"></td>
				<td align="center"><font size="1" face="Verdana" color="#3366CC"><label for="ChangePwd"><%=getadminSingleAccessPwdLngStr("LtxtChangePwdLog")%></label></font></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bordercolor="#C8E7E8" align="center">
		<input type="submit" value="<%=getadminSingleAccessPwdLngStr("DtxtAccept")%>" name="B1" class="OlkBtn"></td>
	</tr>
		<input type="hidden" name="submitCmd" value="SingleUserPwd">
		<input type="hidden" name="UName" value="<%=Request("UName")%>">
	</form>
</table>

<script language="javascript">
function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))return false; else return true;
}
</script>
</body>
<% Else %>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<% If Request("ChangePwd")= "Y" Then ChangePwd = "Y" Else ChangePwd = "N" %>
<script language="javascript">
opener.setNewPwd('<%=Replace(Request("pwd1"), "'", "\'")%>', '<%=ChangePwd%>');
window.close();
</script>
</body>
<% End If %>
</html>