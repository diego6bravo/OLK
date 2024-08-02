<!--#include file="lang/changePwdLogon.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Mobile OLK</title>
<style type="text/css">
.style1 {
	border-style: solid;
	border-width: 0;
}
.style2 {
	font-weight: bold;
	border-style: solid;
	border-width: 0;
}
</style>
</head>

<%

set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
oLic.LicenceServer = licip
oLic.LicencePort = licport

set rs = Server.CreateObject("ADODB.RecordSet")

If Request.Form("doSave") = "Y" Then
	sql = "update OLKAgentsAccess set Password = N'" & oLic.GetEncPwd(Request("txtNewPwd")) & "', ChangePwd = 'N' where SlpCode = " & Session("vendid")
	conn.execute(sql)
	userType = "V"
	
	If Request.Cookies("pwd") <> "" Then
	    Response.cookies("pwd").expires = DateAdd("d",60,now())
    	Response.cookies("pwd").path = "/"
    	Response.cookies("pwd") = Request("txtNewPwd")
    End If

	Response.Redirect "operaciones.asp?cmd=home"
End If
%>
<script language="javascript" src="general.js"></script>
<script type="text/javascript">
<!--
function valFrm()
{
	var newPwdLen = document.frmChangePwd.txtNewPwd.value.length;
	if (newPwdLen < 5 || newPwdLen > 20)
	{
		alert('<%=getchangePwdLogonLngStr("LtxtValPwdChars")%>');
		document.frmChangePwd.txtNewPwd.focus();
		return false;
	}
	else if (document.frmChangePwd.txtNewPwd.value != document.frmChangePwd.txtConfPwd.value)
	{
		alert('<%=getchangePwdLogonLngStr("LtxtValNoMatchNewPwd")%>');
		document.frmChangePwd.txtNewPwd.value = '';
		document.frmChangePwd.txtConfPwd.value = '';
		document.frmChangePwd.txtNewPwd.focus();
		return false;
	}
	return true;
}
//-->
</script>
<body topmargin="0" onload="document.frmChangePwd.txtNewPwd.focus();" <% If Session("rtl") <> "" Then %> xdir="rtl" <% End If %>>

<div align="center">
	<center>
	<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
		<tr>
			<td>
			<p align="center"><b><font face="Verdana" size="1" color="#FDAF2F"><%=getchangePwdLogonLngStr("LtxtChangePwd")%></font></b></p>
			</td>
		</tr>
		<tr>
			<td>
			<p align="center"><img border="0" src="images/pocket_olkicon.gif"></p>
			</td>
		</tr>
		<tr>
			<td style="font-size: 10px">&nbsp;</td>
		</tr>
		<form method="POST" name="frmChangePwd" action="changePwdLogon.asp" onsubmit="javascritp:return valFrm();">
			<tr>
				<td bgcolor="#F0F8FF">
				<p align="center"><font size="1" face="Verdana"><%=mySession.GetCompanyName%></font></p>
				</td>
			</tr>
			<tr>
				<td bgcolor="#F0F8FF">
				<div align="center">
					<center>
					<table border="0" cellpadding="0" bordercolor="#111111" id="AutoNumber2" style="width: 95%">
						<tr>
							<td bgcolor="#DDEFFF" class="style2">
							<font size="1" face="Verdana"><%=getchangePwdLogonLngStr("DtxtUser")%>:</font></td>
							<td bgcolor="#DDEFFF" class="style1"><font size="1" face="Verdana"><%=mySession.GetAgentName%></font>
							</td>
						</tr>
						<tr>
							<td bgcolor="#DDEFFF" class="style2">
							<font size="1" face="Verdana"><%=getchangePwdLogonLngStr("LtxtNewPwd")%>:</font></td>
							<td bgcolor="#DDEFFF" class="style1">
							<input type="password" name="txtNewPwd" maxlength="20" style="font-size: 12px; size= ; float: left; width: 90%" size="12" value="" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
						</tr>
						<tr>
							<td bgcolor="#DDEFFF" class="style2">
							<font size="1" face="Verdana"><%=getchangePwdLogonLngStr("LtxtConfPwd")%>:</font></td>
							<td bgcolor="#DDEFFF" class="style1">
							<input type="password" name="txtConfPwd" maxlength="20" style="font-size: 12px; size= ; float: left; width: 90%" size="12" value="" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
						</tr>
						<tr>
							<td colspan="2" bgcolor="#DDEFFF">
							<p align="center">
							<input type="submit" value="<%=getchangePwdLogonLngStr("DtxtAccept")%>" style="color: #005782; font-family: verdana; font-size: 10px; border: 1px solid #006699; background-color: #C1E1FF" name="btnEnter"></p>
							</td>
						</tr>
					</table>
					</center></div>
				</td>
			</tr>
			<input type="hidden" name="doSave" value="Y">
		</form>
	</table>
	</center></div>

</body>

</html>
