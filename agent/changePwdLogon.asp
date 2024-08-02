<!--#include file="lang/changePwdLogon.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl" <% End If %>="">

<link href="style1.css" rel="stylesheet" type="text/css">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="general.js"></script>
<title>Top Manage OLK - <%=getchangePwdLogonLngStr("LtxtChangePwd")%></title>
<style type="text/css">
.style1 {
	font-family: Tahoma;
	color: silver;
}
.style2 {
	text-align: center;
	background-color: #F7FAFB;
}
</style>
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")

If Request.Form("doSave") = "Y" Then
	sql = "update OLKAgentsAccess set Password = N'" & Request("txtNewPwd") & "', ChangePwd = 'N' where SlpCode = " & Session("vendid")
	conn.execute(sql)
	userType = "V"
	
	If Request.Cookies("pwd") <> "" Then
	    Response.cookies("pwd").expires = DateAdd("d",60,now())
    	Response.cookies("pwd").path = "/"
    	Response.cookies("pwd") = Request("txtNewPwd")
    End If

	Response.Redirect "agent.asp"
End If
%>
<body onload="document.frmChangePwd.txtNewPwd.focus();">
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
<form method="POST" action="changePwdLogon.asp" name="frmChangePwd" onsubmit="javascript:return valFrm();">
	<div align="center">
		<center>
		<table border="0" cellpadding="0" cellspacing="0" width="625">
			<tr>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="41" height="1" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="69" height="1" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="399" height="1" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="76" height="1" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="40" height="1" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></p>
				</td>
			</tr>
			<tr>
				<td colspan="5">
				<p align="center">
				<img name="login_clientsNuevo_r1_c1" src='images/<%=Session("rtl")%>login_clientsNuevo_r1_c1.jpg' width="625" height="29" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="1" height="29" border="0" alt=""></p>
				</td>
			</tr>
			<tr>
				<td rowspan="2" background="images/login_clientsNuevo_r2_c1.jpg">
				<p align="center">
				<img name="login_clientsNuevo_r2_c1" src='images/<%=Session("rtl")%>login_clientsNuevo_r2_c1.jpg' width="41" height="341" border="0" alt=""></p>
				</td>
				<td colspan="3" valign="middle">
				<table border="0" width="100%" cellpadding="0" id="table1">
					<tr>
						<td>
						<div align="center">
							<br>
							<br>
							<table border="0" cellpadding="0" width="90%" id="table2">
								<tr>
									<td colspan="2" class="style2">
									<font face="Tahoma" size="1"><strong><%=getchangePwdLogonLngStr("LtxtChangePwd")%></strong></font></td>
								</tr>
								<tr>
									<td width="169" bgcolor="#E7F0F5">
									<p align="center">
									<font face="Tahoma" size="1"><%=getchangePwdLogonLngStr("DtxtCmp")%></font></p>
									</td>
									<td bgcolor="#E7F0F5">
									<font face="Tahoma" size="1">
									<%=mySession.GetCompanyName%></font></td>
								</tr>
								<tr>
									<td width="169" bgcolor="#E7F0F5">
									<p align="center">
									<font face="Tahoma" size="1"><%=getchangePwdLogonLngStr("DtxtUser")%></font></p>
									</td>
									<td bgcolor="#E7F0F5">
									<font face="Tahoma" size="1"><%=mySession.GetAgentName%></font></td>
								</tr>
								<tr>
									<td width="169" bgcolor="#E7F0F5">
									<p align="center">
									<font size="1" face="Tahoma"><%=getchangePwdLogonLngStr("LtxtNewPwd")%></font></p>
									</td>
									<td bgcolor="#E7F0F5">
									<input type="password" class="input" name="txtNewPwd" size="20" maxlength="20"> 
									<font face="Tahoma" size="1">5 - 20 <%=getchangePwdLogonLngStr("LtxtChars")%></font></td>
								</tr>
								<tr>
									<td width="169" bgcolor="#E7F0F5">
									<p align="center">
									<font size="1" face="Tahoma"><%=getchangePwdLogonLngStr("LtxtConfPwd")%></font></p>
									</td>
									<td bgcolor="#E7F0F5">
									<input type="password" class="input" name="txtConfPwd" size="20" maxlength="20"> </td>
								</tr>
								<tr>
									<td colspan="2" valign="top" bgcolor="#E7F0F5">
									<p align="center">
									<input type="submit" value="<%=getchangePwdLogonLngStr("DtxtAccept")%>" name="btnAccept" style="font-family: Tahoma; font-size: 10px; border: 1px solid #10699C; background-color: #FFFFFF;"></p>
									</td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td><center>
						<p><font size="2" face="Verdana"></font></p>
						</center></td>
					</tr>
				</table>
				</td>
				<td rowspan="2">
				<p align="center">
				<img name="login_clientsNuevo_r2_c5" src='images/<%=Session("rtl")%>login_clientsNuevo_r2_c5.jpg' width="40" height="341" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="1" height="239" border="0" alt=""></p>
				</td>
			</tr>
			<tr>
				<td>
				<p align="center">
				<img name="login_clientsNuevo_r3_c2" src="images/login_clientsNuevo_r3_c2.jpg" width="69" height="102" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img name="login_clientsNuevo_r3_c3" src='images/<%=Session("rtl")%>login_clientsNuevo_r3_c3.jpg' width="399" height="102" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img name="login_clientsNuevo_r3_c4" src='images/<%=Session("rtl")%>login_clientsNuevo_r3_c4.jpg' width="76" height="102" border="0" alt=""></p>
				</td>
				<td>
				<p align="center">
				<img src="images/spacer.gif" width="1" height="102" border="0" alt=""></p>
				</td>
			</tr>
		</table>
		</center></div>
<input type="hidden" name="doSave" value="Y">
</form>
<p align="center"><font color="#C0C0C0" size="1" face="Verdana">
<a href="http://www.topmanage.com.pa/"><span class="style1">TopManage</span></a> 
®</font><font face="Tahoma" color="#c0c0c0" size="1"> 2002 - 2012 - <%=getchangePwdLogonLngStr("DtxtEMail")%>:
<a href="mailto:info@topmanage.com.pa"><font color="#c0c0c0">info@topmanage.com.pa</font></a> 
- <%=getchangePwdLogonLngStr("DtxtPhone")%>: <font color="#c0c0c0">507.300.7200</font></font></p>
<p align="center">&nbsp;</p>

</body>
</html>