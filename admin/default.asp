<% response.buffer = true 
Session.Timeout=30 
Session.LCID = 6154
If Request("logout") = "Y" or Request.Form.Count = 0 Then Session.Abandon %>
<!--#include file="conn.asp" -->
<!--#include file="lang.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
set rm = Server.CreateObject("ADODB.RecordSet")

set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
oLic.LicenceServer = licip
oLic.LicencePort = licport


sap1 = Request.Form("uid")
sap2 = Request.Form("pwd")
If Request.Form.Count > 0 Then
If Err.Number = 0 Then 
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "OLKAdminAccess"
	cmd.ActiveConnection = connCommon
	cmd.Parameters.Refresh
	cmd("@UId") = Request.Form("uid")
	cmd("@Pwd") = oLic.GetEncPwd(Request.Form("pwd"))

	set rs = server.createobject("ADODB.RecordSet")
	set rs = cmd.Execute
	If rs("Verfy") = "Y" Then
		Session("OLKAdmin") = True
		Session("olkdb") = ""
		Session("ID") = ""
		If InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") <> 0 Then
			Session("style") = "ie"
		Else
			Session("style") = "nc"
		End If
		mySession.Login("ADM")
		response.redirect "admin.asp"
	End If
ElseIf Err.Number = -2147217843 Then
	response.redirect "changeCnPwd.asp?rAction=admin"
Else
	response.write "<center>" & Err.Description & "</center>"
End If
End If
%>
<!--#include file="lang/default.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<link href="style1.css" rel="stylesheet" type="text/css">

<head>
<title><%=getdefaultLngStr("LttlAdmin")%></title>
<base target="_self">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	
	If oLic.IsAlive Then
		isNo = False
		Select Case oLic.HasLicence(50)
			Case "NO"
				ErrMsg = getdefaultLngStr("DtxtNoOLKLic")
				isNo = True
			Case "EXP"
				ErrMsg = getdefaultLngStr("DtxtOLKLicExp")
		End Select
		
		If ErrMsg <> "" Then
			Select Case oLic.HasLicence(51)
				Case "YES"
					ErrMsg = ""
				Case "EXP"
					If isNo Then ErrMsg = getdefaultLngStr("DtxtOLKLicExp")
			End Select
		End If
		
		If ErrMsg <> "" Then
			Select Case oLic.HasLicence(52)
				Case "YES"
					ErrMsg = ""
				Case "EXP"
					If isNo Then ErrMsg = getdefaultLngStr("DtxtOLKLicExp")
			End Select
		End If
		
		If ErrMsg <> "" Then
			Select Case oLic.HasLicence(53)
				Case "YES"
					ErrMsg = ""
				Case "EXP"
					If isNo Then ErrMsg = getdefaultLngStr("DtxtOLKLicExp")
			End Select
		End If
	Else
		ErrMsg = getdefaultLngStr("DtxtInactiveLicServer")
	End If


%>
</head>

<body bgcolor="#ffffff" onload="<% If ErrMsg = "" Then %>document.frmLogin.uid.focus()<% End If %>">

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="497">
		<tr>
			<td>
			<img src="images/spacer.gif" width="58" height="1" border="0" alt=""></td>
			<td>
			<img src="images/spacer.gif" width="379" height="1" border="0" alt=""></td>
			<td>
			<img src="images/spacer.gif" width="60" height="1" border="0" alt=""></td>
			<td>
			<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<tr>
			<td colspan="3">
			<img name="admin_art_r1_c1" src="images/admin_art_r1_c1.gif" width="497" height="134" border="0" alt=""></td>
			<td>
			<img src="images/spacer.gif" width="1" height="134" border="0" alt=""></td>
		</tr>
		<tr>
			<td colspan="3" background="images/<%=Session("rtl")%>admin_art_r2_c1.gif">
			<p align="center"><b><font color="#006699" size="2"><%=getdefaultLngStr("LttlAdmin")%></font></b></td>
			<td>
			<img src="images/spacer.gif" width="1" height="23" border="0" alt=""></td>
		</tr>
		<tr>
			<td rowspan="2">
			<img name="admin_art_r3_c1" src="images/<%=Session("rtl")%>admin_art_r3_c1.gif" width="58" height="216" border="0" alt=""></td>
			<td valign="middle">
			<form method="POST" action="default.asp" name="frmLogin">
				<table border="0" width="100%" cellpadding="0" id="table1">
					<%  If ErrMsg = "" Then %>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" width="300" id="table2">
								<tr>
									<td bgcolor="#F7F7F7" style="width: 100px">
									<font face="Tahoma" size="1"><%=getdefaultLngStr("LtxtUserName")%></font>
									</td>
									<td bgcolor="#F7F7F7">
									<input type="text" name="uid" class="input" size="29" style="width: 100%; "></td>
								</tr>
								<tr>
									<td bgcolor="#F7F7F7" style="width: 100px">
									<font face="Tahoma" size="1"><%=getdefaultLngStr("LtxtPwd")%></font>
									</td>
									<td bgcolor="#F7F7F7">
									<input type="password" name="pwd" class="input" size="29" style="width: 100%; "></td>
								</tr>
								<tr>
									<td bgcolor="#F7F7F7" style="width: 100px">
									<font face="Tahoma" size="1"><%=getdefaultLngStr("LtxtLng")%></font></td>
									<td bgcolor="#F7F7F7">
									<select size="1" name="newLng" class="input" onchange="javascript:window.location.href='?newLng=' + this.value">
									<% For i = 0 to UBound(myLanIndex) %>
									<option <% If Request.Cookies("myLng") = myLanIndex(i)(0) Then %>selected<% End If %> value="<%=myLanIndex(i)(0)%>">
									<%=myLanIndex(i)(1)%></option>
									<% Next %>
									</select></td>
								</tr>
								<tr>
									<td colspan="2" valign="top" bgcolor="#F7F7F7">
									<p align="center">
									<input type="submit" value="<%=getdefaultLngStr("LtxtEnter")%>" name="B1" style="font-family: Tahoma; font-size: 10px; border: 1px solid #10699C; background-color: #FFFFFF"></p>
									</td>
								</tr>
							</table>
						</div>
						</td>
					</tr>	
					<tr>
						<td align="center"><font size="1" face="Verdana"><% 
						If sap1 = "" and sap2 = "" Then %>
						<%=getdefaultLngStr("LtxtEntryUidPwd")%>
						<% ElseIf sap1 <> "" and sap2 <> "" then %>
						<font color="#FF0000"><%=getdefaultLngStr("LtxtBadUidPwd")%></font>
						<% End If %></font></td>
					</tr>
					<% Else %>
					<tr>
						<td>
						<p align="center"><font color="#FF0000"><%=ErrMsg%></font></td>
					</tr>
					<% End If %>
				</table>
				<input type="hidden" name="other" value="<%=Request("other")%>">
			</form>
			</td>
			<td rowspan="2">
			<img name="admin_art_r3_c3" src="images/<%=Session("rtl")%>admin_art_r3_c3.gif" width="60" height="216" border="0" alt=""></td>
			<td>
			<img src="images/spacer.gif" width="1" height="199" border="0" alt=""></td>
		</tr>
		<tr>
			<td>
			<img name="admin_art_r4_c2" src="images/admin_art_r4_c2.gif" width="379" height="17" border="0" alt=""></td>
			<td>
			<img src="images/spacer.gif" width="1" height="17" border="0" alt=""></td>
		</tr>
	</table>
	<p><font color="#C0C0C0" size="1" face="Verdana"><a href="http://www.topmanage.com.pa/"><font color="#C0C0C0">TopManage</font></a> &reg;</font><font face="Tahoma" color="#c0c0c0" size="1"> 2002 - 
	2012 - E-mail: <a href="mailto:info@topmanage.com.pa"><font color="#c0c0c0">
	info@topmanage.com.pa</font></a> 
	- <%=getdefaultLngStr("DtxtPhone")%>: 507.300.7200</font></p>
</div>

</body>

</html>
