<%@ Language=VBScript %>
<% Session.Timeout=60 %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

strScriptName = LCase(Request.ServerVariables("SCRIPT_NAME"))
If InStr(strScriptName ,"/") > 0 Then 
	strScriptName = right(strScriptName, len(strScriptName) - InStrRev(strScriptName,"/")) 
End If 

strRootPath = Replace(LCase(Request.ServerVariables("URL")), strScriptName, "")

%>
<!--#include file="conn.asp" -->
<!--#include file="lang.asp"-->
<!--#include file="lang/login.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="general.js"></script>
<title>TopManage OLK Module</title>
<script type="text/javascript" src="jQuery/js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="jQuery/js/jquery-ui-1.7.2.custom.min.js"></script>
<script language="javascript">
var txtValUser = '<%=getloginLngStr("LtxtValUser")%>';
var txtValPwd = '<%=getloginLngStr("LtxtValPwd")%>';
var txtValConfBranch = '<%=getloginLngStr("LtxtValConfBranch")%>';
var txtValUType = '<%=getloginLngStr("LtxtValUType")%>';
</script>
<link href="style1.css" rel="stylesheet" type="text/css">
<style type="text/css">
.style1 {
	font-family: Tahoma;
	color: silver;
}
</style>
</head>

<% 
licAgent = False
licAM = False
If Not Session("noLic") Then
	set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
	oLic.LicenceServer = licip
	oLic.LicencePort = licport
	
	If oLic.IsAlive Then
		isNo = False
		Select Case oLic.HasLicence(51)
			Case "YES"
				licAgent = True
			Case "NO"
				ErrMsg = getloginLngStr("DtxtNoOLKLic")
				isNo = True
			Case "EXP"
				ErrMsg = getloginLngStr("DtxtOLKLicExp")
		End Select
		
		Select Case oLic.HasLicence(53)
			Case "YES"
				licAM = True
				ErrMsg = ""
			Case "EXP"
				If isNo Then ErrMsg = getloginLngStr("DtxtOLKLicExp")
		End Select
	Else
		ErrMsg = getloginLngStr("DtxtInactiveLicServer")
	End If
End If

If Request("logout") = "Y" Then 
	Session.Abandon
	If Request("redir") <> "" Then Response.Redirect Request("redir") Else ReloginAnon()
ElseIf Session("vendid") <> "" Then
	Response.Redirect "agent.asp"
End If

If Err.Number = -2147217843 Then
	response.redirect "admin/changeCnPwd.asp?rAction=c_p"
Else
	response.write Err.Description
End If

If Request.Form("UserName") <> "" then uid = Request.Form("UserName") else uid = Request.Cookies("uid")
If Request.Form("Password") <> "" then pwd = Request.Form("Password") else pwd = Request.Cookies("pwd")
If Request.Form("branch") <> "" Then branch = Request.Form("branch") Else branch = Request.Cookies("branch")

If Request.Form.Count > 0 Then
	If Request("EnableBranchs") = "Y" Then 
		Response.cookies("branch").expires = DateAdd("d",60,now())
		Response.cookies("branch").path = "/"  
		Response.cookies("branch") = Request.Form("branch")
	Else
		Response.cookies("branch") = ""
	End If
	Response.cookies("cmp").expires = DateAdd("d",60,now())
	Response.cookies("cmp").path = "/"  
	Response.cookies("cmp") = Request.Form("dbID")
End If
	    	
If Request.Form.Count > 0 Then
	ID = CInt(Request("dbID"))
	set rConn = server.createobject("ADODB.Recordset")
	cmd.ActiveConnection = connCommon 
	cmd.CommandText = "OLKChangeDB"
	cmd.Parameters.Refresh()
	cmd("@ID") = ID
	set rConn = cmd.execute()
	If Not rConn.Eof Then
		If rConn("Verfy") = "Y" Then
			myApp.LoadDBConfigData ID 
		Else
			mySession.EndDBSession
			Response.redirect "default.asp"
		End If
	Else
			Response.redirect "default.asp"
	End If
	
	myApp.ConnectDB
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKVentasLogon" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@userid") = saveHTMLDecode(Request("UserName"), True)
	cmd("@pass") = saveHTMLDecode(Request("Password"), True)
	cmd("@IP") = Left(Request.ServerVariables("remote_addr"), 15)
	If Request("branch") <> "" Then cmd("@branch") = Request("branch") Else cmd("@branch") = -1
	set rd = cmd.execute()
	If RD("Verfy") = "True" then 
		ChangePwd = rd("ChangePwd") = "Y"
	End If
	If RD("Verfy") = "False" Then
		If RD("Access") <> "D" or IsNull(rd("Access")) Then
			strNote = "<font color=""#FF0000"">&nbsp;" & getloginLngStr("LtxtWrongUidPwd") & "</font>"
		Else
			strNote = "<font color=""#FF0000"">&nbsp;" & getloginLngStr("LtxtDisAcct") & "</font>"
		End If
	ElseIf RD("Verfy") = "IPErr" Then 
			strNote = "<font color=""#FF0000"">&nbsp;" & getloginLngStr("LtxtWrongIP") & "</font>"
	ElseIf RD("Verfy") = "Branch" Then 
			strNote = "<font color=""#FF0000"">&nbsp;" & getloginLngStr("LtxtNoBranchAccess") & "</font>"
	ElseIf RD("Verfy") = "True" Then
		If Not Session("noLic") Then
			hasLic = False
			If licAgent Then
				hasLic = oLic.ConfHasLic(51, 0, Request("UserName"))
			End If
			
			If licAM and not hasLic Then
				hasLic = oLic.ConfHasLic(53, 0, Request("UserName"))
			End If
			
			If hasLic Then 
				ExecuteLogin
			Else
				strNote = "<font color=""#FF0000"">&nbsp;" & getloginLngStr("LtxtNoUserLic") & "</font>"
			End If
		Else
			ExecuteLogin
		End If
	End If
End If

Sub ExecuteLogin
	SetSessionVars
	Session("vendid") = RD("SLPCode")
	Session("useraccess") = rd("Access")
	Session("BranchWhs") = rd("BranchWhs")
	Session("AgentWhs") = rd("AgentWhs")
	Session("AgentLastUpdate") = rd("LastUpdate")
	
	Dim myAut
	set myAut = New clsAuthorization
	myAut.LoadAuthorization rd("SlpCode"), Request.Form("dbID")

	Session("ActiveSearch") = "disabled"
	Session.LCID = myApp.LCID
	savePassword()
	If Request("EnableBranchs") = "Y" Then Session("branch") = Request("branch") Else Session("branch") = -1
	Session("sHeight") = Request("sHeight")
	userType = "V"
	mySession.LoginAgent
	If Not ChangePwd Then
		Response.Redirect "agent.asp"
	Else
		userType = ""
		Response.Redirect "changePwdLogon.asp"
	End If
End Sub

Public Sub savePassword()
	If Request("Save") = "ON" and Request.Form("UserName") <> "" Then
		Response.cookies("uid").expires = DateAdd("d",60,now())
		Response.cookies("uid").path = "/"
		Response.cookies("uid") = Request("UserName")  
		Response.cookies("pwd").expires = DateAdd("d",60,now())
		Response.cookies("pwd").path = "/"
		Response.cookies("pwd") = Request("Password")
	ElseIf Request("Save") <> "ON" and Request.Form("UserName") <> "" Then
		Response.cookies("uid") = ""
		Response.cookies("pwd") = ""
	End If
End Sub

%>
<body onload="document.Form1.sHeight.value = screen.availHeight; if (!document.Form1.UserName.disabled)document.Form1.UserName.focus();">
<form method="POST" action="login.asp" name="Form1" onsubmit="return ValidateForm();">
       <div align="center">
  <center>
<table border="0" cellpadding="0" cellspacing="0" width="625">
  	<tr>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="41" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
	<img src="images/spacer.gif" width="69" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="399" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="76" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="40" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
		<p align="center">
		<img name="login_clientsNuevo_r1_c1" src="images/<%=Session("rtl")%>login_clientsNuevo_r1_c1.jpg" width="625" height="29" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="29" border="0" alt=""></td>
	</tr>
	<tr>
		<td rowspan="2" background="images/login_clientsNuevo_r2_c1.jpg">
		<p align="center">
		<img name="login_clientsNuevo_r2_c1" src="images/<%=Session("rtl")%>login_clientsNuevo_r2_c1.jpg" width="41" height="341" border="0" alt=""></td>
		<td colspan="3" valign="middle">
		<table border="0" width="100%" cellpadding="0" id="table1">
			<% If ErrMsg = "" Then %>
			<tr>
				<td>
				<div align="center">
					<br><br><table border="0" cellpadding="0" width="90%" id="table2">
						<tr>
							<td width="169" colspan="2">
							&nbsp;</td>
						</tr>
						<% 
						If Request.ServerVariables("HTTPS") = "off" Then HTTPStr = "http://" Else HTTPStr = "https://"
						curUrl = HTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")), LCase(Request.ServerVariables("PATH_INFO")), strRootPath)	
						
						set rs = Server.CreateObject("ADODB.RecordSet")
						cmd.ActiveConnection = connCommon
						cmd.CommandText = "OLKGetDBList"
						cmd("@UserType") = "A"
						cmd("@curURL") = curURL
						rs.open cmd, , 3, 1
						If rs.recordcount = 0 then %>
						<tr>
							<td bgcolor="#E7F0F5" colspan="2">
							<p align="center"><font size="1" color="#CC0000">
							<%=getloginLngStr("LtxtNoDBConf")%></font></td>
							</tr>
						<% ElseIf rs.recordcount > 0 then
						dbID = rs("ID") %>
						<tr>
							<td width="169" bgcolor="#E7F0F5">
							<p align="center"><font face="Tahoma" size="1">
							<%=getloginLngStr("DtxtCmp")%></font></td>
							<td bgcolor="#E7F0F5"><% If rs.recordcount > 1 Then %>
                            <select class="input" size="1" name="dbID" style="width: 353px; height: 16px;" onchange="javascript:changeDB()">
                            <% do while not rs.eof %>
							<option value="<%=rs("ID")%>" <% if rs("Verfy") = "Y" and Request.Cookies("cmp") = "" or Request("dbID") = "" and Request.Cookies("cmp") = CStr(rs("ID")) or Request("dbID") = CStr(rs("ID")) then
							dbID = rs("ID") %>selected<% end if %>><%=myHTMLEncode(rs("CmpName"))%> 
							<% If myApp.ShowDbName Then %>(<%=rs("dbName")%>)<% End If %></option>
							<% rs.movenext
							loop 
							%>
							</select><% Else
							dbID = rs("ID") %><font face="Tahoma" size="1"><%=myHTMLEncode(rs("CmpName"))%> 
							<% If myApp.ShowDbName Then %>(<%=rs("dbName")%>)<% End If %></font><input type="hidden" name="dbID" value="<%=dbID%>">
							<% End If %></td>
						</tr>
						<% 
						If dbID <> -1 Then isUpdated = myApp.IsDBUpdated(dbID) %>
						<tr id="trUpdate" <% If isUpdated Then %>style="display: none;"<% End If %>>
							<td class="Update" colspan="2"><%=getloginLngStr("LtxtDBNotUpd")%></td>
						</tr>
						<%  If dbID = -1 Then
								rs.movefirst
								dbID = CInt(rs("ID"))
							End If
						end if
						If dbID <> -1 Then
						EnableBranchs = False
						If isUpdated Then 
							myApp.LoadDBConfigData(dbID)
							EnableBranchs = myApp.EnableBranchs
						End If %>
						<tr>
							<td width="169" bgcolor="#E7F0F5">
							<p align="center"><font face="Tahoma" size="1">
							<%=getloginLngStr("DtxtUser")%></font>
							</td>
							<td bgcolor="#E7F0F5">
                            <input type="text" <% If Not isUpdated Then %>disabled<% End If %> name="UserName" id="UserName" size="33" class="input" value="<% If myApp.AllowSavePwd or not myApp.AllowSavePwd and Request.Form.Count > 0 Then %><%=uid%><% End If %>" onfocus="this.select()"></td>
						</tr>
						<tr>
							<td width="169" bgcolor="#E7F0F5">
							<p align="center"><font size="1" face="Tahoma">
														<%=getloginLngStr("DtxtPwd")%></font></td>
							<td bgcolor="#E7F0F5">
							<input type="password" <% If Not isUpdated Then %>disabled<% End If %> name="Password" id="Password" size="33" value="<% If myApp.AllowSavePwd or not myApp.AllowSavePwd and Request.Form.Count > 0 Then %><%=pwd%><% End If %>" class="input" onfocus="this.select()"></td>
						</tr>
						<tr id="trBranch"<% If not EnableBranchs Then %> style="display: none;"<% End If %>>
							<td width="169" bgcolor="#E7F0F5">
							<p align="center"><font face="Tahoma" size="1">
							<%=getloginLngStr("DtxtBranch")%></font></td>
							<td bgcolor="#E7F0F5">
							<select class="input" size="1" name="branch" style="width: 180px; height: 16px;">
                            <% 
                            If EnableBranchs Then
                            myApp.ConnectDB
                            cmd.ActiveConnection = connCommon
                            cmd.CommandText = "DBOLKGetBranchList" & Session("ID")
                            cmd.Parameters.Refresh()
                            cmd("@LanID") = Session("LanID")
                            set rs = cmd.execute()
                            do while not rs.eof %>
							<option value="<%=rs("branchIndex")%>" <% If CStr(branch) = CStr(rs("branchIndex")) then %>selected<% end if %>><%=myHTMLEncode(rs("branchName"))%></option>
							<% rs.movenext
							loop
							End If %>
							</select></td>
						</tr>
						<tr>
							<td width="169" bgcolor="#E7F0F5">
							<p align="center"><font face="Tahoma" size="1">
							<%=getloginLngStr("LtxtLng")%></font></td>
							<td bgcolor="#E7F0F5">
							<select size="1" name="newLng" class="input" onchange="javascript:window.location.href='?newLng=' + this.value">
							<% For i = 0 to UBound(myLanIndex) %>
							<option <% If Request.Cookies("myLng") = myLanIndex(i)(0) Then %>selected<% End If %> value="<%=myLanIndex(i)(0)%>">
							<%=myLanIndex(i)(1)%></option>
							<% Next %>
							</select></td>
						</tr>
						<% If myApp.AllowSavePwd Then %>
						<tr>
							<td colspan="2" valign="top" bgcolor="#E7F0F5">
							<div align="center">
                              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="200" id="AutoNumber1">
                                <tr>
                                  <td width="163">
                                  <p align="right"><font face="Tahoma" size="1">
                                  <label for="Save"><%=getloginLngStr("LtxtSavePwd")%>:</label></font></td>
                                  <td width="37">
                                  <input type="checkbox" <% If Not isUpdated Then %>disabled<% End If %> name="Save" id="Save" value="ON" <% If Request.Cookies("uid") <> "" then Response.write "checked"%>></td>
                                </tr>
                              </table>
                            </div>
                            </td>
						</tr>
						<% End If %>
						<% If strNote <> "" Then %>
						<tr>
							<td colspan="2" valign="top" bgcolor="#E7F0F5">
							<p align="center"><font size="1" face="Verdana"><%=strNote%>
							</font></td>
						</tr>
						<% End If %>
						<tr>
							<td colspan="2" valign="top" bgcolor="#E7F0F5">
							<p align="center">
							<input type="submit" <% If Not isUpdated Then %>disabled<% End If %> value="<%=getloginLngStr("LtxtEnter")%>" id="btnEnter" name="btnEnter" style="font-family: Tahoma; font-size: 10px; border: 1px solid #10699C; background-color: #FFFFFF; width:60;"></td>
						</tr>
						<% End If %>
						</table>
						<input type="hidden" name="EnableBranchs" id="EnableBranchs" value="<%=GetYN(EnableBranchs)%>">
				</div>
				</td>
			</tr>
			<% Else %>
			<tr>
				<td height="50" valign="middle">
				<p align="center"><font size="1" color="#CC0000">
				<%=ErrMsg%></font></td>
			</tr>
			<% End If %>
			<tr>
				<td><center>
                <p><font size="2" face="Verdana"></font></p>
                </center></td>
			</tr>
		</table></td>
		<td rowspan="2">
		<p align="center">
		<img name="login_clientsNuevo_r2_c5" src="images/<%=Session("rtl")%>login_clientsNuevo_r2_c5.jpg" width="40" height="341" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="239" border="0" alt=""></td>
	</tr>
	<tr>
		<td>
		<p align="center">
		<img name="login_clientsNuevo_r3_c2" src="images/login_clientsNuevo_r3_c2.jpg" width="69" height="102" border="0" alt=""></td>
		<td>
		<p align="center">
		<img name="login_clientsNuevo_r3_c3" src="images/<%=Session("rtl")%>login_clientsNuevo_r3_c3.jpg" width="399" height="102" border="0" alt=""></td>
		<td>
		<p align="center">
		<img name="login_clientsNuevo_r3_c4" src="images/<%=Session("rtl")%>login_clientsNuevo_r3_c4.jpg" width="76" height="102" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="102" border="0" alt=""></td>
	</tr>
</table>
  </center>
</div>
		<input type="hidden" name="sHeight" value="400">
		<input type="hidden" name="other" value="<%=Request("Other")%>">
</form>
<p align="center"><font color="#C0C0C0" size="1" face="Verdana">
<a href="http://www.topmanage.com.pa/"><span class="style1">TopManage</span></a> &reg;</font><font face="Tahoma" color="#c0c0c0" size="1"> 2002 - 2012 - <%=getloginLngStr("DtxtEMail")%>: <a href="mailto:info@topmanage.com.pa"><font color="#c0c0c0">
info@topmanage.com.pa</font></a> - <%=getloginLngStr("DtxtPhone")%>: 507.300.7200</font></p>
<p align="center">&nbsp;</p>
<script language="javascript" src="default.js"></script>
</body>


<% mySession.EndDBSession
conn.close %></html>
<% Sub SetSessionVars
	Response.Cookies("olkdb") = olkdb
End Sub

Sub ReloginAnon()
	If Request.ServerVariables("HTTPS") = "off" Then HTTPStr = "http://" Else HTTPStr = "https://"
	curUrl = HTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")), LCase(Request.ServerVariables("PATH_INFO")), strRootPath)
	set rs = Server.CreateObject("ADODB.RecordSet")
	myApp.ConnectCommon
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "OLKValidateDomain"
	cmd.Parameters.Refresh()
	cmd("@newAddress") = curUrl
	cmd("@curdb") = ""
	set rs = cmd.execute()
	If Not rs.Eof Then
		If rs(0) = "Y" Then
			ID = CInt(rs("ID"))
			myApp.LoadDBConfigData(ID)
			If myApp.EnableAnSesion Then
				Session("UserName") = "-Anon-"
				Session("PriceList") = myApp.AnSesListNum
				Session("vendid") = -1
				userType = "C"
				Session("RetVal") = -1
				Session.LCID = myApp.LCID
				Session("branch") = -1
				
				Response.Cookies("olkdb") = olkdb
				
			    Response.cookies("OLKAnon").expires = DateAdd("d",30,now())
				Response.cookies("OLKAnon") = "Y"
				
				Response.Redirect "default.asp"
			End If
		End If
	End If
End Sub
 %>