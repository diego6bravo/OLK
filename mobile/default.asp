<%@ Language=VBScript %>
<%
Session.Timeout=60
response.buffer = true %>
<!--#include file="lang.asp"-->
<!--#include file="lang/default.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="conn.asp" -->
<!--#include file="authorizationClass.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="mobileoptimized" content="0">
<meta name="viewport" content="width=320,user-scalable=false">
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
licMobile = False
licAM = False

set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
oLic.LicenceServer = licip
oLic.LicencePort = licport

If oLic.IsAlive Then
	isNo = False
	Select Case oLic.HasLicence(52)
		Case "YES"
			licMobile = True
		Case "NO"
			isNo = True
			ErrMsg = "" & getdefaultLngStr("LtxtNoOLKLic") & ""
		Case "EXP"
			ErrMsg = "" & getdefaultLngStr("LtxtOLKLicExp") & ""
	End Select
	
	Select Case oLic.HasLicence(53)
		Case "YES"
			licAM = True
			ErrMsg = ""
		Case "EXP"
			If isNo Then ErrMsg = "" & getdefaultLngStr("LtxtOLKLicExp") & ""
	End Select
Else
	ErrMsg = "" & getdefaultLngStr("LtxtInactiveLicServer") & ""
End If

If Request("logout") = "Y" Then Session.Abandon

If Request.ServerVariables("HTTPS") = "off" Then HTTPStr = "http://" Else HTTPStr = "https://"
curUrl = HTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")), LCase(Request.ServerVariables("PATH_INFO")), strRootPath)	

set rs = Server.CreateObject("ADODB.RecordSet")
cmd.ActiveConnection = connCommon
cmd.CommandText = "OLKGetDBList"
cmd("@UserType") = "M"
cmd("@curURL") = curURL
cmd("@UpdOnly") = "Y"
rs.open cmd, , 3, 1

dbID = -1
          %>
<script language="javascript" src="general.js"></script>
<script type="text/javascript">
var txtValUser = '<%=getdefaultLngStr("LtxtValUser")%>';
var txtValPwd = '<%=getdefaultLngStr("LtxtValPwd")%>';
var txtValConfBranch = '<%=getdefaultLngStr("LtxtValConfBranch")%>';
</script>
<script language="javascript" src="default.js"></script>
<body topmargin="0" onload="javascript:focusLogin();"<% If Session("rtl") <> "" Then %> dir="rtl"<% End If %>>

<div align="center">
  <center>
  <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
    <tr>
      <td>
      <p align="center"><b><font face="Verdana" size="1" color="#FDAF2F">
      <%=getdefaultLngStr("LtxtWelcome")%></font></b></td>
    </tr>
    <tr>
      <td>
      <p align="center">
      <img border="0" src="images/pocket_olkicon.gif"></td>
    </tr>
    <tr>
      <td style="font-size: 10px">
      &nbsp;</td>
    </tr>
    <% If ErrMsg = "" Then %>
      <form method="POST" name="frmLogin" action="default.asp" onsubmit="javascritp:return ValidateForm();">
		<% If rs.recordcount = 0 then %>
		<tr>
			<td bgcolor="#F0F8FF">
			<p align="center">
			<font size="1" color="#CC0000" face="Verdana"><%=getdefaultLngStr("LtxtNoDBConf")%></font></td>
			</tr>
		<% ElseIf rs.recordcount > 0 then %>
		<tr>
			<td bgcolor="#F0F8FF"><p align="center">&nbsp;
			<% If rs.recordcount > 1 Then %>
            <select size="1" name="dbID" style="width: 180; height: 16; font-family:Verdana; font-size:10px" onchange="javascript:changeDB()">
            <% do while not rs.eof %>
			<option value="<%=rs("ID")%>" <% if rs("Verfy") = "Y" and Request.Cookies("cmp") = "" or Request("dbID") = "" and Request.Cookies("cmp") = CStr(rs("ID")) or Request("dbID") = CStr(rs("ID")) then
			dbID = CInt(rs("ID")) %>selected<% end if %>><%=myHTMLEncode(rs("CmpName"))%></option>
			<% rs.movenext
			loop 
			%>
			</select><% Else
			dbID = CInt(rs("ID")) %>
			<font face="Verdana, Geneva, Tahoma, sans-serif" size="1"><%=myHTMLEncode(rs("CmpName"))%></font>
			<input type="hidden" name="dbID" value="<%=dbID%>"><% End If %></td>
		</tr>
		<% If dbID = -1 Then 
			rs.movefirst
			dbID = rs("ID")
		End If
		If dbID <> -1 Then isUpdated = myApp.IsDBUpdated(dbID)
		EnableBranchs = False
		If Not isUpdated Then %>
		<tr bgcolor="#FFD2A6" align="center">
			<td class="Update"><font face="Verdana, Geneva, Tahoma, sans-serif" size="1"><b><%=getdefaultLngStr("LtxtDBNotUpd")%></b></font></td>
		</tr>
		<% Else
			myApp.LoadDBConfigData(dbID)
			EnableBranchs = myApp.EnableBranchs
		End If
		End If %>
        <tr>
      <td bgcolor="#F0F8FF">
        <div align="center">
          <center>
          <table border="0" cellpadding="0"  bordercolor="#111111" style="width: 95%">
            <tr>
              <td bgcolor="#DDEFFF" class="style2"><font size="1" face="Verdana">
              <%=getdefaultLngStr("DtxtUser")%>:</font></td>
              <td bgcolor="#DDEFFF" class="style1"><input name="UserName" <% If Not isUpdated Then %>disabled<% End If %> style="font-size:12px; size=; float:left; width:90%" size="12" value="<% If Request.Form.Count = 0 and myApp.AllowSavePwd Then %><%=Request.Cookies("UId")%><% Else %><%=Request.Form("UserName")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
            </tr>
            <tr>
              <td bgcolor="#DDEFFF" class="style2"><font size="1" face="Verdana">
              <%=getdefaultLngStr("DtxtPwd")%>:</font></td>
              <td bgcolor="#DDEFFF" class="style1"><input type="password" <% If Not isUpdated Then %>disabled<% End If %> name="Password" style="font-size:12px; size=; float:left; width:90%" size="12" value="<% If Request.Form.Count = 0 and myApp.AllowSavePwd Then %><%=Request.Cookies("pwd")%><% Else %><%=Request.Form("Password")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
            </tr>
            <tr>
              <td bgcolor="#DDEFFF" class="style2"><font size="1" face="Verdana">
              <%=getdefaultLngStr("LtxtLng")%>:</font></td>
              <td bgcolor="#DDEFFF" class="style1">
				<select size="1" name="newLng" class="input" onchange="javascript:window.location.href='?newLng=' + this.value" style="font-family:Verdana; font-size:10px">
				<% For i = 0 to UBound(myLanIndex) %>
				<option <% If Request.Cookies("myLng") = myLanIndex(i)(0) Then %>selected<% End If %> value="<%=myLanIndex(i)(0)%>">
				<%=myLanIndex(i)(1)%></option>
				<% Next %>
				</select></td>
            </tr>
			<input type="hidden" name="EnableBranchs" value="<%=GetYN(EnableBranchs)%>">
            <tr id="trBranch"<% If Not EnableBranchs Then %> style="display: none"<% End If %>>
              <td bgcolor="#DDEFFF" class="style2"><font size="1" face="Verdana">
              <%=getdefaultLngStr("DtxtBranch")%>:</font></td>
              <td bgcolor="#DDEFFF" class="style1">
			<select class="input" size="1" name="branch" style="width: 100%; height: 16;font-family :Verdana; font-size:10px">
            <% 
            If EnableBranchs Then
            myApp.ConnectDB
            cmd.ActiveConnection = connCommon
            cmd.CommandText = "DBOLKGetBranchList" & Session("ID")
            cmd.Parameters.Refresh()
            cmd("@LanID") = Session("LanID")
            set rs = cmd.execute()
            do while not rs.eof %>
			<option value="<%=rs("branchIndex")%>" <% if Request.Form("branch") = "" and CStr(Request.Cookies("branch")) = CStr(rs("branchIndex")) or CStr(Request.Form("branch")) = CStr(rs("branchIndex")) then %>selected<% end if %>><%=myHTMLEncode(rs("branchName"))%></option>
			<% rs.movenext
			loop
			End If %>
			</select></td>
            </tr>
            <% If myApp.AllowSavePwd Then %>
            <tr>
              <td colspan="2" bgcolor="#DDEFFF">
              <div align="center">
                <center>
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="121" id="AutoNumber3">
                  <tr>
                    <td width="20">
					<input type="checkbox" name="Save" <% If Not isUpdated Then %>disabled<% End If %> value="ON" id="save" <% If Request.Cookies("uid") <> "" then Response.write "checked"%>></td>
                    <td width="101"><font size="1" face="Arial">&nbsp;<font color="#005782"><label for="save"><%=getdefaultLngStr("LtxtSavePwd")%></label></font></font></td>
                  </tr>
                </table>
                </center>
              </div>
              </td>
            </tr>
            <% End If %>
            <tr>
              <td colspan="2" bgcolor="#DDEFFF">
              <p align="center">
              <input type="submit" value="<%=getdefaultLngStr("LtxtEnter")%>" <% If Not isUpdated Then %>disabled<% End If %> style="color: #005782; font-family: verdana; font-size: 10px; border: 1px solid #006699; background-color: #C1E1FF" name="btnEnter"></td>
            </tr>
          </table>
          </center>
        </div>
        </td></tr><input type="hidden" name="Other" value="<%=Request("Other")%>">
      </form>
      <% End If %>
      </table>
     <% If ErrMsg <> "" Then %>
     <center><b><font color="#FF0000" face="Verdana" size="1">&nbsp;<%=ErrMsg%></font></b></center>
     <% ElseIf Request.Form("UserName") = "" Then %>
    <center><b><font face="Verdana" size="1">&nbsp;<%=getdefaultLngStr("LtxtEnterUidPwd")%></font></b></center>
    <% ElseIf Request.Form("UserName") <> "" and Request("btnEnter") <> "" Then
    		set cmd = Server.CreateObject("ADODB.Command")
    		cmd.ActiveConnection = connCommon
    		cmd.CommandType = adCmdStoredProc
    		cmd.CommandText = "DBOLKVentasLogon" & Session("ID")
    		cmd.Parameters.Refresh()
    		cmd("@userid") = saveHTMLDecode(Request("UserName"), True)
    		cmd("@pass") = oLic.GetEncPwd(Request("Password"))
    		cmd("@IP") = Left(Request.ServerVariables("remote_addr"), 15)
    		If Request("branch") <> "" Then cmd("@branch") = Request("branch") Else cmd("@branch") = -1
    		set rs = cmd.execute()
    	If RS("Verfy") = "False" Then %>
    	<center><b><font color="#FF0000" face="Verdana" size="1">&nbsp;<%=getdefaultLngStr("LtxtWrongUidPwd")%></font></b></center>
    	<% ElseIf rs("Verfy") = "IPErr" Then %>
    	<center><b><font color="#FF0000" face="Verdana" size="1">&nbsp; <%=getdefaultLngStr("LtxtWrongIP")%>.</font></b></center>
    	<% ElseIf rs("Verfy") = "Branch" Then %>
    	<center><b><font color="#FF0000" face="Verdana" size="1">&nbsp; <%=getdefaultLngStr("LtxtNoBranchAccess")%>.</font></b></center>
    	<% ElseIf RS("Verfy") = "True" Then
		
			hasLic = False
			If licMobile Then
				hasLic = oLic.ConfHasLic(52, 0, Request("UserName"))
			End If
			
			If licAM and not hasLic Then
				hasLic = oLic.ConfHasLic(53, 0, Request("UserName"))
			End If
			
			If hasLic Then 
				ExecuteLogin
			Else %><center><b><font color="#FF0000" face="Verdana" size="1">&nbsp; <%=getdefaultLngStr("LtxtNoUserLic")%>.</font></b></center><%
			End If
			
    	
    	End If
    End If
    
    Sub ExecuteLogin
    	Session("vendid") = RS("SLPCode")
    	Session("useraccess") = rs("Access")
		Session("BranchWhs") = rs("BranchWhs")
		Session("AgentWhs") = rs("AgentWhs")
		Session("AgentLastUpdate") = rs("LastUpdate")
    	Session.LCID = 6154 'myApp.LCID
		ChangePwd = rs("ChangePwd") = "Y"
		Dim myAut
		set myAut = New clsAuthorization
		myAut.LoadAuthorization rs("SlpCode"), ""

    	If myApp.EnableBranchs Then Session("branch") = Request("branch") Else Session("branch") = -1
    	
    	sql = 	"select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName)  SLPName, SLPCode, " & _
    			"(select Access from olkagentsaccess where slpcode = oslp.slpcode) Access from oslp where SlpCode = " & Session("vendid")
    	set rs = conn.execute(sql)

    	Session("vendnm") = RS("SLPName")
    	If Request("Save") = "ON" Then
		  Response.Cookies("uid").expires = DateAdd("d",60,now())
		  Response.Cookies("uid").path = "/"
		  Response.Cookies("uid") = Request("UserName")  
		  Response.Cookies("pwd").expires = DateAdd("d",60,now())
		  Response.Cookies("pwd").path = "/"
		  Response.Cookies("pwd") = Request("Password")
    	Else
		  Response.Cookies("uid") = ""
		  Response.Cookies("pwd") = ""
		  Response.cookies("branch") = ""
		End If
		If myApp.EnableBranchs Then 
			Response.cookies("branch").expires = DateAdd("d",60,now())
			Response.cookies("branch").path = "/"  
			Response.cookies("branch") = Request("branch")
		Else
			Response.cookies("branch") = ""
		End If
	  Response.cookies("cmp").expires = DateAdd("d",60,now())
	  Response.cookies("cmp").path = "/"  
	  Response.cookies("cmp") = Request("dbID")
		    userType = "V"
			mySession.LoginAgent
	    	If Not ChangePwd Then
		    	Response.Redirect "operaciones.asp?cmd=home"
		    Else
		    	userType = ""
		    	Response.Redirect "changePwdLogon.asp"
		    End If
    End Sub
%>
</center>
</div>
<script language="javascript">
function changeDB()
{
	document.frmLogin.submit();
}
function EnableBranchs(EnableBranchs)
{
	document.frmLogin.EnableBranchs.value = EnableBranchs;
	if (EnableBranchs == "Y")
	{
		document.getElementById('trBranch').style.display = '';
	}
	else
	{
		document.getElementById('trBranch').style.display = 'none';
	}
}

function getBranchs() { return document.frmLogin.branch }
</script>
<iframe name="addFrame" id="addFrame" src="" style="display: none">
</iframe>
</body>
<% set rs = nothing
conn.close %></html>