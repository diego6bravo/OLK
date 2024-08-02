<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/extAgentsUsers.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl" <% End If %>="">

<!--#include file="../myHTMLEncode.asp"-->

<head>
<%
      set rx = Server.CreateObject("ADODB.recordset")
If userType = "V" Then
	SelDes = "0"
Else
	sql = "select SelDes from OLKCommon"
	set rx = conn.execute(sql)
	SelDes = rx(0)
	rx.close
End If
sql = "SELECT T0.SlpCode AS User_Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, T0.SlpName) U_Name, "

If Request("agentsusers") <> "" Then
	sql = sql & "Case When T0.SlpCode in (" & Request("agentsusers") & ") or '" & Request("agentsusers") & "' = '-2' Then 'Y' Else 'N' End "
Else
	sql = sql & "'N'"
End If

sql = sql & " [Checked] " & _
	  "FROM OSLP T0 " & _
	  "inner join OLKAgentsAccess T1 on T1.SlpCode = T0.SlpCode and T1.Access <> 'D' " & _
	  "WHERE T0.SlpCode not in (-1) " & _
	  "ORDER BY U_Name "
rx.open sql, conn, 3, 1
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="../general.js"></script>
<link rel="stylesheet" type="text/css" href="../design/<%=SelDes%>/style/stylePopUp.css">
<title><%=getextAgentsUsersLngStr("LttlOLKUsers")%></title>
</head>
<script type="text/javascript">
<!--
function doChkAll()
{
	chkUser = document.frm.chkUser;
	if (chkUser.length)
	{
		for (var i = 0;i<chkUser.length;i++)
		{
			chkUser[i].checked = document.frm.chkAll.checked;
		}
	}
	else
	{
		chkUser.checked = document.frm.chkAll.checked;
	}
}
function checkAll()
{
	var check = true;
	
	chkUser = document.frm.chkUser;
	if (chkUser.length)
	{
		for (var i = 0;i<chkUser.length;i++)
		{
			if (!chkUser[i].checked)
			{
				check = false;
				break;
			}
		}
	}
	else
	{
		check = chkUser.checked;
	}	
	
	document.frm.chkAll.checked = check;
}
function doAccept()
{
	var agentCode = '';
	var agentDesc = '';
	
	if (!document.frm.chkAll.checked)
	{
		chkUser = document.frm.chkUser;
		if (chkUser.length)
		{
			for (var i = 0;i<chkUser.length;i++)
			{
				if (chkUser[i].checked)
				{
					if (agentCode != '')
					{
						agentCode += ', ';
						agentDesc += ', ';
					}
					agentCode += chkUser[i].value.split('|')[0];
					agentDesc += chkUser[i].value.split('|')[1];
				}
			}
		}
		else
		{
			if (chkUser.checked)
			{
				agentCode = chkUser.value.split('|')[0];
				agentDesc = chkUser.value.split('|')[1];
			}
		}
	}
	else
	{
		agentCode = '-2';
		agentDesc = '<%=getextAgentsUsersLngStr("DtxtAll")%>';
	}
	
	opener.agentsTo(agentCode, agentDesc);
	window.close();
}
//-->
</script>
<!--#include file="../design/popvars.inc" -->
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<table border="0" cellpadding="0" width="100%" id="table1">
	<form name="frm">
	<tr class="GeneralTlt">
		<td id="tdMyTtl" width="50%"><%=getextAgentsUsersLngStr("LttlOLKUsers")%>:</td>
	</tr>
	<tr>
		<td width="50%" height="42">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table2">
			<% do while not rx.eof %>
			<tr class="GeneralTbl">
				<td width="100%">
				<input type="checkbox" style="border-style: solid; border-width: 0; background: background-image" name="chkUser" <% If rx("Checked") = "Y" Then %>checked<% End If %> value='<%=RX("User_Code")%>|<%=myHTMLEncode(RX("U_Name"))%>' id='chk<%=RX("User_Code")%>' onclick="javascript:checkAll()"><label for='chk<%=RX("User_Code")%>'><%=RX("U_Name")%></label></td>
			</tr>
			<% rx.movenext
				loop
				If rx.recordcount > 1 then %>
			<tr class="GeneralTbl">
				<td width="100%">
				<input type="checkbox" style="border-style: solid; border-width: 0; background: background-image" name="chkAll" <% If Request("agentsusers") = "-2" Then %>checked<% End If %> value="Y" id="chkAll" onclick="javascript:doChkAll();"><label for="chkAll"><%=getextAgentsUsersLngStr("DtxtAll")%></label></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	</form>
</table>
<center><input type="button" value="<%=getextAgentsUsersLngStr("DtxtAccept")%>" name="btnAccept" onclick="javascript:doAccept();"></center>

</body>
</html>