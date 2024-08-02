<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminAgentsAutCopy.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->

<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Request("dbID") = "" Then
	sql = "select IsNull(SlpName, '') SlpName from OSLP T0 where SlpCode = " & Request("SlpCode")
	set rs = conn.execute(sql)
	SlpName = rs("SlpName")
Else
	SlpName = Request("UserName")
	sql = "select dbName from OLKDBA where ID = " & Request("dbID")
	set rs = connCommon.execute(sql)
	dbName = rs(0)
End If
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<title><%=Replace(getadminAgentsAutCopyLngStr("LttlCopyAut"), "{0}", SlpName)%></title>
<style type="text/css">
.style1 {
	color: #2E94D4;
}
.style2 {
	background-color: #E1F3FD;
	font-family: Verdana;
	font-weight: bold;
	font-size: 10px;
	color: #31659C;
	text-align: center;
}
.style3 {
	background-color: #FFFFFF;
}
</style>
<script type="text/javascript" src="general.js"></script>
<script language="javascript">
function setTblSet()
{
	if (browserDetect() == 'msie')
	{
		tblSave.style.top = document.body.offsetHeight-31+document.body.scrollTop;
	}
	else if (browserDetect() == 'opera')
	{
		tblSave.style.top = document.body.offsetHeight-27+document.body.scrollTop;
	}
	else //firefox & others
	{
		tblSave.style.top = window.innerHeight-27+document.body.scrollTop;
	}
}
</script>
</head>

<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();"<% If Request.Form.Count = 0 Then %> onload="setTblSet();" onscroll="setTblSet();"<% End If %>>
<% If Request.Form.Count = 0 Then %>
<script type="text/javascript">
<!--
function valFrm()
{
	if (document.frmCopy.SlpCodeTo)
	{
		if (document.frmCopy.SlpCodeTo.length)
		{
			var found = false;
			for (var i = 0;i<document.frmCopy.SlpCodeTo.length;i++)
			{
				if (document.frmCopy.SlpCodeTo[i].checked)
				{
					found = true;
					break;
				}
			}
			if (!found)
			{
				alert('<%=getadminAgentsAutCopyLngStr("LtxtValSelUser")%>');
				return false;
			}
		}
		else
		{
			if (!document.frmCopy.SlpCodeTo.checked)
			{
				alert('<%=getadminAgentsAutCopyLngStr("LtxtValSelUser")%>');
				return false;
			}
		}
	}
	else
	{
		alert('<%=getadminAgentsAutCopyLngStr("LtxtNoUserToCopy")%>');
		return false;
	}
	return true;
}
//-->
</script>
<form method="POST" name="frmCopy" action="adminAgentsAutCopy.asp" onsubmit="return valFrm();" webbot-action="--WEBBOT-SELF--">
<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<tr>
		<td class="popupTtl" colspan="2"><%=Replace(getadminAgentsAutCopyLngStr("LttlCopyAut"), "{0}", SlpName)%></td>
	</tr>
	<tr>
		<td class="popupTtl"><%=getadminAgentsAutCopyLngStr("LtxtRollFilter")%></td>
		<td class="popupTtl"><select name="cmbRole" onchange="javascript:window.location.href='adminAgentsAutCopy.asp?<% If Request("dbID") <> "" Then %>dbID=<%=Request("dbID")%>&UserName=<%=SlpName%>&<% End If %>SlpCode=<%=Request("SlpCode")%>&Access=<%=Request("Access")%>&role='+this.value;">
		<option value=""><%=getadminAgentsAutCopyLngStr("DtxtAll")%></option>
		<% 
		If Request("dbID") = "" Then
			sql = "select T0.typeID, T0.name from OHTY T0"
		Else
			sql = "select T0.typeID, T0.name from [" & dbName & "]..OHTY T0"
		End If
		set rs = conn.execute(sql) 
		do while not rs.eof %>
		<option <% If Request("role") = CStr(rs("typeID")) Then %>selected<% End If %> value="<%=rs("typeID")%>"><%=rs("name")%></option>
		<% rs.movenext
		loop %>
		</select></td>
	</tr>
	<tr>
		<td class="popupTtl" <% If Request("role") <> "" Then %>colspan="2"<% End If %>><%=getadminAgentsAutCopyLngStr("DtxtAgent")%></td>
		<% If Request("role") = "" Then %>
		<td class="style2"><%=getadminAgentsAutCopyLngStr("DtxtRole")%></td><% End If %>
	</tr>
	<% 
	If Request("dbID") = "" Then
		sql = "select X0.SlpCode, IsNull(SlpName, '') SlpName, T2.name Role " & _
			"from OLKAgentsAccess X0  " & _
			"inner join OSLP T0 on T0.SlpCode = X0.SlpCode " & _
			"left outer join OHEM T1 on T1.salesPrson = T0.SlpCode " & _
			"left outer join OHTY T2 on T2.typeID = IsNull(T1.type, (select roleID from HEM6 where empID = T1.empID and line = (select Min(line) from HEM6 where empID = T1.empID))) " & _
			"where X0.Access = '" & Request("Access") & "' and X0.SlpCode <> " & Request("SlpCode")
		If Request("role") <> "" Then sql = sql & " and exists(select '' from HEM6 where empID = T1.empID and roleID = " & Request("role") & ") "
	Else
		sql = "select X1.SlpCode, IsNull(X0.UserName, '') SlpName, T2.name Role " & _  
			"from OLKCommon..OLKAgentsAccess X0  " & _  
			"inner join OLKCommon..OLKAgentsAccessDB X1 on X1.dbID = " & Request("dbID") & " and X1.UserName = X0.UserName " & _  
			"left outer join [" & dbName & "]..OHEM T1 on T1.salesPrson = X1.SlpCode " & _  
			"left outer join [" & dbName & "]..OHTY T2 on T2.typeID = IsNull(T1.type, (select roleID from [" & dbName & "]..HEM6 where empID = T1.empID and line = (select Min(line) from [" & dbName & "]..HEM6 where empID = T1.empID))) " & _  
			"inner join [" & dbName & "]..OLKAgentsAccess T3 on T3.SlpCode = X1.SlpCode " & _  
			"where T3.Access = '" & Request("Access") & "' and X0.UserName <> N'" & saveHTMLDecode(SlpName, False) & "' "
			
		If Request("role") <> "" Then sql = sql & " and exists(select '' from [" & dbName & "]..HEM6 where empID = T1.empID and roleID = " & Request("role") & ") " 
	End If
	set rs = conn.execute(sql)
	do while not rs.eof %>
	<tr class="popupOptValue">
		<td <% If Request("role") <> "" Then %>colspan="2"<% End If %>>
		<input type="checkbox" name="SlpCodeTo" value="<%=rs("SlpCode")%>" id="copyTo<%=rs("SlpCode")%>" class="noborder"><label for="copyTo<%=rs("SlpCode")%>"><%=rs("SlpName")%></label></td>
		<% If Request("role") = "" Then %>
		<td>
		<%=rs("Role")%></td><% End If %>
	</tr>
	<% rs.movenext
	loop %>
	<tr height="27">
		<td <% If Request("role") <> "" Then %>colspan="2"<% End If %>>&nbsp;</td>
		<% If Request("role") = "" Then %>
		<td>&nbsp;</td><% End If %>
	</tr>
</table>
<table cellpadding="0" border="0" width="100%" id="tblSave" style="position: absolute; " class="style3">
	<tr>
		<td width="75"><input type="submit" name="btnSave" value="<%=getadminAgentsAutCopyLngStr("DtxtSave")%>" class="OlkBtn">
		</td>
		<td><hr size="1" class="style1"></td>
		<td width="75"><input type="button" name="btnCancel" value="<%=getadminAgentsAutCopyLngStr("DtxtCancel")%>" class="OlkBtn" onclick="window.close();">
		</td>
	</tr>
</table>
<% If Request("dbID") <> "" Then %><input type="hidden" name="dbID" value="<%=Request("dbID")%>">
<input type="hidden" name="UserName" value="<%=Server.HTMLEncode(SlpName)%>"><% End If %>
<input type="hidden" name="SlpCode" value="<%=Request("SlpCode")%>">
<input type="hidden" name="Access" value="<%=Request("Access")%>">
</form>
<% Else 
If Request("dbID") = "" Then
	sql = "update OLKAgentsAccess set [Authorization] = (select [Authorization] from OLKAgentsAccess where SlpCode = " & Request("SlpCode")& "), " & _
			"MaxDiscount = (select MaxDiscount from OLKAgentsAccess where SlpCode = " & Request("SlpCode")& "), " & _
			"MaxDocDiscount = (select MaxDocDiscount from OLKAgentsAccess where SlpCode = " & Request("SlpCode")& ") where SlpCode in (" & Request("SlpCodeTo") & ")"
Else
	sql = "update [" & dbName & "]..OLKAgentsAccess set [Authorization] = (select [Authorization] from [" & dbName & "]..OLKAgentsAccess where SlpCode = " & Request("SlpCode")& "), " & _
			"MaxDiscount = (select MaxDiscount from [" & dbName & "]..OLKAgentsAccess where SlpCode = " & Request("SlpCode")& "), " & _
			"MaxDocDiscount = (select MaxDocDiscount from [" & dbName & "]..OLKAgentsAccess where SlpCode = " & Request("SlpCode")& ") where SlpCode in (" & Request("SlpCodeTo") & ")"
End If
conn.execute(sql) %>
<script language="javascript">
window.close();
</script>
<% End If %>
</body>

</html>
<% conn.close
set rs = nothing %>