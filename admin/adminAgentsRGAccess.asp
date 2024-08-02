<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminAgentsRGAccess.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select IsNull(SlpName, '') SlpName, Case When Exists(select 'A' from OLKAgentsAccess where SlpCode = T0.SlpCode) Then 'Y' Else 'N' End VerfyUser from OSLP T0 where SlpCode = " & Request("SlpCode")
set rs = conn.execute(sql)
SlpName = rs("SlpName")
VerfyUser = rs("VerfyUser")
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<title><%=getadminAgentsRGAccessLngStr("LttlAccessGroups")%> - <%=SlpName%></title>
<style type="text/css">
.style1 {
	color: #2E94D4;
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

<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();" onload="setTblSet();" onscroll="setTblSet();">
<% If Request.Form.Count = 0 Then
sql = "select T0.rgIndex, T0.rgName, Case When T1.rgIndex is null Then 'N' Else 'Y' End Verfy " & _
"from OLKRG T0 " & _
"left outer join OLKRGAccess T1 on T1.rgIndex = T0.rgIndex and T1.SlpCode = " & Request("SlpCode") & " " & _
"where T0.SuperUser = 'N' and T0.UserType = 'V' "
set rs = conn.execute(sql)
%>
<form method="POST" action="adminAgentsRGAccess.asp" webbot-action="--WEBBOT-SELF--">
<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<tr>
		<td class="popupTtl"><%=getadminAgentsRGAccessLngStr("LttlAccessGroups")%> - <%=SlpName%></td>
	</tr>
	<% do while not rs.eof
	rgIndex = rs("rgIndex") %>
	<tr class="popupOptValue">
		<td>
		<input type="checkbox" name="rgIndex" <% If rs("Verfy") = "Y" Then %>checked<% End If %> value="<%=rgIndex%>" id="rgIndex<%=rgIndex%>" class="noborder"><label for="rgIndex<%=rgIndex%>"><%=rs("rgName")%></label></td>
	</tr>
	<% rs.movenext
	loop %>
	<tr height="27">
		<td>&nbsp;</td>
	</tr>
</table>
<table cellpadding="0" border="0" width="100%" id="tblSave" style="position: absolute; ">
	<tr>
		<td width="75"><input type="submit" name="btnSave" value="<%=getadminAgentsRGAccessLngStr("DtxtSave")%>" class="OlkBtn">
		</td>
		<td><hr size="1" class="style1"></td>
	</tr>
</table>
<input type="hidden" name="SlpCode" value="<%=Request("SlpCode")%>">
<input type="hidden" name="VerfyUser" value="<%=VerfyUser%>">
</form>
<% Else 
sql = 	"declare @SlpCode int set @SlpCode = " & Request("SlpCode") & " " & _
		"delete OLKRGAccess where SlpCode = @SlpCode "
If Request("rgIndex") <> "" and Request("VerfyUser") = "Y" Then
	sql = sql & "insert OLKRGAccess(SlpCode, rgIndex) select @SlpCode, Value from OLKCommon.dbo.OLKSplit('" & Request("rgIndex") & "', ', ')"
End If
conn.execute(sql) %>
<script language="javascript">
<% If Request("VerfyUser") = "N" Then %>
opener.setNewRGAccess('<%=Request("rgIndex")%>');
<% End If %>
window.close();
</script>
<% End If %>
</body>

</html>
<% conn.close
set rs = nothing %>