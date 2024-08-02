<!--#include file="chkLogin.asp" -->
<!--#include file="lang/selectAgents.asp" -->
<!--#include file="myHTMLEncode.asp" -->

<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getselectAgentsLngStr("LttlSelAgent")%></title>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Request.Form.Count > 0 Then
If Request("SlpCode") <> "" Then
	If Request("C1") <> "ON" Then
		SlpCode = Request("SlpCode")
		sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) SlpName from OSLP where SlpCode in (" & SlpCode & ")"
		set rs = conn.execute(sql)
		SlpName = ""
		do while not rs.eof
			If SlpName <> "" Then SlpName = SlpName & ", "
			SlpName = SlpName & rs(0)
		rs.movenext
		loop
	Else
		SlpCode = "-999"
		SlpName = getselectAgentsLngStr("DtxtAll")
	End If
End If %>
<script language="javascript" src="general.js"></script>
<script language="javascript">
opener.agentsToCode("<%=SlpCode%>", "<%=SlpName%>");
window.close();
</script>
<% Else %>

<%
set rx = Server.CreateObject("ADODB.recordset")
If Request("SlpCode") = "" Then
	sql = "select SlpCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) SlpName, 'N' Verfy from OSLP T0 where exists(select 'A' from OLKAgentsAccess where SlpCode = T0.SlpCode)"
Else
	sql = "select T0.SlpCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, T0.SlpName) SlpName, "
	
	If Request("SlpCode") <> "-999" Then
		sql = sql & "Case When T1.SlpCode is null Then 'N' Else 'Y' End Verfy "
	Else
		sql = sql & "'Y' Verfy "
	End If
	
	sql = sql & "from OSLP T0 "
	
	If Request("SlpCode") <> "-999" Then
		sql = sql & "left outer join OSLP T1 on T1.SlpCode = T0.SlpCode and T1.SlpCode in (" & Request("SlpCode") & ") "
	End If
	
	sql = sql & "where exists(select 'A' from OLKAgentsAccess where SlpCode = T0.SlpCode) "
End If
rx.open sql, conn, 3, 1
%>
<script language="javascript" src="general.js"></script>
<SCRIPT LANGUAGE="JavaScript">
var checkflag = "false";
function check(field) 
{
	<% If rx.recordcount > 1 Then %>
	All = field.checked;
	SlpCode = document.frmSelectAgents.SlpCode;
	for (var i = 0;i<SlpCode.length;i++)
	{
		SlpCode[i].checked = All;
	}
	<% Else %>
	document.frmSelectAgents.SlpCode.checked=field.checked;
	<% End If %>
}

function checkAll()
{
var All = true;
<% If rx.recordcount > 1 Then %>
var SlpCode = document.frmSelectAgents.SlpCode;
for (var i = 0;i<SlpCode.length;i++)
{
	if (!SlpCode[i].checked)
	{
		All = false;
		break;
	}
}
document.frmSelectAgents.C1.checked = All;
<% Else %>
if (!document.frmSelectAgents.SlpCode.checked) { document.frmSelectAgents.C1.checked = false; }
<% End If %>
}
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
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
</head>

<body topmargin="0" leftmargin="0" onload="javascript:setTblSet();checkAll();" onbeforeunload="javascript:opener.clearWin();" onscroll="setTblSet();">
<form method="POST" action="selectAgents.asp" name="frmSelectAgents">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table1">
              <tr>
                <td width="50%" class="popupTtl"><%=getselectAgentsLngStr("LttlSelAgent")%></td>
              </tr>
              <tr>
                <td width="50%">
                <table border="0" cellpadding="0" cellspacing="2" bordercolor="#111111" width="100%" id="table2" class="style1">
              <% do while not rx.eof %>
                  <tr>
                    <td width="100%" class="popupOptValue">
				<input type="checkbox" class="noborder" name="SlpCode" <% If rx("Verfy") = "Y" Then %>checked<% End If %> value="<%=RX("SlpCode")%>" id="fps<%=RX("SlpCode")%>" onclick="javascript:checkAll()"><label for="fps<%=RX("SlpCode")%>"><span class="style2"><%=Server.HTMLEncode(RX("SlpName"))%></span></label></td>
                  </tr>
        <% rx.movenext
        loop %>
                  <tr>
                    <td width="100%" class="popupOptValue">
					<input type="checkbox" class="noborder" name="C1" value="ON" onclick="check(this)" id="fp1"><span class="style2"><label for="fp1"><%=getselectAgentsLngStr("DtxtAll")%></label></span></td>
                  </tr>
				<tr height="27">
					<td>&nbsp;</td>
				</tr>

                </table>
                </td>
              </tr>
              </table>
			<input type="hidden" name="pop" value="Y">
				<table cellpadding="0" border="0" width="100%" style="position: absolute;" id="tblSave" bgcolor="#FFFFFF">
					<tr>
						<td style="width: 75px"><input type="submit" value="<%=getselectAgentsLngStr("DtxtAccept")%>" name="B2" class="OlkBtn"></td>
						<td>
							<hr color="#0D85C6" size="1"></td>
						<td style="width: 75px"><input type="button" value="<%=getselectAgentsLngStr("DtxtClose")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
					</tr>
				</table>
			</form>
</body>

<% End IF %></html><% conn.close
set rx = nothing %>