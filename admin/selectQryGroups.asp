<!--#include file="chkLogin.asp" -->
<!--#include file="lang/selectQryGroups.asp" -->
<!--#include file="myHTMLEncode.asp" -->

<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getselectQryGroupsLngStr("LttlQryGroups")%></title>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


set rx = Server.CreateObject("ADODB.recordset")

sql = "select ItmsTypCod, Convert(nvarchar(2),ItmsTypCod) + ' - ' + ItmsGrpNam ItmsGrpNam, "

If Request("QryGroups") <> "" Then
	sql = sql & "Case When ItmsTypCod in (" & Request("QryGroups") & ") Then 'Y' Else 'N' End "
Else
	sql = sql & "'N' "
End If

sql = sql & "Verfy from OITG order by 1"

rx.open sql, conn, 3, 1
%>
<script language="javascript" src="general.js"></script>
<SCRIPT LANGUAGE="JavaScript">
var checkflag = "false";
function check(field) 
{
	All = field.checked;
	QryGroup = document.frmSelectQryGroups.QryGroup;
	for (var i = 0;i<QryGroup.length;i++)
	{
		QryGroup[i].checked = All;
	}
}

function checkAll()
{
	var All = true;
	var QryGroup = document.frmSelectQryGroups.QryGroup;
	for (var i = 0;i<QryGroup.length;i++)
	{
		if (!QryGroup[i].checked)
		{
			All = false;
			break;
		}
	}
	document.frmSelectQryGroups.C1.checked = All;
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
function doAccept()
{
	var retVal = '';
	var QryGroups = document.frmSelectQryGroups.QryGroup;
	for (var i = 0;i<QryGroups.length;i++)
	{
		if (QryGroups[i].checked)
		{
			retVal += (retVal != '' ? ',' : '') + QryGroups[i].value;
		}
	}
	opener.setQryGroups(retVal);
	window.close();
}
</script>
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
</head>

<body topmargin="0" leftmargin="0" onload="javascript:setTblSet();checkAll();" onbeforeunload="javascript:opener.clearWin();" onscroll="setTblSet();">
<form method="POST" action="selectQryGroups.asp" name="frmSelectQryGroups">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table1">
              <tr>
                <td width="50%" class="popupTtl"><%=getselectQryGroupsLngStr("LttlQryGroups")%></td>
              </tr>
              <tr>
                <td width="50%">
                <table border="0" cellpadding="0" cellspacing="2" bordercolor="#111111" width="100%" id="table2" class="style1">
              <% do while not rx.eof %>
                  <tr>
                    <td width="100%" class="popupOptValue">
				<input type="checkbox" class="noborder" name="QryGroup" <% If rx("Verfy") = "Y" Then %>checked<% End If %> value="<%=RX("ItmsTypCod")%>" id="fps<%=RX("ItmsTypCod")%>" onclick="javascript:checkAll()"><label for="fps<%=RX("ItmsTypCod")%>"><span class="style2"><%=RX("ItmsGrpNam")%></span></label></td>
                  </tr>
		        <% rx.movenext
		        loop %>
                  <tr>
                    <td width="100%" class="popupOptValue">
					<input type="checkbox" class="noborder" name="C1" value="ON" onclick="check(this)" id="fp1"><span class="style2"><label for="fp1"><%=getselectQryGroupsLngStr("DtxtAll")%></label></span></td>
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
						<td style="width: 75px"><input type="button" value="<%=getselectQryGroupsLngStr("DtxtAccept")%>" name="B2" class="OlkBtn" onclick="doAccept();"></td>
						<td>
							<hr color="#0D85C6" size="1"></td>
						<td style="width: 75px"><input type="button" value="<%=getselectQryGroupsLngStr("DtxtClose")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
					</tr>
				</table>
			</form>
</body>

</html><% conn.close
set rx = nothing %>