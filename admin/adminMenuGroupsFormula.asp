<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<!--#include file="lang/adminMenuGroupsFormula.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getadminMenuGroupsFormulaLngStr("LtxtFormula")%></title>
<link rel="stylesheet" href="style/style_pop.css">
<style>
<!--
input        { font-family: Verdana; font-size: 10px; border: 1px solid #73B9B9; background-color: #D8EDEE }
.noborder {
	border-style: solid;
	border-width: 0;
	background: background-image;
}

-->
</style>
</head>
<body bgcolor="#F9FDFF" marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onbeforeunload="javascript:opener.clearWin();">

<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">
		&nbsp;<%=getadminMenuGroupsFormulaLngStr("LtxtFormula")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD" align="center">
		<font size="1" face="Verdana" color="#3366CC"><%=getadminMenuGroupsFormulaLngStr("DtxtType")%>: <% Select Case Request("Type")
			Case "F" %><%=getadminMenuGroupsFormulaLngStr("DtxtField")%>
			<% Case "D" %><%=getadminMenuGroupsFormulaLngStr("DtxtDescription")%>
			<% End Select %></font></td>
	</tr>
	<tr>
		<td bordercolor="#C8E7E8" align="center">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td rowspan="2">
							<textarea name="txtQry" id="txtQry" dir="ltr" rows="5"  style="border:1px solid #68A6C0; color: #3F7B96; font-family: Verdana; font-size: 10px; width: 100%; padding: 0; background-color: #D9F0FD;" onkeydown="javscript:document.getElementById('btnVerfy').src='images/btnValidate.gif';document.getElementById('btnVerfy').style.cursor = 'hand';;document.getElementById('valFormula').value='Y';"><% If Request("Q") <> "" Then %><%=Server.HTMLEncode(Request("Q"))%><% End If %></textarea>
						</td>
						<td valign="top" width="1">
							&nbsp;</td>
					</tr>
					<tr>
						<td valign="bottom" width="1">
							<img src="images/btnValidateDis.gif" id="btnVerfy" alt="<%=getadminMenuGroupsFormulaLngStr("DtxtValidate")%>" onclick="javascript:if (document.getElementById('valFormula').value == 'Y')VerfyFilter();">
							<input type="hidden" name="valFormula" id="valFormula" value="N">
						</td>
					</tr>
				</table>

		</td>
	</tr>
	<tr>
		<td bordercolor="#C8E7E8" align="center">
		<input type="button" value="<%=getadminMenuGroupsFormulaLngStr("DtxtAccept")%>" name="btnAccept" class="OlkBtn" onclick="setValue();"></td>
	</tr>
</table>
<script type="text/javascript">
<!--
function setValue()
{
	if (document.getElementById('valFormula').value == 'N')
	{
		opener.setFormulaVal(document.getElementById('txtQry').value);
		window.close();
	}
	else
	{
		alert('<%=getadminMenuGroupsFormulaLngStr("LtxtVal")%>');
	}
}
function VerfyFilter()
{
	document.frmVerfyQuery.Query.value = document.getElementById('txtQry').value;
	document.frmVerfyQuery.submit();
}
function VerfyQueryVerified()
{
	document.getElementById('btnVerfy').style.cursor = '';
	document.getElementById('btnVerfy').src = 'images/btnValidateDis.gif';
	document.getElementById('valFormula').value = 'N';
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="MenuGroupFormula">
	<input type="hidden" name="subType" value="<%=Request("Type")%>">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
	<input type="hidden" name="TableID" value="<%=Request("TableID")%>">
</form>
</body>

</html>