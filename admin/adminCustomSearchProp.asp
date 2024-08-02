<!--#include file="chkLogin.asp" -->
<!--#include file="lang/adminCustomSearchProp.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getadminCustomSearchPropLngStr("LtxtCustomSearchProp")%></title>
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<script language="javascript" src="general.js"></script>
<script language="javascript" src="js_up_down.js"></script>
</head>

<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")

ID = CInt(Request("ID"))

Select Case CInt(Request("ObjectCode"))
	Case 2
		sql = 	"select T0.GroupCode PropID, T0.GroupName Descr, T1.Active, IsNull(T1.Ordr, T0.GroupCode) Ordr " & _  
				"from OCQG T0 " & _  
				"left outer join OLKCustomSearchProp T1 on T1.ObjectCode = 2 and T1.ID = " & ID & " and T1.PropID = T0.GroupCode " & _
				"order by T1.Ordr"
	Case 4
		sql =	"select T0.ItmsTypCod PropID, T0.ItmsGrpNam Descr, T1.Active, IsNull(T1.Ordr, T0.ItmsTypCod) Ordr " & _  
				"from OITG T0 " & _  
				"left outer join OLKCustomSearchProp T1 on T1.ObjectCode = 4 and T1.ID = " & ID & " and T1.PropID = T0.ItmsTypCod "  & _
				"order by T1.Ordr"
End Select

rs.open sql, conn, 3, 1
%>
<body topmargin="0" leftmargin="0" onbeforeunload="javascript:opener.clearWin();" onload="setTblSet();" onscroll="setTblSet();">

<form name="frm" method="post" action="adminSubmit.asp">
<table style="width: 100%">
				<tr>
								<td class="popupTtl"><%=getadminCustomSearchPropLngStr("LtxtCustomSearchProp")%></td>
				</tr>
				<tr>
								<td>
								<table style="width: 100%" cellpadding="0" cellspacing="0">
												<tr class="popupTtl">
																<td></td>
																<td align="center"><%=getadminCustomSearchPropLngStr("DtxtProp")%></td>
																<td align="center"><%=getadminCustomSearchPropLngStr("DtxtOrder")%></td>
												</tr>
												<% do while not rs.eof
												PropID = rs("PropID") %>
												<tr class="popupOptValue">
																<td style="width: 20px;"><input type="checkbox" class="noborder" name="qryGroup<%=PropID%>" id="qryGroup<%=PropID%>" onclick="doCheck();" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %>></td>
																<td><label for="qryGroup<%=PropID%>"><%=rs("Descr")%></label></td>
																<td align="center">
																<table cellpadding="0" cellspacing="0" border="0">
																	<tr>
																		<td><input id="Ordr<%=PropID%>" name="Ordr<%=PropID%>" class="input" value="<%=rs("Ordr")%>" size="7" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);"></td>
																		<td valign="middle">
																		<table cellpadding="0" cellspacing="0" border="0">
																			<tr>
																				<td><img src="images/img_nud_up.gif" id="btnQryOrder<%=PropID%>Up"></td>
																			</tr>
																			<tr>
																				<td><img src="images/spacer.gif"></td>
																			</tr>
																			<tr>
																				<td><img src="images/img_nud_down.gif" id="btnQryOrder<%=PropID%>Down"></td>
																			</tr>
																		</table>
																		</td>
																	</tr>
																</table>
																<script language="javascript">NumUDAttach('frm', 'Ordr<%=PropID%>', 'btnQryOrder<%=PropID%>Up', 'btnQryOrder<%=PropID%>Down');</script>
																</td>
												</tr>
												<% rs.movenext
												loop
												rs.Filter = "Active = 'Y'" %>
												<tr class="popupOptValue">
																<td><input type="checkbox" name="chkAll" id="chkAll" class="noborder" value="Y" onclick="doCheckAll(this.checked);" <% If rs.recordcount = 64 Then %>checked<% End If %>></td>
																<td colspan="2"><label for="chkAll"><%=getadminCustomSearchPropLngStr("DtxtAll")%></label></td>
												</tr>
								</table>
								</td>
				</tr>
				<tr height="27">
					<td>&nbsp;</td>
				</tr>
</table>
<input type="hidden" name="submitCmd" value="adminCustomSearch">
<input type="hidden" name="cmd" value="Prop">
<input type="hidden" name="ObjID" value="<%=Request("ObjectCode")%>">
<input type="hidden" name="ID" value="<%=Request("ID")%>">
				<table cellpadding="0" border="0" width="100%" style="position: absolute;" id="tblSave" bgcolor="#FFFFFF">
					<tr>
						<td style="width: 75px"><input type="submit" value="<%=getadminCustomSearchPropLngStr("DtxtAccept")%>" name="B2" class="OlkBtn"></td>
						<td>
							<hr color="#0D85C6" size="1"></td>
						<td style="width: 75px"><input type="button" value="<%=getadminCustomSearchPropLngStr("DtxtCancel")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
					</tr>
				</table>
</form>
<script type="text/javascript">
function doCheckAll(check)
{
	for (var i = 1;i<=64;i++)
	{
		document.getElementById('qryGroup' + i).checked = check;
	}
}
function doCheck()
{
	var check = true;
	for (var i = 1;i<=64;i++)
	{
		if (!document.getElementById('qryGroup' + i).checked)
		{
			check = false;
			break;
		}
	}
	document.getElementById('chkAll').checked = check;
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
</body>

</html>
