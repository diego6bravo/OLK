<!--#include file="top.asp" -->
<!--#include file="lang/adminInformer.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<head>
<script language="javascript" src="js_up_down.js"></script>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E1F3FD;
}
.style2 {
	font-weight: bold;
	background-color: #E1F3FD;
}
.style3 {
	text-align: center;
}
.style4 {
				font-weight: bold;
				background-color: #E1F3FD;
				direction: ltr;
}
</style>
</head>
<%

conn.execute("use [" & Session("OLKDB") & "]")

%>
<script type="text/javascript">
<!--
function valFrm()
{
	arrType = document.frmInformer.Type;
	for (var i = 0;i<arrType.length;i++)
	{
		if (arrType[i].value == 'U')
		{
			var myID = arrType[i].value + document.frmInformer.ID[i].value;
			var rowName = document.getElementById('RowName' + myID);
			if (rowName.value == '')
			{
				alert('<%=getadminInformerLngStr("LtxtFldNam")%>');
				rowName.focus();
				return false;
			}
		}
	}
	return true;
}
function doDel(id)
{
	document.frmAction.action = 'adminSubmit.asp';
	document.frmAction.ID.value = id;
	document.frmAction.submit();
}
function doEdit(id)
{
	document.frmAction.action = 'adminInformerEdit.asp';
	document.frmAction.ID.value = id;
	document.frmAction.submit();
}
//-->
</script>
<form name="frmAction" method="post" action="">
<input type="hidden" name="submitCmd" value="adminInformer">
<input type="hidden" name="cmd" value="del">
<input type="hidden" name="ID" value="">
</form>
<table border="0" cellpadding="0" width="100%" id="table3">
<form method="POST" action="adminSubmit.asp" name="frmInformer" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("LttlInformer")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminInformerLngStr("LttlInformerNote")%></font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" style="width: 16px; height: 21px;" class="style1"></td>
				<td align="center" class="style2" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("DtxtName")%></font></td>
				<td align="center" class="style2" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("DtxtOrder")%></font></td>
				<td align="center" class="style2" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("DtxtAlignment")%></font></td>
				<td align="center" class="style2" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("LtxtLinkedReport")%></font></td>
				<td align="center" class="style4" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("DtxtQuery")%></font></td>
				<td align="center" class="style2" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("LtxtHideNull")%></font></td>
				<td align="center" class="style2" style="height: 21px">
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerLngStr("DtxtActive")%></font></td>
				<td align="center" width="16" class="style1" style="height: 21px"></td>
			</tr>
			<% sql = "select T0.[Type], T0.[ID],  " & _  
					"Case T0.[Type] When 'S' Then T1.Description collate database_default When 'C' Then T3.name When 'A' Then T3.name When 'U' Then T0.[Name] End [Name],  " & _  
					"T0.Align, T0.Query, T0.rsIndex, T0.Active, T0.Ordr, T2.rsName, T0.HideNull  " & _  
					"from OLKInformer T0  " & _  
					"left outer join OLKCommon..OLKInformerDesc T1 on T0.[Type] = 'S' and T1.ID = T0.ID and T1.LanID = " & Session("LanID") & "  " & _  
					"left outer join OLKRS T2 on T2.rsIndex = T0.rsIndex  " & _  
					"left outer join OLKOps T3 on T3.ID = T0.ID and T0.[Type] in ('A', 'C')  " & _  
					"order by [Ordr]    " 
			rs.open sql, conn, 3, 1
			do while not rs.eof
			myID = rs("Type") & rs("ID") %>
			<input type="hidden" name="Type" value="<%=rs("Type")%>">
			<input type="hidden" name="ID" value="<%=rs("ID")%>">
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
			  	<% If rs("Type") = "U" Then %><a href="javascript:doEdit('<%=rs("ID")%>');">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a><% End If %></td>
			  <td valign="top">
			  	<% Select Case rs("Type")
			  	Case "S", "A", "C" %>
				<table border="0" cellpadding="0" id="table13" style="width: 100%;">
					<tr>
						<td width="0">
				<font size="1" face="Verdana" color="#4783C5"><%=RS("Name")%><%
			  	Select Case rs("Type")
			  		Case "A" %> - <%=getadminInformerLngStr("LtxtWaitForAut")%><%
			  		Case "C" %> - <%=getadminInformerLngStr("LtxtWaitForConf")%><%
			  	End Select %></font></td>
						<td width="16"><a href="javascript:doFldTrad('Informer', 'Type,ID', 'S,<%=rs("ID")%>', 'AlterName', 'T', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
				</table>
			  	<% Case "U" %>
				<table border="0" cellpadding="0" id="table13" style="width: 100%;">
					<tr>
						<td width="0">
						<input class="input" size="20" value="<%=Server.HTMLEncode(RS("Name"))%>" id="RowName" name="RowName<%=myID%>" maxlength="100" onkeydown="return chkMax(event, this, 100);" style="width: 100%;">
						</td>
						<td width="16"><a href="javascript:doFldTrad('Informer', 'Type,ID', 'U,<%=rs("ID")%>', 'AlterName', 'T', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
				</table>
				<% End Select %>
			  </td>
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder<%=myID%>" id="RowOrder<%=myID%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Ordr")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrder<%=myID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrder<%=myID%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table></td>
				<td valign="top" class="style3">
				<% If rs("Type") = "U" Then %>
				<select size="1" name="RowAlign<%=myID%>">
				<option></option>
				<option <% If rs("Align") = "L" Then %>selected<% End If %> value="L"><%=getadminInformerLngStr("DtxtLeft")%></option>
				<option <% If rs("Align") = "C" Then %>selected<% End If %> value="C"><%=getadminInformerLngStr("DtxtCenter")%></option>
				<option <% If rs("Align") = "R" Then %>selected<% End If %> value="R"><%=getadminInformerLngStr("DtxtRight")%></option>
				</select><% End If %></td>
				<td valign="top" class="style3">
				<font size="1" face="Verdana" color="#4783C5"><%=rs("rsName")%></font></td>
				<td valign="top" class="style3">
				<% If rs("Type") = "U" Then %>
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("Query"))%>"><% End If %></td>
				<td valign="top">
				<p align="center">
				<input type="checkbox" name="RowHideNull<%=myID%>"<% If rs("Type") = "S" or rs("Type") = "A" or rs("Type") = "W" Then %>disabled <% End If %> <% If rs("HideNull") = "Y" Then %>checked<% End If %> value="Y" class="noborder">&nbsp;</td>
				<td valign="top">
				<p align="center">
				<input type="checkbox" name="RowActive<%=myID%>" <% If rs("Active") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
				<td valign="top" width="16">
				<% If rs("Type") = "U" Then %><a href="javascript:if(confirm('<%=getadminInformerLngStr("LtxtConfRem")%>'.replace('{0}', '<%=Replace(Rs("Name"),"'","\'")%>')))doDel('<%=rs("ID")%>');">
				<img border="0" src="images/remove.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
			</tr>
			<script language="javascript">NumUDAttach('frmInformer', 'RowOrder<%=myID%>', 'btnRowOrder<%=myID%>Up', 'btnRowOrder<%=myID%>Down');</script>
			<% rs.movenext
			loop %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table22">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminInformerLngStr("DtxtSave")%>" name="btnUpdate" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminInformerLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="window.location.href='adminInformerEdit.asp'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminInformer">
	<input type="hidden" name="cmd" value="update">
	</form>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<!--#include file="bottom.asp" -->