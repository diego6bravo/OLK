<!--#include file="top.asp" -->
<!--#include file="lang/adminCartMinRep.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<head>
<% 
varx = 0
conn.execute("use [" & Session("OLKDB") & "]")

If Request.Form.Count > 0 and Request.Form("btnUpdate") <> "" Then
	sql = "select RowType, LineIndex from OLKCMREP where RowActive <> 'D'"
	set rs = conn.execute(sql)
	sql = ""
	do while not rs.eof
		lIndex = rs("RowType") & rs("LineIndex")
		If Request("RowActive" & lIndex) = "Y" Then RowActive = "Y" Else RowActive = "N"
		If Request("ShowV" & lIndex) = "Y" Then ShowV = "Y" Else ShowV = "N"
		If Request("ShowC" & lIndex) = "Y" Then ShowC = "Y" Else ShowC = "N"
		If Request("PrintV" & lIndex) = "Y" Then PrintV = "Y" Else PrintV = "N"
		If Request("PrintC" & lIndex) = "Y" Then PrintC = "Y" Else PrintC = "N"
		If Request("RowAlign" & lIndex) <> "" Then Align = "'" & Request("RowAlign" & lIndex) & "'" Else Align = "NULL"
		If Request("RowMain") = lIndex Then Main = "Y" Else Main = "N"
		sql = sql & "update OLKCMREP set RowActive = '" & RowActive & "', RowName = N'" & saveHTMLDecode(Request("RowName" & lIndex), False) & "', " & _
		"Align = " & Align & ", ShowV = '" & ShowV & "', ShowC = '" & ShowC & "', PrintV = '" & PrintV & "', PrintC = '" & PrintC & "', RowOrder = " & Request("RowOrder" & lIndex) & ", Main = '" & Main & "' " & _
		"where RowType = '" & rs("RowType") & "' and LineIndex = " & rs("LineIndex") & " "
	rs.movenext
	loop
	If sql <> "" Then 
		conn.execute(sql)
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGenQry" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@Type") = "CMREP"
		cmd.execute()
	End If
	rs.close
ElseIf Request.Form("go") = "Y" Then
	Select Case Request("goAction")
		Case "mup"
			sql = "declare @RowOrder int set @RowOrder = " & Request("LineIndex") & " " & _
			"update OLKCMREP set RowOrder = -999999 where RowOrder = @RowOrder " & _
			"update OLKCMREP set RowOrder = RowOrder + 1 where RowOrder = @RowOrder - 1 " & _
			"update OLKCMREP set RowOrder = @RowOrder - 1 where RowOrder = -999999"
			conn.execute(sql)
		Case "mdown"
			sql = "declare @RowOrder int set @RowOrder = " & Request("LineIndex") & " " & _
			"update OLKCMREP set RowOrder = -999999 where RowOrder = @RowOrder " & _
			"update OLKCMREP set RowOrder = RowOrder -1 where RowOrder = @RowOrder + 1 " & _
			"update OLKCMREP set RowOrder = @RowOrder + 1 where RowOrder = -999999"
			conn.execute(sql)
		Case "del"
			sql = "update OLKCMREP set RowActive = 'D', RowOrder = -1 where RowType = 'U' and LineIndex = " & Request("LineIndex") & " " & _
			"update OLKCMREP set RowOrder = RowOrder -1 where RowType = 'U' and LineIndex > " & Request("LineIndex")
			conn.execute(sql)
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGenQry" & Session("ID")
			cmd.Parameters.Refresh
			cmd("@Type") = "CMREP"
			cmd.execute()
	End Select
End If
%>
<script language="javascript" src="js_up_down.js"></script>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
function goAction(RowType, LineIndex, goAction)
{
	if (goAction == 'editLine') frmGo.action = 'adminCartMinRepEdit.asp';
	document.frmGo.goAction.value = goAction;
	document.frmGo.RowType.value = RowType;
	document.frmGo.LineIndex.value = LineIndex;
	document.frmGo.submit();
}
</script>
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
</style>
</head>

<form method="post" action="adminCartMinRep.asp" name="frmGo">
<input type="hidden" name="go" value="Y">
<input type="hidden" name="goAction" value="">
<input type="hidden" name="RowType" value="">
<input type="hidden" name="LineIndex" value="">
</form>
<script type="text/javascript">
<!--
function valFrm()
{
	rowName = document.form1.RowName;
	if (rowName.length)
	{
		for (var i = 0;i<rowName.length;i++)
		{
			if (rowName[i].value == '')
			{
				alert('<%=getadminCartMinRepLngStr("LtxtValFldNam")%>');
				rowName[i].focus();
				return false;
			}
		}
	}
	else
	{
		if (rowName.value == '')
		{
			alert('<%=getadminCartMinRepLngStr("LtxtValFldNam")%>');
			rowName.focus();
			return false;
		}
	}
	return true;
}
//-->
</script>
<table border="0" cellpadding="0" width="100%" id="table3">
<form method="POST" action="adminCartMinRep.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("LttlPerCartMinRep")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminCartMinRepLngStr("LttlPerCartMinRepNote")%></font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" style="width: 16px; " class="style1"></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtName")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtOrder")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtOLK")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtSystem")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtAlignment")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtActive")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepLngStr("DtxtMain")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C">
				<%=myHTMLDecode(getadminCartMinRepLngStr("LtxtAgentsVisible"))%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C">
				<%=myHTMLDecode(getadminCartMinRepLngStr("LtxtAgentsPrinting"))%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C">
				<%=myHTMLDecode(getadminCartMinRepLngStr("LtxtClientsVisible"))%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C">
				<%=myHTMLDecode(getadminCartMinRepLngStr("LtxtClientsPrinting"))%></font></td>
				<td align="center" width="16" class="style1" style="height: 21px"></td>
			</tr>
			<% sql = "select RowType, LineIndex, RowName, RowQuery, SystemQuery, RowActive, RowOrder, Align, ShowV, ShowC, PrintV, PrintC, Main " & _
					"from OLKCMREP where RowActive <> 'D' order by RowOrder asc"
			rs.open sql, conn, 3, 1.
			do while not rs.eof
			myID = rs("RowType") & rs("LineIndex") %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
			  	<a href="javascript:goAction('<%=rs("RowType")%>', <%=rs("LineIndex")%>, 'editLine');">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top">
				<table border="0" cellpadding="0" id="table13" style="width: 100%;">
					<tr>
						<td width="0">
						<input class="input" size="20" value="<%=Server.HTMLEncode(RS("RowName"))%>" id="RowName" name="RowName<%=myID%>" onkeydown="return chkMax(event, this, 50);" style="width: 100%;">
						</td>
						<td width="16"><a href="javascript:doFldTrad('CMREP', 'RowType,LineIndex', '<%=rs("RowType")%>,<%=rs("LineIndex")%>', 'AlterRowName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminCartMinRepLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
			  </td>
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder<%=myID%>" id="RowOrder<%=myID%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("RowOrder")%>">
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
				<% If Not IsNull(RS("RowQuery")) Then %><img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("RowQuery"))%>"><% End If %></td>
				<td valign="top" class="style3">
				<% If Not IsNull(RS("SystemQuery")) Then %><img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("SystemQuery"))%>"><% End If %></td>
				<td valign="top" align="center">
				<select size="1" name="RowAlign<%=myID%>">
				<option></option>
				<option <% If rs("Align") = "L" Then %>selected<% End If %> value="L"><%=getadminCartMinRepLngStr("DtxtLeft")%></option>
				<option <% If rs("Align") = "C" Then %>selected<% End If %> value="C"><%=getadminCartMinRepLngStr("DtxtCenter")%></option>
				<option <% If rs("Align") = "R" Then %>selected<% End If %> value="R"><%=getadminCartMinRepLngStr("DtxtRight")%></option>
				</select></td>
				<td valign="top">
				<p align="center">
				<input type="checkbox" name="RowActive<%=myID%>" <% If rs("RowActive") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
				<td valign="top">
				<p align="center">
				<input type="radio" name="RowMain" id="RowMain<%=myID%>" <% If rs("Main") = "Y" Then %>checked<% End If %> value="<%=myID%>" class="noborder"></td>
				<td valign="top" align="center">
				<input type="checkbox" name="ShowV<%=myID%>" <% If rs("ShowV") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
				<td valign="top" align="center">
				<input type="checkbox" name="PrintV<%=myID%>" <% If rs("PrintV") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
				<td valign="top" align="center">
				<input type="checkbox" name="ShowC<%=myID%>" <% If rs("ShowC") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
				<td valign="top" align="center">
				<input type="checkbox" name="PrintC<%=myID%>" <% If rs("PrintC") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
				<td valign="top" width="16">
				<% If rs("RowType") = "U" Then %><a href="javascript:if(confirm('<%=getadminCartMinRepLngStr("LtxtConfRemFld")%>'.replace('{0}', '<%=Replace(Rs("RowName"),"'","\'")%>')))goAction('<%=rs("RowType")%>', <%=rs("LineIndex")%>, 'del');">
				<img border="0" src="images/remove.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
			</tr>
			<script language="javascript">NumUDAttach('form1', 'RowOrder<%=myID%>', 'btnRowOrder<%=myID%>Up', 'btnRowOrder<%=myID%>Down');</script>
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
				<input type="submit" value="<%=getadminCartMinRepLngStr("DtxtSave")%>" name="btnUpdate" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminCartMinRepLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="window.location.href='adminCartMinRepEdit.asp'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	</form>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<!--#include file="bottom.asp" -->