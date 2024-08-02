<!--#include file="top.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<!--#include file="lang/adminInformerEdit.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #F3FBFE;
}
.style3 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>

</head>
	<form name="frmEditMonitor" action="adminSubmit.asp" method="post" onsubmit="return valFrm();">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("ID") <> "" Then %><%=getadminInformerEditLngStr("LttlEditMonitor")%><% Else %><%=getadminInformerEditLngStr("LttlAddMonitor")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminInformerEditLngStr("LttlAddEditMonitor")%></font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
		<% RowActive = ""
		conn.execute("use [" & Session("OLKDB") & "]")
		If Request("ID") <> "" Then
			sql = "select * from OLKInformer where Type = 'U' and ID = " & Request("ID")
			set rs = conn.execute(sql) 
			RowName = rs("Name")
			RowActive = rs("Active")
			HideNull = rs("HideNull")
			RowQuery = rs("Query")
			rowOrder = rs("Ordr")
			rsIndex = rs("rsIndex")
			Align = rs("Align")
			If IsNull(rsIndex) Then rsIndex = ""
		Else
			RowActive = "Y"
			HideNull = "N"
			rowName = ""
			sql = "select IsNull(Max(Ordr)+1, 0) from OLKInformer"
			set rs = conn.execute(sql)
			rowOrder = rs(0) %>
		<input type="hidden" name="rowNameTrad">
		<input type="hidden" name="RowQueryDef">
		<% End If %>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b><font face="Verdana" size="1" color="#31659C"><%=getadminInformerEditLngStr("DtxtName")%></font></b></td>
				<td valign="top" class="style1">
				<table cellpadding="0" cellspacing="0" border="0" width="300">
					<tr>
						<td>
						<input name="rowName" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(RowName)%>" size="20" maxlength="100" onkeydown="return chkMax(event, this, 100);">
						</td>
						<td width="16"><a href="javascript:doFldTrad('Informer', 'ID', '<%=Request("ID")%>', 'AlterName', 'T', <% If Request("ID") = "" Then %>document.frmEditMonitor.rowNameTrad<% Else %>null<% End If %>);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
				</table></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminInformerEditLngStr("DtxtOrder")%></font></b></td>
				<td valign="top" class="style1">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder" id="RowOrder" size="7" style="font-size: 10px; font-family: Verdana; color: #3F7B96; font-weight: bold; border: 1px solid #68A6C0; background-color: #D9F0FD; text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rowOrder%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrderDown"></td>
							</tr>
						</table></td>
					</tr>
				</table></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b>
				<font face="Verdana" size="1" color="#31659C">
				<%=getadminInformerEditLngStr("DtxtAlignment")%></font></b></td>
				<td valign="top" class="style1">
				<select size="1" name="RowAlign">
				<option></option>
				<option <% If Align = "L" Then %>selected<% End If %> value="L"><%=getadminInformerEditLngStr("DtxtLeft")%></option>
				<option <% If Align = "C" Then %>selected<% End If %> value="C"><%=getadminInformerEditLngStr("DtxtCenter")%></option>
				<option <% If Align = "R" Then %>selected<% End If %> value="R"><%=getadminInformerEditLngStr("DtxtRight")%></option>
				</select></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b>
				<font face="Verdana" size="1" color="#31659C">
				<%=getadminInformerEditLngStr("DtxtReport")%></font></b></td>
				<td valign="top" class="style1">
				<select size="1" name="RowReport">
				<option></option>
				<% 
				LastRG = ""
				sql = "select T1.rgName, T0.rsIndex, T0.rsName " & _
						"from OLKRS T0 " & _
						"inner join OLKRG T1 on T1.rgIndex = T0.rgIndex " & _
						"where T1.UserType = 'V' and T0.Active = 'Y' " & _
						"order by T1.rgName, T0.rsName "
				set rs = conn.execute(sql)
				do while not rs.eof
				If LastRG <> rs("rgName") Then
					If LastRG <> "" Then Response.Write "</optgroup>"
					Response.WRite "<optgroup label=""" & myHTMLEncode(rs("rgName")) & """>"
					LastRG = rs("rgName")
				End If %>
				<option <% If CStr(rsIndex) = CStr(rs("rsIndex")) Then %>selected<% End If %> value="<%=rs("rsIndex")%>"><%=myHTMLEncode(rs("rsName"))%></option>
				<% rs.movenext
				loop
				Response.Write "</optgroup>" %>
				</select></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				&nbsp;</td>
				<td valign="top" class="style1">
				<font face="Verdana" size="1" color="#31659C">
				<input type="checkbox" id="RowHideNull" name="RowHideNull" <% If HideNull = "Y" Then %>checked<% End If %> value="Y" class="noborder"><label for="RowHideNull"><%=getadminInformerEditLngStr("LtxtHideNull")%></label></font></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				&nbsp;</td>
				<td valign="top" class="style1">
				<font face="Verdana" size="1" color="#31659C">
				<input type="checkbox" id="RowActive" name="RowActive" <% If RowActive = "Y" Then %>checked<% End If %> value="Y" class="noborder"><label for="RowActive"><%=getadminInformerEditLngStr("DtxtActive")%></label></font></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b>
				<font face="Verdana" size="1" color="#31659C">
				<%=getadminInformerEditLngStr("DtxtQuery")%></font></b></td>
				<td valign="top" class="style1">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td rowspan="2">
									<textarea cols="78" dir="ltr" name="RowQuery" class="input" style="width: 100%; " rows="6" onkeypress="javascript:document.frmEditMonitor.btnVerfy.src='images/btnValidate.gif';document.frmEditMonitor.btnVerfy.style.cursor = 'hand';document.frmEditMonitor.valRowQuery.value='Y';"><%=myHTMLEncode(RowQuery)%></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminInformerEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(21, 'Query', '<%=Request("ID")%>', <% If Request("ID") <> "" Then %>null<% Else %>document.frmEditMonitor.RowQueryDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnVerfy" alt="<%=getadminInformerEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditMonitor.valRowQuery.value == 'Y')VerfyQuery();">
									<input type="hidden" name="valRowQuery" value="N">
								</td>
							</tr>
						</table>
				</td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b>
				<font face="Verdana" size="1" color="#31659C">
				<%=getadminInformerEditLngStr("DtxtVariables")%></font></b></td>
				<td valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5"><span dir="ltr">@SlpCode</span> = 
						<%=getadminInformerEditLngStr("DtxtAgent")%><br>
						<span dir="ltr">@LanID</span> = <%=getadminInformerEditLngStr("DtxtLanID")%></font></td>
			</tr>
			<tr>
				<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
				<b>
				<font face="Verdana" size="1" color="#31659C">
				<%=getadminInformerEditLngStr("DtxtFunctions")%></font></b></td>
				<td valign="top" class="style1">
						<% HideFunctionTitle = True
						functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminInformerEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminInformerEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminInformerEditLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('|L:txtValCanMon|'))window.location.href='adminInformer.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>

	<input type="hidden" name="submitCmd" value="adminInformer">
	<input type="hidden" name="cmd" value="save">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	<tr>
		<td height="15"></td>
	</tr>
</table>
	</form>
<script language="javascript" src="js_up_down.js"></script>
<script language="javascript">
function valFrm()
{
	if (document.frmEditMonitor.rowName.value == '') {
		alert('<%=getadminInformerEditLngStr("LtxtValFldNam")%>');
		document.frmEditMonitor.rowName.focus();
		return false; }
	else if (document.frmEditMonitor.RowQuery.value == '') {
		alert('<%=getadminInformerEditLngStr("LtxtValQry")%>');
		document.frmEditMonitor.RowQuery.focus();
		return false; }
	else if (document.frmEditMonitor.valRowQuery.value == 'Y') {
		alert('<%=getadminInformerEditLngStr("LtxtValQryVal")%>');
		document.frmEditMonitor.btnVerfy.focus();
		return false; } 
	return true;
}


NumUDAttach('frmEditMonitor', 'RowOrder', 'btnRowOrderUp', 'btnRowOrderDown');
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frmEditMonitor.RowQuery.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	document.frmEditMonitor.btnVerfy.src='images/btnValidateDis.gif'
	document.frmEditMonitor.btnVerfy.cursor = '';
	document.frmEditMonitor.valRowQuery.value='N';
}
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="TaskMon">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>

<!--#include file="bottom.asp" -->