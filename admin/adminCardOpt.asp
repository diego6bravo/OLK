<!--#include file="top.asp" -->
<!--#include file="lang/adminCardOpt.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<% conn.execute("use [" & Session("OLKDB") & "]")
varx = 0 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	font-weight: bold;
	background-color: #E1F3FD;
}
.style2 {
	background-color: #E1F3FD;
}
.style3 {
	background-color: #F3FBFE;
}
.style4 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<script language="javascript" src="js_up_down.js"></script>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<link rel="stylesheet" type="text/css" href="style_cal.css">
<script language="javascript">
var LtxtValQryVal = '<%=getadminCardOptLngStr("LtxtValQryVal")%>';
var LtxtValFldNam = '<%=getadminCardOptLngStr("LtxtValFldNam")%>';
var LtxtValFldNam2 = '<%=getadminCardOptLngStr("LtxtValFldNam2")%>';
var LtxtValQry = '<%=getadminCardOptLngStr("LtxtValQry")%>';
</script>
<script language="javascript" src="adminCardOpt.js"></script>
<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("edit") <> "Y" and Request("NewFld") <> "Y" Then %>
<%
sql = "select * from olkCardRep order by rowOrder, colIndex asc"
rs.open sql, conn, 3, 1 %>
<form method="POST" action="adminsubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCardOptLngStr("LttlCrdDet")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminCardOptLngStr("LttlCrdDetNote")%> </font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" class="style1" style="width: 16px">
				&nbsp;</td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminCardOptLngStr("DtxtName")%>&nbsp;</font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCardOptLngStr("DtxtOrder")%></font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C"><%=getadminCardOptLngStr("DtxtPosition2")%></font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCardOptLngStr("DtxtCodification")%></font></td>
				<td align="center" class="style1"><nobr>
				<font size="1" face="Verdana" color="#31659C"><%=getadminCardOptLngStr("DtxtField")%> / 
				<%=getadminCardOptLngStr("DtxtQuery")%>&nbsp;</font></nobr></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminCardOptLngStr("DtxtAccess")%>&nbsp;</font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminCardOptLngStr("DtxtShowAt")%>&nbsp;</font></td>
				<td align="center" width="16" class="style2">&nbsp;</td>
			</tr>
			<%
			If rs.recordcount > 0 then
			do While NOT RS.EOF 
		   	varx = varx + 1 %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
				<a href='adminCardOpt.asp?edit=Y&amp;rI=<%=rs("rowIndex")%>#table20'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input class="input" size="20" style="width: 100%; " value="<%=Server.HTMLEncode(RS("rowName"))%>" name="rowName<%=RS("rowIndex")%>" id="rowName" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td style="width: 16px"><a href="javascript:doFldTrad('CardRep', 'rowIndex', <%=rs("rowIndex")%>, 'alterRowName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminCardOptLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
			  </td>
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder<%=rs("rowIndex")%>" id="RowOrder<%=rs("rowIndex")%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("RowOrder")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrder<%=rs("rowIndex")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrder<%=rs("rowIndex")%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td valign="top">
					<select size="1" name="ColIndex<%=RS("rowIndex")%>" class="input">
					<option value="L" <% if rs("colindex") = "L" then %>selected<% end if %>>
					<%=getadminCardOptLngStr("DtxtLeft")%></option>
					<option value="R" <% if rs("colindex") = "R" then %>selected<% end if %>>
					<%=getadminCardOptLngStr("DtxtRight")%></option>
					</select></td>
				<td>
				<nobr>
				<select size="1" name="rowType<%=RS("rowIndex")%>" class="input">
				<option value="T"<% If rs("rowType") = "T" Then %> selected<% End If %>>
				<%=getadminCardOptLngStr("DtxtDisabled")%></option>
				<option value="L"<% If rs("rowType") = "L" Then %> selected<% End If %>>
				<%=getadminCardOptLngStr("DtxtLow")%></option>
				<option value="M"<% If rs("rowType") = "M" Then %> selected<% End If %>>
				<%=getadminCardOptLngStr("DtxtMedium")%></option>
				<option value="H"<% If rs("rowType") = "H" Then %> selected<% End If %>>
				<%=getadminCardOptLngStr("DtxtHigh")%></option>
				</select>
				<input type="checkbox" <% If rs("rowTypeRnd") = "Y" Then %>checked<% End If %> name="rowTypeRnd<%=rs("rowIndex")%>" id="rowTypeRnd<%=rs("rowIndex")%>" value="ON" class="noborder"><font face="Verdana" size="1" color="#31659C"><label for="rowTypeRnd<%=rs("rowIndex")%>"><%=getadminCardOptLngStr("DtxtRndLtr")%></label></font></nobr></td>
				<td valign="top" align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("rowField"))%>"></td>
				<td valign="top">
				<nobr>
				<select size="1" class="input" name="rowAccess<%=RS("rowIndex")%>">
				<option <% If Rs("rowAccess") = "T" Then %>selected<%end if %> value="T">
				<%=getadminCardOptLngStr("DtxtAll")%></option>
				<option <% If Rs("rowAccess") = "V" Then %>selected<%end if %> value="V">
				<%=getadminCardOptLngStr("DtxtAgents")%></option>
				<option <% If Rs("rowAccess") = "C" Then %>selected<%end if %> value="C">
				<%=getadminCardOptLngStr("DtxtClients")%></option>
				<option <% If Rs("rowAccess") = "D" Then %>selected<%end if %> value="D">
				<%=getadminCardOptLngStr("DtxtDisabled")%></option>
				</select>
				<select size="1" class="input" name="rowOP<%=RS("rowIndex")%>">
				<option <% If Rs("rowOP") = "O" Then %>selected<%end if %> value="O">
				<%=getadminCardOptLngStr("DtxtOLK")%></option>
				<option <% If Rs("rowOP") = "P" Then %>selected<%end if %> value="P">
				<%=getadminCardOptLngStr("DtxtPocket")%></option>
				<option <% If Rs("rowOP") = "T" Then %>selected<%end if %> value="T">
				<%=getadminCardOptLngStr("DtxtOLK")%>/<%=getadminCardOptLngStr("DtxtPocket")%></option>
				</select></nobr></td>
				<td valign="top">
				<select size="1" class="input" name="showAt<%=RS("rowIndex")%>">
				<option <% If Rs("showAt") = "A" Then %>selected<%end if %> value="A">
				<%=getadminCardOptLngStr("DtxtAll")%></option>
				<option <% If Rs("showAt") = "D" Then %>selected<%end if %> value="D">
				<%=getadminCardOptLngStr("DtxtDetail")%></option>
				<option <% If Rs("showAt") = "S" Then %>selected<%end if %> value="S">
				<%=getadminCardOptLngStr("DtxtSearch")%></option>
				</select></td>
				<td valign="middle" width="16">
						<a href="javascript:if(confirm('<%=getadminCardOptLngStr("LtxtConfDelFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("rowName")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&rI=<%=rs("rowIndex")%>&submitCmd=admincrdopt';">
						<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="rowIndex" value="<%=rs("rowIndex")%>">
				<script language="javascript">NumUDAttach('form1', 'RowOrder<%=rs("rowIndex")%>', 'btnRowOrder<%=rs("rowIndex")%>Up', 'btnRowOrder<%=rs("rowIndex")%>Down');</script>
				<% RS.MoveNext
				loop
				Else %>
				<tr>
					<td align="center" class="style1" colspan="8">
					<font size="1" face="Verdana" color="#31659C"><%=getadminCardOptLngStr("DtxtNoData")%></font></td>
				</tr>
				<% End If %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table22">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCardOptLngStr("DtxtSave")%>" <% If rs.recordcount = 0 then %>disabled<% End If %> name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminCardOptLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminCardOpt.asp?NewFld=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="admincrdopt">
<input type="hidden" name="cmd" value="u">
</form>
<% End If %>

<% If Request("edit") = "Y" or Request("NewFld") = "Y" Then %>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("rI") = "" Then %><%=getadminCardOptLngStr("LttlAddFldCrdDet")%><% Else %><%=getadminCardOptLngStr("LttlEditFldCrdDet")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminCardOptLngStr("LttlAddFldCrdDetNote")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("edit") = "Y" Then
			sql = "select rowName, rowAccess, rowField, rowType, rowOP, rowTypeRnd, rowTypeDec, rowOrder, colIndex, ShowAt, RowAlign " & _
			"from olkCardRep where rowIndex = " & Request("rI") 
			set rs = conn.execute(sql) 
			rowName = rs("rowName")
			rowAccess = rs("rowAccess")
			rowField = rs("rowField")
			rowType = rs("rowType")
			rowOP = rs("rowOP")
			rowTypeRnd = rs("rowTypeRnd")
			rowOrder = rs("rowOrder")
			rowTypeDec = rs("rowTypeDec")
			colIndex = rs("colIndex")
			showAt = rs("ShowAt")
			Align = rs("RowAlign")
		Else
			sql = "select IsNull(Max(rowOrder)+1, 0) from olkCardRep"
			set rs = conn.execute(sql)
			rowOrder = rs(0)
			rowTypeDec = "P"
			rowName = ""
			showAt = "A" %>
		<input type="hidden" name="rowNameTrad">
		<input type="hidden" name="customSqlDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td>
				<table border="0" cellpadding="0">
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCardOptLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td valign="top" class="style3" style="width: 260px">
						<p align="center">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input name="rowName" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(rowName)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('CardRep', 'rowIndex', '<%=Request("rI")%>', 'alterRowName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.rowNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCardOptLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCardOptLngStr("DtxtOrder")%></font></b></td>
						<td valign="top" class="style3">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td valign="top">
									<input type="text" name="RowOrder" id="RowOrder" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rowOrder%>">
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
						<td bgcolor="#E2F3FC">
						<font face="Verdana" size="1" color="#31659C"><strong><%=getadminCardOptLngStr("DtxtAlignment")%></strong></font></td>
						<td valign="top" class="style3">
						<select size="1" name="RowAlign">
						<option></option>
						<option <% If Align = "L" Then %>selected<% End If %> value="L"><%=getadminCardOptLngStr("DtxtLeft")%></option>
						<option <% If Align = "C" Then %>selected<% End If %> value="C"><%=getadminCardOptLngStr("DtxtCenter")%></option>
						<option <% If Align = "R" Then %>selected<% End If %> value="R"><%=getadminCardOptLngStr("DtxtRight")%></option>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC">
						<font face="Verdana" size="1" color="#31659C"><strong><%=getadminCardOptLngStr("DtxtPosition2")%></strong></font></td>
						<td valign="top" class="style3">
						<select size="1" name='ColIndex' class="input" style="width: 120; height: 16">
						<option value="L">
						<%=getadminCardOptLngStr("DtxtLeft")%></option>
						<option value="R" <% if colIndex = "R" then %>selected<% end if %>>
						<%=getadminCardOptLngStr("DtxtRight")%></option>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCardOptLngStr("DtxtCodification")%>&nbsp;</font></b></td>
						<td valign="top" class="style3">
						<font face="Verdana" size="1">
						<nobr>
						<select size="1" name="rowType" class="input" style="width: 120; height: 16">
						<option value="T"<% If rowType = "T" Then %> selected<% End If %>>
						<%=getadminCardOptLngStr("DtxtDisabled")%></option>
						<option value="L"<% If rowType = "L" Then %> selected<% End If %>>
						<%=getadminCardOptLngStr("DtxtLow")%></option>
						<option value="M"<% If rowType = "M" Then %> selected<% End If %>>
						<%=getadminCardOptLngStr("DtxtMedium")%></option>
						<option value="H"<% If rowType = "H" Then %> selected<% End If %>>
						<%=getadminCardOptLngStr("DtxtHigh")%></option>
						</select><input type="checkbox" name="rowTypeRnd" <% If rowTypeRnd = "Y" Then %>checked<% End If %> value="ON" id="rowTypeRnd" class="noborder"><font color="#31659C"><label for="rowTypeRnd"><%=getadminCardOptLngStr("DtxtRndLtr")%></label></font></nobr></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCardOptLngStr("DtxtDecimal")%><br>(<%=getadminCardOptLngStr("DtxtCodification")%>)&nbsp;</font></b></td>
						<td valign="top" class="style3">
						<select size="1" name="rowTypeDec" class="input">
						<option <% If rowTypeDec = "S" Then %>selected<% End If %> value="S"><%=getadminCardOptLngStr("DtxtDecSum")%></option>
						<option <% If rowTypeDec = "P" Then %>selected<% End If %> value="P"><%=getadminCardOptLngStr("DtxtDecPrice")%></option>
						<option <% If rowTypeDec = "R" Then %>selected<% End If %> value="R"><%=getadminCardOptLngStr("DtxtDecRate")%></option>
						<option <% If rowTypeDec = "Q" Then %>selected<% End If %> value="Q"><%=getadminCardOptLngStr("DtxtDecQty")%></option>
						<option <% If rowTypeDec = "%" Then %>selected<% End If %> value="%"><%=getadminCardOptLngStr("DtxtDecPercent")%></option>
						<option <% If rowTypeDec = "M" Then %>selected<% End If %> value="M"><%=getadminCardOptLngStr("DtxtDecMeasure")%></option>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCardOptLngStr("DtxtField")%>&nbsp;</font></b></td>
						<td valign="top" class="style3">
						<select <% If Request("Edit") = "Y" Then %>disabled<% End If %> size="1" name="rowField" class="input" onchange="javascript:document.form2.customSql.value=this.value;">
						<option></option>
						<% If Request("Edit") <> "Y" Then
						sql = "select name from syscolumns where id = object_id('OCRD')"
					   set rs = conn.execute(sql)
					   while not rs.eof %>
						<option value="<%=RS("Name")%>"><%=RS("Name")%></option>
						<% rs.movenext
						wend
						Else %>
						<option>--------</option>
						<% End If %>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCardOptLngStr("DtxtAccess")%>&nbsp;</font></b></td>
										<td valign="top" class="style3">
										<select size="1" name="rowAccess" class="input" style="width: 100">
						<option value="T" <% If rowAccess = "T" Then %>selected<% End IF %>>
						<%=getadminCardOptLngStr("DtxtAll")%></option>
						<option value="V" <% If rowAccess = "V" Then %>selected<% End IF %>>
						<%=getadminCardOptLngStr("DtxtAgents")%></option>
						<option value="C" <% If rowAccess = "C" Then %>selected<% End IF %>>
						<%=getadminCardOptLngStr("DtxtClients")%></option>
						<option value="D" <% If rowAccess = "D" Then %>selected<% End IF %>>
						<%=getadminCardOptLngStr("DtxtDisabled")%></option>
						</select><select size="1" class="input" name="rowOP" style="width: 100">
						<option <% If rowOP = "O" Then %>selected<%end if %> value="O">
						<%=getadminCardOptLngStr("DtxtOLK")%></option>
						<option <% If rowOP = "P" Then %>selected<%end if %> value="P">
						<%=getadminCardOptLngStr("DtxtPocket")%></option>
						<option <% If rowOP = "T" Then %>selected<%end if %> value="T">
						<%=getadminCardOptLngStr("DtxtOLK")%>/<%=getadminCardOptLngStr("DtxtPocket")%></option>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCardOptLngStr("DtxtShowAt")%>&nbsp;</font></b></td>
						<td><select size="1" class="input" name="showAt">
						<option <% If showAt = "A" Then %>selected<%end if %> value="A">
						<%=getadminCardOptLngStr("DtxtAll")%></option>
						<option <% If showAt = "D" Then %>selected<%end if %> value="D">
						<%=getadminCardOptLngStr("DtxtDetail")%></option>
						<option <% If showAt = "S" Then %>selected<%end if %> value="S">
						<%=getadminCardOptLngStr("DtxtSearch")%></option>
						</select></td>
					</tr>
				</table>
				<table border="0" cellpadding="0" width="100%" id="table20">
					<tr>
						<td valign="top">
						<table border="0" width="100%" id="table23" cellpadding="0">
							<tr>
								<td valign="top" colspan="2" bgcolor="#E2F3FC">
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="10" style="width: 100%" name="customSql" cols="100" class="input" onkeypress="javascript:document.form2.btnVerfyFilter.src='images/btnValidate.gif';document.form2.btnVerfyFilter.style.cursor = 'hand';;document.form2.valQuery.value='Y';"><%=myHTMLEncode(rowField)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCardOptLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(10, 'customSql', '<%=Request("rI")%>', <% If Request("rI") <> "" Then %>null<% Else %>document.form2.customSqlDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminCardOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top" colspan="2">
								<table cellpadding="0" style="width: 100%">
									<tr>
										<td valign="top" style="width: 119px" bgcolor="#E2F3FC" class="style4">
										<font size="1" face="Verdana">
										<strong><%=getadminCardOptLngStr("DtxtVariables")%></strong></font></td>
										<td class="style3"><font face="Verdana" size="1" color="#4783C5"><span dir="ltr">@CardCode</span> = <%=getadminCardOptLngStr("DtxtClientCode")%><br>
								<span dir="ltr">@SlpCode</span> = <%=getadminCardOptLngStr("DtxtAgentCode")%><br>
								<span dir="ltr">@dbName</span> = <%=getadminCardOptLngStr("DtxtDB")%><br>
								<span dir="ltr">@LanID</span> = <%=getadminCardOptLngStr("DtxtLanID")%></font></td>
									</tr>
									<tr>
										<td valign="top" style="width: 119px" bgcolor="#E2F3FC" class="style4">
										<font size="1" face="Verdana"><strong><%=getadminCardOptLngStr("DtxtFunctions")%></strong></font></td>
										<td class="style3"><% HideFunctionTitle = True
										functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top">
								<img src="images/spacer.gif"></td>
								<td>
								<img src="images/spacer.gif"></td>
							</tr>
							</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCardOptLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminCardOptLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminCardOptLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminCardOptLngStr("DtxtConfCancel")%>'))window.location.href='adminCardOpt.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="rI" value="<%=Request("rI")%>">
	<input type="hidden" name="submitCmd" value="admincrdopt">
	<input type="hidden" name="cmd" value="<% If Request("Edit") = "Y" Then %>e<% Else %>a<% End If %>">
	<% End If %>
	
</table>
<% If Request("edit") = "Y" or Request("NewFld") = "Y" Then %>
</form>
<script language="javascript">
NumUDAttach('form2', 'RowOrder', 'btnRowOrderUp', 'btnRowOrderDown');
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="crdOpt">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<form name="frmGetRSVars" action="getRSVars.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="rsIndex" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %><!--#include file="bottom.asp" -->