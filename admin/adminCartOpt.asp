<!--#include file="top.asp" -->
<!--#include file="lang/adminCartOpt.asp" -->
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
	font-weight: normal;
	background-color: #E2F3FC;
	color: #31659C;
}
.style4 {
	background-color: #F3FBFE;
}
.style5 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style6 {
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
var LtxtNoRSVars = '<%=getadminCartOptLngStr("LtxtNoRSVars")%>';
var LtxtValQryVal = '<%=getadminCartOptLngStr("LtxtValQryVal")%>';
var LtxtValFldNam = '<%=getadminCartOptLngStr("LtxtValFldNam")%>';
var LtxtValFldNam2 = '<%=getadminCartOptLngStr("LtxtValFldNam2")%>';
var LtxtValQry = '<%=getadminCartOptLngStr("LtxtValQry")%>';
var LtxtSelFldVar = '<%=getadminCartOptLngStr("LtxtSelFldVar")%>';
var LtxtValVarVal = '<%=getadminCartOptLngStr("LtxtValVarVal")%>';
var LtxtValVarQry = '<%=getadminCartOptLngStr("LtxtValVarQry")%>';
var LtxtValVarQryBrnVerfy = '<%=getadminCartOptLngStr("LtxtValVarQryBrnVerfy")%>';
var DtxtValNumVal = '|D:txtValNumVal|';
var DtxtVariable = '<%=getadminCartOptLngStr("DtxtVariable")%>';
var DtxtValue = '<%=getadminCartOptLngStr("DtxtValue")%>';
var DtxtQuery = '<%=getadminCartOptLngStr("DtxtQuery")%>';
var DtxtItemCode = '<%=Replace(getadminCartOptLngStr("DtxtItemCode"), "'", "\'")%>';
var DtxtPList = '<%=getadminCartOptLngStr("DtxtPList")%>';
var DtxtClientCode = '<%=getadminCartOptLngStr("DtxtClientCode")%>';
var DtxtAgentCode = '<%=Replace(getadminCartOptLngStr("DtxtAgentCode"), "'", "\'")%>';
var DtxtDB = '<%=getadminCartOptLngStr("DtxtDB")%>';
var DtxtWhsCode = '<%=getadminCartOptLngStr("DtxtWhsCode")%>';
var DtxtQty = '<%=getadminCartOptLngStr("LtxtQtyInUnit")%>';
var DtxtPrice = '<%=getadminCartOptLngStr("DtxtPrice")%>';
var DtxtUnit = '<%=getadminCartOptLngStr("DtxtUnit")%>';
var LtxtValRepLnk = '<%=getadminCartOptLngStr("LtxtValRepLnk")%>';
var DtxtValidate = '<%=getadminCartOptLngStr("DtxtValidate")%>';
var DtxtLogNum = '<%=getadminCartOptLngStr("DtxtLogNum")%>';
var DtxtLineNum = '<%=getadminCartOptLngStr("DtxtLine")%>';
var CalendarFormat = '<%=GetCalendarFormatString%>';
</script>
<script language="javascript" src="adminCartOpt.js"></script>
<table border="0" cellpadding="0" width="100%">
	<% If Request("edit") <> "Y" and Request("NewFld") <> "Y" Then
sql = "select * from olkCartRep order by [Order] asc"
rs.open sql, conn, 3, 1 %>
<form method="POST" action="adminsubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCartOptLngStr("Lttl")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminCartOptLngStr("LttlNote")%> </font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td align="center" class="style1" style="width: 16px">
				&nbsp;</td>
				<td align="center" class="style1" style="width: 200px">
				<font size="1" face="Verdana" color="#31659C"><%=getadminCartOptLngStr("DtxtName")%>&nbsp;</font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCartOptLngStr("DtxtOrder")%></font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCartOptLngStr("DtxtCodification")%></font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminCartOptLngStr("DtxtField")%>&nbsp;/ 
				<%=getadminCartOptLngStr("DtxtQuery")%></font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminCartOptLngStr("DtxtAccess")%>&nbsp;</font></td>
				<td align="center" width="16" class="style2">&nbsp;</td>
			</tr>
			<%
			If rs.recordcount > 0 then
			do While NOT RS.EOF 
			ID = rs("ID")
		   	varx = varx + 1 %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
				<a href='adminCartOpt.asp?edit=Y&amp;rI=<%=ID%>#table20'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top" style="width: 200px">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input class="input" style="width: 100%; " size="20" value="<%=Server.HTMLEncode(rs("Name"))%>" name="Name<%=ID%>" id="Name" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td style="width: 16px"><a href="javascript:doFldTrad('CartRep', 'ID', <%=ID%>, 'alterName', 'T', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
				</table>
			  </td>
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="Order<%=ID%>" id="Order<%=ID%>" size="7" style="text-align:right" class="input"onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Order")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnOrder<%=ID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnOrder<%=ID%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td>
				<nobr>
				<select size="1" name="Type<%=ID%>" class="input">
				<option value="T"<% If rs("Type") = "T" Then %> selected<% End If %>>
				<%=getadminCartOptLngStr("DtxtDisabled")%></option>
				<option value="L"<% If rs("Type") = "L" Then %> selected<% End If %>>
				<%=getadminCartOptLngStr("DtxtLow")%></option>
				<option value="M"<% If rs("Type") = "M" Then %> selected<% End If %>>
				<%=getadminCartOptLngStr("DtxtMedium")%></option>
				<option value="H"<% If rs("Type") = "H" Then %> selected<% End If %>>
				<%=getadminCartOptLngStr("DtxtHigh")%></option>
				</select>
				<input type="checkbox" <% If rs("TypeRnd") = "Y" Then %>checked<% End If %> name="TypeRnd<%=ID%>" id="TypeRnd<%=ID%>" value="ON" class="noborder"><font color="#31659C" face="Verdana" size="1"><label for="TypeRnd<%=ID%>"><%=getadminCartOptLngStr("DtxtRndLtr")%></label></font></nobr></td>
				<td valign="top" align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(rs("Field"))%>"></td>
				<td valign="top">
				<nobr>
				<select size="1" class="input" name="Access<%=ID%>" style="width: 100">
				<option <% If rs("Access") = "T" Then %>selected<%end if %> value="T">
				<%=getadminCartOptLngStr("DtxtAll")%></option>
				<option <% If rs("Access") = "V" Then %>selected<%end if %> value="V">
				<%=getadminCartOptLngStr("DtxtAgents")%></option>
				<option <% If rs("Access") = "C" Then %>selected<%end if %> value="C">
				<%=getadminCartOptLngStr("DtxtClients")%></option>
				<option <% If rs("Access") = "D" Then %>selected<%end if %> value="D">
				<%=getadminCartOptLngStr("DtxtDisabled")%></option>
				</select>
				<select size="1" class="input" name="OP<%=ID%>" style="width: 100">
				<option <% If rs("OP") = "O" Then %>selected<%end if %> value="O">
				<%=getadminCartOptLngStr("DtxtOLK")%></option>
				<option <% If rs("OP") = "P" Then %>selected<%end if %> value="P">
				<%=getadminCartOptLngStr("DtxtPocket")%></option>
				<option <% If rs("OP") = "T" Then %>selected<%end if %> value="T">
				<%=getadminCartOptLngStr("DtxtOLK")%>/<%=getadminCartOptLngStr("DtxtPocket")%></option>
				</select></nobr></td>
				<td valign="middle" width="16">
						<a href="javascript:if(confirm('<%=getadminCartOptLngStr("LtxtConfDelFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(rs("Name")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&rI=<%=ID%>&submitCmd=adminCartOpt';">
						<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="ID" value="<%=ID%>">
				<script language="javascript">NumUDAttach('form1', 'Order<%=ID%>', 'btnOrder<%=ID%>Up', 'btnOrder<%=ID%>Down');</script>
				<% RS.MoveNext
				loop
				Else %>
				<tr>
					<td align="center" class="style1" colspan="7">
					<font size="1" face="Verdana" color="#31659C"><%=getadminCartOptLngStr("DtxtNoData")%></font></td>
				</tr>
				<% End If %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCartOptLngStr("DtxtSave")%>" <% If rs.recordcount = 0 then %>disabled<% End If %> name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminCartOptLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminCartOpt.asp?NewFld=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminCartOpt">
	<input type="hidden" name="cmd" value="u">
</form>
<% End If %>

<% If Request("edit") = "Y" or Request("NewFld") = "Y" Then %>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("rI") = "" Then %><%=getadminCartOptLngStr("LttlAddFldDet")%><% Else %><%=getadminCartOptLngStr("LttlEditFldDet")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminCartOptLngStr("LttlAddFldDetNote")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("edit") = "Y" Then
			sql = "select Name, Access, Field, Type, OP, TypeRnd, TypeDec, [Order], linkActive, IsNull(linkObject, -1) linkObject, Align, [Dynamic] " & _
			"from olkCartRep where ID = " & Request("rI") 
			set rs = conn.execute(sql) 
			Name = rs("Name")
			Access = rs("Access")
			Field = rs("Field")
			myType = rs("Type")
			OP = rs("OP")
			TypeRnd = rs("TypeRnd")
			Order = rs("Order")
			linkActive = rs("linkActive")
			linkObject = rs("linkObject")
			TypeDec = rs("TypeDec")
			Align = rs("Align")
			Dynamic = rs("Dynamic")
		Else
			sql = "select IsNull(Max([Order])+1, 0) from olkCartRep"
			set rs = conn.execute(sql)
			Order = rs(0)
			TypeDec = "P"
			Align = ""
			Name = "" %>
		<input type="hidden" name="NameTrad">
		<input type="hidden" name="customSqlDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td>
				<table border="0" cellpadding="0">
					<tr>
						<td bgcolor="#E2F3FC" style="width: 200px"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCartOptLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td valign="top" style="width: 200px" class="style4">
						<p align="center">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input style="width: 100%; " name="Name" class="input" value="<%=Server.HTMLEncode(Name)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('CartRep', 'ID', '<%=Request("rI")%>', 'alterName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.NameTrad<% End If %>);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCartOptLngStr("DtxtOrder")%></font></b></td>
						<td valign="top" class="style4">
						<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td valign="top">
								<input type="text" name="Order" id="Order" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=Order%>">
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0" style="width: 100%">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnOrderUp"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnOrderDown"></td>
								</tr>
							</table></td>
						</tr>
					</table></td></tr>
					<tr>
						<td bgcolor="#E2F3FC" style="height: 30px"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminCartOptLngStr("DtxtCodification")%>&nbsp;</font></b></td>
						<td valign="top" class="style4" style="height: 30px">
						<select size="1" name="Type" class="input" style="width: 120; height: 16">
						<option value="T"<% If myType = "T" Then %> selected<% End If %>>
						<%=getadminCartOptLngStr("DtxtDisabled")%></option>
						<option value="L"<% If myType = "L" Then %> selected<% End If %>>
						<%=getadminCartOptLngStr("DtxtLow")%></option>
						<option value="M"<% If myType = "M" Then %> selected<% End If %>>
						<%=getadminCartOptLngStr("DtxtMedium")%></option>
						<option value="H"<% If myType = "H" Then %> selected<% End If %>>
						<%=getadminCartOptLngStr("DtxtHigh")%></option>
						</select><input type="checkbox" name="TypeRnd" <% If TypeRnd = "Y" Then %>checked<% End If %> value="ON" id="TypeRnd" class="noborder"><font face="Verdana" size="1" color="#31659C"><label for="TypeRnd"><%=getadminCartOptLngStr("DtxtRndLtr")%></label></td></tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCartOptLngStr("DtxtDecimal")%> (<%=getadminCartOptLngStr("DtxtCodification")%>)&nbsp;</font></b></td>
						<td valign="top" class="style4">
						<select size="1" name="TypeDec" class="input">
						<option <% If TypeDec = "S" Then %>selected<% End If %> value="S"><%=getadminCartOptLngStr("DtxtDecSum")%></option>
						<option <% If TypeDec = "P" Then %>selected<% End If %> value="P"><%=getadminCartOptLngStr("DtxtDecPrice")%></option>
						<option <% If TypeDec = "R" Then %>selected<% End If %> value="R"><%=getadminCartOptLngStr("DtxtDecRate")%></option>
						<option <% If TypeDec = "Q" Then %>selected<% End If %> value="Q"><%=getadminCartOptLngStr("DtxtDecQty")%></option>
						<option <% If TypeDec = "%" Then %>selected<% End If %> value="%"><%=getadminCartOptLngStr("DtxtDecPercent")%></option>
						<option <% If TypeDec = "M" Then %>selected<% End If %> value="M"><%=getadminCartOptLngStr("DtxtDecMeasure")%></option>
						</select></td></tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCartOptLngStr("DtxtField")%>&nbsp;</font></b></td>
						<td valign="top" class="style4">
						<select <% If Request("Edit") = "Y" Then %>disabled<% End If %> size="1" name="Field" class="input" onchange="javascript:document.form2.customSql.value=this.value;">
						<option></option>
						<% If Request("Edit") <> "Y" Then
						sql = "select name from syscolumns where id = object_id('OITM')"
					   set rs = conn.execute(sql)
					   while not rs.eof %>
						<option value="<%=RS("Name")%>"><%=RS("Name")%></option>
						<% rs.movenext
						wend
						Else %>
						<option>--------</option>
						<% End If %>
						</select></td></tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCartOptLngStr("DtxtAccess")%>&nbsp;</font></b></td>
										<td valign="top" class="style4">
										<nobr><select size="1" name="Access" class="input" style="width: 100px;">
						<option value="T" <% If Access = "T" Then %>selected<% End IF %>>
						<%=getadminCartOptLngStr("DtxtAll")%></option>
						<option value="V" <% If Access = "V" Then %>selected<% End IF %>>
						<%=getadminCartOptLngStr("DtxtAgents")%></option>
						<option value="C" <% If Access = "C" Then %>selected<% End IF %>>
						<%=getadminCartOptLngStr("DtxtClients")%></option>
						<option value="D" <% If Access = "D" Then %>selected<% End IF %>>
						<%=getadminCartOptLngStr("DtxtDisabled")%></option>
						</select><select size="1" class="input" name="OP" style="width: 100px;">
						<option <% If OP = "O" Then %>selected<%end if %> value="O">
						<%=getadminCartOptLngStr("DtxtOLK")%></option>
						<option <% If OP = "P" Then %>selected<%end if %> value="P">
						<%=getadminCartOptLngStr("DtxtPocket")%></option>
						<option <% If OP = "T" Then %>selected<%end if %> value="T">
						<%=getadminCartOptLngStr("DtxtOLK")%>/<%=getadminCartOptLngStr("DtxtPocket")%></option>
						</select></nobr></td></tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCartOptLngStr("DtxtAlignment")%>&nbsp;</font></b></td>
						<td valign="top" class="style4">
						<select size="1" name="Align" class="input">
						<option></option>
						<option <% If Align = "L" Then %>selected<% End If %> value="L"><%=getadminCartOptLngStr("DtxtLeft")%></option>
						<option <% If Align = "R" Then %>selected<% End If %> value="R"><%=getadminCartOptLngStr("DtxtRight")%></option>
						<option <% If Align = "C" Then %>selected<% End If %> value="C"><%=getadminCartOptLngStr("DtxtCenter")%></option>
						<option <% If Align = "J" Then %>selected<% End If %> value="J"><%=getadminCartOptLngStr("DtxtJustify")%></option>
						</select></td></tr>

					<tr>
						<td bgcolor="#E2F3FC">&nbsp;</td>
							<td valign="top" class="style4"><b>
						<font size="1" face="Verdana" color="#31659C">
							<input type="checkbox" name="chkDynamic" id="chkDynamic" <% If Dynamic = "Y" Then %>checked<% End If %> class="noborder" value="Y"><label for="chkDynamic"><%=getadminCartOptLngStr("DtxtDynamic")%></label></font></td></tr>
				</table>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td valign="top">
						<table border="0" width="100%" cellpadding="0">
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminCartOptLngStr("DtxtQuery")%> from TDOC, DOC1, OITM</strong></font></td>
								<td valign="top" class="style4">
								<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="20" name="customSql" style="width: 100%;" class="input" onkeydown="return catchTab(this,event)" onkeypress="javascript:document.form2.btnVerfyFilter.src='images/btnValidate.gif';document.form2.btnVerfyFilter.style.cursor = 'hand';;document.form2.valQuery.value='Y';"><%=myHTMLEncode(Field)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="|D:txtDefinition|" onclick="javascript:doFldNote(22, 'customSql', '<%=Request("rI")%>', <% If Request("rI") <> "" Then %>null<% Else %>document.form2.customSqlDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminCartOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminCartOptLngStr("DtxtVariables")%></strong></font></td>
								<td class="style4">
								<font face="Verdana" size="1" color="#4783C5">
								<span dir="ltr">@LogNum</span> = <%=getadminCartOptLngStr("DtxtLogNum")%><br>
								<span dir="ltr">@ItemCode</span> = <%=getadminCartOptLngStr("DtxtItemCode")%><br>
								<span dir="ltr">@PriceList</span> = <%=getadminCartOptLngStr("DtxtPList")%><br>
								<span dir="ltr">@CardCode</span> = <%=getadminCartOptLngStr("DtxtClientCode")%><br>
								<span dir="ltr">@SlpCode</span> = <%=getadminCartOptLngStr("DtxtAgentCode")%><br>
								<span dir="ltr">@dbName</span> = <%=getadminCartOptLngStr("DtxtDB")%><br>
								<span dir="ltr">@WhsCode</span> = <%=getadminCartOptLngStr("DtxtWhsCode")%><br>
								<span dir="ltr">@LanID</span> = <%=getadminCartOptLngStr("DtxtLanID")%><br>
								<span dir="ltr">@Quantity</span> = <%=getadminCartOptLngStr("LtxtQtyInUnit")%><br>
								<span dir="ltr">@Unit</span> = <%=getadminCartOptLngStr("DtxtUnit")%>: 1 = <%=getadminCartOptLngStr("DtxtUnit")%>, 2 = <%=getadminCartOptLngStr("DtxtSalUnit")%>, 3 = <%=getadminCartOptLngStr("DtxtPackUnit")%><br>
								<span dir="ltr">@Price</span> = <%=getadminCartOptLngStr("DtxtPrice")%></font></td>
							</tr>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminCartOptLngStr("DtxtFunctions")%></strong></font></td>
								<td class="style4">
								<% HideFunctionTitle = True
								functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"-->
								</td>
							</tr>
							<tr>
								<td valign="top" style="width: 100px" class="style6">
								<strong>
								<img src="images/spacer.gif"></strong></td>
								<td>
								<img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font face="Verdana" size="1"><strong><%=getadminCartOptLngStr("DtxtLink")%></strong></font></td>
								<td>
								<table border="0" cellpadding="0" cellspacing="1" width="100%" style="border: 1px solid #4783C5; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
									<tr>
										<td width="120" class="style5">
								<font face="Verdana" size="1"><strong><%=getadminCartOptLngStr("DtxtReport")%></strong></font></td>
										<td class="style4">
								<select size="1" name="linkObject" onchange="javascript:changeObject(this.value);">
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
								<option <% If CStr(linkObject) = CStr(rs("rsIndex")) Then %>selected<% End If %> value="<%=rs("rsIndex")%>"><%=myHTMLEncode(rs("rsName"))%></option>
								<% rs.movenext
								loop
								Response.Write "</optgroup>" %>
								</select></td>
										<td class="style4">
								<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
								<font face="Verdana" size="1" color="#4783C5">
								<input type="checkbox" class="noborder" <% If linkActive = "Y" Then %>checked<% End If %> name="linkActive" value="Y" id="linkActive"><label for="linkActive"><%=getadminCartOptLngStr("DtxtActive")%></label></font></td>
									</tr>
									<tr id="trRepVars" <% If linkObject = "-1" Then %>style="display:none;"<% End If %>>
										<td width="120" valign="top" class="style5">
										<font face="Verdana" size="1"><strong><%=getadminCartOptLngStr("DtxtVariables")%></strong></font></td>
										<td colspan="2">
										<table border="0" cellpadding="0" cellspacing="1" width="100%" id="tblLinkVars">
											<tr>
												<td class="style3" colspan="2" style="width: 280px">
												<font face="Verdana" size="1">
												<strong><%=getadminCartOptLngStr("DtxtVariable")%></strong></font></td>
												<td class="style3">
												<font face="Verdana" size="1">
												<strong><%=getadminCartOptLngStr("DtxtValue")%> 
												</strong> </font></td>
											</tr>
											<%
											If linkObject <> "" and linkObject <> -1 Then
											sql = "select T0.varVar, T0.varDataType, IsNull(T1.valBy, 'F') valBy, T1.valValue, T1.valDate, T0.varNotNull " & _
											"from OLKRSVars T0 " & _
											"left outer join OLKCartRepLinksVars T1 on T1.ID = " & Request("rI") & " and T1.varId = T0.varVar " & _
											"where T0.rsIndex = " & linkObject
											set rs = conn.execute(sql)
											do while not rs.eof
											%>
											<tr>
												<td class="style4" style="width: 120px"><nobr><font face="Verdana" size="1" color="#4783C5">
												<span dir="ltr">@<%=rs("varVar")%><% If rs("varNotNull") = "Y" Then %><font color="red">*</font><% End If %></span></font></nobr></td>
												<td class="style4" style="width: 320px"><font face="Verdana" size="1" color="#4783C5">
												<input type="radio" class="OptionButton" style="background:background-image" value="F" <% If rs("valBy") = "F" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" id="rdFld<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','F');">
												<label for="rdFld<%=rs("varVar")%>"><%=getadminCartOptLngStr("DtxtVariable")%></label>
												<input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "V" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="V" id="rdVal<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','V');">
												<label for="rdVal<%=rs("varVar")%>"><%=getadminCartOptLngStr("DtxtValue")%></label>
												<input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "Q" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="Q" id="rdQry<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','Q');">
												<label for="rdQry<%=rs("varVar")%>"><%=getadminCartOptLngStr("DtxtQuery")%></label>
												</font></td>
												<td class="style4"><font color="#4783C5" face="Verdana" size="1">
											<table border="0" id="tblValDat<%=rs("varVar")%>" cellspacing="0" cellpadding="0" style="<% If rs("valBy") = "V" and rs("varDataType") <> "datetime" or rs("valBy") = "F" or rs("valBy") = "Q" Then %>;display: none<% End If %>">
												<tr>
													<td><img border="0" src="images/cal.gif" id="btnValDatImg<%=rs("varVar")%>" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
													<td>
													<input type="text" readonly name="colValDat<%=rs("varVar")%>" size="12" value="<%=FormatDate(rs("valDate"), False)%>" onclick="btnValDatImg<%=rs("varVar")%>.click()"></td>
													<td><img border="0" src="images/remove.gif" style="cursor: hand" onclick="javascript:document.form2.colValDat<%=rs("varVar")%>.value='';"></td>
												</tr>
											</table>
											<input style="<% If rs("valBy") = "F" or rs("varDataType") = "datetime" or rs("valBy") = "Q" Then %>display: none<% End If %>" type="text" name="valValueV<%=rs("varVar")%>" id="valValueV<%=rs("varVar")%>" size="25" value="<% If rs("valBy") = "V" and not IsNull(rs("valValue")) Then Response.Write Server.HTMLEncode(rs("valValue"))%>" onchange="valThis(this,'<%=rs("varVar")%>');"><select style="<% If rs("valBy") = "V" or rs("valBy") = "Q" Then %>display: none<% End If %>" size="1" name="valValueF<%=rs("varVar")%>" id="valValueF<%=rs("varVar")%>">
											<option></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@LogNum" Then Response.Write "selected" %> value="@LogNum"><%=getadminCartOptLngStr("DtxtLogNum")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@LineNum" Then Response.Write "selected" %> value="@LineNum"><%=getadminCartOptLngStr("DtxtLine")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@ItemCode" Then Response.Write "selected" %> value="@ItemCode"><%=getadminCartOptLngStr("DtxtItemCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@PriceList" Then Response.Write "selected" %> value="@PriceList"><%=getadminCartOptLngStr("DtxtPList")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@CardCode" Then Response.Write "selected" %> value="@CardCode"><%=getadminCartOptLngStr("DtxtClientCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@SlpCode" Then Response.Write "selected" %> value="@SlpCode"><%=getadminCartOptLngStr("DtxtAgentCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@WhsCode" Then Response.Write "selected" %> value="@WhsCode"><%=getadminCartOptLngStr("DtxtWhsCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@Quantity" Then Response.Write "selected" %> value="@Quantity"><%=getadminCartOptLngStr("LtxtQtyInUnit")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@Unit" Then Response.Write "selected" %> value="@Unit"><%=getadminCartOptLngStr("DtxtUnit")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@Price" Then Response.Write "selected" %> value="@Price"><%=getadminCartOptLngStr("DtxtPrice")%></option>
											</select></font>
											<table cellpadding="0" cellspacing="0" border="0" width="100%" <% If rs("valBy") <> "Q" Then %>style="display: none;"<% End If %> id="tblQuery<%=rs("varVar")%>">
												<tr>
													<td>
														<textarea dir="ltr" style="width: 100%; " name="valQuery<%=rs("varVar")%>" id="valQuery<%=rs("varVar")%>" onchange="javascript:document.form2.btnVerfyQueryVar<%=rs("varVar")%>.src='images/btnValidate.gif';document.form2.btnVerfyQueryVar<%=rs("varVar")%>.style.cursor = 'hand';;document.form2.valQueryVar<%=rs("varVar")%>.value='Y';"><% If rs("valBy") = "Q" Then %><%=myHTMLEncode(rs("valValue"))%><% End If %></textarea>
													</td>
													<td width="1" valign="bottom">
													<img src="images/btnValidateDis.gif" id="btnVerfyQueryVar<%=rs("varVar")%>" alt="<%=getadminCartOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQueryVar<%=rs("varVar")%>.value == 'Y')VerfyQueryVar('<%=rs("varVar")%>');">
													<input type="hidden" name="valQueryVar<%=rs("varVar")%>" id="valQueryVar<%=rs("varVar")%>" value="N">
													<img src="images/spacer.gif"></td>
												</tr>
											</table>
											</td>
											<input type="hidden" name="varDataType<%=rs("varVar")%>" id="varDataType<%=rs("varVar")%>" value="<%=rs("varDataType")%>">
											<input type="hidden" name="varVar" value="<%=rs("varVar")%>">
											<input type="hidden" name="varNotNull" value="<%=rs("varNotNull")%>">
											</tr>
											<% If rs("varDataType") = "datetime" Then %>
											<script language="javascript">
											Calendar.setup({
											    inputField     :    "colValDat<%=rs("varVar")%>",     // id of the input field
											    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
											    button         :    "btnValDatImg<%=rs("varVar")%>",  // trigger for the calendar (button ID)
											    align          :    "Bl",           // alignment (defaults to "Bl")
											    singleClick    :    true
											});
											</script>
											<% End If %>
											<script language="javascript">rsHasVars = true;</script>
											<% rs.movenext
											loop
											Else %>
											<tr>
												<td colspan="3" class="style4">
												<p align="center">
												<font face="Verdana" size="1" color="#4783C5">
												<%=getadminCartOptLngStr("LtxtNoRSVars")%></font></td>
											</tr>
											<% End If %>
										</table>
										</td>
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
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCartOptLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminCartOptLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminCartOptLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('|D:txtConfCancel|'))window.location.href='adminCartOpt.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="rI" value="<%=Request("rI")%>">
	<input type="hidden" name="submitCmd" value="adminCartOpt">
	<input type="hidden" name="cmd" value="<% If Request("Edit") = "Y" Then %>e<% Else %>a<% End If %>">
	<% End If %>
	
</table>
<% If Request("edit") = "Y" or Request("NewFld") = "Y" Then %>
</form>
<script language="javascript">
NumUDAttach('form2', 'Order', 'btnOrderUp', 'btnOrderDown');
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="cartOpt">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<form name="frmGetRSVars" action="getRSVars.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="rsIndex" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %><!--#include file="bottom.asp" -->