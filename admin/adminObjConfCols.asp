<!--#include file="top.asp" -->
<!--#include file="lang/adminObjConfCols.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
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
.style7 {
	background-color: #F3FBFE;
	direction: ltr;
}
</style>
</head>

<script language="javascript" src="js_up_down.js"></script>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<link rel="stylesheet" type="text/css" href="style_cal.css">
<script language="javascript">
var LtxtNoRSVars = '<%=getadminObjConfColsLngStr("LtxtNoRSVars")%>';
var LtxtValQryVal = '<%=getadminObjConfColsLngStr("LtxtValQryVal")%>';
var LtxtValFldNam = '<%=getadminObjConfColsLngStr("LtxtValFldNam")%>';
var LtxtValFldNam2 = '<%=getadminObjConfColsLngStr("LtxtValFldNam2")%>';
var LtxtValQry = '<%=getadminObjConfColsLngStr("LtxtValQry")%>';
var LtxtSelFldVar = '<%=getadminObjConfColsLngStr("LtxtSelFldVar")%>';
var LtxtValVarVal = '<%=getadminObjConfColsLngStr("LtxtValVarVal")%>';
var LtxtValVarQry = '<%=getadminObjConfColsLngStr("LtxtValVarQry")%>';
var LtxtValVarQryBrnVerfy = '<%=getadminObjConfColsLngStr("LtxtValVarQryBrnVerfy")%>';
var DtxtValNumVal = '|D:txtValNumVal|';
var DtxtVariable = '<%=getadminObjConfColsLngStr("DtxtVariable")%>';
var DtxtValue = '<%=getadminObjConfColsLngStr("DtxtValue")%>';
var DtxtQuery = '<%=getadminObjConfColsLngStr("DtxtQuery")%>';
var LtxtValRepLnk = '<%=getadminObjConfColsLngStr("LtxtValRepLnk")%>';
var DtxtValidate = '<%=getadminObjConfColsLngStr("DtxtValidate")%>';
var DtxtLogNum = '<%=getadminObjConfColsLngStr("DtxtLogNum")%>';
var LtxtActionID = '<%=getadminObjConfColsLngStr("LtxtActionID")%>';
var CalendarFormat = '<%=GetCalendarFormatString%>';
var typeID = '<%=Request("TypeID")%>';
</script>
<script language="javascript" src="adminObjConfCols.js"></script>
<table border="0" cellpadding="0" width="100%" id="table3">
<% If Request("ID") = "" and Request("New") <> "Y" Then
sql = "select T0.ID, T1.Name + ' - ' + T0.Name Name " & _  
		"from OLKOps T0  " & _  
		"inner join OLKOpsGrps T1 on T1.ID = T0.GroupID " & _  
		"where T0.Status <> 'D' " 
set ro = Server.CreateObject("ADODB.RecordSet")
ro.open sql, conn, 3, 1
%>
<form method="POST" action="adminsubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminObjConfColsLngStr("LttlObjConfDet")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminObjConfColsLngStr("LttlObjConfNote")%> </font></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><select name="TypeID" size="1" onchange="window.location.href='?TypeID=' + this.value;">
		<option value=""><%=getadminObjConfColsLngStr("LtxtSelConfType")%></option>
		<option <% If Request("TypeID") = "A" Then %>selected<% End If %> value="A"><%=getadminObjConfColsLngStr("DtxtActions")%></option>
		<option <% If Request("TypeID") = "C" Then %>selected<% End If %> value="C"><%=getadminObjConfColsLngStr("DtxtBPS")%></option>
		<option <% If Request("TypeID") = "I" Then %>selected<% End If %> value="I"><%=getadminObjConfColsLngStr("DtxtItems")%></option>
		<option <% If Request("TypeID") = "D" Then %>selected<% End If %> value="D"><%=getadminObjConfColsLngStr("DtxtComDocs")%></option>
		<option <% If Request("TypeID") = "R" Then %>selected<% End If %> value="R"><%=getadminObjConfColsLngStr("DtxtReceipts")%></option>
		<% do while not ro.eof %>
		<option <% If Request("TypeID") = "OP" & ro("ID") Then %>selected<% End If %> value="OP<%=ro("ID")%>"><%=getadminObjConfColsLngStr("DtxtOp")%> - <%=ro("Name")%></option>
		<% ro.movenext
		loop %>
		</select></td>
	</tr><% 
	If Request("TypeID") <> "" Then 
	set rs = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetObjConfCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TypeID") = Request("TypeID")
	rs.open cmd, , 3, 1 %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" class="style1" style="width: 16px">
				&nbsp;</td>
				<td align="center" class="style1" style="width: 200px">
				<font size="1" face="Verdana" color="#31659C"><%=getadminObjConfColsLngStr("DtxtName")%>&nbsp;</font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminObjConfColsLngStr("DtxtOrder")%></font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminObjConfColsLngStr("DtxtCodification")%></font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminObjConfColsLngStr("DtxtField")%>&nbsp;/ 
				<%=getadminObjConfColsLngStr("DtxtQuery")%></font></td>
				<td align="center" class="style1">
				&nbsp;</td>
				<td align="center" width="16" class="style2">&nbsp;</td>
			</tr>
			<%
			If rs.recordcount > 0 then
			do While NOT RS.EOF 
			ID = rs("ID") %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
				<a href="adminObjConfCols.asp?TypeID=<%=Request("TypeID")%>&ID=<%=ID%>#table20"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top" style="width: 200px">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input class="input" style="width: 100%; " size="20" value="<%=Server.HTMLEncode(RS("Name"))%>" name="rowName<%=ID%>" id="rowName" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td style="width: 16px"><a href="javascript:doFldTrad('ObjConfCols', 'TypeID,ID', '<%=Request("TypeID")%>,<%=ID%>', 'alterName', 'T', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
				</table>
			  </td>
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder<%=ID%>" id="RowOrder<%=ID%>" size="7" style="text-align:right" class="input"onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Order")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrder<%=ID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrder<%=ID%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td>
				<nobr>
				<select size="1" name="rowType<%=ID%>" class="input">
				<option value="T">
				<%=getadminObjConfColsLngStr("DtxtDisabled")%></option>
				<option value="L" <% If rs("Encode") = "L" Then %> selected<% End If %>>
				<%=getadminObjConfColsLngStr("DtxtLow")%></option>
				<option value="M"<% If rs("Encode") = "M" Then %> selected<% End If %>>
				<%=getadminObjConfColsLngStr("DtxtMedium")%></option>
				<option value="H"<% If rs("Encode") = "H" Then %> selected<% End If %>>
				<%=getadminObjConfColsLngStr("DtxtHigh")%></option>
				</select>
				<input type="checkbox" <% If rs("EncodeRnd") = "Y" Then %>checked<% End If %> name="rowTypeRnd<%=ID%>" id="rowTypeRnd<%=ID%>" value="ON" class="noborder"><font color="#31659C" face="Verdana" size="1"><label for="rowTypeRnd<%=ID%>"><%=getadminObjConfColsLngStr("DtxtRndLtr")%></label></font></nobr></td>
				<td valign="top" align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("Query"))%>"></td>
				<td valign="top" align="center">
				<input type="checkbox" name="chkActive<%=ID%>" <% If rs("Active") = "Y" Then %>checked<% End If %> value="Y" id="chkActive<%=ID%>" class="noborder"><font face="Verdana" size="1" color="#31659C"><label for="chkActive<%=ID%>"><%=getadminObjConfColsLngStr("DtxtActive")%></label></font></td>
				<td valign="middle" width="16">
						<a href="javascript:if(confirm('<%=getadminObjConfColsLngStr("LtxtConfDelFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("Name")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&TypeID=<%=Request("TypeID")%>&ID=<%=ID%>&submitCmd=adminObjConfCols';">
						<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="rowIndex" value="<%=ID%>">
				<script language="javascript">NumUDAttach('form1', 'RowOrder<%=ID%>', 'btnRowOrder<%=ID%>Up', 'btnRowOrder<%=ID%>Down');</script>
				<% RS.MoveNext
				loop
				Else %>
				<tr>
					<td align="center" class="style1" colspan="7">
					<font size="1" face="Verdana" color="#31659C"><%=getadminObjConfColsLngStr("DtxtNoData")%></font></td>
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
				<input type="submit" value="<%=getadminObjConfColsLngStr("DtxtSave")%>" <% If rs.recordcount = 0 then %>disabled<% End If %> name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminObjConfColsLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminObjConfCols.asp?TypeID=<%=Request("TypeID")%>&New=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminObjConfCols">
<input type="hidden" name="cmd" value="u">
</form>
<% End If %>
<% End If %>
<% If Request("ID") <> "" or Request("New") = "Y" Then %>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("ID") = "" Then %><%=getadminObjConfColsLngStr("LttlAddFldObjConf")%><% Else %><%=getadminObjConfColsLngStr("LttlEditFldObjConf")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5">
		<%=getadminObjConfColsLngStr("LttlAddFldObjConfNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">
		<font face="Verdana" size="1" color="#4783C5"><strong><%=getadminObjConfColsLngStr("DtxtType")%>:</strong>&nbsp;<% Select Case Request("TypeID")
			Case "A" %><%=getadminObjConfColsLngStr("DtxtActions")%>
		<%	Case "C" %><%=getadminObjConfColsLngStr("DtxtBPS")%>
		<%	Case "I" %><%=getadminObjConfColsLngStr("DtxtItems")%>
		<%	Case "D" %><%=getadminObjConfColsLngStr("DtxtComDocs")%>
		<%	Case "R" %><%=getadminObjConfColsLngStr("DtxtReceipts")%>
		<%	Case Else 
			If Left(Request("TypeID"), 2) = "OP" Then
				sql = "select T1.Name + ' - ' + T0.Name Name, T0.ObjectID " & _  
						"from OLKOps T0  " & _  
						"inner join OLKOpsGrps T1 on T1.ID = T0.GroupID " & _  
						"where T0.ID =  " & Replace(Request("TypeID"), "OP", "")
				set ro = Server.CreateObject("ADODB.RecordSet")
				ro.open sql, conn, 3, 1
				Response.Write ro("Name")
				OpObj = ro("ObjectID")
			End If
		End Select %></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("ID") <> "" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetObjConfColsDetails" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@TypeID") = Request("TypeID")
			cmd("@ID") = Request("ID")
			set rs = cmd.execute()
			rowName = rs("Name")
			rowField = rs("Query")
			rowType = rs("Encode")
			rowTypeRnd = rs("EncodeRnd")
			rowOrder = rs("Order")
			LinkType = rs("LinkType")
			linkActive = rs("LinkActive")
			linkObject = rs("LinkObject")
			rowTypeDec = rs("EncodeFormat")
			LinkLink = rs("LinkLink")
			Active = rs("Active")
		Else
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetObjConfColsNewOrder" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@TypeID") = Request("TypeID")
			set rs = cmd.execute()
			rowOrder = rs(0)
			rowTypeDec = "P"
			Active = "Y"
			LinkType = "R"
			rowName = "" %>
		<input type="hidden" name="rowNameTrad">
		<input type="hidden" name="customSqlDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td>
				<table border="0" cellpadding="0" width="100%" id="table20">
					<tr>
						<td align="center" bgcolor="#E2F3FC" style="width: 200px"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminObjConfColsLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminObjConfColsLngStr("DtxtOrder")%></font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminObjConfColsLngStr("DtxtCodification")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminObjConfColsLngStr("DtxtDecimal")%><br>(<%=getadminObjConfColsLngStr("DtxtCodification")%>)&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC">&nbsp;</td>
					</tr>
					<tr>
						<td valign="top" style="width: 200px" class="style4">
						<p align="center">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input style="width: 100%; " name="rowName" class="input" value="<%=Server.HTMLEncode(rowName)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('ObjConfCols', 'TypeID,ID', '<%=Request("TypeID")%>,<%=Request("ID")%>', 'alterName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.rowNameTrad<% End If %>);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
							</tr>
						</table>
						</td>
						<td valign="top" class="style4">
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
						<td valign="top" class="style4">
						<nobr>
						<select size="1" name="rowType" class="input" style="width: 120; height: 16">
						<option value="T"<% If rowType = "T" Then %> selected<% End If %>>
						<%=getadminObjConfColsLngStr("DtxtDisabled")%></option>
						<option value="L"<% If rowType = "L" Then %> selected<% End If %>>
						<%=getadminObjConfColsLngStr("DtxtLow")%></option>
						<option value="M"<% If rowType = "M" Then %> selected<% End If %>>
						<%=getadminObjConfColsLngStr("DtxtMedium")%></option>
						<option value="H"<% If rowType = "H" Then %> selected<% End If %>>
						<%=getadminObjConfColsLngStr("DtxtHigh")%></option>
						</select><input type="checkbox" name="rowTypeRnd" <% If rowTypeRnd = "Y" Then %>checked<% End If %> value="ON" id="rowTypeRnd" class="noborder"><font face="Verdana" size="1" color="#31659C"><label for="rowTypeRnd"><%=getadminObjConfColsLngStr("DtxtRndLtr")%></label></nobr></td>
						<td valign="top" class="style4">
						<select size="1" name="rowTypeDec" class="input">
						<option <% If rowTypeDec = "S" Then %>selected<% End If %> value="S"><%=getadminObjConfColsLngStr("DtxtDecSum")%></option>
						<option <% If rowTypeDec = "P" Then %>selected<% End If %> value="P"><%=getadminObjConfColsLngStr("DtxtDecPrice")%></option>
						<option <% If rowTypeDec = "R" Then %>selected<% End If %> value="R"><%=getadminObjConfColsLngStr("DtxtDecRate")%></option>
						<option <% If rowTypeDec = "Q" Then %>selected<% End If %> value="Q"><%=getadminObjConfColsLngStr("DtxtDecQty")%></option>
						<option <% If rowTypeDec = "%" Then %>selected<% End If %> value="%"><%=getadminObjConfColsLngStr("DtxtDecPercent")%></option>
						<option <% If rowTypeDec = "M" Then %>selected<% End If %> value="M"><%=getadminObjConfColsLngStr("DtxtDecMeasure")%></option>
						</select></td>
						<td valign="top" class="style4" align="center">
						<input type="checkbox" name="chkActive" <% If Active = "Y" Then %>checked<% End If %> value="Y" id="chkActive" class="noborder"><font face="Verdana" size="1" color="#31659C"><label for="chkActive"><%=getadminObjConfColsLngStr("DtxtActive")%></label></font></td>
					</tr>
				</table>
				<table border="0" cellpadding="0" width="100%" id="table20">
					<tr>
						<td valign="top">
						<table border="0" width="100%" id="table23" cellpadding="0">
							<tr>
								<td valign="top" colspan="2" class="style4">
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="10" style="width: 100%" name="customSql" cols="100" class="input" onkeypress="javascript:document.form2.btnVerfyFilter.src='images/btnValidate.gif';document.form2.btnVerfyFilter.style.cursor = 'hand';;document.form2.valQuery.value='Y';"><%=myHTMLEncode(rowField)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="|D:txtDefinition|" onclick="javascript:doFldNote(21, 'customSql', '<%=Request("TypeID")%>,<%=Request("ID")%>', <% If Request("ID") <> "" Then %>null<% Else %>document.form2.customSqlDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminObjConfColsLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<% If Request("TypeID") <> "A" Then %>
							<tr>
								<td valign="top" bgcolor="#E2F3FC" style="width: 120px" class="style5">
								<font size="1" face="Verdana">
								<strong><%=getadminObjConfColsLngStr("LtxtAvlTables")%></strong></font></td>
								<td class="style4"><label for="fx1">
								<font face="Verdana" size="1" color="#4783C5"><% Select Case Request("TypeID")
									Case "C" %>TCRD = <%=getadminObjConfColsLngStr("DtxtBP")%><%
									Case "I" %>TITM = <%=getadminObjConfColsLngStr("DtxtItem")%><%
									Case "D" %>TDOC = <%=getadminObjConfColsLngStr("DtxtComDocs")%><%
									Case "R" %>TPMT = <%=getadminObjConfColsLngStr("DtxtReceipt")%><%
									Case Else
										Select Case OpObj
											Case 2
												%>TCRD = <%=getadminObjConfColsLngStr("DtxtBP")%><%
											Case 4
												%>TITM = <%=getadminObjConfColsLngStr("DtxtItem")%><%
											Case 24
												%>TPMT = <%=getadminObjConfColsLngStr("DtxtReceipt")%><%
											Case Else
												 %>TDOC = <%=getadminObjConfColsLngStr("DtxtComDocs")%><%
										End Select
								End Select %></font></label></td>
							</tr>
							<% End If %>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminObjConfColsLngStr("DtxtVariables")%></strong></font></td>
								<td class="style7">
								<font face="Verdana" size="1" color="#4783C5">
								<% Select Case Request("TypeID")
									Case "A" %><span dir="ltr">@ID</span> = <%=getadminObjConfColsLngStr("LtxtActionID")%><%
									Case Else %><span dir="ltr">@LogNum</span> = <%=getadminObjConfColsLngStr("DtxtLogNum")%><%
								End Select %></font></td>
							</tr>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminObjConfColsLngStr("DtxtFunctions")%></strong></font></td>
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
								<font face="Verdana" size="1"><strong><%=getadminObjConfColsLngStr("DtxtLink")%></strong></font></td>
								<td>
								<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table25" style="border: 1px solid #4783C5; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
									<tr>
										<td width="120" class="style5">
										<select name="LinkType" size="1" onchange="changeLinkType(this.value);">
										<option value="R"><%=getadminObjConfColsLngStr("DtxtReport")%></option>
										<option value="F" <% If LinkType = "F" Then %>selected<% End If %>><%=getadminObjConfColsLngStr("DtxtForm")%></option>
										</select></td>
										<td class="style4">
										<select size="1" name="linkObjectRS" id="linkObjectRS" <% If LinkType = "F" Then %>style="display: none; "<% End If %> onchange="javascript:changeObject(this.value);">
										<option></option>
										<%
											LastRG = ""
											set cmd = Server.CreateObject("ADODB.Command")
											cmd.ActiveConnection = connCommon
											cmd.CommandType = &H0004
											cmd.CommandText = "DBOLKGetRSList" & Session("ID")
											cmd.Parameters.Refresh()
											cmd("@UserType") = "V"
											set rs = cmd.execute()
											do while not rs.eof
											If LastRG <> rs("rgName") Then
												If LastRG <> "" Then Response.Write "</optgroup>"
												Response.WRite "<optgroup label=""" & myHTMLEncode(rs("rgName")) & """>"
												LastRG = rs("rgName")
											End If %>
											<option <% If CStr(linkObject) = CStr(rs("rsIndex")) and LinkType = "R" Then %>selected<% End If %> value="<%=rs("rsIndex")%>"><%=myHTMLEncode(rs("rsName"))%></option>
											<% rs.movenext
											loop
											Response.Write "</optgroup>" %>
										</select><select size="1" name="linkObjectForm" id="linkObjectForm" <% If LinkType = "R" Then %>style="display: none; "<% End If %> onchange="javascript:changeObject(this.value);">
										<option></option>
										<%  set cmd = Server.CreateObject("ADODB.Command")
											cmd.ActiveConnection = connCommon
											cmd.CommandType = &H0004
											cmd.CommandText = "DBOLKGetSectionsList" & Session("ID")
											cmd.Parameters.Refresh()
											cmd("@UserType") = "A"
											set rs = cmd.execute()
											do while not rs.eof %>
											<option <% If CStr(linkObject) = CStr(rs("SecID")) and LinkType = "F" Then %>selected<% End If %> value="<%=rs("SecID")%>"><%=rs("SecName")%></option>
											<% rs.movenext
											loop %>
										</select></td>
										<td class="style4">
								<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
								<font face="Verdana" size="1" color="#4783C5">
								<input type="checkbox" class="noborder" <% If linkActive = "Y" Then %>checked<% End If %> name="linkActive" value="Y" id="linkActive"><label for="linkActive"><%=getadminObjConfColsLngStr("DtxtActive")%></label></font></td>
									</tr>
									<tr id="trRepVars" <% If linkObject = "-1" or LinkType = "F" Then %>style="display:none;"<% End If %>>
										<td width="120" valign="top" class="style5">
										<font face="Verdana" size="1"><strong><%=getadminObjConfColsLngStr("DtxtVariables")%></strong></font></td>
										<td colspan="2">
										<table border="0" cellpadding="0" cellspacing="1" width="100%" id="tblLinkVars">
											<tr>
												<td class="style3" colspan="2" style="width: 280px">
												<font face="Verdana" size="1">
												<strong><%=getadminObjConfColsLngStr("DtxtVariable")%></strong></font></td>
												<td class="style3">
												<font face="Verdana" size="1">
												<strong><%=getadminObjConfColsLngStr("DtxtValue")%> 
												</strong> </font></td>
											</tr>
											<%
											If linkObject <> "" and linkObject <> "-1" and LinkType = "R" Then
											set cmd = Server.CreateObject("ADODB.Command")
											cmd.ActiveConnection = connCommon
											cmd.CommandType = &H0004
											cmd.CommandText = "DBOLKGetRSVarsObjConf" & Session("ID")
											cmd.Parameters.Refresh()
											cmd("@TypeID") = Request("TypeID")
											cmd("@ID") = Request("ID")
											cmd("@rsIndex") = linkObject
											set rs = cmd.execute()
											do while not rs.eof
											%>
											<tr>
												<td class="style4" style="width: 120px"><nobr><font face="Verdana" size="1" color="#4783C5">
												<span dir="ltr">@<%=rs("varVar")%><% If rs("varNotNull") = "Y" Then %><font color="red">*</font><% End If %></span></font></nobr></td>
												<td class="style4" style="width: 320px"><font face="Verdana" size="1" color="#4783C5">
												<input type="radio" class="OptionButton" style="background:background-image" value="F" <% If rs("valBy") = "F" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" id="rdFld<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','F');">
												<label for="rdFld<%=rs("varVar")%>"><%=getadminObjConfColsLngStr("DtxtVariable")%></label>
												<input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "V" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="V" id="rdVal<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','V');">
												<label for="rdVal<%=rs("varVar")%>"><%=getadminObjConfColsLngStr("DtxtValue")%></label>
												<input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "Q" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="Q" id="rdQry<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','Q');">
												<label for="rdQry<%=rs("varVar")%>"><%=getadminObjConfColsLngStr("DtxtQuery")%></label>
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
											<% Select Case Request("TypeID")
												Case "A" %>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@ID" Then Response.Write "selected" %> value="@ID"><%=getadminObjConfColsLngStr("LtxtActionID")%></option>
											<%	Case Else %>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@LogNum" Then Response.Write "selected" %> value="@LogNum"><%=getadminObjConfColsLngStr("DtxtLogNum")%></option><%
											End Select %>
											</select></font>
											<table cellpadding="0" cellspacing="0" border="0" width="100%" <% If rs("valBy") <> "Q" Then %>style="display: none;"<% End If %> id="tblQuery<%=rs("varVar")%>">
												<tr>
													<td>
														<textarea dir="ltr" style="width: 100%; " name="valQuery<%=rs("varVar")%>" id="valQuery<%=rs("varVar")%>" onchange="javascript:document.form2.btnVerfyQueryVar<%=rs("varVar")%>.src='images/btnValidate.gif';document.form2.btnVerfyQueryVar<%=rs("varVar")%>.style.cursor = 'hand';;document.form2.valQueryVar<%=rs("varVar")%>.value='Y';"><% If rs("valBy") = "Q" Then %><%=myHTMLEncode(rs("valValue"))%><% End If %></textarea>
													</td>
													<td width="1" valign="bottom">
													<img src="images/btnValidateDis.gif" id="btnVerfyQueryVar<%=rs("varVar")%>" alt="<%=getadminObjConfColsLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQueryVar<%=rs("varVar")%>.value == 'Y')VerfyQueryVar('<%=rs("varVar")%>');">
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
												<%=getadminObjConfColsLngStr("LtxtNoRSVars")%></font></td>
											</tr>
											<% End If %>
										</table>
										</td>
									</tr>
									<tbody id="linkLinkData" <% If linkObject = "-1" or LinkType = "R" Then %>style="display: none;"<% End If %>>
									<tr class="style4">
										<td colspan="3">
										<textarea dir="ltr" rows="3" class="input" name="linkLink" style="width: 100%"><%=myHTMLEncode(linkLink)%></textarea>
										</td>
									</tr>
									<tr>
										<td valign="top" valign="top" class="style5">
										<font face="Verdana" size="1"><strong><%=getadminObjConfColsLngStr("DtxtVariables")%></strong></font>
										</td>
										<td colspan="2" class="style4">
										<font color="#4783C5" face="Verdana" size="1">
										<select size="3" name="linkLinkVars" onclick="javascript:if(this.value!=null&&this.value!='')document.form2.linkLink.value+='{' + this.value + '}';">
											<% Select Case Request("TypeID")
												Case "A" %>
											<option value="@ID"><%=getadminObjConfColsLngStr("LtxtActionID")%></option>
											<%	Case Else %>
											<option value="@LogNum"><%=getadminObjConfColsLngStr("DtxtLogNum")%></option><%
											End Select %>
										</select></font></td>
									</tr>
									</tbody>
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
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminObjConfColsLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminObjConfColsLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminObjConfColsLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminObjConfColsLngStr("DtxtConfCancel")%>'))window.location.href='adminObjConfCols.asp?TypeID=<%=Request("TypeID")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="TypeID" value="<%=Request("TypeID")%>">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	<input type="hidden" name="submitCmd" value="adminObjConfCols">
	<input type="hidden" name="cmd" value="<% If Request("ID") <> "" Then %>e<% Else %>a<% End If %>">
	<% End If %>
</table>
<% If Request("ID") <> "" or Request("New") = "Y" Then %>
</form>
<script language="javascript">
NumUDAttach('form2', 'RowOrder', 'btnRowOrderUp', 'btnRowOrderDown');
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="objConfCols">
	<input type="hidden" name="TypeID" value="<%=Request("TypeID")%>">
	<input type="hidden" name="OpObj" value="<%=OpObj%>">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<form name="frmGetRSVars" action="getRSVars.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="rsIndex" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %><!--#include file="bottom.asp" -->