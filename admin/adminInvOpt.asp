<!--#include file="top.asp" -->
<!--#include file="lang/adminInvOpt.asp" -->
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
var LtxtNoRSVars = '<%=getadminInvOptLngStr("LtxtNoRSVars")%>';
var LtxtValQryVal = '<%=getadminInvOptLngStr("LtxtValQryVal")%>';
var LtxtValFldNam = '<%=getadminInvOptLngStr("LtxtValFldNam")%>';
var LtxtValFldNam2 = '<%=getadminInvOptLngStr("LtxtValFldNam2")%>';
var LtxtValQry = '<%=getadminInvOptLngStr("LtxtValQry")%>';
var LtxtSelFldVar = '<%=getadminInvOptLngStr("LtxtSelFldVar")%>';
var LtxtValVarVal = '<%=getadminInvOptLngStr("LtxtValVarVal")%>';
var LtxtValVarQry = '<%=getadminInvOptLngStr("LtxtValVarQry")%>';
var LtxtValVarQryBrnVerfy = '<%=getadminInvOptLngStr("LtxtValVarQryBrnVerfy")%>';
var DtxtValNumVal = '<%=getadminInvOptLngStr("DtxtValNumVal")%>';
var DtxtVariable = '<%=getadminInvOptLngStr("DtxtVariable")%>';
var DtxtValue = '<%=getadminInvOptLngStr("DtxtValue")%>';
var DtxtQuery = '<%=getadminInvOptLngStr("DtxtQuery")%>';
var DtxtItemCode = '<%=Replace(getadminInvOptLngStr("DtxtItemCode"), "'", "\'")%>';
var LtxtPListCartOnly = '<%=getadminInvOptLngStr("LtxtPListCartOnly")%>';
var DtxtClientCode = '<%=getadminInvOptLngStr("DtxtClientCode")%>';
var DtxtAgentCode = '<%=Replace(getadminInvOptLngStr("DtxtAgentCode"), "'", "\'")%>';
var DtxtDB = '<%=getadminInvOptLngStr("DtxtDB")%>'
var DtxtWhsCode = '<%=getadminInvOptLngStr("DtxtWhsCode")%>'
var LtxtNoWPocket = '<%=getadminInvOptLngStr("LtxtNoWPocket")%>'
var LtxtValRepLnk = '<%=getadminInvOptLngStr("LtxtValRepLnk")%>'
var DtxtValidate = '<%=getadminInvOptLngStr("DtxtValidate")%>'
var CalendarFormat = '<%=GetCalendarFormatString%>';
</script>
<script language="javascript" src="adminInvOpt.js"></script>
<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("edit") <> "Y" and Request("NewFld") <> "Y" Then %>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminInvOptLngStr("LttlInvAvlRepOrdr")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"> <%=getadminInvOptLngStr("LttlInvAvlRepOrdrNote")%></font></td>
	</tr>
	<form method="POST" action="adminInvOpt.asp" name="frmORDRActive">
	<tr>
		<td>
		<% 
		If Request("btnSaveItemDispORDR") <> "" Then
			sql = "select OrderCode from OLKItemDispORDR"
			set rs = conn.execute(sql)
			sql = ""
			do while not rs.eof
				If Request("Active" & rs("OrderCode")) = "Y" Then Active = "Y" Else Active = "N"
				sql = sql & "update OLKItemDispORDR set Active = '" & Active & "', OrderIndex = " & Request("DispOrder" & rs("OrderCode")) & " where OrderCode = '" & rs("OrderCode") & "' "
			rs.movenext
			loop
			conn.execute(sql)
			rs.close
		End If
		
		sql = "select * from OLKItemDispORDR order by OrderIndex"
		rs.open sql, conn, 3, 1 %>
		<table border="0" cellpadding="0" id="table24">
			<tr>
				<td width="120" class="style1"><font size="1" face="Verdana" color="#31659C">
				<%=getadminInvOptLngStr("DtxtDescription")%></font></td>
				<td width="60" class="style1">
				<p align="center"><font size="1" face="Verdana" color="#31659C">
				<%=getadminInvOptLngStr("DtxtActive")%></font></td>
				<td width="60" class="style1">
				<p align="center"><font size="1" face="Verdana" color="#31659C">
				<%=getadminInvOptLngStr("DtxtOrder")%></font></td>
			</tr>
		<% do while not rs.eof %>
			<tr>
				<td width="120" bgcolor="#F3FBFE"><font face="Verdana" size="1" color="#4783C5"><b><%
				select case rs("OrderCode")
					Case "SAP" %><%=getadminInvOptLngStr("DtxtSAP")%>
				<%	Case "BDG" %><%=getadminInvOptLngStr("DtxtWarehouse")%>
				<%	Case "OLK" %><%=getadminInvOptLngStr("DtxtOLK")%>
				<% End Select %></b></font></td>
				<td bgcolor="#F3FBFE">
				<p align="center">
				<input type="checkbox" <% If rs("Active") = "Y" Then %>checked<% End If %> name="Active<%=rs("OrderCode")%>" value="Y" class="noborder"></td>
				<td bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0" width="80">
					<tr>
						<td>
							<input type="text" name="DispOrder<%=rs("OrderCode")%>" id="DispOrder<%=rs("OrderCode")%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("OrderIndex")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnDispOrder<%=rs("OrderCode")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnDispOrder<%=rs("OrderCode")%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
			</tr>
			<script language="javascript">NumUDAttach('frmORDRActive', 'DispOrder<%=rs("OrderCode")%>', 'btnDispOrder<%=rs("OrderCode")%>Up', 'btnDispOrder<%=rs("OrderCode")%>Down');</script>
			<% rs.movenext
			loop %>
			</table>
		</td>
		</tr>
		<tr>
			<td>
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td width="75"><input type="submit" value="<%=getadminInvOptLngStr("DtxtSave")%>" name="btnSaveItemDispORDR" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:75; height:22; font-weight:bold">
					</td>
					<td><hr size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		</form>
<%
sql = "select * from olkItemRep order by rowOrder asc"
rs.close
rs.open sql, conn, 3, 1 %>
<form method="POST" action="adminsubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminInvOptLngStr("LttlItmDet")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminInvOptLngStr("LttlItmDetNote")%> </font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" class="style1" style="width: 16px">
				&nbsp;</td>
				<td align="center" class="style1" style="width: 200px">
				<font size="1" face="Verdana" color="#31659C"><%=getadminInvOptLngStr("DtxtName")%>&nbsp;</font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminInvOptLngStr("DtxtOrder")%></font></td>
				<td align="center" class="style1">
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminInvOptLngStr("DtxtCodification")%></font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminInvOptLngStr("DtxtField")%>&nbsp;/ 
				<%=getadminInvOptLngStr("DtxtQuery")%></font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminInvOptLngStr("DtxtAccess")%>&nbsp;</font></td>
				<td align="center" width="16" class="style2">&nbsp;</td>
			</tr>
			<%
			If rs.recordcount > 0 then
			do While NOT RS.EOF 
			rowIndex = Replace(rs("rowIndex"), "-", "_")
		   	varx = varx + 1 %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
				<a href="adminInvOpt.asp?edit=Y&rI=<%=rs("rowIndex")%>#table20"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top" style="width: 200px">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input class="input" style="width: 100%; " size="20" value="<%=Server.HTMLEncode(RS("rowName"))%>" name="rowName<%=rowIndex%>" id="rowName" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td style="width: 16px"><a href="javascript:doFldTrad('ItemRep', 'rowIndex', <%=rs("rowIndex")%>, 'alterRowName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminInvOptLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
			  </td>
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder<%=rowIndex%>" id="RowOrder<%=rowIndex%>" size="7" style="text-align:right" class="input"onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("RowOrder")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrder<%=rowIndex%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrder<%=rowIndex%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td>
				<nobr>
				<select size="1" name="rowType<%=rowIndex%>" class="input">
				<option value="T"<% If rs("rowType") = "T" Then %> selected<% End If %>>
				<%=getadminInvOptLngStr("DtxtDisabled")%></option>
				<option value="L"<% If rs("rowType") = "L" Then %> selected<% End If %>>
				<%=getadminInvOptLngStr("DtxtLow")%></option>
				<option value="M"<% If rs("rowType") = "M" Then %> selected<% End If %>>
				<%=getadminInvOptLngStr("DtxtMedium")%></option>
				<option value="H"<% If rs("rowType") = "H" Then %> selected<% End If %>>
				<%=getadminInvOptLngStr("DtxtHigh")%></option>
				</select>
				<input type="checkbox" <% If rs("rowTypeRnd") = "Y" Then %>checked<% End If %> name="rowTypeRnd<%=rowIndex%>" id="rowTypeRnd<%=rowIndex%>" value="ON" class="noborder"><font color="#31659C" face="Verdana" size="1"><label for="rowTypeRnd<%=rowIndex%>"><%=getadminInvOptLngStr("DtxtRndLtr")%></label></font></nobr></td>
				<td valign="top" align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("rowField"))%>"></td>
				<td valign="top">
				<nobr>
				<select size="1" class="input" name="rowAccess<%=rowIndex%>" style="width: 100">
				<option <% If Rs("rowAccess") = "T" Then %>selected<%end if %> value="T">
				<%=getadminInvOptLngStr("DtxtAll")%></option>
				<option <% If Rs("rowAccess") = "V" Then %>selected<%end if %> value="V">
				<%=getadminInvOptLngStr("DtxtAgents")%></option>
				<option <% If Rs("rowAccess") = "C" Then %>selected<%end if %> value="C">
				<%=getadminInvOptLngStr("DtxtClients")%></option>
				<option <% If Rs("rowAccess") = "D" Then %>selected<%end if %> value="D">
				<%=getadminInvOptLngStr("DtxtDisabled")%></option>
				</select>
				<select size="1" class="input" name="rowOP<%=rowIndex%>" style="width: 100">
				<option <% If Rs("rowOP") = "O" Then %>selected<%end if %> value="O">
				<%=getadminInvOptLngStr("DtxtOLK")%></option>
				<option <% If Rs("rowOP") = "P" Then %>selected<%end if %> value="P">
				<%=getadminInvOptLngStr("DtxtPocket")%></option>
				<option <% If Rs("rowOP") = "T" Then %>selected<%end if %> value="T">
				<%=getadminInvOptLngStr("DtxtOLK")%>/<%=getadminInvOptLngStr("DtxtPocket")%></option>
				</select></nobr></td>
				<td valign="middle" width="16">
						<a href="javascript:if(confirm('<%=getadminInvOptLngStr("LtxtConfDelFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("rowName")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&rI=<%=rs("rowIndex")%>&submitCmd=admininvopt';">
						<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="rowIndex" value="<%=rs("rowIndex")%>">
				<script language="javascript">NumUDAttach('form1', 'RowOrder<%=rowIndex%>', 'btnRowOrder<%=rowIndex%>Up', 'btnRowOrder<%=rowIndex%>Down');</script>
				<% RS.MoveNext
				loop
				Else %>
				<tr>
					<td align="center" class="style1" colspan="7">
					<font size="1" face="Verdana" color="#31659C"><%=getadminInvOptLngStr("DtxtNoData")%></font></td>
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
				<input type="submit" value="<%=getadminInvOptLngStr("DtxtSave")%>" <% If rs.recordcount = 0 then %>disabled<% End If %> name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminInvOptLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminInvOpt.asp?NewFld=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="admininvopt">
<input type="hidden" name="cmd" value="u">
</form>
<% End If %>

<% If Request("edit") = "Y" or Request("NewFld") = "Y" Then %>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("rI") = "" Then %><%=getadminInvOptLngStr("LttlAddFldItmDet")%><% Else %><%=getadminInvOptLngStr("LttlEditFldItmDet")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminInvOptLngStr("LttlAddFldItmDetNote")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("edit") = "Y" Then
			sql = "select rowName, rowAccess, rowField, rowType, rowOP, rowTypeRnd, rowTypeDec, rowOrder, HideNull, linkActive, IsNull(linkObject, -1) linkObject " & _
			"from olkItemRep where rowIndex = " & Request("rI") 
			set rs = conn.execute(sql) 
			rowName = rs("rowName")
			rowAccess = rs("rowAccess")
			rowField = rs("rowField")
			rowType = rs("rowType")
			rowOP = rs("rowOP")
			rowTypeRnd = rs("rowTypeRnd")
			rowOrder = rs("rowOrder")
			linkActive = rs("linkActive")
			linkObject = rs("linkObject")
			rowTypeDec = rs("rowTypeDec")
			HideNull = rs("HideNull")
		Else
			sql = "select IsNull(Max(rowOrder)+1, 0) from olkItemRep"
			set rs = conn.execute(sql)
			rowOrder = rs(0)
			rowTypeDec = "P"
			rowName = "" %>
		<input type="hidden" name="rowNameTrad">
		<input type="hidden" name="customSqlDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%" id="table20">
			<tr>
				<td>
				<table border="0" cellpadding="0">
					<tr>
						<td bgcolor="#E2F3FC" style="width: 200px"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminInvOptLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td valign="top" style="width: 200px" class="style4">
						<p align="center">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input style="width: 100%; " name="rowName" class="input" value="<%=Server.HTMLEncode(rowName)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('ItemRep', 'rowIndex', '<%=Request("rI")%>', 'alterRowName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.rowNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminInvOptLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminInvOptLngStr("DtxtOrder")%></font></b></td>
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
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminInvOptLngStr("DtxtCodification")%>&nbsp;</font></b></td>
						<td valign="top" class="style4">
						<nobr>
						<select size="1" name="rowType" class="input" style="width: 120; height: 16">
						<option value="T"<% If rowType = "T" Then %> selected<% End If %>>
						<%=getadminInvOptLngStr("DtxtDisabled")%></option>
						<option value="L"<% If rowType = "L" Then %> selected<% End If %>>
						<%=getadminInvOptLngStr("DtxtLow")%></option>
						<option value="M"<% If rowType = "M" Then %> selected<% End If %>>
						<%=getadminInvOptLngStr("DtxtMedium")%></option>
						<option value="H"<% If rowType = "H" Then %> selected<% End If %>>
						<%=getadminInvOptLngStr("DtxtHigh")%></option>
						</select><input type="checkbox" name="rowTypeRnd" <% If rowTypeRnd = "Y" Then %>checked<% End If %> value="ON" id="rowTypeRnd" class="noborder"><font face="Verdana" size="1" color="#31659C"><label for="rowTypeRnd"><%=getadminInvOptLngStr("DtxtRndLtr")%></label></nobr></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminInvOptLngStr("DtxtDecimal")%>(<%=getadminInvOptLngStr("DtxtCodification")%>)</font></b></td>
						<td valign="top" class="style4">
						<select size="1" name="rowTypeDec" class="input">
						<option <% If rowTypeDec = "S" Then %>selected<% End If %> value="S"><%=getadminInvOptLngStr("DtxtDecSum")%></option>
						<option <% If rowTypeDec = "P" Then %>selected<% End If %> value="P"><%=getadminInvOptLngStr("DtxtDecPrice")%></option>
						<option <% If rowTypeDec = "R" Then %>selected<% End If %> value="R"><%=getadminInvOptLngStr("DtxtDecRate")%></option>
						<option <% If rowTypeDec = "Q" Then %>selected<% End If %> value="Q"><%=getadminInvOptLngStr("DtxtDecQty")%></option>
						<option <% If rowTypeDec = "%" Then %>selected<% End If %> value="%"><%=getadminInvOptLngStr("DtxtDecPercent")%></option>
						<option <% If rowTypeDec = "M" Then %>selected<% End If %> value="M"><%=getadminInvOptLngStr("DtxtDecMeasure")%></option>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminInvOptLngStr("DtxtField")%>&nbsp;</font></b></td>
						<td valign="top" class="style4">
						<select <% If Request("Edit") = "Y" Then %>disabled<% End If %> size="1" name="rowField" class="input" onchange="javascript:document.form2.customSql.value=this.value;">
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
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminInvOptLngStr("DtxtAccess")%>&nbsp;</font></b></td>
						<td valign="top" class="style4">
						<nobr><select size="1" name="rowAccess" class="input" style="width: 100">
						<option value="T" <% If rowAccess = "T" Then %>selected<% End IF %>>
						<%=getadminInvOptLngStr("DtxtAll")%></option>
						<option value="V" <% If rowAccess = "V" Then %>selected<% End IF %>>
						<%=getadminInvOptLngStr("DtxtAgents")%></option>
						<option value="C" <% If rowAccess = "C" Then %>selected<% End IF %>>
						<%=getadminInvOptLngStr("DtxtClients")%></option>
						<option value="D" <% If rowAccess = "D" Then %>selected<% End IF %>>
						<%=getadminInvOptLngStr("DtxtDisabled")%></option>
						</select><select size="1" class="input" name="rowOP" style="width: 100">
						<option <% If rowOP = "O" Then %>selected<%end if %> value="O">
						<%=getadminInvOptLngStr("DtxtOLK")%></option>
						<option <% If rowOP = "P" Then %>selected<%end if %> value="P">
						<%=getadminInvOptLngStr("DtxtPocket")%></option>
						<option <% If rowOP = "T" Then %>selected<%end if %> value="T">
						<%=getadminInvOptLngStr("DtxtOLK")%>/<%=getadminInvOptLngStr("DtxtPocket")%></option>
						</select></nobr></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC">&nbsp;</td>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<input type="checkbox" name="chkHideNull" class="noborder" id="chkHideNull" <% If HideNull = "Y" Then %>checked<% End If %> value="Y"><label for="chkHideNull"><%=getadminInvOptLngStr("LtxtHideNull")%></label>&nbsp;</font></b></td>
						<td valign="top" class="style4">
						&nbsp;</td>
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
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminInvOptLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(11, 'customSql', '<%=Request("rI")%>', <% If Request("rI") <> "" Then %>null<% Else %>document.form2.customSqlDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminInvOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminInvOptLngStr("DtxtVariables")%></strong></font></td>
								<td class="style4">
								<font face="Verdana" size="1" color="#4783C5">
								<span dir="ltr">@ItemCode</span> = <%=getadminInvOptLngStr("DtxtItemCode")%><br>
								<span dir="ltr">@branchIndex</span> = <%=getadminInvOptLngStr("DtxtBranch")%><br>
								<span dir="ltr">@PriceList</span> = <%=getadminInvOptLngStr("LtxtPListCartOnly")%><br>
								<span dir="ltr">@CardCode</span> = <%=getadminInvOptLngStr("DtxtClientCode")%><br>
								<span dir="ltr">@SlpCode</span> = <%=getadminInvOptLngStr("DtxtAgentCode")%><br>
								<span dir="ltr">@dbName</span> = <%=getadminInvOptLngStr("DtxtDB")%><br>
								<span dir="ltr">@WhsCode</span> = <%=getadminInvOptLngStr("DtxtWhsCode")%> (<%=getadminInvOptLngStr("LtxtNoWPocket")%>)<br>
								<span dir="ltr">@LanID</span> = <%=getadminInvOptLngStr("DtxtLanID")%><br>
						<span dir="ltr">@Quantity</span> = <%=getadminInvOptLngStr("LtxtQtyInUnit")%><br>
						<span dir="ltr">@Unit</span> = <%=getadminInvOptLngStr("DtxtUnit")%>: 1 = <%=getadminInvOptLngStr("DtxtUnit")%>, 2 = <%=getadminInvOptLngStr("DtxtSalUnit")%>, 3 = <%=getadminInvOptLngStr("DtxtPackUnit")%><br>
						<span dir="ltr">@Price</span> = <%=getadminInvOptLngStr("DtxtPrice")%></font></td>
							</tr>
							<tr>
								<td valign="top" class="style5" style="width: 100px">
								<font size="1" face="Verdana"><strong><%=getadminInvOptLngStr("DtxtFunctions")%></strong></font></td>
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
								<font face="Verdana" size="1"><strong><%=getadminInvOptLngStr("DtxtLink")%></strong></font></td>
								<td>
								<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table25" style="border: 1px solid #4783C5; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
									<tr>
										<td width="120" class="style5">
								<font face="Verdana" size="1"><strong><%=getadminInvOptLngStr("LtxtRep")%></strong></font></td>
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
								<input type="checkbox" class="noborder" <% If linkActive = "Y" Then %>checked<% End If %> name="linkActive" value="Y" id="linkActive"><label for="linkActive"><%=getadminInvOptLngStr("DtxtActive")%></label></font></td>
									</tr>
									<tr id="trRepVars" <% If linkObject = "-1" Then %>style="display:none;"<% End If %>>
										<td width="120" valign="top" class="style5">
										<font face="Verdana" size="1"><strong><%=getadminInvOptLngStr("DtxtVariables")%></strong></font></td>
										<td colspan="2">
										<table border="0" cellpadding="0" cellspacing="1" width="100%" id="tblLinkVars">
											<tr>
												<td class="style3" colspan="2" style="width: 280px">
												<font face="Verdana" size="1">
												<strong><%=getadminInvOptLngStr("DtxtVariable")%></strong></font></td>
												<td class="style3">
												<font face="Verdana" size="1">
												<strong><%=getadminInvOptLngStr("DtxtValue")%> 
												</strong> </font></td>
											</tr>
											<%
											If linkObject <> "" and linkObject <> -1 Then
											sql = "select T0.varVar, T0.varDataType, IsNull(T1.valBy, 'F') valBy, T1.valValue, T1.valDate, T0.varNotNull " & _
											"from OLKRSVars T0 " & _
											"left outer join OLKItemRepLinksVars T1 on T1.rowIndex = " & Request("rI") & " and T1.varId = T0.varVar " & _
											"where T0.rsIndex = " & linkObject
											set rs = conn.execute(sql)
											do while not rs.eof
											%>
											<tr>
												<td class="style4" style="width: 120px"><nobr><font face="Verdana" size="1" color="#4783C5">
												<span dir="ltr">@<%=rs("varVar")%><% If rs("varNotNull") = "Y" Then %><font color="red">*</font><% End If %></span></font></nobr></td>
												<td class="style4" style="width: 320px"><font face="Verdana" size="1" color="#4783C5">
												<input type="radio" class="OptionButton" style="background:background-image" value="F" <% If rs("valBy") = "F" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" id="rdFld<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','F');">
												<label for="rdFld<%=rs("varVar")%>"><%=getadminInvOptLngStr("DtxtVariable")%></label>
												<input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "V" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="V" id="rdVal<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','V');">
												<label for="rdVal<%=rs("varVar")%>"><%=getadminInvOptLngStr("DtxtValue")%></label>
												<input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "Q" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="Q" id="rdQry<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','Q');">
												<label for="rdQry<%=rs("varVar")%>"><%=getadminInvOptLngStr("DtxtQuery")%></label>
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
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@ItemCode" Then Response.Write "selected" %> value="@ItemCode"><%=getadminInvOptLngStr("DtxtItemCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@PriceList" Then Response.Write "selected" %> value="@PriceList"><%=getadminInvOptLngStr("LtxtPListCartOnly")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@CardCode" Then Response.Write "selected" %> value="@CardCode"><%=getadminInvOptLngStr("DtxtClientCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@SlpCode" Then Response.Write "selected" %> value="@SlpCode"><%=getadminInvOptLngStr("DtxtAgentCode")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@dbName" Then Response.Write "selected" %> value="@dbName"><%=getadminInvOptLngStr("DtxtDB")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@WhsCode" Then Response.Write "selected" %> value="@WhsCode"><%=getadminInvOptLngStr("DtxtWhsCode")%> (<%=getadminInvOptLngStr("LtxtNoWPocket")%>)</option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@Quantity" Then Response.Write "selected" %> value="@Quantity"><%=getadminInvOptLngStr("LtxtQtyInUnit")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@Unit" Then Response.Write "selected" %> value="@Unit"><%=getadminInvOptLngStr("DtxtUnit")%></option>
											<option <% If rs("valBy") = "F" Then If rs("valValue") = "@Price" Then Response.Write "selected" %> value="@Price"><%=getadminInvOptLngStr("DtxtPrice")%></option>
											</select></font>
											<table cellpadding="0" cellspacing="0" border="0" width="100%" <% If rs("valBy") <> "Q" Then %>style="display: none;"<% End If %> id="tblQuery<%=rs("varVar")%>">
												<tr>
													<td>
														<textarea dir="ltr" style="width: 100%; " name="valQuery<%=rs("varVar")%>" id="valQuery<%=rs("varVar")%>" onchange="javascript:document.form2.btnVerfyQueryVar<%=rs("varVar")%>.src='images/btnValidate.gif';document.form2.btnVerfyQueryVar<%=rs("varVar")%>.style.cursor = 'hand';;document.form2.valQueryVar<%=rs("varVar")%>.value='Y';"><% If rs("valBy") = "Q" Then %><%=myHTMLEncode(rs("valValue"))%><% End If %></textarea>
													</td>
													<td width="1" valign="bottom">
													<img src="images/btnValidateDis.gif" id="btnVerfyQueryVar<%=rs("varVar")%>" alt="<%=getadminInvOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQueryVar<%=rs("varVar")%>.value == 'Y')VerfyQueryVar('<%=rs("varVar")%>');">
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
												<%=getadminInvOptLngStr("LtxtNoRSVars")%></font></td>
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
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminInvOptLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminInvOptLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminInvOptLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminInvOptLngStr("DtxtConfCancel")%>'))window.location.href='adminInvOpt.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="rI" value="<%=Request("rI")%>">
	<input type="hidden" name="submitCmd" value="admininvopt">
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
	<input type="hidden" name="type" value="invOpt">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<form name="frmGetRSVars" action="getRSVars.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="rsIndex" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %><!--#include file="bottom.asp" -->