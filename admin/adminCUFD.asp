<!--#include file="top.asp" -->
<!--#include file="lang/adminCUFD.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style2 {
	background-color: #F3FBFE;
}
.style3 {
	background-color: #E2F3FC;
	color: #31659C;
}
.style4 {
	color: #4783C5;
}
.style5 {
	background-color: #E1F3FD;
	text-align: center;
	color: #31659C;
}
.style6 {
	background-color: #E1F3FD;
	text-align: center;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style7 {
	text-align: center;
}
</style>
</head>

<script language="javascript">
var objField
function Start(page, w, h, s, field) {
objField = field
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}
function setTimeStamp(varDate) {
objField.value = varDate
}
</script>
<script language="javascript" src="js_up_down.js"></script>
<% 
varx = 0
set rd = Server.CreateObject("ADODB.recordset")
set rTest = Server.CreateObject("ADODB.RecordSet") %>
<script language="javascript">
function valFrm()
{
	FieldID = document.frmCartOpt.FieldID;
	for (var i = 0;i<FieldID.length;i++)
	{
		var saveID = FieldID[i].value.replace('-', '_');
		if (document.getElementById('Query' + saveID).checked)
		{
			if (document.getElementById('SqlQuery' + saveID).value == '')
			{
				alert('<%=getadminCUFDLngStr("LtxtValFldQry")%>'.replace('{0}', document.getElementById('FieldName' + saveID).value));
				return false;
			}
			else if (document.getElementById('valQuery' + saveID).value == 'Y')
			{
				alert('<%=getadminCUFDLngStr("LtxtValFldQryVal")%>'.replace('{0}', document.getElementById('FieldName' + saveID).value));
				return false;
			}
		}
	}
	return true;
}
</script>
<form method="POST" name="frmCartOpt" action="adminSubmit.asp" onsubmit="return valFrm();">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCUFDLngStr("LttlUFld")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> 
		<span class="style4"><%=getadminCUFDLngStr("LttlUFldNote")%></span></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<font size="1" color="#4783C5" face="Verdana"><%=getadminCUFDLngStr("DtxtType")%> </font>
				<select size="1" name="TableID" class="input" onchange="javascript:window.location.href='adminCUFD.asp?TableID='+this.value;">
				<option></option>
				<optgroup label="<%=getadminCUFDLngStr("LtxtGeneral")%>">
					<option <% If Request("TableID") = "OCLG" Then %>selected<% End If %> value="OCLG">
					<%=getadminCUFDLngStr("LtxtActivities")%></option>
					<option <% If Request("TableID") = "OITM" Then %>selected<% End If %> value="OITM">
					<%=getadminCUFDLngStr("DtxtItems")%></option>
					<option <% If Request("TableID") = "OCRD" Then %>selected<% End If %> value="OCRD">
					<%=getadminCUFDLngStr("DtxtClients")%></option>
					<option <% If Request("TableID") = "OCPR" Then %>selected<% End If %> value="OCPR">
					<%=getadminCUFDLngStr("DtxtClients")%> - <%=getadminCUFDLngStr("LtxtContacts")%></option>
					<option <% If Request("TableID") = "CRD1" Then %>selected<% End If %> value="CRD1">
					<%=getadminCUFDLngStr("DtxtClients")%> - <%=getadminCUFDLngStr("LtxtAddresses")%></option>
					<option <% If Request("TableID") = "OOPR" Then %>selected<% End If %> value="OOPR">
					<%=getadminCUFDLngStr("DtxtSO")%></option>
				</optgroup>
				<optgroup label="<%=getadminCUFDLngStr("LtxtSale")%> / <%=getadminCUFDLngStr("LtxtPurchase")%>">
					<option <% If Request("TableID") = "OINV" Then %>selected<% End If %> value="OINV">
					<%=getadminCUFDLngStr("DtxtComDocs")%></option>
					<option <% If Request("TableID") = "INV1" Then %>selected<% End If %> value="INV1">
					<%=getadminCUFDLngStr("DtxtComDocs")%> - <%=getadminCUFDLngStr("DtxtLines")%></option>
				</optgroup>
				<optgroup label="<%=getadminCUFDLngStr("LtxtBanks")%>">
					<option <% If Request("TableID") = "ORCT" Then %>selected<% End If %> value="ORCT">
					<%=getadminCUFDLngStr("DtxtReceipt")%></option>
				</optgroup>
				<optgroup label="<%=getadminCUFDLngStr("DtxtService")%>">
					<option <% If Request("TableID") = "OSCL" Then %>selected<% End If %> value="OSCL">
					<%=getadminCUFDLngStr("DtxtServiceCall")%></option>
					<option <% If Request("TableID") = "SCL1" Then %>selected<% End If %> value="SCL1">
					<%=getadminCUFDLngStr("DtxtServiceCall")%> - <%=getadminCUFDLngStr("DtxtSolution")%></option>
					<option <% If Request("TableID") = "SCL5" Then %>selected<% End If %> value="SCL5">
					<%=getadminCUFDLngStr("DtxtServiceCall")%> - <%=getadminCUFDLngStr("LtxtActivities")%></option>
					<option <% If Request("TableID") = "SCL2" Then %>selected<% End If %> value="SCL2">
					<%=getadminCUFDLngStr("DtxtServiceCall")%> - <%=getadminCUFDLngStr("DtxtInvExpns")%></option>
					<option <% If Request("TableID") = "SCL3" Then %>selected<% End If %> value="SCL3">
					<%=getadminCUFDLngStr("DtxtServiceCall")%> - <%=getadminCUFDLngStr("DtxtTravelExpns")%></option>
					<option <% If Request("TableID") = "OINS" Then %>selected<% End If %> value="OINS">
					<%=getadminCUFDLngStr("DtxtEquipmentCard")%></option>
					<option <% If Request("TableID") = "OCTR" Then %>selected<% End If %> value="OCTR">
					<%=getadminCUFDLngStr("DtxtServiceContract")%></option>
					<option <% If Request("TableID") = "CTR1" Then %>selected<% End If %> value="CTR1">
					<%=getadminCUFDLngStr("DtxtServiceContract")%> - <%=getadminCUFDLngStr("DtxtItems")%></option>
					<option <% If Request("TableID") = "OSLT" Then %>selected<% End If %> value="OSLT">
					<%=getadminCUFDLngStr("DtxtSolKnowBase")%></option>
				</optgroup>
				<optgroup label="<%=getadminCUFDLngStr("DtxtInventory")%>">
					<option <% If Request("TableID") = "OWTQ" Then %>selected<% End If %> value="OWTQ">
					<%=getadminCUFDLngStr("DtxtInvTransReq")%></option>
					<option <% If Request("TableID") = "WTQ1" Then %>selected<% End If %> value="WTQ1">
					<%=getadminCUFDLngStr("DtxtInvTransReq")%> - <%=getadminCUFDLngStr("DtxtLines")%></option>
					<option <% If Request("TableID") = "OWTR" Then %>selected<% End If %> value="OWTR">
					<%=getadminCUFDLngStr("DtxtInvTrans")%></option>
					<option <% If Request("TableID") = "WTR1" Then %>selected<% End If %> value="WTR1">
					<%=getadminCUFDLngStr("DtxtInvTrans")%> - <%=getadminCUFDLngStr("DtxtLines")%></option>
				</optgroup>
				</select><input type="hidden" name="LoadFieldID" value="<%=Request("FieldID")%>"></td>
	</tr>
	<tr>
		<td>
		<table border="0" id="table6" cellpadding="0" style="width: 100%">
			<% If Request("TableID") <> "" Then
			TableID = Request("TableID")
			Select Case TableID
				Case "OINV"
					DefID = 1
				Case "INV1"
					DefID = 2
				Case "OITM"
					DefID = 3
				Case "OCRD"
					DefID = 4
				Case "OCPR"
					DefID = 5
				Case "CRD1"
					DefID = 6
				Case "ORCT"
					DefID = 7
				Case "OCLG"
					DefID = 8
				Case "OOPR"
					DefID = 9
			End Select
           sql = "select T0.FieldID, AliasID, Descr, TypeID, IsNull(EditType, '') EditType, SizeID, IsNull(Dflt,'') Dflt, NotNull, IsNull(Pos, 'D') Pos, T2.[Order], " & _
           "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
           "Then 'Y' Else 'N' End As DropDown, NullField, Active, IsNull(Query,'N') Query, SqlQuery, SqlQueryField, IsNull(AType,'V') AType, OP, RTable, " & _
           "IsNull(T1.[Order], IsNull((select Max([Order])+1+(select Count('') from CUFD X0 where TableID = T0.TableID and not (TypeID = 'N' and IsNull(EditType, '') = 'T') and not exists(select '' from OLKCUFD where TableID = X0.TableID and FieldID = X0.FieldID) and FieldID < T0.FieldID) from OLKCUFD where TableID = T0.TableID), 1)) [UDFOrder], IsNull(T1.GroupID, -1) GroupID, 'Y' Config, 'Y' EnableQry " & _
           "from cufd T0 " & _
           "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
           "left outer join OLKCUFDGroups T2 on T2.TableID = T0.TableID and T2.GroupID = IsNull(T1.GroupID, -1) " & _
           "where T0.TableId = '" & TableID & "' and not (TypeID = 'N' and IsNull(EditType, '') = 'T')"
           If Request("TableID") = "INV1" and Not myApp.SVer2007 Then sql = sql & " and T0.AliasID <> 'LineMemo' "
           If Request("FieldID") <> "" Then sql = sql & " and T0.FieldID = " & Request("FieldID") & " "
           
           sql = sql &  "union all select T0.FieldID, AliasID, Descr, null, '' EditType, null SizeID, '' Dflt, null NotNull, null Pos, T0.FieldID [Order], " & _
           "'N' DropDown, NullField, Active, IsNull(Query,'N') Query, SqlQuery, SqlQueryField, IsNull(AType,'V') AType, OP, null RTable, " & _
           "-1 [UDFOrder], -2 GroupID, T0.Config, CASE When (T0.FieldID = -2 and T0.TableID = 'CRD1') Then 'Y' When ((T0.FieldID >= -7 and T0.FieldID <= -3 or T0.FieldID = -20) and T0.TableID = 'INV1') Then 'Y' Else 'N' End EnableQry " & _
           "from (select X0.TableID, X0.FieldID, X0.Name collate database_default AliasID, ISNULL(X1.AlterName, X0.Name) collate database_default Descr, X0.Config from OLKCommon..OLKFLD X0 " & _
			"left outer join OLKCommon..OLKFLDAlterNames X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID and X1.LanID = " & Session("LanID") & " " & _
			"where X0.TableID = '" & TableID & "' and (X0.TableID <> 'CRD1' or X0.TableID = 'CRD1' and (X0.FieldID not in (-8) or X0.FieldID = -8 and N'" & myApp.LawsSet & "' = 'GB'))) T0 " & _
           "left outer join OLKCUFD T1 on T1.TableID = T0.TableID collate database_default and T1.FieldID = T0.FieldID " & _
           "left outer join OLKCUFDGroups T2 on T2.TableID = T0.TableID collate database_default and T2.GroupID = IsNull(T1.GroupID, -1) "
           If Request("FieldID") <> "" Then sql = sql & " and T0.FieldID = " & Request("FieldID") & " "
           If Request("TableID") = "INV1" Then
           		If myApp.MDStyle <> "S" Then
	           		sql = sql & " where T0.FieldID not in (-3,-4,-5,-6,-7) "
	           	Else
	           		sql = sql & " where T0.FieldID not in (-20) and T0.FieldID not in (select CASE DimCode When 1 Then -7 When 2 Then -6 When 3 Then -5 When 4 Then -4 When 5 Then -3 End from ODIM where DimActive = 'N') "
	           	End If
	           	Select Case myApp.LawsSet 
			      	Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
			      		sql = sql  & " and T0.FieldID <> -19"
			      	Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
			      		sql = sql  & " and T0.FieldID <> -18"
			      	Case Else
			      		sql = sql  & " and T0.FieldID not in (-18, -19)"
			      End Select
           End If
           If myApp.LawsSet <> "CL" and Request("TableID") = "OINV" Then 
           	sql = sql & " where T0.FieldID <> -12"
           ElseIf myApp.LawsSet <> "BR" and Request("TableID") = "INV1" Then
           	sql = sql & " and T0.FieldID <> -11"
           End If
           sql = sql & " order by 10, Pos, [UDFOrder] "

           set rs = conn.execute(sql)
           
			set rd = Server.CreateObject("ADODB.RecordSet")
			sql = "select GroupID, GroupName, [Order], (select Count('') from OLKCUFD where TableID = T0.TableID and GroupID = T0.GroupID) [Count] from OLKCUFDGroups T0 where TableID = '" & TableiD & "' order by [Order]"
			set rd = conn.execute(sql) %>
			<tr>
				<td class="style6">
				<font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtGroup")%></strong></font></td>
				<td class="style6">
				<font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtField")%></strong></font></td>
				<td class="style6">
				<font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtDescription")%></strong></font></td>
				<% If Request("TableID") <> "INV1" and Request("TableID") <> "OCPR" and Request("TableID") <> "CRD1" Then %>
				<td class="style6"><font face="Verdana" size="1">
				<strong><%=getadminCUFDLngStr("DtxtPosition2")%></strong></font></td>
				<% End If %>
				<td class="style6"><font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtAccess")%></strong></font></td>
				<td class="style6"><font size="1" face="Verdana">
				<strong>OLK/Mobile</strong></font></td>
				<td class="style5"><font face="Verdana" size="1">
				<strong><%=getadminCUFDLngStr("DtxtOrder")%></strong></font></td>
				<td class="style5"><font face="Verdana" size="1">
				<strong><%=getadminCUFDLngStr("DtxtSize")%></strong></font></td>
				<td class="style5"><font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtQuery")%></strong></font></td>
				<td class="style5"><font face="Verdana" size="1">
				<strong><%=getadminCUFDLngStr("DtxtNotNull")%></strong></font></td>
				<td class="style5"><font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtActive")%></strong></font></td>
			</tr>
			<% if not rs.eof then
			do while not rs.eof
			saveFieldID = Replace(rs("FieldID"), "-", "_") %>
			<tr><input type="hidden" name="FieldID" value="<%=rs("FieldID")%>">
				<input type="hidden" name="FieldName<%=saveFieldID%>" id="FieldName<%=saveFieldID%>" value="<%=rs("Descr")%>">
				<td bgcolor="#F3FBFE">
				<% If rs("FieldID") >= 0 Then %>
				<% If Request("TableID") = "ORCT" Then %><input type="hidden" name="cmbGroup<%=saveFieldID%>" value="-1"><% End If %>
				<select size="1" <% If Request("TableID") <> "ORCT" Then %>name="cmbGroup<%=saveFieldID%>"<% End If %> class="input" <% If Request("TableID") = "ORCT" Then %>disabled<% End If %>>
				<% rd.movefirst
				do while not rd.eof %>
				<option <% If rs("GroupID") = rd("GroupID") Then %>selected<% End If %> value="<%=rd("GroupID")%>"><%=rd("GroupName")%></option>
				<% rd.movenext
				loop %>
				</select><% Else %><font size="1" color="#4783C5" face="Verdana"><%=getadminCUFDLngStr("DtxtSystem")%></font><input type="hidden" name="cmbGroup<%=saveFieldID%>" value="-2"><% End If %></td>
				<td bgcolor="#F3FBFE">
				<font size="1" color="#4783C5" face="Verdana"><%=rs("AliasID")%>&nbsp;</font></td>
				<td bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><font size="1" color="#4783C5" face="Verdana"><%=rs("Descr")%>&nbsp;</font></td>
						<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" style="width: 16px">
						<a href="javascript:doFldTrad('CUFD', 'TableID,FieldID', '<%=TableID%>,<%=rs("FieldID")%>', 'AlterDescr', 'T', null);"><img src="images/trad.gif" alt="<%=getadminCUFDLngStr("DtxtTranslate")%>" border="0"></a>
						</td>
					</tr>
				</table>
				</td>
				<% If Request("TableID") <> "INV1" and Request("TableID") <> "OCPR" and Request("TableID") <> "CRD1" Then %>
				<td bgcolor="#F3FBFE"><select <% If rs("FieldID") < 0 Then %>disabled<% End If %> size="1" name="Pos<%=saveFieldID%>" class="input">
				<option <% If rs("Pos") = "D" Then %>selected<% End If %> value="D">
				<%=getadminCUFDLngStr("DtxtRight")%></option>
				<option <% If rs("Pos") = "I" Then %>selected<% End If %> value="I">
				<%=getadminCUFDLngStr("DtxtLeft")%></option>
				</select></td>
				<% Else %>
				<input type="hidden" name="Pos<%=saveFieldID%>" value="D">
				<% End If %>
				<td bgcolor="#F3FBFE"><select size="1" name="AType<%=saveFieldID%>" <% If rs("Config") = "N" Then %>disabled<% End If %> class="input">
				<option <% If rs("AType") = "C" Then %> selected<% End If %> value="C"><%=getadminCUFDLngStr("DtxtClients")%></option>
				<option <% If rs("AType") = "V" Then %> selected<% End If %> value="V"><%=getadminCUFDLngStr("DtxtAgents")%></option>
				<option <% If rs("AType") = "T" Then %> selected<% End If %> value="T"><%=getadminCUFDLngStr("DtxtAll")%></option>
				</select></td>
				<td bgcolor="#F3FBFE">
				<select size="1" name="OP<%=saveFieldID%>" <% If rs("Config") = "N" Then %>disabled<% End If %> class="input">
				<option <% If rs("OP") = "O" Then %> selected<% End If %> value="O"><%=getadminCUFDLngStr("DtxtOLK")%></option>
				<option <% If rs("OP") = "P" Then %> selected<% End If %> value="P"><%=getadminCUFDLngStr("DtxtPocket")%></option>
				<option <% If rs("OP") = "T" Then %> selected<% End If %> value="T"><%=getadminCUFDLngStr("DtxtAll")%></option>
				</select></td>
				<td bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="UDFOrder<%=saveFieldID%>" <% If rs("FieldID") < 0 or rs("Config") = "N" Then %>disabled<% End If %> id="UDFOrder<%=saveFieldID%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("UDFOrder")%>">
						</td>
						<% If rs("FieldID") >= 0 Then %><td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnUDFOrder<%=saveFieldID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnUDFOrder<%=saveFieldID%>Down"></td>
							</tr>
						</table></td><% End If %>
					</tr>
				</table>
				<% If rs("FieldID") >= 0 Then %><script language="javascript">NumUDAttach('frmCartOpt', 'UDFOrder<%=saveFieldID%>', 'btnUDFOrder<%=saveFieldID%>Up', 'btnUDFOrder<%=saveFieldID%>Down');</script><% End If %>
				</td>
				<td bgcolor="#F3FBFE" class="style7"><font size="1" color="#4783C5"><font size="1" color="#4783C5" face="verdana"><%=rs("SizeID")%></font>&nbsp;</font></td>
				<td align="center" bgcolor="#F3FBFE">
				<font color="#4783C5" size="1" face="Verdana, Arial, Helvetica, sans-serif">
				<input <% If rs("EnableQry") <> "Y" Then %>disabled<% End If %> type="checkbox" <% If rs("EditType") = "I" or Not IsNull(rs("RTable")) Then %>disabled<% End If %> name="Query<%=saveFieldID%>" id="Query<%=saveFieldID%>" <% If rs("DropDown") = "Y" Then %>disabled<% End If %> value="Y" <% If rs("Query") = "Y" and rs("DropDown") <> "Y" Then %>checked<% End If %> onclick="javascript:QueryChecked('<%=saveFieldID%>');" class="noborder"></font></td>
				<td align="center" bgcolor="#F3FBFE">
				<font color="#4783C5" size="1" face="Verdana, Arial, Helvetica, sans-serif">
				<input type="checkbox" <% If rs("Config") = "N" Then %>disabled checked<% End If %> name="Null<%=saveFieldID%>" value="Y" <% If rs("NullField") = "Y" Then %>checked<% End If %> class="noborder"></font></td>
				<td align="center" bgcolor="#F3FBFE">
				<font color="#4783C5" size="1" face="Verdana, Arial, Helvetica, sans-serif">
				<input type="checkbox" <% If rs("Config") = "N" Then %>disabled checked<% End If %> name="Active<%=saveFieldID%>" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %> class="noborder"></font></td>
			</tr>
			<input type="hidden" name="FieldID" value="<%=rs("FieldID")%>">
			<tr id="trQuery<%=saveFieldID%>" <% If rs("Query") <> "Y" Then %>style="display: none"<% End If %>>
				<td valign="top" bgcolor="#F3FBFE">&nbsp;</td>
				<td valign="top" bgcolor="#F3FBFE">&nbsp;</td>
				<td valign="top" bgcolor="#F3FBFE">&nbsp;</td>
				<td colspan="8">
				<table border="0" width="100%" id="table7" cellpadding="0">
					<tr>
						<td class="style3" style="width: 75%">
				<font size="1" face="Verdana">
				<strong><%=getadminCUFDLngStr("DtxtQuery")%> <% If rs("FieldID") = -2 and Request("TableID") = "CRD1" Then %>&nbsp;from OCRY where ...<% End If %></strong></font></td>
						<td valign="top" class="style5" style="width: 25%">
						<% If CInt(rs("FieldID")) >= 0 Then %><font size="1" face="Verdana">
						<strong><%=getadminCUFDLngStr("LtxtSelFld")%></strong></font><% End If %></td>
					</tr>
					<tr>
						<td class="style2" style="width: 75%">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td rowspan="2">
									<textarea rows="6" dir="ltr" name="SqlQuery<%=saveFieldID%>" id="SqlQuery<%=saveFieldID%>" cols="73" class="input" onkeydown="javascript:document.frmCartOpt.btnVerfy<%=saveFieldID%>.src='images/btnValidate.gif';document.frmCartOpt.btnVerfy<%=saveFieldID%>.style.cursor = 'hand';document.frmCartOpt.valQuery<%=saveFieldID%>.value='Y';" style="width: 100%; "><%=rs("sqlQuery")%></textarea>
								</td>
								<td valign="top" width="1" style="height: 20px">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCUFDLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(13, 'SqlQuery', '<%=DefID%><%=saveFieldID%>', null);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnVerfy<%=saveFieldID%>" alt="<%=getadminCUFDLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmCartOpt.valQuery<%=saveFieldID%>.value == 'Y')VerfyQuery(this, document.frmCartOpt.valQuery<%=saveFieldID%>, <%=rs("FieldID")%>, '<%=saveFieldID%>');">
									<input type="hidden" name="valQuery<%=saveFieldID%>" id="valQuery<%=saveFieldID%>" value="N">
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<table style="width: 100%">
									<tr>
										<td width="120" valign="top">
										<font face="Verdana" size="1" color="#4783C5"><strong><%=getadminCUFDLngStr("DtxtVariables")%></strong></font></td>
										<td>
										<font face="Verdana" size="1" color="#4783C5">
										<span dir="ltr">@LogNum</span> = <%=getadminCUFDLngStr("LtxtOLKDocKey")%><br>
										<% If Request("TableID") = "INV1" Then %><span dir="ltr">@LineNum</span> = <%=getadminCUFDLngStr("LtxtLineNum")%><br><% End If %>
										<span dir="ltr">@LanID</span> = <%=getadminCUFDLngStr("DtxtLanID")%><br>
										<span dir="ltr">@SlpCode</span> = <%=getadminCUFDLngStr("DtxtAgentCode")%><br>
										<span dir="ltr">@dbName</span> = <%=getadminCUFDLngStr("DtxtDB")%><br>
										<span dir="ltr">@branch</span> = <%=getadminCUFDLngStr("LtxtBranchCode")%><br>
										<% If Request("TableID") <> "OINV" Then %><span dir="ltr">@CardCode</span> = <%=getadminCUFDLngStr("DtxtClientCode")%><br><% End If %>
										<% If Request("TableID") = "INV1" Then %><span dir="ltr">@ItemCode</span> = <%=getadminCUFDLngStr("DtxtItemCode")%><br>
										<span dir="ltr">@WhsCode</span> = <%=getadminCUFDLngStr("DtxtWhsCode")%><br><% End If %>
										<% If Request("TableID") = "OINV" or Request("TableID") = "INV1" Then %><span dir="ltr">@PriceList</span> = <%=getadminCUFDLngStr("DtxtPriceList")%><% End If %>
										</font></td>
									</tr>
									</table></td>
							</tr>
						</table>
						</td>
						<td valign="top" class="style2" style="width: 25%">
						<% If CInt(rs("FieldID")) >= 0 Then %>
						<select size="1" name="SqlQueryField<%=saveFieldID%>" id="SqlQueryField<%=saveFieldID%>">
						<% If Not IsNull(rs("sqlQuery")) Then 
						
							sql = "declare @LanID int set @LanID = -1 declare @LogNum int set @LogNum = -1 " & _
									"declare @dbName nvarchar(100) set @dbName = '' declare @branch int set @branch = -1 declare @SlpCode int set @SlpCode = -1 "
							
							If Request("TableID") <> "OITM" Then
								sql = sql & "declare @CardCode nvarchar(15) set @CardCode = '' "
							End If
							
							If Request("TableID") = "OINV" or Request("TableID") = "INV1" Then
								sql = sql & "declare @PriceList int set @PriceList = -1 "
							End If
							
							If Request("TableID") = "INV1" Then
								sql = sql & "declare @LineNum int declare @ItemCode nvarchar(15) set @ItemCode = '' declare @WhsCode nvarchar(8) set @WhsCode = '' "
							End If
							
							sql = sql & rs("sqlQuery")

						set rTest = conn.execute(sql)
						For each Field in rTest.Fields %>
						<option <% If rs("SqlQueryField") = Field.Name Then %>selected<% End If %> value="<%=Field.Name%>"><%=Field.Name%></option>
						<% Next
						End If %>
						</select><% End If %></td>
					</tr>
				</table>
				</td>
			</tr>
			<% 
			rs.movenext
			loop
			else %>
			<tr>
				<td colspan="11" class="style2">
				<p align="center"><font face="Verdana" size="1" color="#4783C5">
				<%=getadminCUFDLngStr("LtxtNoAvlFld")%></font></td>
			</tr>
			<% End If %>
			<% Else %>
			<tr>
				<td colspan="11" class="style2">
				<p align="center"><font face="Verdana" size="1" color="#4783C5">
				<%=getadminCUFDLngStr("LtxtSelType")%></font></td>
			</tr>
			<% End If %>
			</table>
		</td>
	</tr>
	</table>
<% If Request("TableID") <> "" Then %>
<table border="0" cellpadding="0" width="100%">
<tr>
	<td width="77">
	<input type="submit" value="<%=getadminCUFDLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
	<td><hr color="#0D85C6" size="1"></td>
</tr>
</table>
<input type="hidden" name="submitCmd" value="adminCUFD">
</form>
<% If Request("TableID") <> "ORCT" Then %>
<script type="text/javascript">
function valFrm2()
{
	if (document.frmGroups.GroupName)
	{
		if (document.frmGroups.GroupName.length)
		{
			for (var i = 0;i<document.frmGroups.GroupName.length;i++)
			{
				if (document.frmGroups.GroupName[i].value == '')
				{
					alert('<%=getadminCUFDLngStr("LtxtValGrpNam")%>');
					document.frmGroups.GroupName[i].focus();
					return false;
				}
			}
		}
		else
		{
			if (document.frmGroups.GroupName.value == '')
			{
				alert('<%=getadminCUFDLngStr("LtxtValGrpNam")%>');
				document.frmGroups.focus();
				return false;
			}
		}
	}
	return true;
}
</script>
<table border="0" cellpadding="0" width="100%" id="tblGroups">

	<form method="POST" action="adminSubmit.asp" name="frmGroups" onsubmit="javascript:return valFrm2();">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminCUFDLngStr("LttlUfdGroups")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminCUFDLngStr("LttlUfdGroupsNote")%></td>
	</tr>
	<tr>
		<td >
		<table border="0" cellpadding="0">
		<tr class="TblRepTltSub">
				<td align="center"><%=getadminCUFDLngStr("DtxtGroup")%></td>
				<td align="center"><%=getadminCUFDLngStr("DtxtOrder")%></td>
				<td align="center" width="1"></td>
			</tr>
			<% NewOrder = 1
			rd.movefirst
			do while not rd.eof
			If CInt(rd("GroupID")) >= 0 Then GroupID = CStr(rd("GroupID")) Else GroupID = "_1" %>
			<input type="hidden" name="GroupID" value="<%=rd("GroupID")%>">
			<tr class="TblRepTbl">
				<td valign="bottom">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input <% If CInt(rd("GroupID")) = -1 Then %>readonly<% End If %> type="text" id="GroupName" name="GroupName<%=GroupID%>" size="50" value="<% If CInt(rd("GroupID")) <> -1 Then %><%=Server.HTMLEncode(rd("GroupName"))%><% Else %><%=getadminCUFDLngStr("DtxtUDF")%><% End If %>" onkeydown="return chkMax(event, this, 50);" maxlength="50"></td>
						<td><a href="javascript:doFldTrad('CUFDGroups', 'TableID,GroupID', '<%=TableID%>,<%=rd("GroupID")%>', 'AlterGroupName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminCUFDLngStr("DtxtTranslate")%>" border="0" <% If CInt(rd("GroupID")) = -1 Then %>style="visibility:hidden;"<% End If %>></a></td>
					</tr>
				</table>
				</td>
				<td valign="top" align="center">

				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="GroupOrder<%=GroupID%>" id="GroupOrder<%=GroupID%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rd("Order")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnGroupOrder<%=GroupID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnGroupOrder<%=GroupID%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				<script language="javascript">NumUDAttach('frmGroups', 'GroupOrder<%=GroupID%>', 'btnGroupOrder<%=GroupID%>Up', 'btnGroupOrder<%=GroupID%>Down');</script>
				</td>
				<td valign="top" style="width: 15px">
				<% If rd("Count") = 0 and rd("GroupID") <> -1 Then %><a href="javascript:if(confirm('<%=getadminCUFDLngStr("LtxtConfDelGrp")%>'.replace('{0}', '<%=Replace(rd("GroupName"), "'", "\'")%>')))window.location.href='adminSubmit.asp?submitCmd=adminCUFDGroups&cmd=delGrp&TableID=<%=TableID%>&id=<%=rd("GroupID")%>'"><img border="0" src="images/remove.gif"></a><% End If %></td>
			</tr>
			<% NewOrder = rd("Order") + 1
			rd.movenext
			loop %>
			<tr class="TblRep<% If Alter Then %>A<% End If %>Tbl">
				<td valign="top">
				<p align="center">
				<input type="hidden" name="GroupNameTrad" value="">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="NewGroupName" size="50" onkeydown="return chkMax(event, this, 50);" maxlength="50"></td>
						<td><a href="javascript:doFldTrad('CUFDGroups', 'TableID,GroupID', '', 'AlterGroupName', 'T', document.frmGroups.GroupNameTrad);"><img src="images/trad.gif" alt="<%=getadminCUFDLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td valign="top" align="center">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="GroupOrder" id="UDFOrder" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=NewOrder%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnGroupOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnGroupOrderDown"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				<script language="javascript">NumUDAttach('frmGroups', 'GroupOrder', 'btnGroupOrderUp', 'btnGroupOrderDown');</script></td>
				<td valign="top" style="width: 15px">
				&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminCUFDLngStr("DtxtSave")%>" name="B2"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="uGrp">
	<input type="hidden" name="TableID" value="<%=TableID%>">
	<input type="hidden" name="submitCmd" value="adminCUFDGroups">
	</form>
</table>
<% End If %>
<% End If %>
<script language="javascript">
function QueryChecked(SaveID)
{
	if (document.getElementById('Query' + SaveID).checked)
	{
		document.getElementById('trQuery' + SaveID).style.display = '';
	}
	else
	{
		document.getElementById('trQuery' + SaveID).style.display = 'none';
	}
}

var btnAfterVerfy;
var hdAfterVerfy;
var cmbSqlQueryField;
function VerfyQuery(btnVerfy, hdVerfy, FieldID, SaveID)
{
	btnAfterVerfy = btnVerfy;
	hdAfterVerfy = hdVerfy;
	document.frmVerfyQuery.FieldID.value = FieldID;
	cmbSqlQueryField = document.getElementById('SqlQueryField' + SaveID);
	document.frmVerfyQuery.Query.value = document.getElementById('SqlQuery' + SaveID).value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	//btnAfterVerfy.disabled = true;
	btnAfterVerfy.src='images/btnValidateDis.gif'
	btnAfterVerfy.style.cursor = '';
	hdAfterVerfy.value='N';
}
function getSqlQueryField()
{
	return cmbSqlQueryField;
}
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="Type" value="CUFD">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
	<input type="hidden" name="TableID" value="<%=Request("TableID")%>">
	<input type="hidden" name="FieldID" value="">
</form>

<% 
set rTest = Nothing
%><!--#include file="bottom.asp" -->