<!--#include file="top.asp" -->
<!--#include file="lang/adminCustomSearchEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("OLKDB") & "]")
isEdit = False
isSystem = False
ObjID = CInt(Request("ObjID"))
If Request("ID") <> "" Then
	ID = Request("ID")
	isEdit = True
	If CInt(ID) < 0 Then isSystem = True
	
	sql = "select * from OLKCustomSearch where ObjectCode = " & ObjID & " and ID = " & Request("ID")
	set rs = conn.execute(sql)
	Name = rs("Name")
	IgnoreGeneralFilter = rs("IgnoreGeneralFilter") = "Y"
	Query = rs("Query")
	CatType = rs("CatType")
	Order1 = rs("Order1")
	Order2 = rs("Order2")
	Status = rs("Status")
	Ordr = rs("Ordr")

	rs.close
	
	sql = "select SessionID from OLKCustomSearchSession where ObjectCode = " & ObjID & " and ID = " & Request("ID")
	set rs = conn.execute(sql)
	do while not rs.eof
		Select Case rs(0)
			Case "C"
				SessionC = True
			Case "A"
				SessionA = True
			Case "P"
				SessionP = True
		End Select
	rs.movenext
	loop
	rs.close
Else
	sql = "select IsNull((select Max(Ordr)+1 from OLKCustomSearch where ObjectCode = " & ObjID & " and Status <> 'D'), 0) "
	set rs = conn.execute(sql)
	CatType = "S"
	Ordr = rs(0)
	Query = ""
	rs.close
End If %>

<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript" src="js_up_down.js"></script>
<script type="text/javascript">
function valFrm()
{
	if (document.frmEditSearchEdit.searchName.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValAlrNoNam")%>');
		document.frmEditSearchEdit.searchName.focus();
		return false;
	}
	<% If isEdit Then %>
	else if (document.frmEditSearchEdit.chkActive.checked && <% If ObjID = 4 Then %>!document.frmEditSearchEdit.SessionC.checked &&<% End If %> !document.frmEditSearchEdit.SessionA.checked && !document.frmEditSearchEdit.SessionP.checked)
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValNoSes")%>');
		return false;
	}
	else if (document.frmEditSearchEdit.txtQry.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValQry")%>');
		document.frmEditSearchEdit.txtQry.focus();
		return false;
	}
	else if (document.frmEditSearchEdit.valRSQuery.value == 'Y')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValQryVal2")%>');
		document.frmEditSearchEdit.btnVerfy.focus();
		return false;
	}
	<% End If %>
	return true;
}
var btnAfterVerfy;
var hdAfterVerfy;
function VerfyQuery()
{
	btnAfterVerfy = document.frmEditSearchEdit.btnVerfy;
	hdAfterVerfy = document.frmEditSearchEdit.valRSQuery;
	document.frmVerfyQuery.type.value = 'editCustomSearch';
	document.frmVerfyQuery.Query.value = document.frmEditSearchEdit.txtQry.value;
	document.frmVerfyQuery.DataType.value = '';
	document.frmVerfyQuery.QueryField.value = '';
	document.frmVerfyQuery.submit();
}
function VerfyQueryVerified()
{
	btnAfterVerfy.src='images/btnValidateDis.gif'
	btnAfterVerfy.style.cursor = '';
	hdAfterVerfy.value='N';
}
function editProp()
{
	OpenWin = window.open('adminCustomSearchProp.asp?ObjectCode=<%=Request("ObjID")%>&ID=<%=Request("ID")%>&pop=Y','OpenWin', 'width=400,height=480,scrollbars=yes');
}
</script>
</head>

<table border="0" cellpadding="0" width="100%">
<form name="frmEditSearchEdit" action="adminSubmit.asp" method="POST" onsubmit="javascript:return valFrm();">
<% If ID = "" Then %>
<input type="hidden" name="varNameTrad">
<input type="hidden" name="varQueryDef">
<% End If %>
	<% Select Case ObjID
	Case 2 %>
	<tr>
		<td class="TblRepTlt"><% If not isEdit Then %><%=getadminCustomSearchEditLngStr("LttlAddSearchCL")%><% Else %><%=getadminCustomSearchEditLngStr("LttlEditSearchCL")%><% End If %></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminCustomSearchEditLngStr("LttlAddEditSearchNotC")%></td>
	</tr>
	<% Case 4 %>
	<tr>
		<td class="TblRepTlt"><% If not isEdit Then %><%=getadminCustomSearchEditLngStr("LttlAddSearch")%><% Else %><%=getadminCustomSearchEditLngStr("LttlEditSearch")%><% End If %></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminCustomSearchEditLngStr("LttlAddEditSearchNote")%></td>
	</tr>
	<% End Select %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td style="width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtName")%></td>
				<td class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="searchName" size="50" maxlength="50" value="<%=Server.HTMLEncode(Name)%>" onkeydown="return chkMax(event, this, 50);"></td>
						<td valign="bottom"><a href="javascript:doFldTrad('CustomSearch', 'ObjectCode,ID', '<%=ObjID%>,<%=ID%>', 'alterName', 'T', <% If ID <> "" Then %>null<% Else %>document.frmEditSearchEdit.varNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCustomSearchEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
			    </td>
			</tr>
			<tr>
				<td style="width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtOrder")%></td>
				<td class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<input type="text" name="RowOrder" id="RowOrder" size="7" style="text-align: right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value='<%=Ordr%>'>
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td>
									<img src="images/img_nud_up.gif" id="btnRowOrderUp"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td>
									<img src="images/img_nud_down.gif" id="btnRowOrderDown"></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					<script language="javascript">NumUDAttach('frmEditSearchEdit', 'RowOrder', 'btnRowOrderUp', 'btnRowOrderDown');</script></td>
			</tr>
			<% If isEdit Then %>
			<tr>
				<td style="width: 98px" class="TblRepTlt">
				&nbsp;</td>
				<td class="TblRepNrm">
				<input type="checkbox" <% If Status = "Y" Then %>checked<% End If %> name="chkActive" id="chkActive" class="noborder" value="Y"><label for="chkActive"><%=getadminCustomSearchEditLngStr("DtxtActive")%></label></td>
			</tr>
			<% End If %>
			<tr>
				<td style="vertical-align: top;width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("LtxtSession")%></td>
				<td class="TblRepNrm">
				<% If ObjID <> 2 Then %><input type="checkbox" <% If SessionC Then %>checked<% End If %> name="Session" id="SessionC" value="C" class="noborder"><label for="SessionC"><%=getadminCustomSearchEditLngStr("DtxtClient")%></label><br><% End If %>
				<input type="checkbox" <% If SessionA Then %>checked<% End If %> name="Session" id="SessionA" value="A" class="noborder"><label for="SessionA"><%=getadminCustomSearchEditLngStr("DtxtAgent")%></label><br>
				<input type="checkbox" <% If SessionP Then %>checked<% End If %> name="Session" id="SessionP" value="P" class="noborder"><label for="SessionP"><%=getadminCustomSearchEditLngStr("DtxtPocket")%></label></td>
			</tr>
			<% If ObjID = 4 Then %>
			<tr>
				<td style="vertical-align: top;width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("LtxtCatalogType")%></td>
				<td class="TblRepNrm">
				<input type="radio" <% If CatType = "S" Then %>checked<% End If %> name="CatType" id="CatTypeS" value="S" class="noborder"><label for="CatTypeS"><%=getadminCustomSearchEditLngStr("DtxtStore")%></label><br>
				<input type="radio" <% If CatType = "C" Then %>checked<% End If %> name="CatType" id="CatTypeC" value="C" class="noborder"><label for="CatTypeC"><%=getadminCustomSearchEditLngStr("DtxtCat")%></label><br>
				<input type="radio" <% If CatType = "L" Then %>checked<% End If %> name="CatType" id="CatTypeL" value="L" class="noborder"><label for="CatTypeL"><%=getadminCustomSearchEditLngStr("DtxtList")%></label></td>
			</tr>
			<% Else %>
			<input type="hidden" name="CatType" value="<%=CatType%>">
			<% End If %>
			<tr>
				<td style="vertical-align: top;width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtOrder")%></td>
				<td class="TblRepNrm">
				<select size="1" name="Order1" style="width: 140px;">
				<option value=""><%=getadminCustomSearchEditLngStr("DtxtDefault")%></option>
				<% Select Case ObjID
				Case 2 %>
					<option <% If Order1 = "CardType" Then %>selected<% End If %> value="CardType"><%=getadminCustomSearchEditLngStr("DtxtType")%></option>
					<option <% If Order1 = "CardCode" or Order1 = "" Then %>selected<% End If %> selected value="CardCode"><%=getadminCustomSearchEditLngStr("DtxtCode")%></option>
					<option <% If Order1 = "CardName" Then %>selected<% End If %> value="CardName"><%=getadminCustomSearchEditLngStr("DtxtName")%></option>
					<option <% If Order1 = "CntctPrsn" Then %>selected<% End If %> value="CntctPrsn"><%=getadminCustomSearchEditLngStr("DtxtContact")%></option>
					<option <% If Order1 = "Balance" Then %>selected<% End If %> value="Balance"><%=getadminCustomSearchEditLngStr("DtxtBalance")%></option>
					<option <% If Order1 = "GroupName" Then %>selected<% End If %> value="GroupName"><%=getadminCustomSearchEditLngStr("DtxtGroup")%></option>
					<option <% If Order1 = "Name" Then %>selected<% End If %> value="Name"><%=getadminCustomSearchEditLngStr("DtxtCountry")%></option>
			<%  Case 4 %>
				<option <% If Order1 = "OITM.ItemCode" Then %>selected<% End If %> value="OITM.ItemCode"><%=getadminCustomSearchEditLngStr("DtxtCode")%></option>
				<option <% If Order1 = "ItemName" Then %>selected<% End If %> value="ItemName"><%=getadminCustomSearchEditLngStr("DtxtDescription")%></option>
				<option <% If Order1 = "Price" Then %>selected<% End If %> value="Price"><%=getadminCustomSearchEditLngStr("DtxtPrice")%></option>
				<% End Select %>
				</select>
				<select size="1" name="Order2" style="width: 100px;">
				<option value=""><%=getadminCustomSearchEditLngStr("DtxtDefault")%></option>
				<option <% If Order2 = "A" Then %>selected<% End If %> value="A"><%=getadminCustomSearchEditLngStr("DtxtAsc")%></option>
	            <option <% If Order2 = "D" Then %>selected<% End If %> value="D"><%=getadminCustomSearchEditLngStr("DtxtDesc")%></option>
	            </select>
				</td>
			</tr>
			<% If isEdit Then %>
			<tr>
				<td style="vertical-align: top;width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtQuery")%><br>
				<span style="font-weight: normal;">from <% Select Case ObjID
				Case 2 %>OCRD, OCRY, OCRG<% Case 4 %>OITM, OITW, OMRC, OITB<% End Select %><br>
				where .....</span></td>
				<td class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td rowspan="2">
							<textarea dir="ltr" rows="16" name="txtQry" cols="86" style="width: 100%" onkeydown="javascript:document.frmEditSearchEdit.btnVerfy.src='images/btnValidate.gif';document.frmEditSearchEdit.btnVerfy.style.cursor = 'hand';document.frmEditSearchEdit.valRSQuery.value='Y';"><%=myHTMLEncode(Query)%></textarea>
						</td>
						<td valign="top" width="1">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCustomSearchEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(20, 'Query', '<%=ObjID%><%=Request("ID")%>', <% If ID <> "" Then %>null<% Else %>document.frmEditSearchEdit.varQueryDef<% End If %>);">
						</td>
					</tr>
					<tr>
						<td valign="bottom" width="1">
							<img src="images/btnValidateDis.gif" id="btnVerfy" alt="<%=getadminCustomSearchEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditSearchEdit.valRSQuery.value == 'Y')VerfyQuery();">
							<input type="hidden" name="valRSQuery" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td valign="top" style="width: 100px" class="TblRepTlt">
				&nbsp;</td>
				<td class="TblRepNrm">
				<input type="checkbox" name="chkIgnoreGeneralFilter" id="chkIgnoreGeneralFilter" <% If IgnoreGeneralFilter Then %>checked<% End If %> value="Y" class="noborder"><label for="chkIgnoreGeneralFilter"><%=getadminCustomSearchEditLngStr("LtxtIgnoreGeneralFilt")%></label></td>
			</tr>
			<tr>
				<td valign="top" style="width: 100px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtVariables")%></td>
				<td class="TblRepNrm">
				<span dir="ltr">@branch</span> = <%=getadminCustomSearchEditLngStr("DtxtBranch")%><br>
				<span dir="ltr">@SlpCode</span> = <%=getadminCustomSearchEditLngStr("LtxtACodeDesc")%><br>
				<span dir="ltr">@LanID</span> = <%=getadminCustomSearchEditLngStr("DtxtLanID")%><% If ObjID = 4 Then %><br>
				<span dir="ltr">@CardCode</span> = <%=getadminCustomSearchEditLngStr("DtxtBP")%><% End If %>
				</td>
			</tr>
			<% End If %>
			<tr>
				<td style="vertical-align: top;width: 98px" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtNote")%></td>
				<td class="TblRepNrm">
				<%=getadminCustomSearchEditLngStr("LtxtVarNote")%></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminCustomSearchEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<% If isEdit Then %><td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminCustomSearchEditLngStr("DtxtSave")%>" name="btnSave"></td><% End If %>
				<% If 1 = 2 Then %><td width="100">
				<input class="BtnRep" type="button" value="|D:txtSaveAs|" name="btnSaveAs" onclick="javascript:doSaveAs();" style="width: 100px; "></td><% End If %>
				<td><hr size="1"></td>
				<% If isSystem Then %>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminCustomSearchEditLngStr("DtxtRestore")%>" name="btnRestore" onclick="javascript:if(confirm('<%=getadminCustomSearchEditLngStr("LtxtConfRestore")%>'))window.location.href='adminSubmit.asp?submitCmd=adminCustomSearch&cmd=restore&ObjID=<%=ObjID%>&ID=<%=ID%>'"></td><% End If %>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminCustomSearchEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminCustomSearchEditLngStr("DtxtConfCancel")%>'))window.location.href='adminCustomSearch.asp?ObjID=<%=ObjID%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminCustomSearch">
	<input type="hidden" name="cmd" value="save">
	<input type="hidden" name="ObjID" value="<%=ObjID%>">
	<input type="hidden" name="ID" value="<%=ID%>">
	</form>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td></td>
			</tr>
		</table>
		</td>
	</tr>
	<% If isEdit Then %>
	<tr class="TblRepNrm">
		<td><%=getadminCustomSearchEditLngStr("DtxtVariables")%></td>
	</tr>
	<% 	sql = 	"select VarID, Name, Variable, " & _
 		"Case DataType When 'S' Then " & _
 		"	Case Variable When 'Search' Then N'" & getadminCustomSearchEditLngStr("LtxtSearch") & "' When 'ItmsGrpCod' Then N'" & getadminCustomSearchEditLngStr("DtxtGroup") & "' When 'FirmCode' Then N'" & getadminCustomSearchEditLngStr("DtxtFirm") & "' when 'Order' Then N'" & getadminCustomSearchEditLngStr("DtxtOrder") & "' " & _
 		"	When 'PriceRange' Then N'" & getadminCustomSearchEditLngStr("LtxtPriceRange") & "' When 'Inventory' Then N'" & getadminCustomSearchEditLngStr("LtxtInv") & "' When 'ItemWithImg' Then N'" & getadminCustomSearchEditLngStr("LtxtItemWithImg") & "' When 'ItemNew' Then N'" & getadminCustomSearchEditLngStr("LtxtItemNew") & "' " & _
 		"	When 'ItemProm' Then N'" & getadminCustomSearchEditLngStr("LtxtItemProm") & "' When 'WishList' Then N'" & getadminCustomSearchEditLngStr("LtxtWishList") & "' When 'CatType' Then N'" & getadminCustomSearchEditLngStr("LtxtCatalogType") & "' " & _
 		"	When 'ItemRange' Then N'" & getadminCustomSearchEditLngStr("LtxtItemRange") & "' When 'ItmsGrpRange' Then N'" & Replace(getadminCustomSearchEditLngStr("LtxtItmsGrpRange"), "'", "''") & "' When 'FirmRange' Then N'" & getadminCustomSearchEditLngStr("LtxtFirmRange") & "' When 'InvRange' Then N'" & getadminCustomSearchEditLngStr("LtxtInvRange") & "' " & _
		"	When 'CardType' Then N'" & getadminCustomSearchEditLngStr("DtxtType") & "' When 'BPRange' Then N'" & getadminCustomSearchEditLngStr("LtxtBPRange") & "' When 'BPGrpRange' Then N'" & Replace(getadminCustomSearchEditLngStr("LtxtBPGroupRange"), "'", "''") & "' " & _
		"	When 'BPCntRange' Then N'" & getadminCustomSearchEditLngStr("LtxtCountriesRange") & "' When 'BPProp' Then N'" & getadminCustomSearchEditLngStr("LtxtBPProp") & "' When 'ItmProp' Then N'" & getadminCustomSearchEditLngStr("LItmProp") & "' End " & _
 		"Else " & _ 
 		"	Case [Type] When 'T' Then N'" & getadminCustomSearchEditLngStr("LtxtText") & "' When 'Q' Then N'" & getadminCustomSearchEditLngStr("DtxtQuery") & "' When 'DD' Then '" & getadminCustomSearchEditLngStr("LtxtCmb") & "' When 'DP' Then N'" & getadminCustomSearchEditLngStr("DtxtDate") & "' When 'L' Then N'" & getadminCustomSearchEditLngStr("LtxtList") & "' When 'CL' Then N'" & getadminCustomSearchEditLngStr("LtxtChkList") & "' End " & _
  		"End [Type], " & _
 		"DataType, Ordr from OLKCustomSearchVars where ObjectCode = " & ObjID & " and ID = " & Request("ID") & " order by Ordr asc"
		rs.open sql, conn, 3, 1 %>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> <%=getadminCustomSearchEditLngStr("LttlVarsNote")%></td>
	</tr>
	<form name="frmVars" action="adminSubmit.asp" method="post" onsubmit="return valFrmVars();">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="tblRepVars">
			<tr class="TblRepTltSub">
				<td width="14"></td>
				<td width="200" class="style1"><%=getadminCustomSearchEditLngStr("DtxtName")%></td>
				<td width="120" class="style1"><%=getadminCustomSearchEditLngStr("DtxtVariable")%></td>
				<td class="style1" style="width: 240px"><%=getadminCustomSearchEditLngStr("LtxtFormat")%></td>
				<td width="120" class="style1"><%=getadminCustomSearchEditLngStr("DtxtType")%></td>
				<td class="style1" style="width: 80px"><%=getadminCustomSearchEditLngStr("DtxtOrder")%></td>
				<td></td>
			</tr>
			<% do while not rs.eof %>
			<input type="hidden" name="VarID" value="<%=rs("VarID")%>">
			<tr class="TblRepTbl">
				<td width="14"><% If rs("DataType") <> "S" Then %>
				<a href="adminCustomSearchEdit.asp?ObjID=<%=ObjID%>&ID=<%=Request("ID")%>&VarID=<%=rs("VarID")%>&#editVar"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a>
				<% End If %></td>
				<td width="200">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="TblRepTbl">
						<td><input type="text" name="varName<%=rs("VarID")%>" <% If rs("DataType") = "S" Then %>readonly<% End If %> id="varName" size="20" style="width: 100%;<% If rs("DataType") = "S" Then %>background-color: #D4D0C8;<% End If %>" value="<% If rs("DataType") <> "S" Then %><%=Server.HTMLEncode(rs("Name"))%><% Else %><%=rs("Type")%><% End If %>" onkeydown="return chkMax(event, this, 50);"></td>
						<td width="16"><% If rs("DataType") <> "S" Then %><a href="javascript:doFldTrad('CustomSearchVars', 'ObjectCode,ID,VarID', '<%=ObjID%>,<%=Request("ID")%>,<%=rs("VarID")%>', 'alterName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminCustomSearchEditLngStr("DtxtTranslate")%>" border="0"></a><% Else %>&nbsp;<% End If %></td>
					</tr>
				</table>
				</td>
				<td width="120"><nobr><span dir="ltr">@<% If rs("DataType") <> "S" Then %><%=rs("Variable")%><% Else %>SystemFilters<% End If %>&nbsp;</span></nobr></td>
				<td style="width: 240px"><%=rs("Type")%>&nbsp;<% If rs("DataType") = "S" and (rs("Variable") = "ItmProp" or rs("Variable") = "BPProp") Then %><input type="button" name="btnEditProp" value="..." onclick="editProp();"><% End If %></td>
				<td width="120"><% If rs("DataType") <> "S" Then %><%=rs("DataType")%><% Else %><%=getadminCustomSearchEditLngStr("DtxtSystem")%><% End If %></td>
				<td align="center" style="width: 80px"><table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="Ordr<%=rs("VarID")%>" name="Ordr<%=rs("VarID")%>" value="<%=rs("Ordr")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("Ordr")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="Ordr<%=rs("VarID")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="Ordr<%=rs("VarID")%>Down"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('frmVars', 'Ordr<%=rs("VarID")%>', 'Ordr<%=rs("VarID")%>Up', 'Ordr<%=rs("VarID")%>Down', 0);</script></td>
				<td>
				<a href="javascript:if(confirm('<%=getadminCustomSearchEditLngStr("LtxtConfDelVar")%>'.replace('{0}', '<%=Replace(rs("Name"), "'", "\'")%>')))window.location.href='adminSubmit.asp?submitCmd=adminCustomSearch&cmd=delVar&ObjID=<%=ObjID%>&ID=<%=Request("ID")%>&VarID=<%=rs("VarID")%>'">
			<img border="0" src="images/remove.gif"></a></td>
			</tr>
			<% rs.movenext
			loop
			rs.close %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminCustomSearchEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminCustomSearch">
	<input type="hidden" name="cmd" value="uVars">
	<input type="hidden" name="ObjID" value="<%=ObjID%>">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
</form>
	<tr class="TblRepTlt" id="editVar">
		<td>&nbsp;<% If Request("VarID") = "" Then %><%=getadminCustomSearchEditLngStr("LttlAddVar")%><% Else %><%=getadminCustomSearchEditLngStr("LttlEditVar")%><% End If %></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminCustomSearchEditLngStr("LttlVarNote")%></td>
	</tr>
	<% If Request("VarID") = "" Then %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="200" class="TblRepTlt">
					<%=getadminCustomSearchEditLngStr("LtxtVarType")%></td>
				<td class="TblRepNrm">
			    <select size="1" name="addType" onchange="javascript:doAddType(this.value);">
			    <option></option>
			    <option <% If Request("new") = "Y" Then %>selected<% End If %> value="Custom"><%=getadminCustomSearchEditLngStr("LtxtCustom")%></option>
			    <%
			    set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetCustomSearchSysAvlVars" & Session("ID")
				cmd.Parameters.Refresh()
			    cmd("@ObjID") = ObjID
			    cmd("@ID") = ID
			    set rs = Server.CreateObject("ADODB.RecordSet")
			    rs.open cmd, , 3, 1
				do while not rs.eof %><option value="<%=rs("Variable")%>"><%
				Select Case rs("Variable")
					Case "Search"
						Response.Write getadminCustomSearchEditLngStr("LtxtSearch")
					Case "ItmsGrpCod"
						Response.Write getadminCustomSearchEditLngStr("DtxtGroup")
					Case "FirmCode"
						Response.Write getadminCustomSearchEditLngStr("DtxtFirm") 
					Case "Order"
						Response.Write getadminCustomSearchEditLngStr("DtxtOrder")
					Case "PriceRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtPriceRange") 
					Case "Inventory"
						Response.Write getadminCustomSearchEditLngStr("LtxtInv")
					Case "ItemWithImg"
						Response.Write getadminCustomSearchEditLngStr("LtxtItemWithImg")
					Case "ItemNew"
						Response.Write getadminCustomSearchEditLngStr("LtxtItemNew")
					Case "ItemProm"
						Response.Write getadminCustomSearchEditLngStr("LtxtItemProm")
					Case "WishList"
						Response.Write getadminCustomSearchEditLngStr("LtxtWishList")
					Case "CatType"
						Response.Write getadminCustomSearchEditLngStr("LtxtCatalogType")
					Case "ItemRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtItemRange")
					Case "ItmsGrpRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtItmsGrpRange")
					Case "FirmRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtFirmRange")
					Case "InvRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtInvRange")
					Case "CardType"
						Response.Write getadminCustomSearchEditLngStr("DtxtType")
					Case "BPRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtBPRange")
					Case "BPGrpRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtBPGroupRange")
					Case "BPCntRange"
						Response.Write getadminCustomSearchEditLngStr("LtxtCountriesRange")
					Case "BPProp"
						Response.Write getadminCustomSearchEditLngStr("LtxtBPProp")
					Case "ItmProp"
						Response.Write getadminCustomSearchEditLngStr("LItmProp")
				End Select %></option><%
				rs.movenext
				loop
				%>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
<%	End If
If Request("VarID") <> "" or Request("new") = "Y" Then
If Request("VarID") <> "" Then
	VarID = Request("VarID")
	
	sql = "select *, " & _
	"Case When Exists(select 'A' from OLKCustomSearchVarsBase where ObjectCode = T0.ObjectCode and ID = T0.ID and VarID = T0.VarID) Then 'Y' Else 'N' End IsTarget " & _
	"from OLKCustomSearchVars T0 where ObjectCode = " & ObjID & " and ID = " & ID & " and VarID = " & VarID
	set rs = conn.execute(sql)
	varName = rs("Name")
	varVar = rs("Variable")
	vType = rs("Type")
	varDataType = rs("DataType")
	varMaxChar = rs("MaxChar")
	varDefVars = rs("DefVars")
	varNotNull = rs("NotNull")
	varDefBy = rs("DefValBy")
	varDefValue = rs("DefValValue")
	varDefDate = rs("DefValDate")
	Ordr = rs("Ordr")
	IsTarget = rs("IsTarget") = "Y"
	Select Case varDefVars
		Case "Q"
			varQuery = rs("Query")
			varQueryField = rs("QueryField")
		Case "F"
			sql = "select Value + + ',' + + Description As 'Line' from OLKCustomSearchVarsVals where ObjectCode = " & ObjID & " and ID = " & ID & " and VarID = " & VarID
			rs.close
			rs.open sql, conn, 3, 1
			do while not rs.eof
			if rs.bookmark > 1 then varQuery = varQuery & VbNewLine
			varQuery = varQuery & rs("Line")
			rs.movenext
			loop
			rs.close
	End Select
Else
	sql = "select IsNull((select Max(Ordr)+1 from OLKCustomSearchVars where ObjectCode = " & ObjID & " and ID = " & ID & "), 0) Ordr"
	set rs = conn.execute(sql)
	Ordr = rs("Ordr")
	varName = ""
	varDefBy = "N"
	IsTarget = False
End IF
%>
<form name="frmEditVar" action="adminSubmit.asp" method="POST" onsubmit="return valEditVar();">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8">
			<tr>
				<td width="97" class="TblRepTlt">
					<%=getadminCustomSearchEditLngStr("DtxtName")%></td>
				<td width="224" class="TblRepNrm">
				
    			<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="varName" size="20" value="<%=Server.HTMLEncode(varName)%>" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td width="16"><a href="javascript:doFldTrad('CustomSearchVars', 'ObjectCode,ID,VarID', '<%=ObjID%>,<%=Request("ID")%>,<%=Request("VarID")%>', 'alterName', 'T', <% If Request("VarID") <> "" Then %>null<% Else %>document.frmEditVar.varNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCustomSearchEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
    			
    			</td>
				<td width="131" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("LtxtFormat")%></td>
				<td class="TblRepNrm">
			    <select size="1" name="varType" onchange="changeType(this.value)">
				<option <% If vType = "T" Then %>selected<% End If %> value="T">
				<%=getadminCustomSearchEditLngStr("LtxtText")%></option>
				<option <% If vType = "Q" Then %>selected<% End If %> value="Q">
				<%=getadminCustomSearchEditLngStr("DtxtQuery")%></option>
				<option <% If vType = "DD" Then %>selected<% End If %> value="DD">
				<%=getadminCustomSearchEditLngStr("LtxtCmb")%></option>
				<option <% If vType = "DP" Then %>selected<% End If %> value="DP">
				<%=getadminCustomSearchEditLngStr("DtxtDate")%></option>
				<option <% If vType = "L" Then %>selected<% End If %> value="L">
				<%=getadminCustomSearchEditLngStr("LtxtList")%></option>
				<option <% If vType = "CL" Then %>selected<% End If %> value="CL">
				<%=getadminCustomSearchEditLngStr("LtxtChkList")%>
				</option>
				</select></td>
			</tr>
			<tr>
				<td width="97" class="TblRepTlt"><span dir="ltr">@<%=getadminCustomSearchEditLngStr("DtxtVariable")%></span></td>
				<td width="224" class="TblRepNrm">
			    <input type="text" name="varVar" size="20" value="<%=varVar%>" maxlength="50" onkeydown="javascript:document.frmEditVar.btnVerfyVarVar.src='images/btnValidate.gif';document.frmEditVar.btnVerfyVarVar.style.cursor = 'hand';document.frmEditVar.valVarVar.value='Y';">
				<img src="images/btnValidateDis.gif" id="btnVerfyVarVar" alt="<%=getadminCustomSearchEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditVar.valVarVar.value == 'Y')VerfyVarVar();">
				<input type="hidden" name="valVarVar" value="N">
				</td>
				<td width="131" class="TblRepTlt">
				<%=getadminCustomSearchEditLngStr("DtxtType")%></td>
				<td class="TblRepNrm">
			    <select size="1" name="varDataType" onchange="chkTypeDate(this)">
				<option <% If varDataType = "nvarchar" Then %>selected<% End If %> value="nvarchar"><%=getadminCustomSearchEditLngStr("LtxtText")%></option>
				<option <% If varDataType = "datetime" Then %>selected<% End If %> value="datetime"><%=getadminCustomSearchEditLngStr("DtxtDate")%></option>
				<option <% If varDataType = "float" or varDataType = "numeric" Then %>selected<% End If %> value="numeric"><%=getadminCustomSearchEditLngStr("DtxtNumeric")%></option>
				<option <% If varDataType = "int" Then %>selected<% End If %> value="int"><%=getadminCustomSearchEditLngStr("LtxtNumWhole")%></option>
				</select></td>
			</tr>
			<tr>
				<td class="TblRepNrm">
				&nbsp;
				</td>
				<td class="TblRepNrm">
				<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<input type="checkbox" name="varNotNull" value="Y" <% If varNotNull = "Y" or varNotNull = "" Then %>checked<% End If %> id="varNotNull" class="OptionButton" style="background:background-image" onclick="javascript:chkNotNull(this);"></td>
						<td class="TblRepNrm"><label for="varNotNull"><%=getadminCustomSearchEditLngStr("DtxtNotNull")%></label></td>
					</tr>
				</table>
				</td>
				<td class="TblRepTlt" width="131">
				<%=getadminCustomSearchEditLngStr("LtxtMaxChar")%></td>
				<td class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="varMaxChar" name="varMaxChar" value="<%=varMaxChar%>" size="6" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnVarMaxCharUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnVarMaxCharDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
			    </td>
			</tr>
			<tr class="TblRepNrm">
				<td>&nbsp;</td>
				<td>
				</td>
				<td class="TblRepTlt" width="131">
				<%=getadminCustomSearchEditLngStr("DtxtOrder")%></td>
				<td><table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="Ordr" name="Ordr" value="<%=Ordr%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=Ordr%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="OrdrUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="OrdrDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('frmEditVar', 'Ordr', 'OrdrUp', 'OrdrDown', 1);</script></td>
			</tr>
			<tr>
				<td valign="top" colspan="4">
				<table border="0" width="100%" id="table12" cellspacing="2" cellpadding="0">
					<tr>
						<td class="TblRepNrm">
				<input <% If varDefVars = "Q" or varDefVars = "" Then %>checked<% End If %> type="radio" value="Q" name="varQueryBy" id="fp1" class="OptionButton" style="background:background-image" onclick="doChangeDefVars('Q');"><label for="fp1"><%=getadminCustomSearchEditLngStr("DtxtQuery")%></label><input <% If varDefVars = "F" Then %>checked<% End If %> type="radio" name="varQueryBy" value="F" id="fp2" class="OptionButton" style="background:background-image" onclick="doChangeDefVars('F');"><label for="fp2"><%=getadminCustomSearchEditLngStr("LtxtFixVals")%> </label>(<%=getadminCustomSearchEditLngStr("LtxtValText")%>)</td>
						<td width="200" class="TblRepTlt"><%=getadminCustomSearchEditLngStr("LtxtBaseVars")%></td>
						<td class="TblRepTlt"><%=getadminCustomSearchEditLngStr("LtxtSelFld")%></td>
					</tr>
					<tr>
						<td width="440" class="TblRepNrm">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td rowspan="2">
									<textarea dir="ltr" rows="8" name="varQuery" id="varQuery" cols="83" class="input" <% If (vType <> "DD" and vType <> "L" and vType <> "Q" and vType <> "CL") or vType = "" Then %>disabled style="background-color: #CCCCCC"<% End If %> onkeydown="javascript:document.frmEditVar.btnVerfyVar.src='images/btnValidate.gif';document.frmEditVar.btnVerfyVar.style.cursor = 'hand';document.frmEditVar.valVarQuery.value='Y';"><%=myHTMLEncode(varQuery)%></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCustomSearchEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(20, 'varDef', '<%=ObjID%><%=Request("ID")%><%=Request("VarID")%>', <% If Request("VarID") <> "" Then %>null<% Else %>document.frmEditVar.varQueryDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnVerfyVar" alt="<%=getadminCustomSearchEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditVar.valVarQuery.value == 'Y')VerfyVarQuery();">
									<input type="hidden" name="valVarQuery" value="N">
								</td>
							</tr>
						</table>
    					</td>
						<td valign="top" width="200" class="TblRepNrm">
						<%
				sql = "select T0.VarID, '@' + T0.Variable Variable, [Type], T0.Name, "
				
				If Request("VarID") = "" Then 
					sql = sql & " 'N' Verfy "
				Else
					sql = sql & " Case When Exists(select 'A' from OLKCustomSearchVarsBase where ObjectCode = " & ObjID & " and ID = " & ID & " and VarID = " & Request("VarID") & " and BaseID = T0.VarID) Then 'Y' Else 'N' End Verfy, " & _
					"DataType, MaxChar "
				End If
				
				sql = sql & " " & _
						"from OLKCustomSearchVars T0 " & _
						"where ObjectCode = " & ObjID & " and ID = " & ID 
				If Request("VarID") <> "" Then sql = sql & " and T0.VarID <> " & Request("VarID")
					set rs = Server.CreateObject("ADODB.RecordSet")
					rs.open sql, conn, 3, 1
					if not rs.eof then
					rs.Filter = "Type <> 'S'"
					If Not rs.Eof Then enabledBaseVars = True %>
				<table cellpadding="0" cellspacing="0" border="0" style="border: 1px solid #4783C5">
				<% do while not rs.eof %>
					<tr class="TblRepNrm">
						<td>
							<input style="background:background-image" onclick="javascript:chkBase(this);" <% If Request("VarID") = "" or Request("VarID") <> "" and ((vType <> "Q" and vType <> "L" and vType <> "DD") or vType = "") and (varDefVars <> "Q" or varDefVars <> "") Then %>disabled<% End If %> class="OptionButton" type="checkbox" name="baseVar" id="baseVars<%=rs("VarID")%>" value="<%=rs("VarID")%>" <% If rs("Verfy") = "Y" Then %>checked<% End If %> onclick="javascript:document.frmEditVar.btnVerfyVar.disabled=false;"><label for="baseVars<%=rs("VarID")%>"><span dir="ltr"><%=rs("Variable")%></span> 
							- <%=rs("Name")%></label>&nbsp;&nbsp;
							<input type="hidden" name="BaseID" value="<%=rs("VarID")%>">
						</td>
					</tr>
					<% rs.movenext
					loop %>
				</table>
				<% Else
				enabledBaseVars = False
				End If %>
						&nbsp;</td>
						<td valign="top" class="TblRepNrm">
						<select size="1" name="varQueryField" <% If ((vType <> "Q") or vType = "") and (varDefVars <> "Q" or varDefVars <> "") Then %>disabled<% End If %> onchange="javascript:document.frmEditVar.btnVerfyVar.disabled=false;">
						<% 
						If varQuery <> "" and varDefVars = "Q" Then
							set rTest = Server.CreateObject("ADODB.RecordSet")
							testQuery = "declare @LanID int set @LanID = " & Session("LanID") & " "
							testQuery = testQuery & "declare @CardCode nvarchar(15) "
							testQuery = testQuery & "declare @SlpCode int "
							testQuery = testQuery & "declare @branch int "

							If enabledBaseVars Then
								rs.movefirst
								do while not rs.eof
									If rs("DataType") = "nvarchar" Then 
										MaxVar = "(" & rs("MaxChar") & ")"
									ElseIf rs("DataType") = "numeric" Then
										MaxVar = "(19,6)"
									Else
										MaxVar = ""
									End If
									testQuery = testQuery & "declare " & rs("Variable") & " " & rs("DataType") & " " & MaxChar & " "
								rs.movenext
								loop
							End If
							testQuery = testQuery & varQuery
							set rTest = conn.execute(testQuery)
							For each itm in rTest.Fields
							If itm.Name <> "" Then %>
							<option <% If varQueryField = itm.Name Then %>selected<% End If %> value="<%=myHTMLEncode(itm.Name)%>"><%=myHTMLEncode(itm.Name)%></option>
							<% End If
							Next
							set rTest = Nothing
						End If %>
						</select></td>
					</tr>
					<tr>
						<td>
						<table cellpadding="0" cellspacing="2" border="0" width="100%">
							<tr>
								<td class="TblRepTlt"><%=getadminCustomSearchEditLngStr("LtxtDefValue")%></td>
								<td class="TblRepNrm">
								<input type="radio" name="varDefBy" id="varDefByN" value="N" <% If varDefBy = "N" Then %>checked<% End If %> class="OptionButton" style="background:background-image" onclick="javascript:changeDefVarBy(this.value);"><label for="varDefByN"><%=getadminCustomSearchEditLngStr("LtxtNone")%></label>
								<input <% If IsTarget Then %>disabled<% End If %> type="radio" name="varDefBy" id="varDefByV" value="V" <% If varDefBy = "V" Then %>checked<% End If %> class="OptionButton" style="background:background-image" onclick="javascript:changeDefVarBy(this.value);"><label for="varDefByV"><%=getadminCustomSearchEditLngStr("DtxtValue")%></label>
								<input <% If IsTarget Then %>disabled<% End If %> type="radio" name="varDefBy" id="varDefByQ" value="Q" <% If varDefBy = "Q" Then %>checked<% End If %> class="OptionButton" style="background:background-image" onclick="javascript:changeDefVarBy(this.value);"><label for="varDefByQ"><%=getadminCustomSearchEditLngStr("DtxtQuery")%></label>
								</td>
							</tr>
						</table>
						</td>
						<td valign="top" width="200" class="TblRepNrm">
						&nbsp;</td>
						<td valign="top" class="TblRepNrm">
						&nbsp;</td>
					</tr>
					<tr class="TblRepNrm">
						<td width="440">
						<table cellpadding="0" cellpadding="0" border="0" id="tblDefValValue" style="<% If varDefBy <> "V" or varDefBy = "V" and vType = "DP" Then %>display:none;<% End If %>">
							<tr>
								<td><input type="text" name="varDefValValue" id="varDefValValue" size="20" value="<% If varDefBy = "V" and varDataType <> "datetime" and not IsNull(varDefValue) Then %><%=Server.HTMLEncode(varDefValue)%><% End If %>"></td>
							</tr>
						</table>
						<table cellpadding="0" cellpadding="0" border="0" id="tblDefValDate" style="<% If varDefBy <> "V"  or varDefBy = "V" and vType <> "DP" Then %>display:none;<% End If %>">
							<tr>
								<td><img border="0" src="images/cal.gif" id="btnDefValDateImg" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
								<td><input type="text" name="varDefValDate" id="varDefValDate" readonly size="20" onclick="btnDefValDateImg.click();" value="<% If varDefBy = "V" and varDataType = "datetime" Then %><%=FormatDate(varDefDate, False)%><% End If %>"></td>
							</tr>
						</table>
						<table cellpadding="0" cellspacing="0" border="0" width="100%" id="tblDefValQuery" style="<% If varDefBy <> "Q" Then %>display:none;<% End If %>">
							<tr>
								<td rowspan="2">
									<textarea dir="ltr" rows="4" name="varDefValQuery" cols="83" onkeydown="javascript:document.frmEditVar.btnDefValQuery.src='images/btnValidate.gif';document.frmEditVar.btnDefValQuery.style.cursor = 'hand';document.frmEditVar.valDefValQuery.value='Y';"><% If varDefBy = "Q" Then %><%=myHTMLEncode(varDefValue)%><% End If %></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCustomSearchEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(20, 'DefValue', '<%=ObjID%><%=Request("ID")%><%=Request("VarID")%>', <% If Request("VarID") <> "" Then %>null<% Else %>document.frmEditVar.varDefValueDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnDefValQuery" alt="<%=getadminCustomSearchEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditVar.valDefValQuery.value == 'Y')VerfyVarDefQuery();">
									<input type="hidden" name="valDefValQuery" value="N">
								</td>
							</tr>
						</table>
						</td>
						<td valign="top" width="200">
						&nbsp;</td>
						<td valign="top">
						&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminCustomSearchEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminCustomSearchEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminCustomSearchEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminCustomSearchEditLngStr("DtxtConfCancel")%>'))window.location.href='adminCustomSearchEdit.asp?ObjID=<%=Request("ObjID")%>&ID=<%=Request("ID")%>&#editVar'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<% If Request("VarID") = "" Then %>
	<input type="hidden" name="varNameTrad">
	<input type="hidden" name="varQueryDef">
	<input type="hidden" name="varDefValueDef">
	<% End If %>
	<input type="hidden" name="submitCmd" value="adminCustomSearch">
	<input type="hidden" name="cmd" value="editVar">
	<input type="hidden" name="ObjID" value="<%=ObjID%>">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="VarID" value="<%=Request("VarID")%>">
	</form>
	<% End If %>
<% End If %>
</table>
<% If isEdit Then %>
<script language="javascript">
//Add Variable Type
function doAddType(value)
{
	switch (value)
	{
		case 'Custom':
			window.location.href='adminCustomSearchEdit.asp?ObjID=<%=ObjID%>&ID=<%=Request("ID")%>&new=Y';
			break;
		case '':
			window.location.href='adminCustomSearchEdit.asp?ObjID=<%=ObjID%>&ID=<%=Request("ID")%>';
			break;
		default:
			window.location.href='adminSubmit.asp?submitCmd=adminCustomSearch&cmd=addSysVar&ObjID=<%=ObjID%>&ID=<%=Request("ID")%>&varType=' + value;
			break;
	}
}
//Variables List
function valFrmVars()
{
	var varName = document.frmVars.varName;
	if (varName)
	{
		if (varName.length)
		{
			for (var i = 0;i<varName.length;i++)
			{
				if (varName[i].value == '')
				{
					alert('<%=getadminCustomSearchEditLngStr("LtxtValVarNam")%>');
					varName[i].focus();
					return false;
				}
			}
		}
		else
		{
			if (varName.value == '')
			{
				alert('<%=getadminCustomSearchEditLngStr("LtxtValVarNam")%>');
				varName.focus();
				return false;
			}
		}
	}
	return true;
}
// Add / Edit Variables
<% If Request("VarID") <> "" or Request("new") = "Y" Then %>
NumUDAttachMin('frmEditVar', 'varMaxChar', 'btnVarMaxCharUp', 'btnVarMaxCharDown', 1);

Calendar.setup({
    inputField     :    "varDefValDate",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnDefValDateImg",  // trigger for the calendar (button ID)
    align          :    "Tr",           // alignment (defaults to "Bl")
    singleClick    :    true
});
function chkBase(chk)
{
	if (getDefBy() != 'N')
	{
		if (confirm('<%=getadminCustomSearchEditLngStr("LtxtChkBaseVars")%>'))
		{
			document.frmEditVar.varDefByN.checked = true;
			changeDefVarBy(getDefBy());
			document.frmEditVar.varDefByV.disabled = true;
			document.frmEditVar.varDefByQ.disabled = true;
		}
		else
		{
			chk.checked = false;
		}
	}
	else
	{
		document.frmEditVar.varDefByV.disabled = getBaseID() != '';
		document.frmEditVar.varDefByQ.disabled = getBaseID() != '';
	}
}
function changeDefVarBy(val)
{
	varType = document.frmEditVar.varType.value;
	document.getElementById('tblDefValValue').style.display = val == 'V' && varType != 'DP' ? '' : 'none';
	document.getElementById('tblDefValDate').style.display = val == 'V' && varType == 'DP' ? '' : 'none';
	document.getElementById('tblDefValQuery').style.display = val == 'Q' ? '' : 'none';
}
function VerfyVarVar()
{
	btnAfterVerfy = document.frmEditVar.btnVerfyVarVar;
	hdAfterVerfy = document.frmEditVar.valVarVar;
	document.frmVerfyQuery.type.value = 'CustSearchVar';
	document.frmVerfyQuery.ID.value = '<%=Request("ID")%>';
	document.frmVerfyQuery.Query.value = document.frmEditVar.varVar.value;
	document.frmVerfyQuery.BaseID.value = '';
	document.frmVerfyQuery.submit();
}

function getBaseID()
{

	var retVal = '';
	<% If enabledBaseVars Then %>
		if (document.frmEditVar.BaseID.length)
		{
			for (var i = 0;i<document.frmEditVar.BaseID.length;i++)
			{
				if (document.getElementById('baseVars' + document.frmEditVar.BaseID[i].value).checked)
				{
					if (retVal != '') retVal += ', ';
					retVal += document.getElementById('baseVars' + document.frmEditVar.BaseID[i].value).value;
				}
			}
		}
		else
		{
			if (document.getElementById('baseVars' + document.frmEditVar.BaseID.value).checked) retVal = document.getElementById('baseVars' + document.frmEditVar.BaseID.value).value;
		}
	<% End If %>
	return retVal;
}

function VerfyVarDefQuery()
{
	btnAfterVerfy = document.frmEditVar.btnDefValQuery;
	hdAfterVerfy = document.frmEditVar.valDefValQuery;
	document.frmVerfyQuery.type.value = 'CustVarDefVal';
	document.frmVerfyQuery.ID.value = '<%=Request("ID")%>';
	document.frmVerfyQuery.Query.value = document.frmEditVar.varDefValQuery.value;
	document.frmVerfyQuery.BaseID.value = getBaseID();
	document.frmVerfyQuery.DataType.value = document.frmEditVar.varDataType.value;
	document.frmVerfyQuery.QueryField.value = '';
	document.frmVerfyQuery.submit();
}
function VerfyVarQuery()
{
	varQueryBy = document.frmEditVar.varQueryBy;
	if (varQueryBy[0].checked)
	{
		btnAfterVerfy = document.frmEditVar.btnVerfyVar;
		hdAfterVerfy = document.frmEditVar.valVarQuery;
		document.frmVerfyQuery.type.value = 'CustSearchVarQry';
		document.frmVerfyQuery.ID.value = '<%=Request("ID")%>';
		document.frmVerfyQuery.Query.value = document.frmEditVar.varQuery.value;
		document.frmVerfyQuery.BaseID.value = getBaseID();
		document.frmVerfyQuery.DataType.value = document.frmEditVar.varDataType.value;
		document.frmVerfyQuery.QueryField.value = document.frmEditVar.varType.value == 'Q' ? document.frmEditVar.varQueryField.value : '';
		document.frmVerfyQuery.submit();
	}
	else
	{
		var myLines = document.frmEditVar.varQuery.value.split('\n');
		for (var i = 0;i<myLines.length;i++)
		{
			if (myLines[i].length == 1) {
				alert('<%=getadminCustomSearchEditLngStr("LtxtValLineNoData")%>'.replace('{0}', (i+1)));
				return; }
			else if (myLines[i].split(',').length != 2) {
				alert('<%=getadminCustomSearchEditLngStr("LtxtLineValsWrongData")%>'.replace('{0}', (i+1)));
				return; }
		}
		document.frmEditVar.btnVerfyVar.src='images/btnValidateDis.gif';
		document.frmEditVar.btnVerfyVar.style.cursor = '';
		document.frmEditVar.valVarQuery.value='N';
	}	
}
function getSqlQueryField()
{
	return document.frmEditVar.varQueryField;
}
function doChangeDefVars(DefVar)
{
	var dType = document.frmEditVar.varType.value;
	document.frmEditVar.varQueryField.disabled = !(dType == 'Q' && DefVar == 'Q');
	<% If enabledBaseVars Then %>
		if (document.frmEditVar.BaseID.length)
		{
			for (var i = 0;i<document.frmEditVar.BaseID.length;i++)
			{
				document.getElementById('baseVars' + document.frmEditVar.BaseID[i].value).disabled = !((dType == 'Q' || dType == 'L' || dType == 'DD' || dType == 'CL') && DefVar == 'Q');
			}
		}
		else
		{
			document.getElementById('baseVars' + document.frmEditVar.BaseID.value).disabled = !((dType == 'Q' || dType == 'L' || dType == 'DD' || dType == 'CL') && DefVar == 'Q');
		}
	<% End If %>
}
function valEditVar()
{
	var varDefBy = getDefBy();
	var varType = document.frmEditVar.varType.value;
	if (document.frmEditVar.varName.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValVarNam")%>');
		document.frmEditVar.varName.focus();
		return false;
	}
	else if (document.frmEditVar.varVar.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValVariable")%>');
		document.frmEditVar.varVar.focus();
		return false;
	}
	else if (document.frmEditVar.varDataType.value == 'nvarchar' && document.frmEditVar.varMaxChar.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValMaxChar")%>');
		document.frmEditVar.varMaxChar.focus();
		return false;
	}
	else if (document.frmEditVar.varQuery.value == '' && (document.frmEditVar.varType.value == 'Q' || document.frmEditVar.varType.value == 'L' ||
														document.frmEditVar.varType.value == 'DD' || document.frmEditVar.varType.value == 'CL'))
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValQryValues")%>');
		document.frmEditVar.varQuery.focus();
		return false;
	}
	else if (document.frmEditVar.valVarQuery.value == 'Y')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValQryVal")%>');
		document.frmEditVar.btnVerfyVar.focus();
		return false;
	}
	else if (document.frmEditVar.valVarVar.value == 'Y')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValVarVar")%>');
		document.frmEditVar.btnVerfyVarVar.focus();
		return false;
	}
	else if (varDefBy == 'V' && varType == 'DP' && document.frmEditVar.varDefValDate.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValDefVarDat")%>');
		document.frmEditVar.varDefValDate.focus();
		return false;
	}
	else if (varDefBy == 'V' && varType != 'DP' && document.frmEditVar.varDefValValue.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValDefVarVal")%>');
		document.frmEditVar.varDefValValue.focus();
		return false;
	}
	else if (varDefBy == 'Q' && document.frmEditVar.varDefValQuery.value == '')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValDefVarQry")%>');
		document.frmEditVar.varDefValQuery.focus();
		return false;
	}
	else if (varDefBy == 'Q' && document.frmEditVar.valDefValQuery.value == 'Y')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValDefVarVerfy")%>');
		document.frmEditVar.btnDefValQuery.focus();
		return false;
	}
	return true;
}

function changeType(dType)
{
	if (dType == 'Q' || dType == 'L' || dType == 'DD' || dType == 'CL')
	{
		document.frmEditVar.varQuery.style.backgroundColor = "";
		document.frmEditVar.varQuery.disabled = false;
		document.frmEditVar.varQueryField.disabled = !(document.frmEditVar.varQueryBy[0].checked && dType == 'Q');
		enableBranchIndex(!document.frmEditVar.varQueryBy[0].checked);
	}
	else
	{
		document.frmEditVar.varQuery.style.backgroundColor = "#CCCCCC";
		document.frmEditVar.varQuery.disabled = true;
		document.frmEditVar.varQueryField.disabled = true;
		enableBranchIndex(false);
	}
	
	if (dType != 'DP' && document.frmEditVar.varDataType.selectedIndex == 1)
	{
		document.frmEditVar.varDataType.selectedIndex = 0;
		enableVarVerfy();
	}
	else if (dType == 'DP' && document.frmEditVar.varDataType.selectedIndex != 1)
	{
		document.frmEditVar.varDataType.selectedIndex = 1;
		enableVarVerfy();
	}
	
	if (dType == 'CL') document.frmEditVar.varNotNull.checked = true;
	
	changeDefVarBy(getDefBy());
}
function chkNotNull(chk)
{
	if (!chk.checked && document.frmEditVar.varType.value == 'CL')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValCLNotNull")%>');
		chk.checked = true;
	}
}
function enableBranchIndex(enable)
{
	<% If enabledBaseVars Then %>
		if (document.frmEditVar.BaseID.length)
		{
			for (var i = 0;i<document.frmEditVar.BaseID.length;i++)
			{
				document.getElementById('baseVars' + document.frmEditVar.BaseID[i].value).disabled = !enable;
			}
		}
		else
		{
			document.getElementById('baseVars' + document.frmEditVar.BaseID.value).disabled = !enable;
		}
	<% End If %>
}
function getDefBy()
{
	var retVal = '';
	for (var i = 0;i<document.frmEditVar.varDefBy.length;i++)
	{
		if (document.frmEditVar.varDefBy[i].checked)
		{
			retVal = document.frmEditVar.varDefBy[i].value;
			break;
		}
	}
	return retVal;
}
function chkTypeDate(val)
{
	if (val.value == 'datetime' && document.frmEditVar.varType.value != 'DP')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValDatFormat")%>');
		val.selectedIndex = 0;
		return;
	}
	else if (val.value != 'datetime' && document.frmEditVar.varType.value == 'DP')
	{
		alert('<%=getadminCustomSearchEditLngStr("LtxtValTypeDatFormat")%>')
		val.selectedIndex = 1;
		return;
	}
	enableVarVerfy();
}
function enableVarVerfy()
{
	if (document.frmEditVar.varQuery.disabled || !document.frmEditVar.varQuery.disabled && document.frmEditVar.varQuery.value == '')
	{
		document.frmEditVar.btnVerfyVar.src='images/btnValidateDis.gif';
		document.frmEditVar.btnVerfyVar.style.cursor = '';
		document.frmEditVar.valVarQuery.value = 'N';
	}
	else
	{
		document.frmEditVar.btnVerfyVar.src='images/btnValidate.gif';
		document.frmEditVar.btnVerfyVar.style.cursor = 'hand';
		document.frmEditVar.valVarQuery.value = 'Y';
	}
	
	if (document.frmEditVar.varDefValQuery.disabled || !document.frmEditVar.varDefValQuery.disabled && document.frmEditVar.varDefValQuery.value == '')
	{
		document.frmEditVar.btnDefValQuery.src='images/btnValidateDis.gif';
		document.frmEditVar.btnDefValQuery.style.cursor = '';
		document.frmEditVar.valDefValQuery.value = 'N';
	}
	else
	{
		document.frmEditVar.btnDefValQuery.src='images/btnValidate.gif';
		document.frmEditVar.btnDefValQuery.style.cursor = 'hand';
		document.frmEditVar.valDefValQuery.value = 'Y';
	}
}
<% End If %>
</script>
<% End If %>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="type" value="CustSearchQry">
	<input type="hidden" name="ObjID" value="<%=ObjID%>">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="BaseID" value="">
	<input type="hidden" name="parent" value="Y">
	<input type="hidden" name="DataType" value="">
	<input type="hidden" name="QueryField" value="">
</form>
<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
<!--#include file="bottom.asp" -->