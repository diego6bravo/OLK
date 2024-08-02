<!--#include file="top.asp" -->
<!--#include file="lang/adminOpsEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<script type="text/javascript">
var txtReadOnly = '<%=getadminOpsEditLngStr("DtxtReadOnly")%>';
var txtWriteRegular = '<%=getadminOpsEditLngStr("LtxtWriteRegular")%>';
var txtCheckBox = '<%=getadminOpsEditLngStr("LtxtCheckBox")%>';
var txtVendorSelector = '<%=getadminOpsEditLngStr("LtxtVendorSelector")%>';
var valOpName = '<%=getadminOpsEditLngStr("LvalOpName")%>';
var valOpFilter = '<%=getadminOpsEditLngStr("LvalOpFilter")%>';
var valFldDesc = '<%=getadminOpsEditLngStr("LvalFldDesc")%>';
var valFld = '<%=getadminOpsEditLngStr("LvalFld")%>';
var valFldQry = '<%=getadminOpsEditLngStr("LvalFldQry")%>';
</script>
<script type="text/javascript" src="adminOps.js"></script>
<script type="text/javascript" src="js_up_down.js"></script>
<% 
ID = CLng(Request("ID"))
set rs = Server.CreateObject("ADODB.RecordSet")

If ID >= 0 Then
	sql = "select * from OLKOps where ID = " & ID
	set rs = conn.execute(sql)
	opName = rs("Name")
	GroupID = rs("GroupID")
	ObjectID = rs("ObjectID")
	Operation = rs("Operation")
	TargetObjectID = rs("TrgtObjID")
	opStatus = rs("Status") = "A"
	GenNewDoc = rs("GenNewDoc") = "Y"
	'opFilter = rs("Filter")
	
	If IsNull(TargetObjectID) Then TargetObjectID = -1
	
Else
	GroupID = -1
	TargetObjectID = -1
End If
%>
<form name="frmOps" method="post" action="adminOpsSubmit.asp" onsubmit="return valFrm();">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<% If ID = -1 Then %><%=getadminOpsEditLngStr("LttlNewOpt")%><% Else %><%=getadminOpsEditLngStr("LttlEditOpt")%><% End If %></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminOpsEditLngStr("LttlOptNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminOpsEditLngStr("DtxtName")%></td>
				<td colspan="7" class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="opName" size="100" value="<%=opName%>" max="100"></td>
						<td><a href="javascript:doFldTrad('Ops', 'ID', '<%=ID%>', 'alterName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminOpsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						<td>&nbsp;</td>
						<td class="TblRepNrm"><input type="checkbox" name="Status" class="noborder" value="Y" id="Status" <% If opStatus Then %>checked<% End if %>><label for="Status"><%=getadminOpsEditLngStr("DtxtActive")%></label></td>
					</tr>
				</table>
			    </td>
			</tr>
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminOpsEditLngStr("DtxtGroup")%></td>
				<% set rd = server.createobject("ADODB.RecordSet")
				sql = "select ID, Name from OLKOpsGrps order by 2 asc"
				set rd = conn.execute(sql)%>
				<td colspan="7" class="TblRepNrm"><select name="GroupID" size="1">
				<% do while not rd.eof %>
				<option <% If CLng(GroupID) = CLng(rd("ID")) Then %>selected<% End If %> value="<%=rd("ID")%>"><%=myHTMLEncode(rd("Name"))%></option>
				<% rd.movenext
				loop %>
				</select></td>
			</tr>
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminOpsEditLngStr("LtxtOperation")%></td>
				<td colspan="7" class="TblRepNrm"><select name="Operation" onchange="changeOp(this.value);" size="1" <% If ID <> - 1 Then %>disabled<% End If %>>
				<% If 1 = 2 Then %><option <% If CLng(Operation) = 0 Then %>selected<% End If %> value="0"><%=getadminOpsEditLngStr("DtxtReadOnly")%></option><% End If %>
				<option <% If CLng(Operation) = 1 Then %>selected<% End If %> value="1"><%=getadminOpsEditLngStr("LtxtUpdOnly")%></option>
				<option <% If CLng(Operation) = 6 Then %>selected<% End If %> value="6"><%=getadminOpsEditLngStr("LtxtMultUpdOnly")%></option>
				<% If 1 = 2 THen %><option <% If CLng(Operation) = 2 Then %>selected<% End If %> value="2"><%=getadminOpsEditLngStr("LtxtConvToObj")%></option><% End If %>
				<% If 1 = 2 Then %><option <% If CLng(Operation) = 3 Then %>selected<% End If %> value="3"><%=getadminOpsEditLngStr("LtxtGenNewDoc")%></option><% End If %>
				<% If 1 = 2 Then %><option <% If CLng(Operation) = 4 Then %>selected<% End If %> value="4">|L:txtGenJrnl|</option><% End If %>
				<% If 1 = 2 Then %><option <% If CLng(Operation) = 5 Then %>selected<% End If %> value="5"><%=getadminOpsEditLngStr("LtxtDraftAproval")%></option><% End If %>
				</select></td>
			</tr>
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminOpsEditLngStr("LtxtSourceObj")%></td>
				<td colspan="7" class="TblRepNrm"><select name="ObjectID" size="1" <% If ID <> - 1 Then %>disabled<% End If %>>
				<% If 1 = 2 Then %>
				<optgroup label="<%=getadminOpsEditLngStr("LtxtGeneral")%>">
					<option <% If CLng(ObjectID) = 2 Then %>selected<% End If %> value="2"><%=getadminOpsEditLngStr("DtxtClient")%></option>
					<option <% If CLng(ObjectID) = 4 Then %>selected<% End If %> value="4"><%=getadminOpsEditLngStr("DtxtItem")%></option>
				</optgroup>
				<% End If %>
				<optgroup label="<%=getadminOpsEditLngStr("LtxtSale")%>">
					<option <% If CLng(ObjectID) = 23 Then %>selected<% End If %> value="23"><%=getadminOpsEditLngStr("DtxtQuote")%></option>
					<option <% If CLng(ObjectID) = 17 Then %>selected<% End If %> value="17"><%=getadminOpsEditLngStr("DtxtSalesOrder")%></option>
					<option <% If CLng(ObjectID) = 15 Then %>selected<% End If %> value="15"><%=getadminOpsEditLngStr("DtxtDelivery")%></option>
					<option <% If CLng(ObjectID) = 16 Then %>selected<% End If %> value="16"><%=getadminOpsEditLngStr("DtxtReturn")%></option>
					<option <% If CLng(ObjectID) = 13 Then %>selected<% End If %> value="13"><%=getadminOpsEditLngStr("DtxtInvoice")%></option>
					<option <% If CLng(ObjectID) = -13 Then %>selected<% End If %> value="-13"><%=getadminOpsEditLngStr("DtxtInvoice")%> - <%=getadminOpsEditLngStr("DtxtReserved")%></option>
					<option <% If CLng(ObjectID) = 14 Then %>selected<% End If %> value="14"><%=getadminOpsEditLngStr("DtxtCredNote")%></option>
					<option <% If CLng(ObjectID) = 203 Then %>selected<% End If %> value="203"><%=getadminOpsEditLngStr("DtxtDownPayReq")%></option>
					<option <% If CLng(ObjectID) = 204 Then %>selected<% End If %> value="204"><%=getadminOpsEditLngStr("DtxtDownPayInv")%></option>
				</optgroup>
				<optgroup label="<%=getadminOpsEditLngStr("LtxtPur")%>">
					<option <% If CLng(ObjectID) = 540000006 Then %>selected<% End If %> value="540000006"><%=getadminOpsEditLngStr("DtxtPurQuote")%></option>
					<option <% If CLng(ObjectID) = 22 Then %>selected<% End If %> value="22"><%=getadminOpsEditLngStr("DtxtPurOrder")%></option>
					<option <% If CLng(ObjectID) = 20 Then %>selected<% End If %> value="20"><%=getadminOpsEditLngStr("DtxtGoodRecPO")%></option>
					<option <% If CLng(ObjectID) = 21 Then %>selected<% End If %> value="21"><%=getadminOpsEditLngStr("DtxtPurReturn")%></option>
					<option <% If CLng(ObjectID) = 18 Then %>selected<% End If %> value="18"><%=getadminOpsEditLngStr("DtxtPurInv")%></option>
					<option <% If CLng(ObjectID) = 19 Then %>selected<% End If %> value="19"><%=getadminOpsEditLngStr("DtxtCredMemPO")%></option>
				</optgroup>
				</select></td>
			</tr>
			<tr>
				<td class="TblRepTlt" style="width: 100px;vertical-align : top; padding-top: 2px;">
				<%=getadminOpsEditLngStr("LtxtFilter")%></td>
				<td colspan="7" class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td rowspan="2">
							<textarea rows="10" name="Filter" dir="ltr" cols="87" class="input" onkeydown="javascript:document.frmOps.btnVerfyFilter.src='images/btnValidate.gif';document.frmOps.btnVerfyFilter.style.cursor = 'hand';;document.frmOps.valFilter.value='Y';"><%=opFilter%></textarea>
						</td>
						<td valign="top">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminOpsEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(24, 'Filter', <%=ID%>, null);">
						</td>
					</tr>
					<tr>
						<td valign="bottom">
							<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminOpsEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmOps.valFilter.value == 'Y')VerfyFilter();">
							<input type="hidden" name="valFilter" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr id="trGenNewDoc" <% If not (CLng(Operation) = 6) Then %>style="display:none;"<% End If %>>
				<td style="width: 100px" class="TblRepTlt">
				&nbsp;</td>
				<td colspan="7" class="TblRepNrm"><input type="checkbox" class="noborder" <% If GenNewDoc Then %>checked<% End If %> name="chkGenNewDoc" id="chkGenNewDoc" value="Y" <% If ID <> -1 Then %>disabled<% End If %>><label for="chkGenNewDoc"><%=getadminOpsEditLngStr("LtxtGenNewDoc")%></label></td>
			</tr>
			<tr id="trTargetObj" <% If not (CLng(Operation) = 2 or CLng(Operation) = 3 or CLng(Operation) = 6) Then %>style="display:none;"<% End If %>>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminOpsEditLngStr("LtxtTargetObj")%></td>
				<td colspan="7" class="TblRepNrm"><select name="TargetObjectID" size="1" <% If ID <> - 1 Then %>disabled<% End If %>>
				<optgroup label="Sales">
					<option <% If CLng(TargetObjectID) = 23 Then %>selected<% End If %> value="23"><%=getadminOpsEditLngStr("DtxtQuote")%></option>
					<option <% If CLng(TargetObjectID) = 17 Then %>selected<% End If %> value="17"><%=getadminOpsEditLngStr("DtxtSalesOrder")%></option>
					<option <% If CLng(TargetObjectID) = 15 Then %>selected<% End If %> value="15"><%=getadminOpsEditLngStr("DtxtDelivery")%></option>
					<option <% If CLng(TargetObjectID) = 16 Then %>selected<% End If %> value="16"><%=getadminOpsEditLngStr("DtxtReturn")%></option>
					<option <% If CLng(TargetObjectID) = 13 Then %>selected<% End If %> value="13"><%=getadminOpsEditLngStr("DtxtInvoice")%></option>
					<option <% If CLng(TargetObjectID) = -13 Then %>selected<% End If %> value="-13"><%=getadminOpsEditLngStr("DtxtInvoice")%> - <%=getadminOpsEditLngStr("DtxtReserved")%></option>
					<option <% If CLng(TargetObjectID) = 14 Then %>selected<% End If %> value="14"><%=getadminOpsEditLngStr("DtxtCredNote")%></option>
					<option <% If CLng(TargetObjectID) = 203 Then %>selected<% End If %> value="203"><%=getadminOpsEditLngStr("DtxtDownPayReq")%></option>
					<option <% If CLng(TargetObjectID) = 204 Then %>selected<% End If %> value="204"><%=getadminOpsEditLngStr("DtxtDownPayInv")%></option>
				</optgroup>
				<optgroup label="Purchase">
					<option <% If CLng(TargetObjectID) = 540000006 Then %>selected<% End If %> value="540000006"><%=getadminOpsEditLngStr("DtxtPurQuote")%></option>
					<option <% If CLng(TargetObjectID) = 22 Then %>selected<% End If %> value="22"><%=getadminOpsEditLngStr("DtxtPurOrder")%></option>
					<option <% If CLng(TargetObjectID) = 20 Then %>selected<% End If %> value="20"><%=getadminOpsEditLngStr("DtxtGoodRecPO")%></option>
					<option <% If CLng(TargetObjectID) = 21 Then %>selected<% End If %> value="21"><%=getadminOpsEditLngStr("DtxtPurReturn")%></option>
					<option <% If CLng(TargetObjectID) = 18 Then %>selected<% End If %> value="18"><%=getadminOpsEditLngStr("DtxtPurInv")%></option>
					<option <% If CLng(TargetObjectID) = 19 Then %>selected<% End If %> value="19"><%=getadminOpsEditLngStr("DtxtCredMemPO")%></option>
				</optgroup>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminOpsEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminOpsEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminOpsEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminOpsEditLngStr("DtxtConfCancel")%>'))window.location.href='adminOps.asp';"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="editOP">
<input type="hidden" name="ID" value="<%=ID%>">
</form>
<% If ID <> -1 Then %>
<form name="frmOpsDet" method="post" action="adminOpsSubmit.asp" onsubmit="return valFrmDet();">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminOpsEditLngStr("LtxtOpDet")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminOpsEditLngStr("LtxtOpDetNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<%
			
			cmbType = -1
			TableID = ""
			If Request("Type") <> "" Then cmbType = CInt(Request("Type"))
			Select Case CLng(ObjectID)
				Case 2
				Case 4
				Case 30
				Case 33
				Case 23, 17, 15, 16, 13, -13, 14, 203, 204, 540000006, 22, 20, 21, 18, 19
					Select Case cmbType
						Case 0, 2
							TableID = "OINV"
						Case 1
							TableID = "INV1"
					End Select
			End Select
			If Operation = 6 Then 
				cmbType = 0
				TableID = ""
			Else %>
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminOpsEditLngStr("LtxtSection")%></td>
				<td colspan="7" class="TblRepNrm">
				<select id="cmbType" onchange="window.location.href='adminOpsEdit.asp?ID=<%=ID%>&Type=' + this.value;">
				<option></option>
				<% 
				Select Case CLng(ObjectID)
					Case 2
					Case 4
					Case 30
					Case 33
					Case 23, 17, 15, 16, 13, -13, 14, 203, 204, 540000006, 22, 20, 21, 18, 19
						Select Case cmbType
							Case 0, 2
								TableID = "OINV"
							Case 1
								TableID = "INV1"
						End Select %>
					<option <% If cmbType = 0 Then %>selected<% End If %> value="0"><%=getadminOpsEditLngStr("LtxtHeader")%></option>
					<option <% If cmbType = 1 Then %>selected<% End If %> value="1"><%=getadminOpsEditLngStr("LtxtLines")%></option>
					<option <% If cmbType = 2 Then %>selected<% End If %> value="2"><%=getadminOpsEditLngStr("LtxtFooter")%></option>
					<%
				End Select %>
				</select>
			    </td>
			</tr><% End If %>
		</table>
		<%
		If cmbType <> -1 Then %>
			<table border="0" cellpadding="0" width="100%">
				<tr class="TblRepTlt">
					<td style="width: 16px;">&nbsp;</td>
					<td><%=getadminOpsEditLngStr("DtxtType")%></td>
					<td><%=getadminOpsEditLngStr("DtxtField")%>/<%=getadminOpsEditLngStr("DtxtQuery")%></td>
					<td><%=getadminOpsEditLngStr("DtxtDescription")%></td>
					<td><%=getadminOpsEditLngStr("LtxtStyle")%></td>
					<td <% If TableID = "INV1" or Operation = 6 Then %>style="display: none;"<% End If %>><%=getadminOpsEditLngStr("DtxtCol")%></td>
					<td><%=getadminOpsEditLngStr("DtxtOrder")%></td>
					<td style="width: 16px;">&nbsp;</td>
				</tr>
				<% 
				set rs = Server.CreateObject("ADODB.RecordSet")
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetOpsLinesData" & Session("ID")
				cmd.Parameters.Refresh
				cmd("@LanID") = Session("LanID")
				cmd("@ID") = ID
				cmd("@TypeID") = cmbType
				cmd("@TableID") = TableID
				cmd("@Operation") = Operation
				rs.open cmd, , 3, 1
				do while not rs.eof
				If Not IsNull(rs("alterType")) Then cmbType = CInt(rs("alterType")) %>
				<tr class="TblRepNrm">
					<td style="width: 16px;"><% If rs("StyleID") = 4 Then %>
					<a href='adminOpsEdit.asp?ID=<%=ID%>&amp;Type=<%=cmbType%>&amp;LineID=<%=rs("LineID")%>'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a><% End If %></td>
					<td><% If Left(rs("AliasID"), 2) = "U_" Then %><%=getadminOpsEditLngStr("DtxtUDF")%><% ElseIf rs("StyleID") = 4 Then %><%=getadminOpsEditLngStr("DtxtCustomized")%><% Else %><%=getadminOpsEditLngStr("DtxtSystem")%><% End If %><%
					If Operation = 6 Then
						Select Case rs("alterType")
							Case 0 
								Response.Write " - " & getadminOpsEditLngStr("LtxtHeader") 
							Case 1
								Response.Write " - " & getadminOpsEditLngStr("LtxtLines") 
						End Select
					End If %></td>
					<td><img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=myHTMLEncode(rs("AliasID"))%>"></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
						<tr class="TblRepNrm">
							<td><% If rs("StyleID") <> 4 Then %><%=rs("AliasDesc")%><% Else %><input id="aliasDesc<%=rs("LineID")%>_<%=rs("alterType")%>" name="aliasDesc<%=rs("LineID")%>_<%=rs("alterType")%>" value="<%=Server.HTMLEncode(rs("AliasDesc"))%>" size="50"><% End If %></td>
							<td style="width: 16px;"><a href="javascript:doFldTrad('OpsLines', 'ID,TypeID,LineID', '<%=ID%>,<%=cmbType%>,<%=rs("LineID")%>', 'AlterAliasDesc', 'T', null);"><img src="images/trad.gif" alt="<%=getadminOpsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>
					<td><% If rs("StyleID") = 4 Then %><%=getadminOpsEditLngStr("DtxtCustomized")%><input type="hidden" name="styleID<%=rs("LineID")%>_<%=rs("alterType")%>" value="4"><%
					Else %><select name="styleID<%=rs("LineID")%>_<%=rs("alterType")%>">
					<option value="0"><%=getadminOpsEditLngStr("DtxtReadOnly")%></option>
					<% If CLng(Operation) <> 5 Then %>
					<option <% If rs("StyleID") = 1 Then %>selected<% End If %> value="1"><%=getadminOpsEditLngStr("LtxtWriteRegular")%></option>
					<% If rs("EnableCheckBox") = "Y" Then %><option <% If rs("StyleID") = 2 Then %>selected<% End If %> value="2"><%=getadminOpsEditLngStr("LtxtCheckBox")%></option><% End If %>
					<% If rs("TypeID") = "A" Then %><option <% If rs("StyleID") = 3 Then %>selected<% End If %> value="3"><%=getadminOpsEditLngStr("LtxtVendorSelector")%></option><% End If %>
					<% End If %>
					</select>
					<% End If %></td>
					<td <% If TableID = "INV1" or Operation = 6 Then %>style="display: none;"<% End If %>>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<input id="colID<%=rs("LineID")%>_<%=rs("alterType")%>" name="colID<%=rs("LineID")%>_<%=rs("alterType")%>" value="<%=rs("ColID")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("ColID")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnColUp<%=rs("LineID")%>_<%=rs("alterType")%>"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnColDown<%=rs("LineID")%>_<%=rs("alterType")%>"></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<input id="orderID<%=rs("LineID")%>" name="orderID<%=rs("LineID")%>_<%=rs("alterType")%>" value="<%=rs("Ordr")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("Ordr")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnOrdrUp<%=rs("LineID")%>_<%=rs("alterType")%>"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnOrdrDown<%=rs("LineID")%>_<%=rs("alterType")%>"></td>
								</tr>
							</table>
							</td>
						</tr>
					</table></td>
					<td style="width: 16px">
					<a href="javascript:if(confirm('<%=getadminOpsEditLngStr("LtxtConfDelLine")%>'))window.location.href='adminOpsSubmit.asp?cmd=delLine&ID=<%=Request("ID")%>&Type=<%=cmbType%>&LineID=<%=rs("LineID")%>'">
					<img border="0" src="images/remove.gif" width="16" height="16"></a>
					<input type="hidden" name="LineID" value="<%=rs("LineID")%>_<%=rs("alterType")%>">
					<script type="text/javascript">NumUDAttachMin('frmOpsDet', 'colID<%=rs("LineID")%>_<%=rs("alterType")%>', 'btnColUp<%=rs("LineID")%>_<%=rs("alterType")%>', 'btnColDown<%=rs("LineID")%>_<%=rs("alterType")%>', 0);
					NumUDAttachMin('frmOpsDet', 'orderID<%=rs("LineID")%>_<%=rs("alterType")%>', 'btnOrdrUp<%=rs("LineID")%>_<%=rs("alterType")%>', 'btnOrdrDown<%=rs("LineID")%>_<%=rs("alterType")%>', 0);</script></td>
				</tr>
				<% rs.movenext
				loop %>
			</table>
			<%
		End If %>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminOpsEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="editOPDet">
<input type="hidden" name="ID" value="<%=ID%>">
<input type="hidden" name="TypeID" value="<%=cmbType%>">
</form>
<% If cmbType <> -1 Then
If Request("LineID") <> "" Then
	sql = "select AliasID, AliasDesc, StyleID, ColID, Ordr from OLKOpsLines where ID = " & ID & " and TypeID = " & cmbType & " and LineID = " & Request("LineID")
	set rs = conn.execute(sql)
	AliasID = rs("AliasID")
	AliasDesc = rs("AliasDesc")
	StyleID = rs("StyleID")
	ColID = rs("ColID")
	Ordr = rs("Ordr")
Else
	sql = "select IsNull(Max(ColID), 0) ColID from OLKOpsLines T0 where ID = " & ID & " and TypeID = " & cmbType
	set rs = conn.execute(sql)
	If Not rs.Eof Then
		ColID = rs("ColID")
		sql = "select IsNull(Max(Ordr)+1, 0) Ordr from OLKOpsLines where ID = " & ID & " and TypeID = " & cmbType
		set rs = conn.execute(sql)
		Ordr = rs("Ordr")
	Else
		ColID = 0
		Ordr = 0
	End If
End If %>
<form name="frmOpsFld" method="post" action="adminOpsSubmit.asp" onsubmit="return valFrmFld();">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<% If Request("LineID") = "" Then %><%=getadminOpsEditLngStr("LtxtAddFld")%><% Else %><%=getadminOpsEditLngStr("LtxtEditFld")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table cellpadding="0" cellspacing="0" border="0">
			<tr class="TblRepTlt">
				<td><%=getadminOpsEditLngStr("DtxtField")%></td>
				<td><%=getadminOpsEditLngStr("DtxtDescription")%></td>
				<td><%=getadminOpsEditLngStr("LtxtStyle")%></td>
				<td <% If TableID = "INV1" or Operation = 6 Then %>style="display: none;"<% End If %>><%=getadminOpsEditLngStr("DtxtCol")%></td>
				<td><%=getadminOpsEditLngStr("DtxtOrder")%></td>
			</tr>
			<tr class="TblRepNrm">
				<td><select name="fldID" onchange="changeFld(this.value);" <% If Request("LineID") <> "" Then %>disabled<% End If %>>
				<option></option><%=getadminOpsEditLngStr("LtxtHeader")%><%
				If Request("LineID") = "" Then
					If Operation <> 6 Then
						ShowFields TableID, "", ""
					Else
						ShowFields "OINV", getadminOpsEditLngStr("LtxtHeader") & " - ", 0
						ShowFields "INV1", getadminOpsEditLngStr("LtxtLines") & " - ", 1
					End If
				End If
				
				Sub ShowFields(ByVal strTable, ByVal strAddDesc, ByVal alterType)
					If alterType = "" Then alterType = cmbType
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetOpsLinesAvlFld" & Session("ID")
					cmd.Parameters.Refresh
					cmd("@LanID") = Session("LanID")
					cmd("@ID") = ID
					cmd("@TypeID") = alterType
					cmd("@TableID") = strTable
					cmd("@LawsSet") = myApp.LawsSet
					set rs = cmd.execute()
					do while not rs.eof
					%><option value="<%=rs("FieldID")%>{S}<%=rs("EnableCheckBox")%>{S}<%=rs("TypeID")%>{S}<%=rs("AliasID")%>{S}<%=alterType%>"><% 
					If rs("FieldID") < 0 Then %><%=getadminOpsEditLngStr("DtxtSystem")%><% Else %><%=getadminOpsEditLngStr("DtxtUDF")%><% End If %> - <%=strAddDesc%><%=rs("Name")%></option><% rs.movenext
					loop
				End Sub %>
				<option <% If Request("LineID") <> "" Then %>selected<% End If %> value="Custom"><%=getadminOpsEditLngStr("DtxtCustomized")%></option>
				</select></td>
				<td><input name="AliasDesc" value="<%=myHTMLEncode(AliasDesc)%>" id="AliasDesc" <% If Request("LineID") = "" Then %>disabled="disabled" style="border-color: #848284;"<% End If %> size="50" type="text">&nbsp;</td>
				<td><select name="StyleID" id="StyleID" disabled="disabled">
				<option value="0"><%=getadminOpsEditLngStr("DtxtReadOnly")%></option>
				<% If CLng(Operation) <> 5 Then %>
				<option value="1"><%=getadminOpsEditLngStr("LtxtWriteRegular")%></option>
				<% End If %>
				</select></td>
				<td align="center" <% If TableID = "INV1" or Operation = 6 Then %>style="display: none;"<% End If %>>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="colID" name="colID" value="<%=ColID%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=ColID%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnColUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnColDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
				<td align="center">
				<table>
					<tr>
						<td>
						<input id="orderID" name="orderID" value="<%=Ordr%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=orderID%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnOrdrUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnOrdrDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table><script type="text/javascript">NumUDAttachMin('frmOpsFld', 'colID', 'btnColUp', 'btnColDown', 0);
					NumUDAttachMin('frmOpsFld', 'orderID', 'btnOrdrUp', 'btnOrdrDown', 0);</script></td>
			</tr>
			<tr class="TblRepNrm" id="trFldQry"<% If Request("LineID") = "" Then %> style="display: none;"<% End If %>>
				<td colspan="5">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td rowspan="2">
							<textarea rows="10" name="AliasID" dir="ltr" cols="150" class="input" onkeydown="javascript:document.frmOpsFld.btnVerfyFld.src='images/btnValidate.gif';document.frmOpsFld.btnVerfyFld.style.cursor = 'hand';;document.frmOpsFld.valFld.value='Y';"><%=AliasID%></textarea>
						</td>
						<td valign="top">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFld" alt="<%=getadminOpsEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(24, 'Fld', '<%=ID%>,<%=cmbType%>,<%=Request("LineID")%>', null);">
						</td>
					</tr>
					<tr>
						<td valign="bottom">
							<img src="images/btnValidateDis.gif" id="btnVerfyFld" alt="<%=getadminOpsEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmOpsFld.valFld.value == 'Y')VerfyFld();">	
							<input type="hidden" name="valFld" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" align="center">
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77" <% If Request("LineID") = "" Then %>style="display: none;"<% End If %> id="tdApply">
				<input type="submit" value="<%=getadminOpsEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminOpsEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminOpsEditLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:window.location.href='adminOpsEdit.asp?ID=<%=ID%>&Type=<%=cmbType%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="editOPFldDet">
<input type="hidden" name="ID" value="<%=ID%>">
<input type="hidden" name="TypeID" value="<%=cmbType%>">
<input type="hidden" name="LineID" value="<%=Request("LineID")%>">
</form>
<% End If %>
<% End If %>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="TypeID" value="<%=cmbType%>">
	<input type="hidden" name="Operation" value="">
	<input type="hidden" name="ObjectID" value="">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>

<!--#include file="bottom.asp" -->