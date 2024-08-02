<% addLngPathStr = "addCard/" %>
<!--#include file="lang/addCard.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<% 
set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SysID") = "OCRD"
cmd("@ObsID") = "TCRD"
set rs = cmd.execute()
If rs(0) = "Y" Then Response.Redirect "configErr.asp?errCmd=Card"

set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCrdData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("CrdRetVal")
set rs = cmd.execute()
EnableSDK = rs("EnableSDK") = "Y"
Confirm = myAut.GetCardProperty(rs("CardType"), "C")

Select Case rs("CardType")
	Case "S"
		GrpCardType = "S"
	Case "C", "L"
		GrpCardType = "C"
End Select

set rsdf = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetUDFSystemCols" & Session("ID")
cmd.Parameters.Refresh
cmd("@LanID") = Session("LanID")
cmd("@UserType") = userType
cmd("@TableID") = "OCRD"
cmd("@OP") = "O"
rsdf.open cmd, , 3, 1

isUpdate = rs("Command") = "U" %>
<script language="javascript">
var DfVendTerm = <%=rs("DfVendTerm")%>;
var DfCustTerm = <%=rs("DfCustTerm")%>;
var selDes = '<%=SelDes%>';
var dbName = '<%=Session("olkdb")%>';
var DtxtAdd = '<%=getaddCardLngStr("DtxtAdd")%>';
var DtxtConfirm = '<%=getaddCardLngStr("DtxtConfirm")%>';
var ConfirmC = <%=LCase(myAut.GetCardProperty("C", "C"))%>;
var ConfirmS = <%=LCase(myAut.GetCardProperty("S", "C"))%>;
var ConfirmL = <%=LCase(myAut.GetCardProperty("L", "C"))%>;
var lawsSet = '<%=myApp.LawsSet%>';
var txtErrSaveData = '<%=getaddCardLngStr("DtxtErrSaveData")%>';
var txtValCod = '<%=getaddCardLngStr("LtxtValCod")%>';
var txtSelGrp = '<%=getaddCardLngStr("LtxtSelGrp")%>';
var txtValRFC = '<%=getaddCardLngStr("LtxtValRFC")%>';
var txtValRFCLen = '<%=getaddCardLngStr("LtxtValRFCLen")%>';
function valUDF()
{
	<% If EnableSDK Then 
	cmd.CommandText = "DBOLKGetUDFNotNull" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = "V"
	cmd("@TableID") = "OCRD"
	cmd("@OP") = "O"
	set rd = cmd.execute()
	do while not rd.eof %>
	if (document.frmAddCard.U_<%=rd("AliasID")%>.value == "")
	{
		alert('<%=getaddCardLngStr("LtxtConfFld")%>'.replace('{0}', '<%=Replace(rd("Descr"), "'", "\'")%>'));
		showUDF(<%=rd("GroupID")%>);
		document.frmAddCard.U_<%=rd("AliasID")%>.focus();
		return false;
	}
	<% rd.movenext
	loop 
	End If %>
	return true;
}
</script>
<script type="text/javascript" src="addCard/addCard.js"></script>
<form method="POST" action="agentClientSubmit.asp" name="frmAddCard" onsubmit="return valFrm();">
<%
CardType = rs("CardType")
%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td colspan="2"><% If Not isUpdate Then %><%=getaddCardLngStr("LttlNewClient")%><% Else %><%=getaddCardLngStr("LttlEditClient")%><% End If %></td>
	</tr>
	<tr>
		<td colspan="2">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddCardLngStr("DtxtCode")%></td>
				<td <% If InStr("MX, CR, GT, US, CA", myApp.LawsSet) = 0 Then %>colspan="2"<% End If %>>
				<table cellpadding="0" cellspacing="2" border="0">
					<tr><% If isUpdate or Not isUpdate and not myApp.AutoGenOCRD Then CardCodeValue = myHTMLEncode(rs("CardCode")) Else CardCodeValue = getaddCardLngStr("DtxtAutomatic") %>
						<td><input type="text" <% If isUpdate or not isUpdate and myApp.AutoGenOCRD Then %>readonly class="inputDis"<% End If %> name="CardCode" size="40" value="<%=CardCodeValue%>" style="<% If rs("VerfyCardCode") = "Y" and not myApp.AutoGenOCRD Then %>background-color: #FFD2A6; <% End if %>" onkeydown="return chkMax(event, this, 15);" maxlength="15" onchange="doProc('CardCode', 'S', this.value);<% If Not isUpdate and not myApp.AutoGenOCRD Then %>chkBP(this.value);<% End If %>"></td>
						<td style="padding-right: 2px; padding-left: 2px;"><img src="images/icon_alert.gif" alt="<%=getaddCardLngStr("DtxtCodeExists")%>" id="dvCodeErr" style="<% If rs("VerfyCardCode") = "N" Then %>display: none; <% End if %>"></td>
						<td>
						<select size="1" name="CardType" onchange="javascript:changeCardType(this.value);">
				      	<% If myAut.HasAuthorization(45) and (not isUpdate or isUpdate and (CardType = "C" or CardType = "L")) Then %><option <% If CardType = "C" Then %>selected<% End If %> value="C"><% If 1 = 2 Then %><%=getaddCardLngStr("DtxtClient")%><% Else %><%=myHTMLEncode(txtClient)%><% End If %></option><% End If %>
						<% If myAut.HasAuthorization(78) and (not isUpdate or isUpdate and CardType = "S") Then %><option <% If CardType = "S" Then %>selected<% End If %> value="S">
						<%=getaddCardLngStr("DtxtSupplier")%></option><% End If %>
						<% If myAut.HasAuthorization(77) and (not isUpdate or isUpdate and (CardType = "C" or CardType = "L")) Then %><option <% If CardType = "L" Then %>selected<% End If %> value="L">
						<%=getaddCardLngStr("DtxtLead")%></option><% End If %>
						</select>
						</td>
					</tr>
				</table>
				<input type="hidden" name="prevCardType" id="prevCardType" value="<%=CardType%>"></td>
				<% If InStr("MX, CR, GT, US, CA", myApp.LawsSet) > 0 Then %>
				<td class="GeneralTblBold2" width="41%"><%=getaddCardLngStr("DtxtType")%>
					<select size="1" name="CmpPrivate" onchange="changeCmpPrivate(this.value);">
				    <option value="C"><%=getaddCardLngStr("DtxtCmp")%></option>
					<option <% If rs("CmpPrivate") = "I" Then %>selected<% End If %> value="I">
					<%=getaddCardLngStr("LtxtNatPer")%></option>
				    </select></td>
				<% Else %>
				<input type="hidden" name="CmpPrivate" value="C">
				<% End If %>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddCardLngStr("DtxtName")%></td>
				<td colspan="2"><input type="text" name="CardName" size="84" value="<%=myHTMLEncode(rs("CardName"))%>" onkeydown="return chkMax(event, this, 100);" maxlength="100" onchange="doProc('CardName', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddCardLngStr("DtxtGroup")%></td>
				<td colspan="2"><select size="1" name="GroupCode" onchange="doProc('GroupCode', 'N', this.value);">
				<% 
				cmd.CommandText = "DBOLKGetCrdGroups" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@CardType") = GrpCardType
				set rd = cmd.execute()
				do While NOT rd.EOF %>
				<option <% If CStr(rd(0)) = CStr(rs("GroupCode")) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
			</tr>
			<% rsdf.Filter = "FieldID = -1"
			If Not rsdf.Eof Then %>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><% 
				Select Case myApp.LawsSet
					Case "PA", "IL", "US", "CA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "ZA" %><%=getaddCardLngStr("DtxtLicTradNum")%><% 
					Case "MX", "CR", "GT" %>RFC<% 
					Case "GB" %><%=getaddCardLngStr("DtxtVatNum")%><%
					Case "CL" %>RUT<% 
				End Select
				MaxLength = 32
				If myApp.LawsSet = "MX" Then
					Select Case rs("CmpPrivate")
						Case "I"
							MaxLength = 13
						Case Else
							MaxLength = 12
					End Select
				End If %></td>
				<td colspan="2"><input type="text" name="LicTradNum" id="LicTradNum" size="84" value="<%=myHTMLEncode(rs("LicTradNum"))%>" onkeydown="return chkMax(event, this, <%=MaxLength%>);" maxlength="<%=MaxLength%>"  onchange="doProc('LicTradNum', 'S', this.value);"></td>
			</tr>
			<% End If %>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddCardLngStr("DtxtAgent")%></td>
				<td colspan="2"><select size="1" name="SlpCode" onchange="doProc('SlpCode', 'N', this.value);">
				<%
				cmd.CommandText = "DBOLKGetAgents" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = cmd.execute()
                do while not rd.eof %>
                <option value="<%=rd("SlpCode")%>" <% If rd("SlpCode") = -1 and IsNull(rs("SlpCode")) or CInt(rs("SlpCode")) = CInt(rd("SlpCode")) Then %>selected<% End If %>><%=myHTMLEncode(rd("SlpName"))%></option>
                <% rd.movenext
                loop %>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td colspan="2">
		<p align="center"><%=getaddCardLngStr("DtxtAddData")%></td>
	</tr>
	<tr>
		<td colspan="2">
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getaddCardLngStr("LtxtPhone1")%></td>
				<td width="25%"><input type="text" name="Phone1" size="35" value="<%=myHTMLEncode(rs("Phone1"))%>" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProc('Phone1', 'S', this.value);"></td>
				<td width="6%" class="GeneralTblBold2"><%=getaddCardLngStr("LtxtCelular")%></td>
				<td width="57%"><input type="text" name="Cellular" size="35" value="<%=myHTMLEncode(rs("Cellular"))%>" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProc('Cellular', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getaddCardLngStr("LtxtPhone2")%></td>
				<td width="25%"><input type="text" name="Phone2" size="35" value="<%=myHTMLEncode(rs("Phone2"))%>" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProc('Phone2', 'S', this.value);"></td>
				<td width="6%" class="GeneralTblBold2"><%=getaddCardLngStr("LtxtFax")%></td>
				<td width="57%"><input type="text" name="Fax" size="35" value="<%=myHTMLEncode(rs("Fax"))%>" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProc('Fax', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getaddCardLngStr("DtxtEMail")%></td>
				<td colspan="3"><input type="text" name="E_Mail" size="69" value="<%=myHTMLEncode(rs("E_Mail"))%>" onkeydown="return chkMax(event, this, 100);" maxlength="100" onchange="doProc('E_Mail', 'S', this.value);"></td>
			</tr>
			<% rsdf.Filter = "FieldID = -2"
			If Not rsdf.Eof Then %>
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getaddCardLngStr("LtxtPymntCod")%></td>
				<td width="25%"><select size="1" class="input" name="cmbGroupNum" id="cmbGroupNum" style="font-size:10px; font-family:Verdana; Width:100%" onchange="doProc('GroupNum', 'N', this.value);">
			    <% 
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetPaymentGroups" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
			    set rd = cmd.execute()
			    do while not rd.eof %>
		        <option value="<%=rd("GroupNum")%>" <% If rd("GroupNum") = rs("GroupNum") then response.write "selected" %>><%=myHTMLEncode(rd("PymntGroup"))%></option>
		        <% rd.movenext
		        loop %>
		        </select></td>
				<td width="6%" class="GeneralTblBold2">&nbsp;</td>
				<td width="57%">&nbsp;</td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
		<table border="0" cellpadding="0" width="100%" id="table6">
			<tr class="GeneralTbl">
				<td valign="top" width="200">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8">
					<tr class="generalTblBold2">
						<td><%=getaddCardLngStr("DtxtImage")%></td>
					</tr>
					<tr>
						<td>
						<table border="0" cellpadding="2" cellpadding="0" id="table12" cellspacing="0">
							<tr>
								<td>
				                <% If rs("Picture") <> "" Then Picture = rs("Picture") Else Picture = "pcard.gif" %>
				                <img id="CardImg" src="pic.aspx?filename=<%=Picture%>&MaxSize=223&dbName=<%=Session("olkdb")%>" border="1" name="ItemImg"></td>
								<td valign="bottom"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddCard.Picture.value = ''; document.frmAddCard.ItemImg.src='pic.aspx?filename=pcard.gif&MaxSize=223&dbName=<%=Session("olkdb")%>';doProc('Picture', 'S', '');" style="cursor: hand">
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<p align="center">
								<input type="button" value="<%=getaddCardLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmAddCard.Picture,document.frmAddCard.CardImg, 223);">
								</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
				<td valign="top">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7">
					<tr class="generalTblBold2">
						<td><%=getaddCardLngStr("DtxtObservations")%></td>
					</tr>
					<tr>
						<td><textarea rows="4" name="Notes" cols="74" onkeydown="return chkMax(event, this, 100);" maxlength="100" onchange="doProc('Notes', 'S', this.value);"><%=myHTMLEncode(rs("Notes"))%></textarea></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<% If EnableSDK Then
	
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCRD"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rg = cmd.execute()

	set rSdk = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCRD"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rSdk.open cmd, , 3, 1

	set rd = Server.CreateObject("ADODB.RecordSet")
	
	do while not rg.eof
	If CInt(rg("GroupID")) < 0 Then GroupID = "_1" Else GroupID = rg("GroupID")
	 %>
	<tr class="GeneralTblBold2">
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2" style="cursor: hand; " onclick="showHideSection(tdShowUDF<%=GroupID%>, trUDF<%=GroupID%>);">
				<td align="center"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getaddCardLngStr("DtxtUDF")%><%
				Case Else
					Response.Write rg("GroupName")
				End Select %></td>
				<td width="20" id="tdShowUDF<%=GroupID%>" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">[+]</td>
			</tr>
		</table>
		</td>
	</tr>
      <tr id="trUDF<%=GroupID%>" style="display: none; ">
        <td width="100%">
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
			<% 
			arrPos = Split("I,D", ",")
			For i = 0 to 1
			rSdk.Filter = "GroupID = " & rg("GroupID") & " and Pos = '" & arrPos(i) & "'"
			If not rSdk.eof then %>
				<td width="50%" valign="top">
			        <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
			        <% do while not rSdk.eof
			        ShowAddCardUFD()
			        rSdk.movenext
			        loop
			        rSdk.movefirst %>
			        </table>
				</td>
			<% End If
			Next %>
			</tr>
		</table>
		</td>
      </tr>
      <% rg.movenext
      loop
      End If %>
	<tr class="GeneralTbl" align="center">
		<td colspan="2">
		<p align="center">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td>
					  <input type="button" <% If rs("VerfyCardCode") = "Y" Then %>disabled<% End If %> value="<% If Not isUpdate Then %><% If Not Confirm Then %><%=getaddCardLngStr("DtxtAdd")%><% Else %><%=getaddCardLngStr("DtxtConfirm")%><% End If %><% Else %><%=getaddCardLngStr("DtxtSave")%><% End If %>" name="btnAdd" onclick="if(valFrm()) { setCardFlow(<%=Session("CrdRetVal")%>);doFlowAlert(); }"></td>
					<td>
					  <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					  <input type="button" value="<%=getaddCardLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%If Not isUpdate Then %><%=getaddCardLngStr("LtxtConfCancel")%><% Else %><%=getaddCardLngStr("LtxtConfNoSave")%><% End If %>'))window.location.href='cardCancel.asp?isUpdate=<%=JBool(isUpdate)%>'"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="newCardSubmit">
<input type="hidden" name="Picture" value="<%=rs("Picture")%>">
<input type="hidden" name="Confirm" value="">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="isUpdate" value="<%=isUpdate%>">
<input type="hidden" name="doSubmit" value="Y">
<input type="hidden" name="doSubmitAdd" value="Y">
</form>
<script  type="text/javascript">
function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('<%=getaddCardLngStr("DtxtValFldMaxChar")%>'.replace('{0}', FSize));
				Field.value = Field.value.subString(0, FSize);
			}
			break;
		case 'N':
			switch (EditType)
			{
				case '':
					if (Field.value != '')
					{
						if (!MyIsNumeric(getNumericVB(Field.value)))
						{
							Field.value = '';
							alert('<%=getaddCardLngStr("DtxtValNumVal")%>');
						}
						else if (parseInt(getNumericVB(Field.value)) < 1)
						{
							Field.value = '';
							alert('<%=getaddCardLngStr("DtxtValNumMinVal")%>'.replace('{0}', '1'));
						}
						else if (parseInt(getNumericVB(Field.value)) > 2147483647)
						{
							alert('<%=getaddCardLngStr("DtxtValNumMaxVal")%>'.replace('{0}', '2147483647'));
							Field.value = 2147483647;
						}
						else if (Field.value.indexOf('<%=GetFormatDec%>') > -1)
						{
							Field.value = '';
							alert('<%=getaddCardLngStr("DtxtValNumValWhole")%>');
						}
					}
					break;
			}
			break;
		case 'B':
			if (Field.value != '')
			{
				if (!MyIsNumeric(getNumericVB(Field.value)))
				{
					Field.value = '';
					alert('<%=getaddCardLngStr("DtxtValNumVal")%>');
				}
				else
				{
					if (parseFloat(getNumericVB(Field.value)) > 1000000000000)
					{
						Field.value = 999999999999;
					}
					else if (parseFloat(getNumericVB(Field.value)) < -1000000000000)
					{
						Field.value = -999999999999;
					}
					
					switch (EditType)
					{
						case 'R':
							Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)), <%=myApp.RateDec%>);
							break;
						case 'S':
							Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)), <%=myApp.SumDec%>);
							break;
						case 'P':
							Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)), <%=myApp.PriceDec%>);
							break;
						case 'Q':
							Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)), <%=myApp.QtyDec%>);
							break;
						case '%':
							Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)), <%=myApp.PercentDec%>);
							break;
						case 'M':
							Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)), <%=myApp.MeasureDec%>);
							break;
					}
				}
			}
			break;
	}
}
</script>
<% Sub ShowAddCardUFD()
	InsertID = rSdk("InsertID")
	FldVal = rs(InsertID)
	Select Case rSdk("TypeID")
		Case "B", "N"
			ProcType = "N"
		Case "M", "A"
			ProcType = "S"
		Case "D"
			ProcType = "D"
	End Select
			 %>
				<tr class="generalTbl">
			            <td bgcolor="#EAF5FF" class="GeneralTblBold2" width="25%">
			              <table border="0" cellpadding="0" cellspacing="0" width="100%">
			                <tr>
			            	  <td>
			            	    <b><font size="1" face="Verdana"><%=rSdk("Descr")%><% If rSdk("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b>
			            	  </td>
			            	    <% If (rSdk("Query") = "Y" or rSdk("TypeID") = "D") and IsNull(rSdk("RTable")) Then %>
			            	    <td width="16">
			            	    	<img border="0" src="images/<% If rSdk("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rSdk("AliasID")%>" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Card&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',400,250,'yes', 'yes', document.frmAddCard.U_<%=rSdk("AliasID")%>, '<%=ProcType%>')"<% End If %>>
			            	    </td>
			            	    <% End If %>
			            	</tr>
			              </table>
			            </td>
			            <td dir="ltr" bgcolor="#EAF5FF" width="75%"><% If rSdk("DropDown") = "Y" or not IsNull(rSdk("RTable")) then 
			            	set rd = Server.CreateObject("ADODB.RecordSet")
							cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@TableID") = "OCRD"
							cmd("@FieldID") = rSdk("FieldID")
							rd.open cmd, , 3, 1
							 %><select size="1" name="U_<%=rSdk("AliasID")%>" class="input" style="width: 99%" onchange="doProc(this.name, '<%=ProcType%>', this.value);">
								<option></option>
								<% do while not rd.eof %>
								<option <% If Not IsNull(rs(InsertID)) Then If CStr(rs(InsertID)) = CStr(rd(0)) Then Response.Write "Selected" %> value="<%=rd(0)%>" <% If rSdk("Dflt")= rd(0) Then %>selected<% End If %>><%=myHTMLEncode(rd(1))%></option>
								<% rd.movenext
								loop
								rd.close %>
							</select>
					<% ElseIf rSdk("TypeID") = "M" and Trim(rSdk("EditType")) = "" or rSdk("TypeID") = "A" and rSdk("EditType") = "?" Then %>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
						<table width="100%" cellspacing="0" cellpadding="0">
						  <tr>
						    <td>
						<% End If %>
						<textarea <% If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" class="input" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>);doProc(this.name, '<%=ProcType%>', this.value);" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Card&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> rows="3" onfocus="this.select()" style="width: 100%" cols="1"><% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %></textarea>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							</td>
							<td width="16">
								<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddCard.U_<%=rSdk("AliasID")%>.value = '';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
					<% ElseIf rSdk("TypeID") = "A" and rSdk("EditType") = "I" Then %>
						<table cellpadding="2" cellspacing="0" border="0">
							<tr>
								<td><img src="pic.aspx?filename=<% If IsNull(rs(InsertID)) Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rSdk("AliasID")%>" border="1">
								<input type="hidden" name="U_<%=rSdk("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
								<td width="16" valign="bottom"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddCard.U_<%=rSdk("AliasID")%>.value = '';document.frmAddCard.imgU_<%=rSdk("AliasID")%>.src='pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';" style="cursor: hand"></td>
							</tr>
							<tr>
								<td colspan="2" height="22">
								<p align="center">
								<input type="button" value="<%=getaddCardLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmAddCard.U_<%=rSdk("AliasID")%>, document.frmAddCard.imgU_<%=rSdk("AliasID")%>,180);"></td>
							</tr>
						</table>
						<% Else
						If Not IsNull(rs(InsertID)) Then 
							If rSdk("TypeID") = "B" Then
				        	Select Case rSdk("EditType")
								Case "R"
									FldVal = FormatNumber(CDbl(FldVal),myApp.RateDec)
								Case "S"
									FldVal = FormatNumber(CDbl(FldVal),myApp.SumDec)
								Case "P"
									FldVal = FormatNumber(CDbl(FldVal),myApp.PriceDec)
								Case "Q"
									FldVal = FormatNumber(CDbl(FldVal),myApp.QtyDec)
								Case "%"
									FldVal = FormatNumber(CDbl(FldVal),myApp.PercentDec)
								Case "M"
									FldVal = FormatNumber(CDbl(FldVal),myApp.MeasureDec)
				        	End Select
				        	End If
						Else
							FldVal = ""
						End If %>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							<table width="100%" cellspacing="0" cellpadding="0">
							  <tr>
							    <td>
							<% End If %>
							<% 
							If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then readOnly = True Else readOnly = False
							If rSdk("TypeID") = "D" Then FldVal = FormatDate(FldVal, False)
							If rSdk("TypeID") = "A" Then fldSize = 43 Else fldSize = 12
							If rSdk("TypeID") = "B" or rSdk("TypeID") = "A" Then
								If rSdk("TypeID") = "B" Then MaxSize = 21 Else MaxSize = rSdk("SizeID")
								isMaxSize = True
							Else
								isMaxSize = False
							End If %>
							<input <% If readOnly Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" id="U_<%=rSdk("AliasID")%>" size="<%=fldSize%>" class="input" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click();"<% End If %> onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>);doProc(this.name, '<%=ProcType%>', this.value);" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click()"<% End If %> <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Card&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> value="<% If Not IsNull(FldVal) Then %><%=FldVal%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %>" <% If rSdk("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> style="width: 100%" <% If isMaxSize Then %> onkeydown="return chkMax(event, this, <%=MaxSize%>);" maxlength="<%=MaxSize%>"<% End if %>>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
								</td>
								<td width="16">
									<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddCard.U_<%=rSdk("AliasID")%>.value = '';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');">
								</td>
							  </tr>
							</table>
							<% End If %>
						<% End If %>
			            </td>
			          </tr>
<% End Sub %>
<script language="javascript">
<% 
If EnableSDK Then
	rSdk.Filter = "TypeID = 'D'"
	If rSdk.recordcount > 0 Then rSdk.movefirst
	do while not rSdk.eof %>
	    Calendar.setup({
	        inputField     :    "U_<%=rSdk("AliasID")%>",     // id of the input field
	        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
	        button         :    "btn<%=rSdk("AliasID")%>",  // trigger for the calendar (button ID)
	        align          :    "Bl",           // alignment (defaults to "Bl")
	        singleClick    :    true
	    });
	<% rSdk.movenext
	loop
End If %>
</script>
