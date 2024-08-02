<% addLngPathStr = "addItem/" %>
<!--#include file="lang/addItem.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script type="text/javascript" src="Controls/ctlLock.js"></script>
<% 

set myLock = new clsLock

ItmRetVal = CLng(Session("ItmRetVal"))

set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SysID") = "OITM"
cmd("@ObsID") = "TITM"
set rs = cmd.execute()
If rs(0) = "Y" Then Response.Redirect "configErr.asp?errCmd=Item"

cmd.CommandText = "DBOLKCheckUDFHasVals" & Session("ID")
cmd.Parameters.Refresh()
cmd("@TableID") = "OITM"
set rs = cmd.execute()
hasUDFTable = rs(0) = "Y"

set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetItmData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = ItmRetVal
set rs = cmd.execute()
EnableSDK = rs("EnableSDK") = "Y"
isUpdate = rs("Command") = "U"

Confirm = myAut.GetObjectProperty(4, "C")
%>
<script language="javascript">
var selDec = '<%=SelDes%>';
var dbName = '<%=Session("olkdb")%>';
var txtErrSaveData = '<%=getaddItemLngStr("DtxtErrSaveData")%>';
var txtAlternative = '<%=getaddItemLngStr("LtxtAlternative")%>';
var txtDefault = '<%=getaddItemLngStr("DtxtDefault")%>';
var txtConfRemComp = '<%=getaddItemLngStr("LtxtConfRemComp")%>';
var txtError = '<%=getaddItemLngStr("DtxtError")%>';
var itmRetVal = <%=ItmRetVal%>;
</script>
<script type="text/javascript" src="addItem/addItem.js"></script>
<form method="POST" action="agentItemSubmit.asp" name="frmAddItem" onsubmit="return valFrm();">
<p align="center">
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><% If Not isUpdate Then %><%=getaddItemLngStr("LttlNewItm")%><% Else %><%=getaddItemLngStr("LttlEditItm")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="GeneralTbl">
				<td width="15%" class="GeneralTblBold2"><%=getaddItemLngStr("DtxtCode")%></td>
				<td width="29%"><% If isUpdate or not isUpdate and not myApp.AutoGenOITM Then ItemCodeValue = myHTMLEncode(rs("ItemCode")) Else ItemCodeValue = getaddItemLngStr("DtxtAutomatic") %>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="ItemCode" size="36" value="<%=ItemCodeValue%>" <% If isUpdate or not isUpdate and myApp.AutoGenOITM Then %> readonly class="inputDis" <% End If %> style="<% If rs("VerfyItemCode") = "Y" and not myApp.AutoGenOITM Then %>background-color: #FFD2A6; <% End if %>" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProcItem(this.value);"></td>
						<td style="padding-right: 2px; padding-left: 2px;"><img src="images/icon_alert.gif" alt="<%=getaddItemLngStr("DtxtCodeExists")%>" id="dvCodeErr" style="<% If rs("VerfyItemCode") = "N" Then %>display: none; <% End if %>"></td>
					</tr>
				</table>
				</td>
				<td width="54%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
            <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" id="table6">
              <tr>
                <td>
                <input type="checkbox"class="noborder" name="PrchseItem" id="chkPrchseItem" <% If rs("VirtualCombo") = "Y" Then %>disabled<% End If %> value="Y"<% If rs("PrchseItem") = "Y" Then %> checked <% End If %> onclick="doProc('PrchseItem', 'S', GetYesNo(this.checked));"></td>
                <td><b><font size="1" face="Verdana"><label for="chkPrchseItem">
				<%=getaddItemLngStr("LtxtPurItem")%></label></font></b></td>
                <td>
                <input type="checkbox"class="noborder" name="SellItem" id="chkSellItem" <% If rs("EnableCombo") = "Y" Then %>disabled<% End If %> value="Y"<% If rs("SellItem") = "Y" Then %> checked <% End If %> onclick="doProc('SellItem', 'S', GetYesNo(this.checked));"></td>
                <td><b><font size="1" face="Verdana"><label for="chkSellItem">
				<%=getaddItemLngStr("LtxtSalItem")%></label></font></b></td>
                <td>
                <input type="checkbox"class="noborder" name="InvntItem" id="chkInvntItem" <% If rs("VirtualCombo") = "Y" Then %>disabled<% End If %> value="Y"<% If rs("InvntItem") = "Y" Then %> checked <% End If %> onclick="doProc('InvntItem', 'S', GetYesNo(this.checked));"></td>
                <td><b><font size="1" face="Verdana"><label for="chkInvntItem">
				<%=getaddItemLngStr("LtxtInvItem")%></label></font></b></td>
				<% If myApp.EnableCombos Then %>
                <td>
                <input type="checkbox"class="noborder" name="OlkCombo" id="chkOlkCombo" value="Y" <% If rs("VirtualCombo") = "Y" Then %>disabled<% End If %> <% If rs("EnableCombo") = "Y" Then %> checked <% End If %> onclick="enableCombos(this.checked, true);"></td>
                <td><b><font size="1" face="Verdana"><label for="chkOlkCombo">
				<%=getaddItemLngStr("LtxtOlkCmb")%></label></font></b></td>
                <td>
                <input type="checkbox"class="noborder" name="OlkVirtual" id="chkOlkVirtual" value="Y" <% If rs("VirtualCombo") = "Y" Then %> checked <% End If %> onclick="enableCombosVirtual(this.checked);"></td>
                <td><b><font size="1" face="Verdana"><label for="chkOlkVirtual">
				<%=getaddItemLngStr("LtxtVirtualCmb")%></label></font></b></td><% End If %>
              </tr>
            </table>
            	</td>
			</tr>
			<tr class="GeneralTbl">
				<td width="15%" class="GeneralTblBold2"><%=getaddItemLngStr("LtxtDesc1")%></td>
				<td colspan="2"><input type="text" name="ItemName" size="61" value="<%=myHTMLEncode(rs("ItemName"))%>" onkeydown="return chkMax(event, this, 100);" style="width: 100%" maxlength="100" onchange="doProc('ItemName', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td width="15%" class="GeneralTblBold2"><%=getaddItemLngStr("LtxtDesc2")%></td>
				<td colspan="2"><input type="text" name="FrgnName" size="61" value="<%=myHTMLEncode(rs("FrgnName"))%>" onkeydown="return chkMax(event, this, 100);" style="width: 100%" maxlength="100" onchange="doProc('FrgnName', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td width="15%" class="GeneralTblBold2"><%=Server.HTMLEncode(txtAlterGrp)%></td>
				<td width="29%">
            <select size="1" name="ItmsGrpCod" onchange="doProc('ItmsGrpCod', 'N', this.value);">
            <option value="100" selected>- <%=getaddItemLngStr("LtxtNoGrp")%> -</option>
			<% set rGroup = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetItemGroups" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			rGroup.open cmd, , 3, 1
			do while not rGroup.eof	%>
			<option <% If rs("ItmsGrpCod") = rGroup("Codigo") Then %>selected<% End If %> value="<%=rGroup("Codigo")%>"><%=myHTMLEncode(rGroup("Name"))%></option>
			<% rGroup.movenext
			loop %>            
			</select></td>
				<td width="54%">
            <table border="0" cellspacing="0" width="300" id="table13">
				<tr class="GeneralTbl">
					<td class="GeneralTblBold2"><%=Server.HTMLEncode(txtAlterFrm)%></td>
					<td>
            <select size="1" name="FirmCode" onchange="doProc('FirmCode', 'N', this.value);">
			<%
			set rFirm = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetItemFirms" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			rFirm.open cmd, , 3, 1
			do while not rFirm.eof
			%>
			<option <% If rs("FirmCode") = rFirm("Code") Then %>selected<% End If %> value="<%=rFirm("Code")%>"><%=myHTMLEncode(rFirm("Name"))%></option>
			<% rFirm.movenext
			loop %>            
			</select></td>
				</tr>
			</table>
				</td>
			</tr>
			<tr class="GeneralTbl">
				<td width="15%" style="height: 24px" class="GeneralTblBold2"><%=getaddItemLngStr("LtxtBarCod")%></td>
				<td width="83%" colspan="2" style="height: 24px">
				<input type="text" name="CodeBars" size="36" value="<%=myHTMLEncode(rs("CodeBars"))%>" onkeydown="return chkMax(event, this, 16);" maxlength="16" onchange="doProc('CodeBars', 'S', this.value);"></td>
			</tr>
			</table>
		</td>
	</tr>
	<% If myApp.EnableCombos Then %>
	<tr class="GeneralTblBold2" id="ttlCmbData" <% If rs("EnableCombo") = "N" Then %>style="display: none;"<% End If %>>
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2">
				<td align="center"><%=getaddItemLngStr("LtxtCmbData")%></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id="trCmbData" <% If rs("EnableCombo") = "N" Then %>style="display: none;"<% End If %>>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2">&nbsp;</td>
				<td><table cellpadding="0" cellspacing="0">
				<tr class="GeneralTbl">
					<td><input type="checkbox" name="chkCmbShowComp"class="noborder" id="chkCmbShowComp" <% If rs("VirtualCombo") = "N" Then %>disabled<% End If %> value="Y" <% If rs("ShowComp") = "Y" Then %>checked<% End If %> onclick="doProcShowComp(this.checked);"></td>
					<td><font size="1"><label for="chkCmbShowComp"><%=getaddItemLngStr("LtxtCmbShowComp")%></label></font></td>
				</tr>
				</table>
				</td>
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtCmbManCstBy")%></td>
				<td><select size="1" name="cmbManCostPrc" id="cmbManCostPrc" <% If rs("VirtualCombo") = "N" Then %>disabled<% End If %> onchange="doProcCmb('ManCostPrc', 'S', this.value);">
				<option value="F"><%=getaddItemLngStr("DtxtFather")%></option>
				<option <% IF rs("ManCostPrc") = "C" Then %>selected<% End If %> value="C"><%=getaddItemLngStr("DtxtComponent")%></option>
				</select></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtFatherPrice")%></td>
				<td><table cellpadding="0" cellspacing="0">
				<tr class="GeneralTbl">
					<td><input type="checkbox" name="chkFatherShowPrice" class="noborder" id="chkFatherShowPrice" value="Y" <% If rs("VirtualCombo") = "N" or rs("VirtualCombo") = "Y" and rs("ShowComp") = "N" Then %>disabled<% End If %> <% If rs("ShowFatherPrice") = "Y" Then %> checked<% End If %> onclick="doProcShowPrice('Father', this.checked);"></td>
					<td><font size="1"><label for="chkFatherShowPrice"><%=getaddItemLngStr("DtxtShow")%></label></font></td>
					<td><input type="checkbox" name="chkAllowChangeFatherPrice"class="noborder" id="chkAllowChangeFatherPrice" value="Y" <% If rs("AllowChangeFatherPrice") = "Y" Then %>checked <% End If %> <% If rs("ShowFatherPrice") = "N" Then %> disabled<% End If %> onclick="doProcCmb('AllowChangeFatherPrice', 'S', GetYesNo(this.checked));"></td>
					<td><font size="1"><label for="chkAllowChangeFatherPrice"><%=getaddItemLngStr("DtxtAllowChange")%></label></font></td>
				</tr>
				</table>
				</td>
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtCostPriceList")%></td>
				<td><select size="1" name="cmbCostPrc" onchange="doProcCmb('CostPrc', 'I', this.value);">
				<option value=""><%=getaddItemLngStr("LtxtSystemCost")%></option>
				<option <% If rs("CostPrc") = -1 Then %>selected<% End If %> value="-1"><%=getaddItemLngStr("DtxtLastPurPrice")%></option>
				<option <% If rs("CostPrc") = -2 Then %>selected<% End If %> value="-2"><%=getaddItemLngStr("DtxtLastEvalPrice")%></option>
				<% 
				cmd.CommandText = "DBOLKGetPriceList" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = Server.CreateObject("ADODB.RecordSet")
				rd.open cmd, , 3, 1
				do while not rd.eof %>
				<option <% If rs("CostPrc") = rd(0) Then %>selected<% End If %> value="<%=rd(0)%>"><%=Server.HTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtCompPrice")%></td>
				<td><table cellpadding="0" cellspacing="0">
				<tr class="GeneralTbl">
					<td><input type="checkbox" name="chkShowCompPrice"class="noborder" id="chkShowCompPrice" value="Y" <% If rs("VirtualCombo") = "N" Then %>disabled<% End If %> <% If rs("ShowCompPrice") = "Y" Then %> checked<% End If %> onclick="doProcShowPrice('Comp', this.checked);"></td>
					<td><font size="1"><label for="chkShowCompPrice"><%=getaddItemLngStr("DtxtShow")%></label></font></td>
					<td><input type="checkbox" name="chkAllowChangeCompPrice"class="noborder" id="chkAllowChangeCompPrice" value="Y" <% If rs("AllowChangeCompPrice") = "Y" Then %>checked <% End If %> <% If rs("ShowCompPrice") = "N" Then %> disabled<% End If %> onclick="doProcCmb('AllowChangeCompPrice', 'S', GetYesNo(this.checked));"></td>
					<td><font size="1"><label for="chkAllowChangeCompPrice"><%=getaddItemLngStr("DtxtAllowChange")%></label></font></td>
				</tr>
				</table>
				</td>
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtSalePriceList")%></td>
				<td><select size="1" name="cmbSalePrc" onchange="doProcCmb('SalePrc', 'I', this.value);">
				<option value=""><%=getaddItemLngStr("LtxtClientPList")%></option>
				<% rd.movefirst
				do while not rd.eof %>
				<option <% If rs("SalePrc") = rd(0) Then %>selected<% End If %> value="<%=rd(0)%>"><%=Server.HTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2" id="ttlCmbComps" <% If rs("EnableCombo") = "N" Then %>style="display: none;"<% End If %>>
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2">
				<td align="center"><%=getaddItemLngStr("LtxtCmbComps")%></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id="trCmbCompsData" <% If rs("EnableCombo") = "N" Then %>style="display: none;"<% End If %>>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTblBold2" style="text-align: center;">
				<td><%=getaddItemLngStr("DtxtCode")%></td>
				<td><%=getaddItemLngStr("DtxtDescription")%></td>
				<td><%=getaddItemLngStr("LtxtMandatory")%></td>
				<td><%=getaddItemLngStr("DtxtQty")%></td>
				<td style="display: none;"><%=getaddItemLngStr("DtxtLines")%></td>
				<td><%=getaddItemLngStr("DtxtWarehouse")%></td>
				<td><%=getaddItemLngStr("LtxtCostPriceList")%></td>
				<td><%=getaddItemLngStr("LtxtSalePriceList")%></td>
				<td style="width: 80px;"><%=getaddItemLngStr("LtxtDiscPrcnt")%></td>
				<td width="16">&nbsp;</td>
			</tr>
			<tBody id="tbCmbData">
			<% 
			If rs("VerfyComp") = "Y" Then
			set rc = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetItmCombosComp" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = ItmRetVal
			rc.open cmd, , 3, 1
			
			set rw = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			rw.open cmd, , 3, 1
			
			do while not rc.eof
			LineID = rc("LineID") %>
			<tr class="GeneralTbl" id="trComp<%=LineID%>"><input type="hidden" name="CompLineID" value="<%=LineID%>">
				<td><table cellpadding="0" cellspacing="0">
				<tr>
					<td><input type="text" id="txtCompItem<%=LineID%>" value="<%=Server.HTMLEncode(rc("ItemCode"))%>" size="20" maxlength="20" onchange="javascript:getItem(<%=LineID%>, this, document.getElementById('txtCompDesc<%=LineID%>'));" onfocus="this.select();">
					</td>
					<td>
					<%
					myLock.ID = "Itm" & LineID
					myLock.Value = rc("LockItm") = "Y"
					myLock.OnClick = "doProcCompLock(" & LineID & ", 'Itm');"
					myLock.GenerateLock
					%>
					</td>
				</tr>
				</table></td>
				<td><input type="text" id="txtCompDesc<%=LineID%>" value="<%=Server.HTMLEncode(rc("ItemName"))%>" size="40" readonly class="inputDes"></td>
				<td style="text-align: center;"><input type="checkbox" id="chkCompLocked<%=LineID%>" <% If rc("Locked") = "Y" Then %>checked<% End If %> value="Y" style="border: solid 0px; background: background-image;" onclick="doProcCmbComp(<%=LineID%>, 'Locked', 'S', GetYesNo(this.checked));"></td>
				<td style="text-align: right;">
				<table cellpadding="0" cellspacing="0">
				<tr>
					<td><input type="text" id="txtCompQty<%=LineID%>" value="<%=rc("Quantity")%>" size="5" style="text-align: right;" onchange="doProcCompNum(<%=LineID%>, this, 'Quantity');" onfocus="this.select();" onkeydown="return valKeyNum(event);">
					</td>
					<td>
					<%
					myLock.ID = "Qty" & LineID
					myLock.Value = rc("LockQty") = "Y"
					myLock.OnClick = "doProcCompLock(" & LineID & ", 'Qty');"
					myLock.GenerateLock
					%>
					</td>
				</tr>
				</table>
				</td>
				<td style="text-align: right; display: none;"><input type="text" id="txtCompLines<%=LineID%>" value="<%=rc("Lines")%>" size="2" style="text-align: right;" onchange="doProcCompNum(<%=LineID%>, this, 'Lines');" onfocus="this.select();"></td>
				<td><select id="cmbCompWhs<%=LineID%>" size="1" onchange="doProcCmbComp(<%=LineID%>, 'WhsCode', 'S', this.value);">
				<option value=""><%=getaddItemLngStr("DtxtDefault")%></option>
				<% If rw.recordcount > 0 then rw.movefirst
				do while not rw.eof %><option <% If rc("WhsCode") = rw(0) Then %>selected<% End If %> value="<%=Server.HTMLEncode(rw(0))%>"><%=Server.HTMLEncode(rw(1))%></option><%
				rw.movenext
				loop %>
				</select></td>
				<td><select id="cmbCompCstPrcLst<%=LineID%>" size="1" onchange="doProcCmbComp(<%=LineID%>, 'AlterCostPrcList', 'I', this.value);">
				<option value=""><%=getaddItemLngStr("DtxtDefault")%></option>
				<% rd.movefirst
				do while not rd.eof %>
				<option <% If rc("AlterCostPrcList") = rd(0) Then %>selected<% End If %> value="<%=rd(0)%>"><%=Server.HTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
				<td><select id="cmbCompSalPrcLst<%=LineID%>" size="1" onchange="doProcCmbComp(<%=LineID%>, 'AlterSalePrcList', 'I', this.value);">
				<option value=""><%=getaddItemLngStr("DtxtDefault")%></option>
				<% rd.movefirst
				do while not rd.eof %>
				<option <% If rc("AlterSalePrcList") = rd(0) Then %>selected<% End If %> value="<%=rd(0)%>"><%=Server.HTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
				<td style="width: 80px;">
				<table cellpadding="0" cellspacing="0">
				<tr>
					<td><input type="text" id="txtDiscPrcnt<%=LineID%>" value="<%=FormatNumber(CDbl(rc("DiscPrcnt")), myApp.PercentDec)%>" size="5" style="text-align: right;" onchange="doProcCompNum(<%=LineID%>, this, 'DiscPrcnt');" onfocus="this.select();" onkeydown="return valKeyNum(event);">
					</td>
					<td>
					<%
					myLock.ID = "Disc" & LineID
					myLock.Value = rc("LockDisc") = "Y"
					myLock.OnClick = "doProcCompLock(" & LineID & ", 'Disc');"
					myLock.GenerateLock
					%>
					</td>
				</tr>
				</table>
				</td>
				<td width="16"><img id="btnDel<%=LineID%>" src="ventas/images/cancel_x.gif" style="cursor: pointer;" onclick="delComp(<%=LineID%>);"></td>
			</tr>
			<% rc.movenext
			loop
			End If %>
			</tBody>
			<tr class="GeneralTbl">
				<td colspan="10" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="button" name="btnAddComp" value="<%=getaddItemLngStr("LtxtAddComp")%>" onclick="addComp();"></td>
			</tr>
		</table>
		</td>
	</tr>
	<% End If %>
	<tr class="GeneralTblBold2">
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2" style="cursor: hand; " onclick="showHideSection(tdShowData, trData);">
				<td align="center"><%=getaddItemLngStr("DtxtAddData")%></td>
				<td width="20" id="tdShowData" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">[+]</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id="trData" style="display: none; ">
		<td>
		<table border="0" cellpadding="0" width="100%" id="table3">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtPurUn")%> / <%=getaddItemLngStr("DtxtQty")%></td>
				<td><font size="1">
				<input onfocus="this.select()" size="20" value="<%=myHTMLEncode(rs("BuyUnitMsr"))%>" name="BuyUnitMsr" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProc('BuyUnitMsr', 'S', this.value);"> 
				- </font> &nbsp;<font size="1"><input onfocus="this.select()" onchange="chkThis(this, 'N', '', '', ''); if(this.value=='')this.value=1;doProc('NumInBuy', 'N', this.value);" size="10" value="<%=rs("NumInBuy")%>" name="NumInBuy" onkeydown="return valKeyNum(event);"></font></td>
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtPurPackUn")%> / <%=getaddItemLngStr("DtxtQty")%></td>
				<td><font size="1">
            <input onfocus="this.select()" size="8" value="<%=myHTMLEncode(rs("PurPackMsr"))%>" name="PurPackMsr" onkeydown="return chkMax(event, this, 8);" maxlength="8" onchange="doProc('PurPackMsr', 'S', this.value);"></font> 
				-
				<font size="1">
            <input onfocus="this.select()" onchange="chkThis(this, 'N', '', '', ''); if(this.value=='')this.value=1;doProc('PurPackUn', 'N', this.value);" size="10" value="<%=rs("PurPackUn")%>" name="PurPackUn" onkeydown="return valKeyNum(event);"></font></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtSalUn")%> / <%=getaddItemLngStr("DtxtQty")%></td>
				<td><font size="1">
				<input onfocus="this.select()" size="20" value="<%=myHTMLEncode(rs("SalUnitMsr"))%>" name="SalUnitMsr" onkeydown="return chkMax(event, this, 20);" maxlength="20" onchange="doProc('SalUnitMsr', 'S', this.value);"> 
				- </font> &nbsp;<font size="1"><input onfocus="this.select()" onchange="chkThis(this, 'N', '', '', ''); if(this.value=='')this.value=1;doProc('NumInSale', 'N', this.value);" size="10" value="<%=rs("NumInSale")%>" name="NumInSale" onkeydown="return valKeyNum(event);"></font></td>
				<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtSalPackUn")%> / <%=getaddItemLngStr("DtxtQty")%></td>
				<td><font size="1">
            <input onfocus="this.select()" size="8" value="<%=myHTMLEncode(rs("SalPackMsr"))%>" name="SalPackMsr" onkeydown="return chkMax(event, this, 8);" maxlength="8" onchange="doProc('SalPackMsr', 'S', this.value);"></font> 
				-
				<font size="1">
            <input onfocus="this.select()" onchange="chkThis(this, 'N', '', '', ''); if(this.value=='')this.value=1;doProc('SalPackUn', 'N', this.value);" size="10" value="<%=rs("SalPackUn")%>" name="SalPackUn" onkeydown="return valKeyNum(event);"></font></td>
			</tr>
			<tr>
				<td colspan="4">
				<table cellpadding="0" cellspacing="2" border="0" width="100%">
					<tr class="GeneralTbl">
						<td class="GeneralTblBold2" width="25%"><%=getaddItemLngStr("DtxtImage")%></td>
						<td class="GeneralTblBold2" width="75%"><%=getaddItemLngStr("DtxtObservations")%></td>
					</tr>
					<tr class="GeneralTbl">
						<td width="25%">
						<table border="0" cellpadding="0" width="100%" id="table5">
							<tr>
								<td> 
								<p align="center"> <% If rs("PicturName") <> "" Then PicturName = rs("PicturName") Else PicturName = "n_a.gif" %>
		        				<table border="0" cellpadding="0" width="100%" id="table12" cellspacing="0">
									<tr>
										<td>
		        						<p align="center">
		        						<img id="ItemImg" src="pic.aspx?filename=<%=PicturName%>&MaxSize=223&dbName=<%=Session("olkdb")%>" border="1" name="ItemImg"></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td>
								<p align="center">
								<input type="button" value="<%=getaddItemLngStr("DtxtAddImg")%>" name="btnChangeImg" onclick="javascript:getImg(document.frmAddItem.picturName,document.frmAddItem.ItemImg, 223);">&nbsp;
								<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddItem.picturName.value = ''; document.frmAddItem.ItemImg.src='pic.aspx?filename=n_a.gif&MaxSize=223&dbName=<%=Session("olkdb")%>';doProc('Picture', 'S', '');" style="cursor: hand">
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
							</tr>
						</table>
						</td>
						<td valign="top" width="75%">
						<textarea rows="11" name="UserText" cols="50" style="width: 100%;" onchange="doProc('UserText', 'S', this.value);"><% If Not IsNull(rs("UserText")) Then %><%=Server.HTMLEncode(rs("UserText"))%><% End If %></textarea>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<% If 1 = 2 Then %>
	<tr class="GeneralTblBold2">
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2" style="cursor: hand; " onclick="showHideSection(tdShowAlterFilters, trShowAlterFilters);">
				<td align="center"><%=getaddItemLngStr("LtxtAlterFilters")%></td>
				<td width="20" id="tdShowAlterFilters" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">[+]</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id="trShowAlterFilters" style="display: none; ">
		<td>
		<table style="width: 100%">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2">&nbsp;</td>
				<td><table cellpadding="0" cellspacing="0"><tr class="GeneralTbl"><td><input class="noborder" <% If rs("ShowImage") = "Y" Then %>checked<% End If %> type="checkbox" id="chkShowImage" value="Y" onclick="doProc('ShowImage', 'ShowImage', 'S', GetYesNo(this.checked));"></td><td><label for="chkShowImage"><%=getaddItemLngStr("LtxtAlterImages")%></label></td></tr></table></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px; width: 120px; cursor: pointer;" onclick="showHideFilter('Group');"><%=Server.HTMLEncode(txtAlterGrp)%>&nbsp;<img id="imgGroup" src="images/<%=Session("rtl")%>right.gif"></td>
				<td class="GeneralTbl"><div style="height: 100px; overflow: auto; overflow-x: none; " id="dvGroup">
				<table cellpadding="0" cellspacing="0" width="100%">
					<% rGroup.close
					cmd.CommandText = "DBOLKGetItemGroupsDraftFilters" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = ItmRetVal
					rGroup.open cmd, , 3, 1
					do while not rGroup.eof %>
					<tr class="GeneralTbl">
						<td width="10"><input type="checkbox" name="chkGroup" <% If rGroup("Verfy") = "Y" Then %>checked<% End If %> id="chkGroup<%=rGroup(0)%>" value="<%=rGroup(0)%>" class="noborder" onclick="doProcFilter(1, this);"></td>
						<td><label for="chkGroup<%=rGroup(0)%>"><%=rGroup(1)%></label></td>
					</tr>
					<% rGroup.movenext
					loop %>
				</table>
				</div></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px; width: 120px; cursor: pointer;" onclick="showHideFilter('Firm');"><%=Server.HTMLEncode(txtAlterFrm)%>&nbsp;<img  id="imgFirm" src="images/<%=Session("rtl")%>right.gif"></td>
				<td class="GeneralTbl"><div style="height: 100px; overflow: auto; overflow-x: none;" id="dvFirm">
				<table cellpadding="0" cellspacing="0" width="100%">
					<% rFirm.close
					cmd.CommandText = "DBOLKGetItemFirmsDraftFilters" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = ItmRetVal
					rFirm.open cmd, , 3, 1
					do while not rFirm.eof %>
					<tr class="GeneralTbl">
						<td width="10"><input type="checkbox" name="chkFirm" <% If rFirm("Verfy") = "Y" Then %>checked<% End If %> id="chkFirm<%=rFirm(0)%>" value="<%=rFirm(0)%>" class="noborder" onclick="doProcFilter(2, this);"></td>
						<td><label for="chkFirm<%=rFirm(0)%>"><%=rFirm(1)%></label></td>
					</tr>
					<% rFirm.movenext
					loop %>
				</table>
				</div></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px; width: 120px; cursor: pointer;" onclick="showHideFilter('Prop');"><%=getaddItemLngStr("DtxtProp")%>&nbsp;<img id="imgProp" src="images/<%=Session("rtl")%>right.gif"></td>
				<td class="GeneralTbl"><div style="height: 100px; overflow: auto; overflow-x: none;" id="dvProp">
				<table cellpadding="0" cellspacing="0" width="100%">
					<% 
					set rQryGroup = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetItemQryGroupsDraftFilters" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = ItmRetVal
					rQryGroup.open cmd, , 3, 1
					do while not rQryGroup.eof %>
					<tr class="GeneralTbl">
						<td width="10"><input type="checkbox" name="chkProp" <% If rQryGroup("Verfy") = "Y" Then %>checked<% End If %> id="chkProp<%=rQryGroup(0)%>" value="<%=rQryGroup(0)%>" class="noborder" onclick="doProcFilter(3, this);"></td>
						<td><label for="chkProp<%=rQryGroup(0)%>"><%=rQryGroup(1)%></label></td>
					</tr>
					<% rQryGroup.movenext
					loop %>
				</table>
				</div></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px; width: 120px; cursor: pointer;" onclick="showHideFilter('Crd');"><%=getaddItemLngStr("DtxtSupplier")%>&nbsp;<img id="imgCrd" src="images/<%=Session("rtl")%>right.gif"></td>
				<td class="GeneralTbl"><div style="height: 100px; overflow: auto; overflow-x: none;" id="dvCrd">
				<table cellpadding="0" cellspacing="0" width="100%">
					<% 
					set rSupp = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetItemSuppDraftFilters" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = ItmRetVal
					rSupp.open cmd, , 3, 1
					do while not rSupp.eof %>
					<tr class="GeneralTbl">
						<td width="10"><input type="checkbox" name="chkCrd" <% If rSupp("Verfy") = "Y" Then %>checked<% End If %> id="chkCrd<%=rSupp.bookmark%>" value="<%=Server.HTMLEncode(rSupp(0))%>" class="noborder" onclick="doProcFilter(4, this);"></td>
						<td><label for="chkCrd<%=rSupp.bookmark%>"><%=rSupp(1)%></label></td>
					</tr>
					<% rSupp.movenext
					loop %>
				</table>
				</div></td>
			</tr>
			<% If hasUDFTable Then
			set rUdfFilter = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetItemUDFDraftFilters" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = ItmRetVal
			rUdfFilter.open cmd, , 3, 1 %>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px; width: 120px; cursor: pointer;" onclick="showHideFilter('Udf');"><%=getaddItemLngStr("DtxtUDF")%>&nbsp;<img id="imgUdf" src="images/<%=Session("rtl")%>right.gif"></td>
				<% rUdfFilter.Filter = "Verfy = 'Y'" %>
				<td><div style="height: 200px; <% If rUdfFilter.eof Then %>display: none;<% End If %>" id="dvUdf">
				<table cellpadding="0" cellspacing="0" width="100%">
					<% rUdfFilter.Filter = "Verfy = 'N'"
					If Not rUdfFilter.eof Then %>
					<tr id="trAddField" class="GeneralTbl">
						<td class="GeneralTblBold2"><%=getaddItemLngStr("LtxtAddField")%></td>
						<td><select id="cmbAddField" size="1" onchange="doLoadField(this);">
						<option></option>
						<% do while not rUdfFilter.eof
						rUdfFilterDesc = rUdfFilter("AliasID") & " - " & rUdfFilter("Descr") %>
						<option value="<%=rUdfFilter("FieldID")%>"><%=rUdfFilterDesc%></option>
						<% rUdfFilter.movenext
						loop %>
						</select></td>
					</tr>
					<% End If %>
					<tbody id="tbFieldData">
					<% rUdfFilter.Filter = "Verfy = 'Y'"
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					do while not rUdfFilter.eof
					rUdfFilterDesc = rUdfFilter("AliasID") & " - " & rUdfFilter("Descr")
					fieldID = rUdfFilter("FieldID")
					set rUdfFilterData = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetItemUDFDraftFiltersData" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LogNum") = ItmRetVal
					cmd("@FieldID") = fieldID
					rUdfFilterData.open cmd, , 3, 1 %>
					<tr>
						<td colspan="2"><hr style="height: 1px;"></td>
					</tr>
					<tr class="GeneralTbl">
						<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px;"><%=rUdfFilterDesc%></td>
						<td>
							<div style="height: 100px; overflow: auto; overflow-x: none;">
							<table cellpadding="0" cellspacing="0" width="100%">
							<% j = 0
							do while not rUdfFilterData.eof %>
							<tr class="GeneralTbl">
								<td width="10"><input type="checkbox" class="noborder" onclick="doChkUdfVal('<%=fieldID%>', this);" id="chkUdfVal<%=fieldID%>_<%=j%>" value="<%=Replace(rUdfFilterData(0), """", """""")%>" <% If rUdfFilterData(2) = "Y" Then %>checked<% End If %>></td>
								<td><label for="chkUdfVal<%=fieldID%>_<%=j%>"><%=rUdfFilterData(1)%></label></td></tr>
							<% j = j + 1
							rUdfFilterData.movenext
							loop %>
						</table></div></td></tr>
					<% rUdfFilter.movenext
					loop %>
					</tbody>
				</table>
				</div></td>
			</tr>
			<% End If %>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px; width: 120px; cursor: pointer;" onclick="showHideFilter('Qry');"><%=getaddItemLngStr("DtxtQuery")%>&nbsp;<img id="imgQry" src="images/<%=Session("rtl")%>right.gif"><br>
				<span id="txtOITMFilter" style="display: none;">(<em>from OITM where ...</em>)</span></td>
				<td><div <% If IsNull(rs("Filter")) Then %>style="display: none;"<% End If %> id="dvQry">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><textarea dir="ltr" rows="10" style="width: 100%; font-size: xx-small;" id="txtFilterQry" class="input" onkeypress="javascript:document.getElementById('btnVerfyFilter').src='images/btnValidate.gif';document.getElementById('btnVerfyFilter').style.cursor = 'hand';document.getElementById('valFilterQuery').value='Y';"><% If Not IsNull(rs("Filter")) Then %><%=Server.HTMLEncode(rs("Filter"))%><% End If %></textarea>
						</td>
						<td width="24" valign="bottom">
						<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getaddItemLngStr("DtxtValidate")%>" onclick="javascript:if (document.getElementById('valFilterQuery').value == 'Y')VerfyQuery();">
						<input type="hidden" id="valFilterQuery" value="N"></td>
					</tr>
				</table>
				</div></td>
			</tr>
		</table>
		</td>
	</tr>
	<% End If %>
      <% If EnableSDK Then
	
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OITM"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rg = cmd.execute()

	set rUfd = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OITM"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rUfd.open cmd, , 3, 1

		set rd = Server.CreateObject("ADODB.RecordSet")
		do while not rg.eof
		If CInt(rg("GroupID")) < 0 Then GroupID = "_1" Else GroupID = rg("GroupID") %>
	<tr class="GeneralTblBold2">
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2" style="cursor: hand; " onclick="showHideSection(tdShowUDF<%=GroupID%>, trUDF<%=GroupID%>);">
				<td align="center"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getaddItemLngStr("DtxtUDF")%><%
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
			rUfd.Filter = "GroupID = " & rg("GroupID") & " and Pos = '" & arrPos(i) & "'"
			If not rUfd.eof then %>
				<td width="50%" valign="top">
			        <table border="0" cellpadding="0" cellspacing="2" bordercolor="#111111" width="100%">
			        <% do while not rUfd.eof
			        ShowItemNewUFD()
			        rUfd.movenext
			        loop
			        rUfd.movefirst %>
			        </table>
				</td>
			<% End If
			Next %>
			</tr>
		</table>
		</td>
      </tr>
      <% rg.movenext
      loop %>
      <% End If %>
	<tr class="GeneralTblBold2">
		<td align="center">
			<table border="0" cellpadding="0" width="100%" id="table11">
				<tr>
					<td>
					  <input type="button" <% If rs("VerfyItemCode") = "Y" Then %>disabled<% End If %> value="<% If Not isUpdate Then %><% If Not Confirm Then %><%=getaddItemLngStr("DtxtAdd")%><% Else %><%=getaddItemLngStr("DtxtConfirm")%><% End If %><% Else %><%=getaddItemLngStr("DtxtSave")%><% End If %>" name="btnAdd" onclick="valFrm();"></td>
					<td>
					  <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					  <input type="button" value="<%=getaddItemLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getaddItemLngStr("LtxtConfCancelItm")%>'))window.location.href='itemCancel.asp'"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="newItemSubmit">
<input type="hidden" name="picturName" value="<%=rs("PicturName")%>">
<input type="hidden" name="Confirm" value="">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="isUpdate" value="<%=isUpdate%>">
<input type="hidden" name="doSubmit" value="Y">
<input type="hidden" name="doSubmitAdd" value="Y">
</form>
<script language="javascript">

function valFrm ()
{
	setItemFlow(itmRetVal);
	
	if (document.frmAddItem.ItemCode.value == '')
	{
		alert('<%=getaddItemLngStr("LtxtValItmCod")%>');
		document.frmAddItem.ItemCode.focus();
		return false;
	}
	
	<% If myApp.EnableCombos Then %>
	if (document.frmAddItem.OlkCombo.checked)
	{
		if (!document.frmAddItem.chkFatherShowPrice.checked && !document.frmAddItem.chkShowCompPrice.checked)
		{
			alert('<%=getaddItemLngStr("LtxtValShowPrice")%>');
			return false;
		}
	}
	<% End If %>
	
	<% If EnableSDK Then 
	cmd.CommandText = "DBOLKGetUDFNotNull" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = "V"
	cmd("@TableID") = "OITM"
	cmd("@OP") = "O"
	set rd = Server.CreateObject("ADODB.RecordSet")
	rd.open cmd, , 3, 1
	do while not rd.eof %>
	if (document.frmAddItem.U_<%=rd("AliasID")%>.value == "")
	{
		alert('<%=getaddItemLngStr("LtxtValFld")%>'.replace('{0}', '<%=rd("Descr")%>'));
		showUDF(<%=rd("GroupID")%>);
		document.frmAddItem.U_<%=rd("AliasID")%>.focus;
		return false;
	}
	<% rd.movenext
	loop 
	End If %>
	
	<% If myApp.EnableCombos Then %>
	if (document.getElementById('chkOlkCombo').checked)
	{
		if (!document.getElementById('chkFatherShowPrice').checked && !document.getElementById('chkShowCompPrice').checked)
		{
			alert('<%=getaddItemLngStr("LtxtValFatComPrc")%>');
			return false;
		}
		
		$.post('addItem/addItemComboProcess.asp?d=' + (new Date()).toString(), { CmdType: 'IsValid' }, function(data)
		{
			var arrData = data.split('{S}');
			
			if (arrData[0] == 'N')
			{
				alert('<%=getaddItemLngStr("LtxtValComp")%>');
				return false;
			}
			
			if (arrData[1] == 'N')
			{
				alert('<%=getaddItemLngStr("LtxtValCompVals")%>');
				return false;
			}
			
			doFlowAlert();
		});
	}
	else
	{
		doFlowAlert();
	}
	<% Else %>
	doFlowAlert();
	<% End If %>
	
	
}
function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('<%=getaddItemLngStr("DtxtValFldMaxChar")%>'.replace('{0}', FSize));
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
							alert('<%=getaddItemLngStr("DtxtValNumVal")%>');
						}
						else if (parseInt(getNumericVB(Field.value)) < 1)
						{
							Field.value = '';
							alert('<%=getaddItemLngStr("DtxtValNumMinVal")%>'.replace('{0}', '1'));
						}
						else if (parseInt(getNumericVB(Field.value)) > 2147483647)
						{
							alert('<%=getaddItemLngStr("DtxtValNumMaxVal")%>'.replace('{0}', '2147483647'));
							Field.value = 2147483647;
						}
						else if (Field.value.indexOf('<%=GetFormatDec%>') > -1)
						{
							Field.value = '';
							alert('<%=getaddItemLngStr("DtxtValNumValWhole")%>');
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
					alert('<%=getaddItemLngStr("DtxtValNumVal")%>');
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

function GetFormatValue(FormatValue, EditType)
{
	switch (EditType)
	{
		case "R":
			return OLKFormatNumber(FormatValue, <%=myApp.RateDec%>);
			break;
		case "S":
			return OLKFormatNumber(FormatValue, <%=myApp.SumDec%>);
			break;
		case "P":
			return OLKFormatNumber(FormatValue, <%=myApp.PriceDec%>);
			break;
		case "Q":
			return OLKFormatNumber(FormatValue, <%=myApp.QtyDec%>);
			break;
		case "%":
			return OLKFormatNumber(FormatValue, <%=myApp.PercentDec%>);
			break;
		case "M":
			return OLKFormatNumber(FormatValue, <%=myApp.MeasureDec%>);
			break;
	}
}
</script>

<% Sub ShowItemNewUFD()
	InsertID = rUfd("InsertID")
	FldVal = rs(InsertID)
	Select Case rUfd("TypeID")
		Case "B", "N"
			ProcType = "N"
		Case "M", "A"
			ProcType = "S"
		Case "D"
			ProcType = "D"
	End Select
	%>
  <tr class="GeneralTbl">
    <td width="100" class="GeneralTblBold2">
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
    	  <td valign="top" class="GeneralTblBold2">
    	    <%=rUfd("Descr")%><% If rUfd("NullField") = "Y" Then %><font color="red">*</font><% End If %>
    	  </td>
    	    <% If (rUfd("Query") = "Y" or rUfd("TypeID") = "D") and IsNull(rUfd("RTable")) Then %>
    	    <td width="16">
    	    	<img border="0" id="btn<%=rUfd("AliasID")%>" src="images/<% If rUfd("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" <% If rUfd("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Item&FieldID=<%=rUfd("FieldID")%>&pop=Y<% If rUfd("TypeID") = "A" Then %>&MaxSize=<%=rUfd("SizeID")%><% End If %>',500,300,'yes', 'yes', document.frmAddItem.U_<%=rUfd("AliasID")%>, '<%=ProcType%>')"<% End If %>>
    	    </td>
    	    <% End If %>
    	</tr>
      </table>
    </td>
    <td dir="ltr">
	<% If rUfd("DropDown") = "Y" or Not IsNull(rUfd("RTable")) then 
		set rd = Server.CreateObject("ADODB.RecordSet")
		cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@TableID") = "OITM"
		cmd("@FieldID") = rUfd("FieldID")
		rd.open cmd, , 3, 1 %>
		<select size="1" name="U_<%=rUfd("AliasID")%>" style="width: 99%" onchange="doProc(this.name, '<%=ProcType%>', this.value);">
			<option></option>
			<% do while not rd.eof %>
			<option <% If Not IsNull(rs(InsertID)) Then If CStr(rs(InsertID)) = CStr(rd(0)) Then Response.Write "selected"%> value="<%=rd(0)%>" <% If rUfd("Dflt")= rd(0) Then %>selected<% End If %>><%=myHTMLEncode(rd(1))%></option>
			<% rd.movenext
			loop
			rd.close %>
		</select>
	<% ElseIf rUfd("TypeID") = "M" and Trim(rUfd("EditType")) = "" or rUfd("TypeID") = "A" and rUfd("EditType") = "?" Then %>
		<% If rUfd("Query") = "Y" or rUfd("TypeID") = "D" Then %>
		<table width="100%" cellspacing="0" cellpadding="0">
		  <tr>
		    <td>
		<% End If %>
		<textarea <% If rUfd("TypeID") = "D" or rUfd("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rUfd("AliasID")%>" onchange="chkThis(this, '<%=rUfd("TypeID")%>', '<%=rUfd("EditType")%>', <%=rUfd("SizeID")%>, '');doProc(this.name, '<%=ProcType%>', this.value);" <% If rUfd("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Item&FieldID=<%=rUfd("FieldID")%>&pop=Y<% If rUfd("TypeID") = "A" Then %>&MaxSize=<%=rUfd("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> rows="3" onfocus="this.select()" style="width: 100%"><% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not isNull(rUfd("Dflt")) Then %><%=myHTMLEncode(rUfd("Dflt"))%><% End If %><% End If %></textarea>
		<% If rUfd("Query") = "Y" or rUfd("TypeID") = "D" Then %>
			</td>
			<td width="16">
				<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddItem.U_<%=rUfd("AliasID")%>.value = '';doProc('U_<%=rUfd("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand">
			</td>
		  </tr>
		</table>
		<% End If %>
	<% ElseIf rUfd("TypeID") = "A" and rUfd("EditType") = "I" Then %>
		<table cellpadding="0" cellspacing="2" border="0">
			<tr>
				<td><img src="pic.aspx?filename=<% If IsNull(FldVal) or FldVal = "" Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rUfd("AliasID")%>" border="1">
				<input type="hidden" name="U_<%=rUfd("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
				<td width="16" valign="bottom"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddItem.U_<%=rUfd("AliasID")%>.value = '';document.frmAddItem.imgU_<%=rUfd("AliasID")%>.src='pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';" style="cursor: hand"></td>
			</tr>
			<tr>
				<td colspan="2" height="22">
				<p align="center">
				<input type="button" value="<%=getaddItemLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmAddItem.U_<%=rUfd("AliasID")%>, document.frmAddItem.imgU_<%=rUfd("AliasID")%>,180);"></td>
			</tr>
		</table>
	<% Else
		If Not IsNull(rs(InsertID)) Then 
			If rUfd("TypeID") = "B" Then
	    	Select Case rUfd("EditType")
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
		<% If rUfd("Query") = "Y" or rUfd("TypeID") = "D" Then %>
		<table width="100%" cellspacing="0" cellpadding="0">
		  <tr>
		    <td>
		<% End If %>
		<% 
		If rUfd("TypeID") = "D" or rUfd("Query") = "Y" Then readOnly = True Else readOnly = False
		If rUfd("TypeID") = "D" Then FldVal = FormatDate(FldVal, False)
		If rUfd("TypeID") = "A" Then fldSize = 43 Else fldSize = 12
		If rUfd("TypeID") = "B" or rUfd("TypeID") = "A" Then
			If rUfd("TypeID") = "B" Then MaxSize = 21 Else MaxSize = rUfd("SizeID")
			isMaxSize = True
		Else
			isMaxSize = False
		End If %>
		<input <% If readOnly Then %>readonly<% End If %> type="text" name="U_<%=rUfd("AliasID")%>" id="U_<%=rUfd("AliasID")%>" size="<%=fldSize%>" <% If rUfd("TypeID") = "D" Then %>onclick="btn<%=rUfd("AliasID")%>.click();"<% End If %> onchange="chkThis(this, '<%=rUfd("TypeID")%>', '<%=rUfd("EditType")%>', <%=rUfd("SizeID")%>, '');doProc(this.name, '<%=ProcType%>', this.value);" <% If rUfd("TypeID") = "D" Then %>onclick="btn<%=rUfd("AliasID")%>.click()"<% End If %> <% If rUfd("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Item&FieldID=<%=rUfd("FieldID")%>&pop=Y<% If rUfd("TypeID") = "A" Then %>&MaxSize=<%=rUfd("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> value="<% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not isNull(rUfd("Dflt")) Then %><%=rUfd("Dflt")%><% End If %><% End If %>" <% If rUfd("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> style="width: 100%" <% If isMaxSize Then %> onkeydown="return chkMax(event, this, <%=MaxSize%>);" maxlength="<%=MaxSize%>"<% End if %>>
		<% If rUfd("Query") = "Y" or rUfd("TypeID") = "D" Then %>
			</td>
			<td width="16">
				<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddItem.U_<%=rUfd("AliasID")%>.value = '';doProc('U_<%=rUfd("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand">
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
	rUfd.Filter = "TypeID = 'D'"
	If rUfd.recordcount > 0 Then rUfd.movefirst
	do while not rUfd.eof %>
	    Calendar.setup({
	        inputField     :    "U_<%=rUfd("AliasID")%>",     // id of the input field
	        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
	        button         :    "btn<%=rUfd("AliasID")%>",  // trigger for the calendar (button ID)
	        align          :    "Bl",           // alignment (defaults to "Bl")
	        singleClick    :    true
	    });
	<% rUfd.movenext
	loop 
End If %>

</script>

