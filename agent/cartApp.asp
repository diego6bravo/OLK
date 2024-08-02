<% addLngPathStr = "" %>
<!--#include file="lang/cartApp.asp" -->
<% 
LtxtErrItmInv = Replace(Replace(getcartAppLngStr("LtxtErrItmInv"), " {0}", ""), """", """""")

Session("ShowError") = False
set rctn = Server.CreateObject("ADODB.recordset")
set rd = Server.CreateObject("ADODB.recordset")
set rs = Server.CreateObject("ADODB.recordset")
set rg = Server.CreateObject("ADODB.recordset")

set rx = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCartRepRead" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@UserType") = userType
cmd("@OP") = "O"
rx.open cmd, , 3, 1

set rcOpt = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@TableID") = "OINV"
cmd("@UserType") = userType
cmd("@OP") = "O"
rcOpt.open cmd, , 3, 1
chkOpt = rcOpt.recordcount > 0

If Session("PayCart") Then PayLogNum = Session("PayRetVal") Else PayLogNum = -1

TreePricOn = myApp.TreePricOn
If Session("CartGroup") = "" Then Session("CartGroup") = myApp.CartGroup

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCartLinesCount" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("RetVal")
set rCount = Server.CreateObject("ADODB.RecordSet")
set rCount = cmd.execute()
myLinesCount = rCount(0)
rCount.close

MainCur = myApp.MainCur
VatPrcnt = myApp.VatPrcnt
			
Select Case userType 
	Case "V" 
		EnableDiscount = myAut.HasAuthorization(68)
		ShowPriceBefDiscount = myApp.ShowPriceBefDiscount
		ShowLineDiscount = myApp.ShowLineDiscount
		If not myApp.ApplyMaxDiscToSU and Session("useraccess") = "P" Then MaxDiscount = 100
		If Session("useraccess") = "U" Then
			MaxDiscount = mySession.MaxDocDiscount
			MaxLineDisc = mySession.MaxLineDiscount
		Else
			MaxDiscount = myApp.MaxDiscount
			MaxLineDisc = MaxDiscount
		End If
		AllowPartSuppSel = myAut.HasAuthorization(104)
	Case "C"
		MaxDiscount = 0
		MaxLineDisc = 0
		AllowPartSuppSel = myApp.AllowClientPartSuppSel
End Select
			
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCartInfo" & Session("ID")
cmd.Parameters.Refresh()
cmd("@lognum") = Session("RetVal")
cmd("@LanID") = Session("LanID")
set rs = cmd.execute()

DocCur = rs("DocCur")
ShowCurncy = rs("ShowCurncy")
EnableMC = rs("EnableMC")
DocDate = rs("DocDate")
SaldoFuera = CDbl(rs("SaldoFuera"))
SaldoChecked = rs("SaldoChecked")
PayDocCur = rs("PayDocCur")
ClientHasLineUDF = rs("ClientHasLineUDF") = "Y"

chkLineSumQty = myApp.CartSumQty < myLinesCount
			
If rs("Verfy") = "True" then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetLinesInfo" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@LogNum") = Session("RetVal")
	cmd("@MainCur") = myApp.MainCur
	cmd("@SumDec") = myApp.SumDec
	cmd("@DirectRate") = GetYN(myApp.DirectRate)
	cmd("@LawsSet")= myApp.LawsSet
	cmd("@OptProm") = GetYN(optProm)
	cmd("@UserType") = userType
	cmd("@CardCode") = Session("UserName")
	cmd("@Enable3dx") = GetYN(myApp.Enable3dx)
	cmd("@SDKLineMemo") = GetYN(myApp.SDKLineMemo)
	cmd("@Object") = rs("Object")
	cmd("@PriceList") = Session("PriceList")
	cmd("@SlpCode") = Session("vendid")
	
	If myApp.EnableCartSum and Request("ViewMode") = "" Then
		If myApp.CartSumQty < myLinesCount and Request("document") <> "B" Then
			cmd("@CartSumQty") = myApp.CartSumQty
		ElseIf Request("document") = "B" and Request("String") <> "" Then
			cmd("@SearchStr") = Request("String")
		End If
	End If
	set rd = cmd.execute()
	rdMoveFirst = False
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetCartExpenses" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@LogNum") = Session("RetVal")
    rg.open cmd, , 3, 1
End If 	
If myApp.SVer >= "6" Then CartExpAddStr = "Exp"

cartJsSrc = "cart.js.asp?Verfy=" & rs("Verfy") & "&PayDocCur=" & PayDocCur & _
"&Balance=" & rs("BalanceMC") & "&EnableMC=" & EnableMC & "&myLinesCount=" & myLinesCount & _
"&CreditLine=" & rs("CreditLine") & "&chkOpt=" & GetBoolStr(chkOpt) & "&txtClient=" & txtClient & "&txtInv=" & txtInv & _
"&object=" & rs("object") & "&cmd=" & Request("cmd") & _
"&EnableDiscount=" & GetBoolStr(EnableDiscount) & "&SelDes=" & SelDes & "&MaxLineDisc=" & MaxLineDisc & _
"&document=" & Request("document") & "&ViewMode=" & Request("ViewMode") & "&txtOfert=" & txtOfert & "&chkLineSumQty=" & JBool(chkLineSumQty) & "&string=" & Server.URLEncode(Request("String"))
			
cartVbSrc = "cart.vbs.asp?MaxDiscount=" & MaxDiscount 

addColSpan = 0
If myApp.GetShowRef Then addColSpan = addColSpan + 1
If myApp.GetShowSalUn Then addColSpan = addColSpan + 1
If EnableDiscount Then
	If ShowLineDiscount Then addColSpan = addColSpan + 1
	If ShowPriceBefDiscount Then addColSpan = addColSpan + 1
End If 
%>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript" src="<%=cartJsSrc%>"></script>
<SCRIPT LANGUAGE="JavaScript">
var txtDocDueDateLimit = '<%=getcartAppLngStr("LtxtDocDueDateLimit")%>';
var PList = <%=Session("PriceList")%>;
var UserType = '<%=userType%>';
var DocCur = '<%=DocCur%>';
var MainCur = '<%=MainCur%>';
var oldDocCur = '<%=DocCur%>';
var cartObject = <%=RS("object")%>;

function ValidateForm(frm) {
	if (document.frmCart.DocDueDate.value == '')
	{
		<% If myApp.EnableHideCartHdr Then %>if (document.frmCart.isHDHidden.value == 'Y') showHeaderDet(this, document.frmCart.isHDHidden, '<%=cartBtnSHAddStr%>');<% End if %>
		var fldDesc = '';
		switch (cartObject)
		{
			case 13:
				fldDesc = DueDate13;
				break;
			case 15:
				fldDesc = DueDate15;
				break;
			case 17:
				fldDesc = DueDate17;
				break;
			case 23:
				fldDesc = DueDate23;
				break;
		}
		alert('<%=getcartAppLngStr("LtxtValFld")%>'.replace('{0}', fldDesc));
		document.frmCart.DocDueDate.focus;
		return false;
	}
	<% If userType = "C" and myApp.ClientType = "C" Then %>
	else if (document.frmCart.D1.selectedIndex == 0)
	{
		<% If myApp.EnableHideCartHdr Then %>if (document.frmCart.isHDHidden.value == 'Y') showHeaderDet(this, document.frmCart.isHDHidden, '<%=cartBtnSHAddStr%>');<% End If %>
		alert('<%=getcartAppLngStr("LtxtValSelCnt")%>');
		return false;
	}
	<% End If %>
	
	<% If chkOpt then 
	cmd.CommandText = "DBOLKGetUDFNotNull" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = userType
	cmd("@TableID") = "OINV"
	cmd("@OP") = "O"
	set rctn = cmd.execute()
	do while not rctn.eof %>
	if (document.frmCart.U_<%=rctn("AliasID")%>.value == '') {
		<% If myApp.EnableHideCartHdr Then %>if (document.frmCart.isHDHidden.value == 'Y') showHeaderDet(this, document.frmCart.isHDHidden, '<%=cartBtnSHAddStr%>');<% End If %>
		alert('<%=getcartAppLngStr("LtxtValFld")%>'.replace('{0}', '<%=rctn("Descr")%>'));
		document.frmCart.U_<%=rctn("AliasID")%>.focus;
		return false;
	}
	<% 
	rctn.movenext
	loop
	End If %>
	
	<% If userType = "V" Then %>
	if (!val3dx()) return false;
	if (!valBtch()) return false;
	<% End If %>
	return true;
}


</SCRIPT>
<form method="POST" action="cart/cartupdate.asp" name="frmCart">
<table border="0" cellpadding="0" width="100%">
	<% If tblCustTtl = "" Then %>
	<tr>
		<td id="tdMyTtl" class="TablasTituloSec">&nbsp;<%=getcartAppLngStr("LttlShopCart")%></td>
	</tr>
	<% Else %>
	<%=Replace(tblCustTtl, "{txtTitle}", getcartAppLngStr("LttlShopCart"))%>
	<% End If %>
	<tr class="CanastaTitle2">
		<td>
		<%=CmpName%>&nbsp;</td>
	</tr>
	<% If userType = "V" and rs("Source") = "C" Then %>
	<tr>
		<td id="tdMyTtl" class="TablasTituloDraft">
		<%=getcartAppLngStr("LtxtClientDoc")%></td>
	</tr>
	<% End If %>
	<% If myApp.EnableHideCartHdr Then %>
	<tr class="CanastaTitle2">
		<td>
		<input type="hidden" id="isHDHidden" name="isHDHidden" value="Y">
		<input type="button" class="BtnMore" style="width: 240px;" value="<% If cartBtnSHAddStr = "" Then %>+ <% Else %><%=cartBtnSHAddStr%><% End If %><%=getcartAppLngStr("LtxtShowHdrDet")%>" id="btnSHDetails" name="btnSHDetails" onclick="javascript:showHeaderDet(this, isHDHidden, '<%=cartBtnSHAddStr%>');"></td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr class="CanastaTblResaltada"<% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
				<td colspan="2">
				<p align="center"><%=getcartAppLngStr("DtxtLogNum")%>&nbsp;<%=Session("RetVal")%></td>
			</tr>
			<tr>
				<td width="50%" valign="top">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("LtxtTo")%></nobr></td>
						<td class="CanastaTbl" width="70%">
						<% If userType = "V" Then %><input type="text" class="input" style="width: 100%;" name="CardName" size="27" value="<%=Replace(myHTMLEncode(rs("CardName")), """", "&quot;")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 100);" onchange="doProc(this.name, 'S', this.value);"><% Else %><% If Not isNull(rs("CardName")) Then %><%=rs("CardName")%><% End If %><% End If %></td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada" valign="top"><nobr><%=getcartAppLngStr("LtxtShipAdd")%></nobr></td>
						<td class="CanastaTbl" width="70%">
						<table border="0" cellpadding="0" width="100%" cellspacing="0">
							<tr>
								<td>
								<select class="input" size="1" name="ShipToCode" style="font-size:10px; font-family:Verdana; width:100%" onchange="javascript:doProc('ShipToCode', 'S', this.value);">
								<% 
								set cmd = Server.CreateObject("ADODB.Command")
								cmd.ActiveConnection = connCommon
								cmd.CommandType = &H0004
								cmd.CommandText = "DBOLKGetBPAdds" & Session("ID")
								cmd.Parameters.Refresh()
								cmd("@CardCode") = Session("UserName")
								cmd("@Type") = "S"
								set rctn = cmd.execute()
								do while not rctn.eof %>
								<option value="<%=myHTMLEncode(rctn(0))%>" <% If rs("ShipToCode") = rctn(0) Then %>selected<% End If %>><%=myHTMLEncode(rctn(0))%></option>
								<% rctn.movenext
								loop %>
								</select></td>
							</tr>
							<tr>
								<td class="CanastaTbl"><span id="txtShipAddress"><%=RS("ShipAddress")%></span></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada" valign="top"><nobr><%=getcartAppLngStr("LtxtPayAdd")%></nobr></td>
						<td class="CanastaTbl" width="70%">
						<table border="0" cellpadding="0" width="100%" cellspacing="0">
							<tr>
								<td>
								<select class="input" size="1" name="PayToCode" style="font-size:10px; font-family:Verdana; width:100%" onchange="javascript:doProc('PayToCode', 'S', this.value);">
								<% 
								cmd("@Type") = "B"
								set rctn = cmd.execute()
								do while not rctn.eof %>
								<option value="<%=myHTMLEncode(rctn(0))%>" <% If rs("PayToCode") = rctn(0) Then %>selected<% End If %>><%=myHTMLEncode(rctn(0))%></option>
								<% rctn.movenext
								loop %>
								</select></td>
							</tr>
							<tr>
								<td class="CanastaTbl"><span id="txtPayAddress"><%=RS("PayAddress")%></span></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("DtxtPhone")%></nobr></td>
						<td class="CanastaTbl" width="70%"><%=RS("Phone1")%></td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("DtxtFax")%></nobr></td>
						<td class="CanastaTbl" width="70%"><%=RS("fax")%></td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("DtxtEMail")%></nobr></td>
						<td class="CanastaTbl" width="70%"><%=RS("e_mail")%></td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("DtxtContact")%><% If userType = "C" and myApp.ClientType = "C" Then %><font color="red">*</font><% End If %></nobr></td>
						<td class="CanastaTbl" width="70%">
				         <%
				         set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						cmd("@CardCode") = Session("UserName")
						set rctn = cmd.execute()
						%>
						<select class="input" size="1" name="D1" style="font-size:10px; font-family:Verdana; width:100%" onchange="doProc('CntctCode', 'N', this.value);">
						<% If myApp.ClientType = "C" and userType = "C" Then %>
						<option></option>
						<% End If %>
				        <% do while not rctn.eof %>
				        <option <% If rctn("CntctCode") = rs("CntctCode") Then %>selected<% End If %> value="<%=rctn("cntctcode")%>"><%=myHTMLEncode(rctn("name"))%></option>
				        <% rctn.movenext
				        loop
				        %>
				        </select></td>
					</tr>
					<% 
	
					rcOpt.Filter = "Pos = 'I'"
					do while not rcOpt.eof 
                    ShowAddCartUFD()
                    rcOpt.movenext
                    loop %>
					<% If userType = "V" or userType = "C" and myApp.EnCSelDoc Then %>
					<tr class="CanastaTblResaltada">
						<td><%=getcartAppLngStr("DtxtDocType")%></td>
						<td class="CanastaTbl">
						<% If Not Session("PayCart") Then %>
							<% If rs("object") <> 22 Then %>
							<select name="R1" size="1" class="input" onchange="javascript:changeDocType(this.value);">
							<% If rs("object") <> 203 and rs("object") <> 204 Then %>
                            <% If myApp.EnableOQUT or RS("object") = "23" Then %><option value="23" <% If RS("object") = "23" Then Response.Write "selected" %>><%=txtQuote%></option><% End If %>
                            <% If myApp.EnableORDR or RS("object") = "17" Then %><option value="17" <% If RS("object") = "17" Then Response.Write "selected" %>><%=txtOrdr%></option><% End If %>
                            <% If myApp.EnableODLN or rs("object") = "15" Then %><option value="15" <% If RS("object") = "15" Then Response.Write "selected" %>><%=txtOdln%></option><% End If %>
                            <% If (myApp.EnableOINV and userType = "V") or rs("object") = "13" and rs("ReserveInvoice") <> "Y" then %><option value="13" <% If RS("object") = "13" and rs("ReserveInvoice") <> "Y" Then Response.Write "selected" %>><%=txtInv%></option><% end if %>
                            <% If (myApp.EnableOINVRes and userType = "V") or rs("object") = "13" and rs("ReserveInvoice") = "Y" then %><option value="-13" <% If RS("object") = "13" and rs("ReserveInvoice") = "Y" Then Response.Write "selected" %>><%=txtInvRes%></option><% end if %>
                            <% Else %>
                            <% If myApp.EnableODPIReq or rs("object") = "203" Then %><option value="203" <% If RS("object") = "203" Then Response.Write "selected" %>><%=txtODPIReq%></option><% End If %>
                            <% If myApp.EnableODPIInv or rs("object") = "204" Then %><option value="204" <% If RS("object") = "204" Then Response.Write "selected" %>><%=txtODPIInv%></option><% End If %>
                            <% End If %>
                            </select>
                            <% Else %>
                           	<%=txtOpor%>
                            <input type="hidden" name="R1" value="22">
                            <% End If %>
                        <% Else %>
                        <input type="hidden" name="R1" value="13">
                        <% If 1 = 2 Then %>txtInv<% Else %><%=txtInv%><% End If %>/<% If 1 = 2 Then %>txtRct<% Else %><%=txtRct%><% End If %>
                        <% End If %>
						</td>
					</tr>
                    <% Else %>
                    <input type="hidden" name="R1" value="<%=rs("object")%>">
                    <% End If %>
                    <% If AllowPartSuppSel Then %>
					<tr class="CanastaTblResaltada" id="trPartSupply" <% If rs("Object") <> 17 Then %>style="display: none; "<% End If %>>
						<td colspan="2">
						<input type="checkbox" name="PartSupply" style="border-style:solid; border-width:0; background:background-image" <% If rs("PartSupply") = "Y" Then %>checked<% End If %> id="PartSupply" value="Y" onclick="doProc(this.name, 'S', GetYesNo(this.checked));"><label for="PartSupply"><%=getcartAppLngStr("LtxtPartSupply")%></label>
						</td>
					</tr>
                    <% Else %>
                    <input type="hidden" name="PartSupply" value="<%=rs("PartSupply")%>">
                    <% End If %>
					</table>
				</td>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td class="CanastaTblResaltada">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="CanastaTblResaltada">
								<td><%=getcartAppLngStr("DtxtDate")%></td><% If userType = "V" Then %>
								<td align="right" width="16"><img border="0" src="images/cal.gif" id="btnDocDate"></td><% End If %>
							</tr>
						</table>
						</td>
						<td class="CanastaTbl">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input readonly class="<% If userType = "C" Then %>InputDes<% Else %>input<% End If %>" type="text" name="DocDate" size="12" value="<%=FormatDate(RS("DocDate"), False)%>" <% If userType = "V" Then %>onclick="btnDocDate.click()" onchange="changeDocDate();"<% End If %>></td>
								<td><img src="images/icon_alert.gif" id="DocDateAlert" alt="<%=getcartAppLngStr("LtxtDocDateLimit")%>" <% If rs("VerfyDocDate") = "N" Then %>style="display: none;"<% End If %>></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD"<% If rs("VerfyDocDueDate") = "N" Then %> style="display: none"<% End If %><% End If %>>
						<td class="CanastaTblResaltada">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="CanastaTblResaltada">
								<td id="txtDocDueDate"><nobr><% Select Case rs("Object") 
		                      	Case 15
		                      		txtDueDate = getcartAppLngStr("LtxtDelDate")
		                      	Case 17
		                      		txtDueDate = getcartAppLngStr("LtxtDelDate")
		                      	Case 23
		                      		txtDueDate = getcartAppLngStr("LtxtComDate")
		                      	Case Else
		                      		txtDueDate = getcartAppLngStr("LtxtPymntDue")
		                      	End Select %>
								<%=txtDueDate%><font color="red">*</font></nobr></td>
								<td align="right" width="16"><img border="0" src="images/cal.gif" id="btnDocDueDate"></td>
							</tr>
						</table>
						</td>
						<td class="CanastaTbl">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input readonly class="input" type="text" name="DocDueDate" id="DocDueDate" size="12" value="<%=FormatDate(RS("DocDueDate"), False)%>" onclick="btnDocDueDate.click()" onchange="doProc(this.name, 'D', this.value);"></td>
								<td><img src="images/icon_alert.gif" id="DocDueDateAlert" alt="<%=Replace(getcartAppLngStr("LtxtDocDueDateLimit"), "{0}", LCase(txtDueDate))%>" <% If rs("VerfyDocDueDate") = "N" Then %>style="display: none;"<% End If %>></td>
							</tr>
						</table>
						</td>
					</tr>
					<% If ShowCurncy = "Y" Then %>
					<tr>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("LtxtCurr")%></nobr></td>
						<td class="CanastaTbl">
						<% If EnableMC = "N" Then %><input type="text" size="3" class="InputDes" readonly name="DocCur" value="<%=myHTMLEncode(DocCur)%>"><% Else
					     %><select size="1" name="DocCur" class="input" onchange="javascript:changeCur()"><%
						 
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetCurrencies" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						set rctn = cmd.execute()
						do while not rctn.eof %>
						<option <% If DocCur = rctn(0) Then %>selected<% End If %> value="<%=myHTMLEncode(rctn(0))%>"><%=myHTMLEncode(rctn(1))%></option>
						<% rctn.movenext
						loop %>
						</select><% End If
						CurRate = 0
						If DocCur <> MainCur Then
							set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							cmd.CommandText = "DBOLKGetCartCurrRate" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@LogNum") = Session("RetVal")
							set rctn = cmd.execute()
							CurRate = CDbl(rctn("Rate"))
						End If %><input type="text" size="12" class="InputDes" readonly name="DocCurRate" id="DocCurRate" style="text-align:right;<% If DocCur = MainCur Then %>display: none;<% End If %>" value="<%=FormatNumber(CurRate, myApp.RateDec)%>" ></td>
					</tr>
                    <% Else %>
					<input type="hidden" name="DocCur" value="<%=myHTMLEncode(DocCur)%>">
                    <% End If %>
                    <tr <% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada"><nobr><% If 1 = 2 Then %>txtRef2<% Else %><%=txtRef2%><% End If %></nobr></td>
						<td class="CanastaTbl">
						<input class="input" type="text" name="marca" size="27" value="<%=myHTMLEncode(RS("NumAtCard"))%>" maxlength="100" onchange="doProc('NumAtCard', 'S', this.value);"></td>
					</tr>
                    <% If userType = "V" Then %>
                    <tr>
						<td class="CanastaTblResaltada"><nobr><% If 1 = 2 Then %>txtAgent<% Else %><%=txtAgent%><% End If %></nobr></td>
						<td class="CanastaTbl">
						<% 
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetAgents" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						If myAut.HasAuthorization(96) Then %>
						<select class="input" size="1" name="SlpCode" onchange="doProc('SlpCode', 'N', this.value);">
                      <% set rctn = cmd.execute()
                    	do while not rctn.eof %>
                      <option value="<%=rctn("SlpCode")%>" <% If rs("SlpCode") = rctn("SLPCode") Then %>selected<% End If %>><%=myHTMLEncode(rctn("SlpName"))%></option>
                      <% rctn.movenext
                      loop %>
                      </select><% Else
                      cmd("@Filter") = rs("SlpCode")
                      set rctn = cmd.execute()
                      %><%=rctn("SlpName")%><input type="hidden" name="SlpCode" value="<%=rs("SlpCode")%>"><% End If %></td>
					</tr>
					<% Else %>
                    <input type="hidden" name="SlpCode" value="<%=rs("SlpCode")%>">
					<% End If %>
					<% If myApp.EnableDocPrjSel Then %>
                    <tr>
						<td class="CanastaTblResaltada"><nobr><%=getcartAppLngStr("DtxtProject")%></nobr></td>
						<td class="CanastaTbl">
						<% 
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetProjects" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID") %>
						<select class="input" size="1" name="SlpCode" onchange="doProc('Project', 'S', this.value);">
						<option></option>
                      <% set rp = Server.CreateObject("ADODB.RecordSet")
                      	set rp = cmd.execute()
                    	do while not rp.eof %>
                      <option value="<%=rp("PrjCode")%>" <% If rs("Project") = rp("PrjCode") Then %>selected<% End If %>><%=myHTMLEncode(rp("PrjName"))%></option>
                      <% rp.movenext
                      loop %>
                      </select></td>
					</tr>
					<% End If %>
                    <% rcOpt.Filter = "Pos = 'D'"
                    do while not rcOpt.eof 
                    ShowAddCartUFD()
                    rcOpt.movenext
                    loop
                    rcOpt.Filter = ""  %>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="FirmTlt3">
				<td align="center" width="20">&nbsp;</td>
				<% If myApp.GetShowRef Then %><td align="center"><%=getcartAppLngStr("DtxtCode")%></td><% End If %>
				<td align="center"><%=getcartAppLngStr("LtxtProd")%></td>
				<% do while not rx.eof %><td align="center" <% If rx("linkActive") = "Y" Then %>colspan="2"<% End If %>><%=rx("Name")%></td><% rx.movenext
				loop %>
				<td align="center"><%=getcartAppLngStr("DtxtQty")%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><%=getcartAppLngStr("DtxtUnit")%></td><% End If %>
				<% If EnableDiscount Then %>
				<% If ShowPriceBefDiscount Then %>
				<td align="center"><%=getcartAppLngStr("LtxtUnitPrice")%></td>
				<% End If %>
				<% If ShowLineDiscount Then %>
				<td align="center"><%=getcartAppLngStr("DtxtDiscount")%></td>
				<% End If %>
				<% End If %>
				<td align="center"><% If Not EnableDiscount or EnableDiscount and not ShowPriceBefDiscount Then %><%=getcartAppLngStr("DtxtPrice")%><% Else %><%=getcartAppLngStr("LtxtPriceAfterDisc")%><% End If %></td>
				<td align="center" width="120"><%=getcartAppLngStr("DtxtTotal")%></td>
			</tr>
 			<% 
			  If RS("Verfy") = "True" Then
			  If myApp.EnableCartSum and (myApp.CartSumQty < myLinesCount or Request("document") = "B") and Request("ViewMode") = "" Then

				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetLinesSumInfo" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@LogNum") = Session("RetVal")
				cmd("@MainCur") = myApp.MainCur
				cmd("@SumDec") = myApp.SumDec
				cmd("@DirectRate") = GetYN(myApp.DirectRate)
				cmd("@LawsSet")= myApp.LawsSet
				cmd("@OptProm") = GetYN(optProm)
				cmd("@UserType") = userType
				cmd("@CardCode") = Session("UserName")
				cmd("@Enable3dx") = GetYN(myApp.Enable3dx)
				cmd("@Object") = rs("Object")
				cmd("@PriceList") = Session("PriceList")
				
				If myApp.EnableCartSum and Request("ViewMode") = "" Then
					If myApp.CartSumQty < myLinesCount and Request("document") <> "B" Then
						cmd("@CartSumQty") = myApp.CartSumQty
					ElseIf Request("document") = "B" and Request("String") <> "" Then
						cmd("@SearchStr") = Request("String")
					End If
				End If
				set rctn = cmd.execute()
			    %>
			<tr>
				<td colspan="<%=3+addColSpan%>" class="CanastaTbl">
				<p align="center" class="TablasTituloDraft"><% If Request("document") <> "B" Then %><%=getcartAppLngStr("LtxtSumLines")%> <span dir="ltr">1 - <%=myLinesCount-myApp.CartSumQty%></span>&nbsp;<img src="images/icon_alert.gif" alt="<%=Replace(getcartAppLngStr("LtxtErrSumItmInv"), "{0}", rctn("ChkInv"))%>" id="iconAlertLineSum" <% If rctn("ChkInv") = 0 Then %> style="display: none;"<% End If %>><% Else %>-- <%=getcartAppLngStr("LtxtFilterLines")%> --<% End If %></td>
				<td class="CanastaTblResaltada">&nbsp;</td>
				<td class="CanastaTbl" align="right" dir="ltr" id="SummTotal">
				<%=DocCur%>&nbsp;<%=FormatNumber(rctn(0),myApp.SumDec)%></td>
			</tr>
			    <script language="javascript">
			    var txtErrSumItmInv = '<%=getcartAppLngStr("LtxtErrSumItmInv")%>';
			    <% If myApp.Enable3dx Then %>
				var Inc3dx = <%=rctn("Inc3dx")%>;
				<% End If %>
				var IncBtch = <%=rctn("IncBtch")%>;
				</script>
			<% End If
			If userType = "V" Then cssHighlight = "CanastaTblChild" Else cssHighlight = "CanastaTblResaltada"
			EnableLineMore = userType = "V" and myAut.HasAuthorization(92) or userType = "C" and ClientHasLineUDF
			LineNumDOC1 = ""
			  rdLineID = 0
			  Do While NOT Rd.EOF
			  	isLineMore = EnableLineMore and rd("TreeType") <> "C"
			  	LineNum = rd("LineNum")
			  	LineNumDoc1 = myApp.ConcValue(LineNumDoc1, LineNum)
			  	TreeType = rd("TreeType")
			  	ItemCode = rd("ItemCode")
			  	TreePricOn = rd("TreePricOn") = "Y" 
			  	Select Case TreeType
			  		Case "C"
			  			cssClass = cssHighlight
			  		Case Else
			  			cssClass = "CanastaTbl"
			  	End Select
			  	
			  	If TreeType = "S" Then treeFatherID = LineNum
			  	
				If CDbl(rd("SalPackUn")) > 1 and rd("SaleType") = 3 Then FormatQty = 0 Else FormatQty = myApp.QtyDec %>
			<tr style="<% Select Case TreeType 
							Case "S" %>font-weight: bold;<% 
							Case "C" %>font-weight: normal; font-style: italic;text-decoration:none;<%
						End Select %>">
				<td class="FirmTlt3" width="20" style="text-align: center;">
				<% If TreeType <> "C" Then %><img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=LineNum%>);">
				<input type="checkbox" id="DelLine<%=LineNum%>" name="DelLine" value="<%=LineNum%>" style="display: none;"><% Else %>&nbsp;<% End If %></td>
				<% If myApp.GetShowRef Then %><td class="<%=cssClass%>"><a class="<%=cssClass%>" href="#" onclick="javascript:goViewItem(<%=LineNum%>, '<%=CleanItem(ItemCode)%>');"><nobr><%=ItemCode%></nobr></a>&nbsp;</td><% End If %>
				<td class="<%=cssClass%>"><span dir="ltr"><a class="<%=cssClass%>" href="#" onclick="javascript:goViewItem(<%=LineNum%>, '<%=CleanItem(ItemCode)%>');"><% If Not IsNull(RD("ItemName")) Then %><%=RD("ItemName")%><% End If %></a>&nbsp;</span></td>
				<% If rx.recordcount > 0 Then
				rx.movefirst
				do while not rx.eof %><% If rx("linkActive") = "Y" Then %><td class="<%=cssClass%>" width="11"><img style="cursor: pointer;" border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" onclick="javascript:doLineLink(<%=rx("ID")%>, <%=LineNum%>);"></td><% End If %><td class="<%=cssClass%>" <% If rx("Align") <> "" Then %> style="text-align: <% Select Case rx("Align")
					Case "J" %>justify<% Case "C" %>center<% Case "L" %>left<% Case "R" %>right<% End Select %>;"<% End If %>><span id="ci<%=LineNum%>_<%=rx("ID")%>"><%=RD("col" & rx("ID"))%></span></td><%
				rx.movenext
				loop 
				End If
				%><td class="<%=cssClass%>" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" width="100">
				<table cellspacing="0" cellpadding="0" border="0">
					<tr>
						<% If rd("HasVolDisc") = "Y" Then %><td width="23"><img src="<% If alterFocoIcon = "" Then %>images/foco_icon.gif<% Else %><%=alterFocoIcon%><% End If %>" width="23" height="22" style="vertical-align: middle" onclick="showVolRep(this, '<%=ItemCode%>', event, <%=LineNum%>);" onmouseover="showVolRep(this, '<%=ItemCode%>', event, <%=LineNum%>);" onmouseout="clearVolRep();"></td><% End If %>
						<% If isLineMore Then %><td width="16"><a class="<%=cssClass%>" href="javascript:Start2('cart/cartEditLine.asp?LineNum=<%=LineNum%>&cmd=u&redir=cart&pop=Y&AddPath=', 400, 1)"><img border="0" id="btnLineMore<%=LineNum%>" src="images/expand<% If Not IsNull(rd("LineMemo")) Then %>blue<%end if%>.gif" align="center"></a></td><% End if %>
						<td><input <% If TreeType = "C" and IsNull(rd("LockQty")) or rd("LockQty") = "Y" or rd("Locked") = "Y" Then %> readonly class="<%=cssClass%>"<% End If %> size="8" value="<%=FormatNumber(CDbl(RD("Quantity")),FormatQty)%>" name="Qty<%=LineNum%>" id="Qty<%=LineNum%>" onfocus="this.select()" onchange="setLineQty(<%=LineNum%>);" onkeydown="return chkKeyDown('Q', event, <%=rdLineID%>);" class="input" style="text-align: right"></td>
						<td width="16" style="<% If rd("ChkInv") = "Y" Then %>display: none; <% End if %>" id="InvErr<%=LineNum%>"><img src="images/icon_alert.gif" alt="<%=LtxtErrItmInv%>"></td>
					</tr>
				</table>
				</td>
				<% If myApp.GetShowSalUn Then %><td class="<%=cssClass%>"><% If userType = "V" Then %>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td class="<%=cssClass%>" valign="middle">
							<% If TreeType = "N" and rd("Locked") = "N" Then %>
							<select size="1" name="selUn<%=LineNum%>" id="selUn<%=LineNum%>" onchange="javascript:setLineUn(<%=LineNum%>);" class="input" style="width: 120px;">
							<option value="1" <% If RD("SaleType") = "1" Then Response.Write "selected"%>><%=getcartAppLngStr("LtxtUn")%></option>
							<option value="2" <% If RD("SaleType") = "2" Then Response.Write "selected"%>><%=myHTMLEncode(RD("SalUnitMsr"))%><% If myApp.GetShowQtyInUn Then %>(<%=RD("NumInSale")%>)<% End If %></option>
							<option value="3" <% If RD("SaleType") = "3" Then Response.Write "selected"%>><%=myHTMLEncode(RD("SalPackMsr"))%><% If myApp.GetShowQtyInUn Then %>(<%=RD("SalPackUn")%>)<% End If %></option>
							</select>
							<% Else
								Select Case rd("SaleType")
									Case 1 %><%=getcartAppLngStr("LtxtUn")%><% 
									Case 2 %><%=myHTMLEncode(RD("SalUnitMsr"))%>(<%=RD("NumInSale")%>)<% 
									Case 3 %><%=myHTMLEncode(RD("SalPackMsr"))%>(<%=RD("SalPackUn")%>)<%
								End Select 
							End If %>
						</td>
						<% 
							sbType = ""
							sbImg = ""
							If rd("ManSerNum") = "Y" Then
						   		sbType = "S"
						   		If CDbl(rd("SerQty")) = 0 Then
						   			sbAlt = getcartAppLngStr("LtxtIncompleteSeries") 'Cuando los series estan vacios
							   		sbImg = "serial_nocheck"
							   	ElseIf CDbl(rd("SerQty")) = CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtViewSerSum") 'Ver resumen de los series
							   		sbImg = "serial_check"
							   	ElseIf CDbl(rd("SerQty")) < CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtIncompleteSeries") 'Cuando esta incompleto la cantidad de series
							   		sbImg = "serial_checkGris"
							   	ElseIf CDbl(rd("SerQty")) > CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtSerExceeds") 'Cuando la cantidad de series sobrepasa cantidad requerida
							   		sbImg = "serial_x"
							   	End If
						   	ElseIf rd("ManBtchNum") = "Y" and rd("Man3dx") = "N" Then
						   		sbType = "B"
						   		If CDbl(rd("BtchQty")) = 0 Then
						   			sbAlt = getcartAppLngStr("LtxtIncompleteBatchs") 'Cuando los lotes estan vacios
							   		sbImg = "batch_nocheck"
							   	ElseIf CDbl(rd("BtchQty")) = CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtViewBtchSum") 'Ver resumen de los lotes
							   		sbImg = "batch_check"
							   	ElseIf CDbl(rd("BtchQty")) < CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtIncompleteBatchs") 'Cuando esta incompleto la cantidad de lotes
							   		sbImg = "batch_checkGris"
							   	ElseIf CDbl(rd("BtchQty")) > CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtBtchExceeds") 'Cuando la cantidad de lotes sobrepasa cantidad requerida
							   		sbImg = "batch_x"
							   	End If
						   	ElseIf rd("ManBtchNum") = "Y" and rd("Man3dx") = "Y" Then
						   		sbType = "3"
						   		If CDbl(rd("BtchQty")) = 0 Then
						   			sbAlt = getcartAppLngStr("LtxtIncompletedx") 'Cuando los 3dx estan vacios
							   		sbImg = "3dx_nocheck"
							   	ElseIf CDbl(rd("BtchQty")) = CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtView3dxSum") 'Ver resumen de los 3dx
							   		sbImg = "3dx_check"
							   	ElseIf CDbl(rd("BtchQty")) < CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("LtxtIncompletedx") 'Cuando esta incompleto la cantidad de 3dx
							   		sbImg = "3dx_checkGris"
							   	ElseIf CDbl(rd("BtchQty")) > CDbl(rd("UnitQty")) Then
						   			sbAlt = getcartAppLngStr("Ltxt3dxExceeds") 'Cuando la cantidad de 3dx sobrepasa cantidad requerida
							   		sbImg = "3dx_x"
							   	End If
						   	End If
						 If sbImg <> "" Then %>
						<td id="tdSetLnk" style="width: 20px;<% If rs("Object") = 23 or rs("Object") = 22 Then %>display: none;<% End If %>">
						<img border="0" src="images/<%=sbImg%>.gif" id="btnS<%=sbType%>" align="right" alt="<%=sbAlt%>" style="cursor: hand" onclick="doSB(this, '<%=CleanItem(ItemCode)%>', '<%=sbType%>', <%=LineNum%>);">
						</td>
						<% End If %>
					</tr>
				</table>
				
				<% else %>
				<span dir="ltr">
			    <nobr><% Select Case rd("SaleType")
		  			Case 1
		  				Response.write getcartAppLngStr("DtxtUnit")
		  			Case 2
		  				Response.write rd("SalUnitMsr") 
		  				If myApp.GetShowQtyInUn Then Response.Write "(" & rd("NumInSale") & ")"
		  			Case 3
		  				Response.write rd("SalPackMsr") 
		  				If myApp.GetShowQtyInUn Then Response.Write "(" & rd("SalPackUn") & ")"
				  		If myApp.UnEmbPriceSet Then
					  		Response.Write " x " & rd("SalUnitMsr") 
					  		If myApp.GetShowQtyInUn Then Response.Write "(" & rd("NumInSale") & ")"
				  		End If
		  			End Select
			    %></nobr></span><input type="hidden" name="selUn<%=LineNum%>" id="selUn<%=LineNum%>" value="<%=RD("SaleType")%>"><% end if %></td><% End If %>
				<% If EnableDiscount Then %>
				<% If ShowPriceBefDiscount Then %>
				<td class="<%=cssClass%>"><% If Not (TreeType = "S" and Not TreePricOn or TreeType = "C" and (TreePricOn or not IsNull(rd("ShowCompPrice")))) or (TreeType = "S" and rd("ShowFatherPrice") = "Y" or TreeType = "C" and rd("ShowCompPrice") = "Y") Then %><p align="right">
				<input readonly class="<%=cssClass%>" style="background-color: transparent; text-align:right; border-width: 0px; <% Select Case TreeType 
				Case "S" %>font-weight: bold;<% 
				Case "C" %>font-style: italic;<%
				End Select %>" value="<%=myHTMLEncode(RD("Currency"))%>&nbsp;<% If Not myApp.UnEmbPriceSet and rd("SaleType") = 3 Then %><%=FormatNumber(CDbl(RD("UnitPrice"))*CDbl(rd("SalPackUn")),myApp.PriceDec)%><% Else %><%=FormatNumber(RD("UnitPrice"),myApp.PriceDec)%><% End If %>" name="UnitPrice<%=LineNum%>" id="UnitPrice<%=LineNum%>" size="16" dir="ltr" onfocus="this.select()"><% End If %></td>
				<% End If %>
				<% If ShowLineDiscount Then %>
				<td class="<%=cssClass%>"><% If Not (TreeType = "S" and (Not TreePricOn or rd("ShowFatherPrice") = "Y") or TreeType = "C" and (TreePricOn or not IsNull(rd("ShowCompPrice")))) or (TreeType = "S" and rd("ShowFatherPrice") = "Y" or TreeType = "C" and rd("ShowCompPrice") = "Y") Then %><p align="right">
				<input style="text-align:right" value="%&nbsp;<%=FormatNumber(RD("DiscPrcnt"),myApp.PercentDec)%>" <% If rd("Locked") = "Y" or rd("LockDisc") = "Y" or (TreeType = "S" and rd("AllowChangeFatherPrice") = "N" or TreeType = "C" and rd("AllowChangeCompPrice") = "N") or (Not myAut.HasAuthorization(68) or (TreeType = "C" and Not TreePricOn and IsNull(rd("AllowChangeCompPrice")))) Then %> readonly class="<%=cssClass%>"<% Else %>class="input"<% End If %> name="DiscPrcnt<%=LineNum%>" id="DiscPrcnt<%=LineNum%>" size="16" dir="ltr" onfocus="this.select()" onchange="javascript:setLineDisc(<%=LineNum%>);" onkeydown="return chkKeyDown('D', event, <%=rdLineID%>);">
				<% End If %></td>
				<% End If %>
				<% End If %>			    
				<td class="<%=cssClass%>">
				<% If Not (TreeType = "S" and (Not TreePricOn or rd("ShowFatherPrice") = "Y") or TreeType = "C" and (TreePricOn or not IsNull(rd("ShowCompPrice")))) or (TreeType = "S" and rd("ShowFatherPrice") = "Y" or TreeType = "C" and rd("ShowCompPrice") = "Y") Then %>
				<p align="right">
			    <input <% If userType = "C" or userType = "V" and (rd("LockDisc") = "Y" or TreeType = "C" and Not TreePricOn and IsNull(rd("AllowChangeCompPrice"))) or rd("Locked") = "Y" or (TreeType = "S" and rd("AllowChangeFatherPrice") = "N" or TreeType = "C" and rd("AllowChangeCompPrice") = "N") Then %>readonly class="<%=cssClass%>"<% Else %><% If Not myAut.HasAuthorization(68) Then %> readonly class="<%=cssClass%>"<% Else %>class="input"<% End If %><% End If %> style="text-align:right<% If userType = "C" Then %>; border-width: 0px;<% End If %>" value="<%=myHTMLEncode(RD("Currency"))%>&nbsp;<% If Not myApp.UnEmbPriceSet and rd("SaleType") = 3 Then %><%=FormatNumber(CDbl(RD("Price"))*CDbl(rd("SalPackUn")),myApp.PriceDec)%><% Else %><%=FormatNumber(RD("Price"),myApp.PriceDec)%><% End If %>" name="price<%=LineNum%>" id="price<%=LineNum%>" size="16" dir="ltr" onfocus="this.select()" onchange="setLinePrice(<%=LineNum%>, '<%=myHTMLEncode(RD("Currency"))%>');" onkeydown="return chkKeyDown('P', event, <%=rdLineID%>);">
			    <% End If %>
			    </td>
				<td class="<%=cssClass%>"><% If Not (TreeType = "S" and Not TreePricOn or TreeType = "C" and (TreePricOn or not IsNull(rd("ShowCompPrice")))) or (TreeType = "S" and rd("ShowFatherPrice") = "Y" or TreeType = "C" and rd("ShowCompPrice") = "Y") Then %>
				<p align="right">
				<input class="<%=cssClass%>" readonly value="<%=myHTMLEncode(DocCur)%>&nbsp;<%=FormatNumber(RD("LineTotal"),myApp.SumDec)%>" name="LineTotal<%=LineNum%>" id="LineTotal<%=LineNum%>" size="26" dir="ltr" style="background-color: transparent; border-width: 0px; text-align: right; <% Select Case TreeType 
				Case "S" %>font-weight: bold;<% 
				Case "C" %>font-style: italic;<%
				End Select %>"><% End If %></td>
			</tr>
			<% rdLineID = rdLineID + 1
			Rd.MoveNext
		     loop
		     Else

		     custFldSpan = rx.recordcount
		     rx.Filter = "linkActive = 'Y'"
		     custFldSpan = custFldSpan + rx.recordcount %>
			<tr>
				<td colspan="<%=5+addColSpan+custFldSpan%>" class="CanastaTbl">
				<p align="center"><%=getcartAppLngStr("LtxtEmptyCart")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
				<tr>
					<td>
						<table cellpadding="0">
							<tr>
								<td width="70"><% If rs("Verfy") Then %><input type="submit" class="BtnEliminar<% If Session("rtl") <> "" Then %>Rtl<% End If %>" value="<%=cartBtnAddStr%><%=getcartAppLngStr("LtxtDelete")%>" name="btnDelLines" onclick="return valDel();"><% End If %></td>
							  <% If Request("document") = "B" Then %>
							  <td width="100">
							  <input type="button" name="btnClearFilter" class="btnClearFilter" value="<%=getcartAppLngStr("LtxtClearFilter")%>" onclick="window.location.href='?cmd=cart'">
							  </td>
							  <% End If %>
								<td><% If myApp.EnableCartSum and myApp.CartSumQty < myLinesCount Then
			                  If Request("ViewMode") = "" Then
			                  	viewCmd = "all"
			                  	viewBtnStr = getcartAppLngStr("LtxtViewAll")
			                  	viewBtnCss = "BtnMore"
			                  Else
			                  	If Request("ViewMode") = "all" Then
				                  	viewCmd = ""
				                  	viewBtnStr = getcartAppLngStr("LbtnViewSummary")
				                  	viewBtnCss = "BtnLess"
				                Else
				                  	viewCmd = "all"
				                  	viewBtnStr = getcartAppLngStr("LtxtViewAll")
				                  	viewBtnCss = "BtnMore"
				                End If
			                  End If %><input type="hidden" name="oldViewMode" value="<%=Request("ViewMode")%>"><input type="submit" class="<%=viewBtnCss%>" value="<%=cartBtnAddStr%><%=viewBtnStr%>" name="btnViewAll" onclick="javascript:goViewAll();"><input type="hidden" name="ViewMode" value="<%=Request("ViewMode")%>"><% End If %></td>
							</tr>
						</table>
					</td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<table border="0" cellpadding="0" cellspacing="1">
					<tr>
						<td class="CanastaTblResaltada"><%=getcartAppLngStr("LtxtSubtotal")%></td>
						<td class="CanastaTbl" style="width: 140">
						<p align="right"><%
						set rt = Server.CreateObject("ADODB.RecordSet")
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LogNum") = Session("RetVal")
						cmd("@MC") = "Y"
						rt.open cmd, , 3, 1
		                %><input class="CanastaTbl" readonly name="SubTotal" size="20" dir="ltr" value="<%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rt("SubTotal")), myApp.SumDec)%>" style="background-color: transparent; border-width: 0px; width: 100%; text-align: right"></td>
					</tr>
					<% If userType = "V" or CDbl(rs("DiscPrcnt")) <> 0 Then %>
					<tr id="trDocDiscPrcnt" <% If rs("Object") = 203 or rs("Object") = 204 Then %>style="display: none;"<% End If %>>
						<td class="CanastaTblResaltada">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td class="CanastaTblResaltada"><%=getcartAppLngStr("DtxtDiscount")%></td>
								<td width="1" style="padding-left: 2px; padding-right: 2px;"><input name="DiscPrcnt" onkeydown="return valKeyNumDec(event);" <% If userType = "C" or userType = "V" and Not myAut.HasAuthorization(91) Then %>readonly class="CanastaTbl"<% Else %>class="input"<% End If %> size="6" value="<%=FormatNumber(rs("DiscPrcnt"), myApp.PercentDec)%>" onfocus="this.select()" onchange="setDocDisc();"  style="text-align: right; "></td>
								<td class="CanastaTblResaltada" width="1">%</td>
							</tr>
						</table></td>
						<td class="CanastaTbl" style="width: 140">
						<p align="right">
						<input name="DiscPrcntVal" class="CanastaTbl" readonly value="<%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rt("Discount")), myApp.SumDec)%>" size="5" style="background-color: transparent; border-width: 0px; text-align: right; width: 100%;" dir="ltr">
		                </td>
					</tr>
					<% Else %>
					<input type="hidden" name="DiscPrcnt" value="%<%=FormatNumber(rs("DiscPrcnt"), myApp.PercentDec)%>">
					<% End If %>
					<% If userType = "V" Then %>
					<tr id="trDocDPM" <% If rs("Object") <> 203 and rs("Object") <> 204 Then %>style="display: none;"<% End If %>>
						<td class="CanastaTblResaltada">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td class="CanastaTblResaltada"><%=getcartAppLngStr("DtxtDPM")%></td>
								<td width="1" style="padding-left: 2px; padding-right: 2px;"><input name="DpmPrcnt" onkeydown="return valKeyNumDec(event);" class="input" size="6" value="<%=FormatNumber(rs("DpmPrcnt"), myApp.PercentDec)%>" onfocus="this.select()" onchange="setDocDPM();"  style="text-align: right; "></td>
								<td class="CanastaTblResaltada" width="1">%</td>
							</tr>
						</table></td>
						<td class="CanastaTbl" style="width: 140px;">
						<p align="right">
						<input name="DPMVal" class="CanastaTbl" readonly value="<%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rt("DPM")), myApp.SumDec)%>" size="5" style="background-color: transparent; border-width: 0px; text-align: right; width: 100%;" dir="ltr">
		                </td>
					</tr>
					<% End If %>
					<% If rs("Verfy") = "True" Then
					do while not rg.eof %>
					<tr>
						<td class="<%=cssHighlight%>">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
		                  	<tr>
		                  		<td class="<% If userType = "V" Then %>CanastaTblExpense<% Else %>CanastaTblResaltada<% End If %>"><%=rg("ItemName")%></td>
		                  		<% If userType = "V" Then %><td width="16"><a class="LinkNoticiasMas" href="javascript:delExp(<%=Rg("LineNum")%>, '<%=Server.HTMLEncode(Rg("ItemName"))%>');"><img src="images/<%=Session("rtl")%>remove.gif" border="0"></a></td>
		                  		<td width="16"><% If myApp.SVer >= "6" Then %>
		                  	<a href="javascript:Start2('cart/cartEditExpLine.asp?LineNum=<%=Rg("LineNum")%>&cmd=u&redir=cart&pop=Y&AddPath=', 140, 0)" ;>
							<img border="0" src="images/expand<% If myApp.SDKLineMemo Then %><% If rg("Comments") <> "" Then %>blue<%end if%><%end if%>.gif" align="center"></a><% End If %></td><% End If %>
		                  	</tr>
		                  </table>
						</td>
						<td class="CanastaTbl" style="width: 140px;">
						  <p align="right">
						  <input onkeydown="return valKeyNumDec(event);" <% If userType = "V" Then %>class="input"<% Else %>class="CanastaTbl" readonly<% End If %> name="<%=CartExpAddStr%>Price<%=rg("LineNum")%>" size="18" value="<%=myHTMLEncode(DocCur)%>&nbsp;<%=FormatNumber(Rg("Price"),myApp.SumDec)%>" onfocus="this.select()"  onchange="setExpVal(<%=rg("LineNum")%>, this);" style="text-align: right; width: 100%;<% If userType = "C" Then %>background-color: transparent; border-width: 0px;<% End If %>" dir="ltr"></td>
					</tr>
					<% rg.movenext
		                loop 
		                End If %>
					<tr>
						<td class="CanastaTblResaltada"><% If 1 = 2 Then %>txtTax<% Else %><% If Not IsNull(txtTax) Then %><%=Server.HTMLEncode(txtTax)%><% End If %><% End If %><% If myApp.LawsSet = "IL" Then %><% If Session("myLng") = "he" Then %><span style="font-size: xx-small; "><% End If %>&nbsp;%<%=FormatNumber(VatPrcnt, myApp.PercentDec)%><% If Session("myLng") = "he" Then %></span><% End If %><% End If %></td>
						<td class="CanastaTbl" style="width: 140">
						<p align="right">
		                <input class="CanastaTbl" readonly dir="ltr" type="text" value="<%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rt("Tax")), myApp.SumDec)%>" name="ITBM" size="20" style="background-color: transparent; border-width: 0px; width: 100%; text-align: right"></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada"><%=getcartAppLngStr("DtxtTotal")%></td>
						<td class="CanastaTbl" style="width: 140">
						<p align="right">
		                  <input class="CanastaTbl" readonly dir="ltr" type="text" name="importe" size="20" value="<%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rt("DocTotal")), myApp.SumDec)%>" style="background-color: transparent; border-width: 0px; width: 100%; text-align: right"></td>
					</tr>
					<% If Session("PayCart") Then %>
					<tr id="trTotalMC"<% If EnableMC = "N" Then %> style="display: none;"<% End If %>>
						<td class="CanastaTblResaltada"><%=getcartAppLngStr("LtxtTotalToPay")%></td>
						<td class="CanastaTbl" style="width: 140">
		                  <p align="right">
		                  <% If Not IsNull(rt("DocTotalMC")) Then DocTotalMC = CDbl(rt("DocTotalMC")) Else DocTotalMC = CDbl(rt("DocTotal")) %>
		                  <input class="CanastaTbl" readonly dir="ltr" type="text" name="TotalMC" size="20" value="<%=PayDocCur & " " %><%=FormatNumber(DocTotalMC, myApp.SumDec)%>" style="background-color: transparent; border-width: 0px; width: 100%; text-align: right"></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada"><%=getcartAppLngStr("DtxtPaid")%></td>
						<td class="CanastaTbl" style="width: 140">
		                  <p align="right">
		                  <input class="CanastaTbl" readonly dir="ltr" name="pagado" size="18" value="<%=PayDocCur%>&nbsp;<%=FormatNumber(Rs("Pagado"),myApp.SumDec)%>" style="background-color: transparent; border-width: 0px; text-align: right; width: 100%;"></td>
					</tr>
					<% end if %>
					</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr>
				<td width="261" valign="top">
				<table border="0" cellpadding="0" width="100%">
					<tr class="CanastaTblResaltada">
						<td class=""><%=getcartAppLngStr("LtxtPymntCod")%></td>
					</tr>
					<tr class="CanastaTbl">
					<td><% 
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetPaymentGroups" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					If userType = "V" and myAut.HasAuthorization(88) Then %><select size="1" class="input" name="cpago" style="font-size:10px; font-family:Verdana; Width:100%" onchange="chkListNum(this.value, this.selectedIndex);">
				    <% 
				    If rctn.state = adStateOpen Then rctn.close
				    rctn.open cmd, , 3, 1
				    while not rctn.eof %>
			        <option value="<%=RCTN("GroupNum")%>" <% If RCTN("GroupNum") = RS("GroupNum") then response.write "selected" %>><%=myHTMLEncode(RCTN("PymntGroup"))%></option>
			        <% rctn.movenext
			        wend %>
			        </select>
			        <% rctn.movefirst %>
			        <script language="javascript">
			        var ExtraDays = new Array(<%=rctn.recordcount%>);
			        var ListNum = new Array(<%=rctn.recordcount%>);
			        <% do while not rctn.eof %>
			        ExtraDays[<%=rctn.bookmark-1%>] = <%=rctn("ExtraDays")%>;
			        ListNum[<%=rctn.bookmark-1%>] = <%=rctn("ListNum")%>;
			        <% rctn.movenext
			        loop %>
			        </script>
			        <% else 
			        cmd("@Filter") = rs("GroupNum")
			        set rctn = cmd.execute() %>
			        <%=myHTMLEncode(RCTN("PymntGroup"))%>
			        <input type="hidden" value="<%=RS("GroupNum")%>" name="cpago">
			        <script language="javascript">
			        var ExtraDays = new Array(1);
			        var ListNum = new Array(1);
			        ExtraDays[0] = <%=rctn("ExtraDays")%>;
			        ListNum[0] = <%=rctn("ListNum")%>;
			        </script>
			        <% end if %></td>
					</tr>
					<tr class="CanastaTblResaltada">
						<td><%=getcartAppLngStr("DtxtObservations")%></td>
					</tr>
					<tr>
						<td>
						<textarea rows="5" name="S1" class="input" style="width: 100%" onkeydown="return chkMax(event, this, 254);" onchange="doProc('comments', 'S', this.value);"><%=myHTMLEncode(RS("memo"))%></textarea></td>
					</tr>
				</table>
				</td>
				<td valign="top">
				<% 
				If myApp.ExpItems and userType = "V" and rs("object") <> 203 and rs("object") <> 204 Then
					varx = 0
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetCartAvlExpns" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = Session("RetVal")
					MinArtTitle = LtxtExpenses 'Gastos
				set rd = cmd.execute() %>
              <table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="84%"><b><font face="Verdana" size="1"><span id="MinArtTitle"><%=MinArtTitle%></span></font></b></td>
					<td width="16%">
					<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<select size="1" name="cmbAddExp" class="input" onchange="javascript:if(this.selectedIndex > 0) goAddSmallList(this.value);">
					<option value=""><%=getcartAppLngStr("LtxtExpenses")%></option>
					<% do while not rd.eof %>
					<option value="<%=myHTMLEncode(rd("ItemCode"))%>"><%=myHTMLEncode(rd("DispItem"))%></option>
					<% rd.movenext
					loop %>
					</select></td>
				</tr>
				</table>
				<% ElseIf userType = "C" Then %>
				<%=rs("CCartNote")%>
				<% Else %>&nbsp;
				<% End If %></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<% If rs("Verfy") = "True" then
				confObj = CInt(rs("object"))
				If confObj = 13 and rs("ReserveInvoice") = "Y" Then confObj = -13
				Confirm = myAut.GetObjectProperty(confObj, "C") or userType = "C" %>
				<td width="100">
				<input type="button" <% If rs("LockAdd") = "Y" or rs("VerfyDocDate") = "Y" or rs("VerfyDocDueDate") = "Y" Then %>disabled<% End If %> class="BtnAgregar<% If Session("rtl") <> "" Then %>Rtl<% End If %>" value="<%=cartBtnAddStr%><% If not Confirm Then%><%=getcartAppLngStr("DtxtAdd")%><%else%><%=getcartAppLngStr("DtxtConfirm")%><%end if%>" name="I2" onclick="javascript:if(ValidateForm(this))chkCart('I<% If Not Session("PayCart") then %>2<% ElseIf Session("PayCart") then %>3<% end if %>')"></td>
				<% End If %>
				<% If myApp.GetEnableCartImp Then %>
				<td width="100">
				<input type="button" class="BtnImport<% If Session("rtl") <> "" Then %>Rtl<% End If %>" value="<%=cartBtnAddStr%><%=getcartAppLngStr("DtxtImport")%>" name="btnImport" onclick="javascript:Start('cart/cartImport.asp?pop=Y&AddPath=')"></td>
				<% End If %>
				<td>
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<input type="button" class="BtnCancel<% If Session("rtl") <> "" Then %>Rtl<% End If %>" value="<%=cartBtnAddStr%><%=getcartAppLngStr("DtxtCancel")%>" name="I3" onclick="javascript:if(confirm('<%=getcartAppLngStr("LtxtConfCancelDoc")%>'))window.location.href='cartCancel.asp?RetVal=<%=Session("RetVal")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="cartSubmit" value="I1">
<input type="hidden" name="redir" value="<%=Request("cmd")%>">
<input type="hidden" name="LineNumDOC1" value="<%=LineNumDOC1%>">
<input type="hidden" name="AddPath" value="">
<input type="hidden" name="Confirm" value="N">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="document" value="<%=Request("document")%>">
<input type="hidden" name="String" value="<%=Request("String")%>">
<input type="hidden" name="Draft" value="">
<input type="hidden" name="Authorize" value="">
<% If userType = "V" Then %>
<input type="hidden" name="NewPList" value="">
<input type="hidden" name="ApplyPListLines" value="">
<% End If %>
</form>
<iframe id="iSetData" name="iSetData" style="display: none" src=""></iframe>
<% rx.Filter = "linkActive = 'Y'"
If Not rx.Eof Then %>
<form name="frmRSLink" id="frmRSLink" method="post" action="viewReportPrint.asp" target="_blank">
</form>
<% End If %>
<% set rs = nothing
set rctn = nothing
set rd = nothing

					Sub ShowAddCartUFD()
					AliasID = rcOpt("InsertID")
					Select Case rcOpt("TypeID")
						Case "B", "N"
							ProcType = "N"
						Case "M", "A"
							ProcType = "S"
						Case "D"
							ProcType = "D"
					End Select %>
                    <tr<% If myApp.EnableHideCartHdr Then %> id="trCartHD" style="display: none; "<% End If %>>
						<td class="CanastaTblResaltada">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="CanastaTblResaltada"><nobr><%=rcOpt("Descr")%><% If rcOpt("NullField") = "Y" Then %><font color="red">*</font><% End If %></nobr></td><% If (rcOpt("Query") = "Y" or rcOpt("TypeID") = "D") and IsNull(rcOpt("RTable")) Then %><td width="16"><img border="0" src="<% If rcOpt("TypeID") <> "D" Then %>design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selecB.gif<% Else %>images/cal.gif<% End If %>" id="btn<%=rcOpt("AliasID")%>" <% If rcOpt("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Doc&FieldID=<%=rcOpt("FieldID")%>&pop=Y<% If rcOpt("TypeID") = "A" Then %>&MaxSize=<%=rcOpt("SizeID")%><% End If %>',500,300,'yes', 'yes', document.frmCart.U_<%=rcOpt("AliasID")%>, '<%=ProcType%>')"<% End If %>></td><% End If %></tr></table></td>
                      <td class="CanastaTbl">
                      <% If rcOpt("DropDown") = "Y" or Not IsNull(rcOpt("RTable")) then 
                      If rcOpt("DropDown") = "Y" Then
						cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						cmd("@TableID") = "OINV"
						cmd("@FieldID") = rcOpt("FieldId")
						set rctn = cmd.execute()
					  Else
					  	sql = "select Code, Name from [@" & rcOpt("RTable") & "] order by 2"
					  	set rctn = conn.execute(sql)
					  End If
					 %>
					<font color="#4783C5">
					<select size="1" name="U_<%=rcOpt("AliasID")%>" class="input" style="width: 99%" onchange="doProc(this.name, '<%=ProcType%>', this.value);">
					<option></option>
					<% do while not rctn.eof %>
					<option value="<%=rctn(0)%>" <% If rs(AliasID) = rctn(0) Then %>selected<% ElseIf rctn(0) = rcOpt("Dflt") and IsNull(rs(AliasID)) Then %>selected<% End If %>><%=myHTMLEncode(rctn(1))%></option>
					<% rctn.movenext
					loop
					rctn.close %></select></font><font size="1" color="#4783C5">
					<% ElseIf rcOpt("TypeID") = "M" and Trim(rcOpt("EditType")) = "" or rcOpt("TypeID") = "A" and rcOpt("EditType") = "?" Then %>
						<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>
						<table width="100%" cellspacing="0" cellpadding="0">
						  <tr>
						    <td>
						<% End If %>
						<font size="1" color="#4783C5">
						<textarea <% If rcOpt("TypeID") = "D" or rcOpt("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" class="input" onchange="chkThis(this, '<%=rcOpt("TypeID")%>', '<%=rcOpt("EditType")%>', <%=rcOpt("SizeID")%>);doProc(this.name, '<%=ProcType%>', this.value);" <% If rcOpt("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Doc&FieldID=<%=rcOpt("FieldID")%>&pop=Y<% If rcOpt("TypeID") = "A" Then %>&MaxSize=<%=rcOpt("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> rows="3" onfocus="this.select()" style="width: 100%" cols="20"><% If rs(AliasID) <> "" Then %><%=myHTMLEncode(rs(AliasID))%><% Else %><% If Not IsNull(rcOpt("Dflt")) Then %><%=myHTMLEncode(rcOpt("Dflt"))%><% End If %><% End If %></textarea></font>
						<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>
							</td>
							<td width="16">
								<img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>x_icon.gif" onclick="document.frmCart.U_<%=rcOpt("AliasID")%>.value = '';doProc('U_<%=rcOpt("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
				<% ElseIf rcOpt("TypeID") = "A" and rcOpt("EditType") = "I" Then %>
					<table cellpadding="2" cellspacing="0" border="0">
						<tr>
							<td>
							<p align="center"><img src="pic.aspx?filename=<% If rs(AliasID) <> "" Then %><%=rs(AliasID)%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rcOpt("AliasID")%>" border="1">
							<input type="hidden" name="U_<%=rcOpt("AliasID")%>" value="<%=rs(AliasID)%>"></td>
							<td width="16" valign="bottom"><img border="0" src="<% If userType = "C" Then %>design/<%=SelDes%>/images/<%=Session("rtl")%>x_icon.gif<% Else %>images/<%=Session("rtl")%>remove.gif<% End If %>" onclick="javascript:document.frmCart.U_<%=rcOpt("AliasID")%>.value = '';document.frmCart.imgU_<%=rcOpt("AliasID")%>.src='pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';doProc('U_<%=rcOpt("AliasID")%>', 'S', '');" style="cursor: hand"></td>
						</tr>
						<tr>
							<td colspan="2" height="22">
							<p align="center"><input type="button" value="<%=getcartAppLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmCart.U_<%=rcOpt("AliasID")%>, document.frmCart.imgU_<%=rcOpt("AliasID")%>,180);"></td>
						</tr>
					</table>
					<% Else 
					If rs(AliasID) <> "" Then 
						strVal = rs(AliasID)
						If rcOpt("TypeID") = "B" Then
			        	Select Case rcOpt("EditType")
							Case "R"
								strVal = FormatNumber(CDbl(strVal),myApp.RateDec)
							Case "S"
								strVal = FormatNumber(CDbl(strVal),myApp.SumDec)
							Case "P"
								strVal = FormatNumber(CDbl(strVal),myApp.PriceDec)
							Case "Q"
								strVal = FormatNumber(CDbl(strVal),myApp.QtyDec)
							Case "%"
								strVal = FormatNumber(CDbl(strVal),myApp.PercentDec)
							Case "M"
								strVal = FormatNumber(CDbl(strVal),myApp.MeasureDec)
			        	End Select
			        	End If
					Else
						strVal = ""
					End If
					
					If rcOpt("TypeID") = "D" Then strVal = FormatDate(strVal, False) %>
					<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td><% End If %>
					<input class="input" <% If rcOpt("TypeID") = "D" or rcOpt("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" size="<% If rcOpt("TypeID") = "A" Then %>43<% Else %>12<% End If %>" onchange="chkThis(this, '<%=rcOpt("TypeID")%>', '<%=rcOpt("EditType")%>', <%=rcOpt("SizeID")%>);doProc(this.name, '<%=ProcType%>', this.value);" <% If rcOpt("TypeID") = "D" Then %>onclick="btn<%=rcOpt("AliasID")%>.click()"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Doc&FieldID=<%=rcOpt("FieldID")%>&pop=Y<% If rcOpt("TypeID") = "A" Then %>&MaxSize=<%=rcOpt("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> value="<%=myHTMLEncode(strVal)%>" onfocus="this.select()" <% If rcOpt("TypeID") <> "D" Then %>style="width: 100%"<% End If %> <% If rcOpt("TypeID") = "A" Then %>maxlength="<%=rcOpt("SizeID")%>"<% End If %>>
					<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %></td><td width="16"><img border="0" src="<% If userType = "C" Then %>design/<%=SelDes%>/images/<%=Session("rtl")%>x_icon.gif<% Else %>images/<%=Session("rtl")%>remove.gif<% End If %>" onclick="document.frmCart.U_<%=rcOpt("AliasID")%>.value = '';doProc('U_<%=rcOpt("AliasID")%>', '<%=ProcType%>', '');"></td></tr></table><% End If %><% End If %></td>
                    </tr>
<% End Sub  %>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "DocDueDate",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnDocDueDate",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    <% If userType = "V" Then %>
    Calendar.setup({
        inputField     :    "DocDate",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnDocDate",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    <% End If %>
<% rcOpt.Filter = "TypeID = 'D'"
If rcOpt.recordcount > 0 Then rcOpt.movefirst
do while not rcOpt.eof %>
    Calendar.setup({
        inputField     :    "U_<%=rcOpt("AliasID")%>",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btn<%=rcOpt("AliasID")%>",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
<% rcOpt.movenext
loop 
set rcOpt = Nothing %>
refreshAfterAdd = true;

function goViewAll()
{
document.frmCart.ViewMode.value='<%=viewCmd%>'
document.frmCart.submit();
}
</script>
<!--#include file="itemDetails.inc"-->