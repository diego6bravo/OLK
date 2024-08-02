<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../lcidReturn.inc"-->
<!--#include file="lang/addresses.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl" <% End If %>="">

<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

set rs = Server.CreateObject("ADODB.RecordSet")
set rd = Server.CreateObject("ADODB.RecordSet")

isDefault = False
If Request("Op") <> "" and Request("Op") <> "add" Then
	cmd.CommandText = "DBOLKGetCrdAddData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("CrdRetVal")
	cmd("@LineNum") = CInt(Request("Op"))
	set rs = cmd.execute()
	NewAddress = rs("NewAddress")
	Street = rs("Street")
	Block = rs("Block")
	City = rs("City")
	ZipCode = rs("ZipCode")
	County = rs("County")
	Country = rs("Country")
	State = rs("State")
	TaxCode = rs("TaxCode")
	IsDefault = rs("IsDefault") = "Y"
	IsUpdate = rs("Command") = "U"
	EnableSDK  = rs("EnableSDK") = "Y"
ElseIf Request("Op") = "add" Then
	cmd.CommandText = "DBOLKGetEnableSDK" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = "CRD1"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rs = cmd.execute()
	EnableSDK = rs(0) = "Y"
End If
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link type="text/css" href="../design/0/jquery-ui-1.7.2.custom.css" rel="stylesheet" >	
<script type="text/javascript" src="../jQuery/js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../jQuery/js/jquery-ui-1.7.2.custom.min.js"></script>
<title><% If Request("AdresType") = "S" Then %><%=getaddressesLngStr("LtxtShipAdd")%><% ElseIf Request("AdresType") = "B" Then %><%=getaddressesLngStr("LtxtBillAdd")%><% End If %>
</title>
<script language="javascript">
function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))return false; else return true;
}
var OpenWin = null;
var Field
function chkWin() { if (OpenWin != null) if (!OpenWin.closed) OpenWin.focus() }

function Start(o, page, w, h, s, r) {
Field = o
OpenWin = this.open(page, "queryWin", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}

function setTimeStamp(Nothing, var1) {
	Field.value = var1;
	if (Field.onchange != null) Field.onchange();
}

function chkNum(fld, dType)
{
	if (dType != 'nvarchar')
	{
		if (!MyIsNumeric(fld.value))
		{
			alert('<%=getaddressesLngStr("DtxtValNumVal")%>');
			fld.value = '';
			fld.focus();
		}
		else if (dType == 'int')
		{
			fld.value = parseInt(fld.value);
		}
	}
}
var SaveImgField
var SaveImgImage
var SaveImgMaxSize
function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	Start('../upload/fileupload.aspx?ID=<%=Session("ID")%>&style=../design/0/style/stylePopUp.css',300,100,'no')
}

function Start(page, w, h, s) {
OpenWin = this.open(page, "ContactPopUp", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}

function changepic(img_src) {
SaveImgField.value = img_src;
SaveImgImage.src = "../pic.aspx?filename=" + img_src + "&MaxSize=" + SaveImgMaxSize + '&dbName=<%=Session("olkdb")%>';
}

</script>
<script type="text/javascript" src="../scr/calendar.js"></script>
<script type="text/javascript" src="../scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="../scr/calendar-setup.js"></script>
<script language="javascript" src="../generalData.js.asp"></script>
<script language="javascript" src="../general.js"></script>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<link rel="stylesheet" type="text/css" media="all" href="../design/0/style/style_cal.css" title="winter" />
</head>

<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onfocus="javascript:chkWin();">

<div align="center">
	<table border="0" cellpadding="0" width="580" id="table1">
		<tr class="GeneralTlt">
			<td><% If Request("AdresType") = "S" Then %><%=getaddressesLngStr("LtxtShipAdd")%><% ElseIf Request("AdresType") = "B" Then %><%=getaddressesLngStr("LtxtBillAdd")%><% End If %></td>
		</tr>
		<form method="POST" action="addresses.asp" name="FormTop">
			<input type="hidden" name="pop" value="Y">
			<input type="hidden" name="AddPath" value="../">
			<% 
			set rd = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetCrdAdds" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = Session("CrdRetVal")
			cmd("@AdresType") = Request("AdresType")
			rd.open cmd, , 3, 1
			If Request("Op") <> "" Then %>
			<tr class="GeneralTlt">
				<td><%=getaddressesLngStr("DtxtAddress")%>:
				<select size="1" name="Op" onchange="submit()">
				<option value=""><%=getaddressesLngStr("LtxtDoSel")%></option>
				<option <% If Request("Op") = "add" then response.write "selected "%>value="add">
				<%=getaddressesLngStr("LtxtAddAddress")%></option>
				<% enableChkDef = False
				do while not rd.eof
				enableChkDef = True %>
				<option <% If CStr(Request("Op")) = CStr(rd("LineNum")) then response.write "selected "%>value='<%=rd("LineNum")%>'>
				<%=myHTMLEncode(rd("NewAddress"))%><% If rd("IsDefault") = "Y" Then %>&nbsp;(<%=getaddressesLngStr("DtxtDefault")%>)<% End If %></option>
				<% rd.movenext
			loop %></select>
				</td>
			</tr>
			<% Else %>
			<tr>
				<td>
				<table style="width: 100%">
					<tr class="GeneralTblBold2">
						<td>&nbsp;</td>
						<td><%=getaddressesLngStr("DtxtName")%></td>
						<td><%=getaddressesLngStr("DtxtAddress")%></td>
					</tr>
					<% enableChkDef = False
					do while not rd.eof
					enableChkDef = True %>
					<tr class="<% If rd("IsDefault") = "N" Then %>GeneralTbl<% Else %>CanastaTblExpense<% End If %>" style="vertical-align: top; padding-top: 2px;">
						<td width="15"><a href="#" onclick="javascript:document.FormTop.Op.value='<%=rd("LineNum")%>';submit();"><img border="0" src="../design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
						<td><%=rd("NewAddress")%></td>
						<td><%=rd("FormatAddress")%></td>
					</tr>
					<% rd.movenext
					loop %>
				</table>
				</td>
			</tr>
			<input type="hidden" name="Op" value="">
			<% End If %>
			<input type="hidden" name="AdresType" value='<%=Request("AdresType")%>'>
		</form>
		<% If Request("Op") <> "" Then %>
		<form method="POST" action="submitAddresses.asp" name="frmAddress">
			<tr>
				<td>
				<table border="0" cellpadding="0" width="100%" id="table2">
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("DtxtName")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="NewAddress" size="50" maxlength="50" value="<%=myHTMLEncode(NewAddress)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("LtxtStreet")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Street" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(Street)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("LtxtBlock")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Block" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(Block)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("DtxtCity")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="City" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(City)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("LtxtPostalCode")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="ZipCode" size="20" maxlength="20" value="<%=myHTMLEncode(ZipCode)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("LtxtCounty")%></td>
						<td height="18" class="GeneralTbl">
						<p align="center">
						<input type="text" name="County" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(County)%>"></p>
						</td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("DtxtCountry")%></td>
						<td height="18" class="GeneralTbl">
						<select size="1" name="Country" onchange="changeCountry()" style="width: 50%; ">
						<option value=""></option>
						<%
						set rd = Server.CreateObject("ADODB.RecordSet")
						cmd.CommandText = "DBOLKGetCountries" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						rd.open cmd, , 3, 1
						do while not rd.eof %>
						<option <% If Country = rd("Code") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=rd("Name")%></option>
						<% rd.movenext
						loop %>
						</select></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("DtxtState")%></td>
						<td height="18" class="GeneralTbl">
						<select size="1" name="State" style="height: 16px; width: 50%; ">
						<option value=""></option>
						<% If Country <> "" Then
							set rd = Server.CreateObject("ADODB.RecordSet")
							cmd.CommandText = "DBOLKGetCountryStates" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@Code") = Country
							rd.open cmd, , 3, 1
							do while not rd.eof %>
							<option <% If State = rd("Code") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=rd("Name")%></option>
							<% rd.movenext
							loop
							End If %>
						</select></td>
					</tr>
					<% If (myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA") and Request("AdresType") = "S" Then %>
					<tr class="GeneralTblBold2">
						<td><%=getaddressesLngStr("LtxtTaxCode")%></td>
						<td height="18" class="GeneralTbl">
						<select size="1" name="TaxCode" style="height: 16px; width: 50%; ">
						<option value=""></option>
						<%  
							set rd = Server.CreateObject("ADODB.RecordSet")
							cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							rd.open cmd, , 3, 1
							do while not rd.eof %>
							<option <% If TaxCode = rd("Code") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=rd("Name")%></option>
							<% rd.movenext
							loop %>
						</select></td>
					</tr>
					<% End If %>
					<tr class="GeneralTblBold2">
						<td>&nbsp;</td>
						<td height="18" class="GeneralTbl">
						<input type="checkbox" name="SetDef" id="SetDef" <% If Not enableChkDef or isDefault Then %>checked disabled<% End IF %> value="Y" style="background: background-image; border: 0px solid"><label for="SetDef"><%=getaddressesLngStr("LtxtSetAsDef")%></label></td>
					</tr>
		<% 
		set rSdk = Server.CreateObject("ADODB.RecordSet")
		If EnableSDK Then 
			set rg = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@TableID") = "CRD1"
			cmd("@UserType") = "V"
			cmd("@OP") = "O"
			set rg = cmd.execute()
		
			set rSdk = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@TableID") = "CRD1"
			cmd("@UserType") = "V"
			cmd("@OP") = "O"
			rSdk.open cmd, , 3, 1

			set rd = Server.CreateObject("ADODB.RecordSet")
			
			do while not rg.eof %>
			<tr class="GeneralTlt">
				<td colspan="2"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getaddressesLngStr("DtxtUDF")%><%
				Case Else
					Response.Write rg("GroupName")
				End Select %></td>
			</tr><%
			rSdk.Filter = "GroupID = " & rg("GroupID")
			do while not rSdk.eof
				InsertID = rSdk("InsertID")
				If Request("Op") <> "add" Then
					FldVal = rs(InsertID)
				Else
					FldVal = rSdk("Dflt")
				End If %>
				<tr class="GeneralTblBold2">
			            <td class="GeneralTblBold2">
			              <table border="0" cellpadding="0" cellspacing="0" width="100%">
			                <tr class="GeneralTblBold2">
			            	  <td>
			            	    <b><font size="1" face="Verdana"><%=rSdk("Descr")%><% If rSdk("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b>
			            	  </td>
			            	    <% If (rSdk("Query") = "Y" or rSdk("TypeID") = "D") and IsNull(rSdk("RTable")) Then %>
			            	    <td width="16">
			            	    	<img border="0" src="../images/<% If rSdk("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rSdk("AliasID")%>" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=Addr&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',400,250,'yes', 'yes', document.frmAddress.U_<%=rSdk("AliasID")%>)"<% End If %>>
			            	    </td>
			            	    <% End If %>
			            	</tr>
			              </table>
			            </td>
			            <td dir="ltr" class="GeneralTbl"><% If rSdk("DropDown") = "Y" or not IsNull(rSdk("RTable")) then 
			            	set rd = Server.CreateObject("ADODB.RecordSet")
							cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@TableID") = "CRD1"
							cmd("@FieldID") = rSdk("FieldID")
							rd.open cmd, , 3, 1
							 %><select size="1" name="U_<%=rSdk("AliasID")%>" class="input" style="width: 99%">
								<option></option>
								<% do while not rd.eof %>
								<option <% If Not IsNull(FldVal) Then If CStr(FldVal) = CStr(rd(0)) Then Response.Write "Selected" %> value="<%=rd(0)%>" <% If rSdk("Dflt")= rd(0) Then %>selected<% End If %>><%=myHTMLEncode(rd(1))%></option>
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
						<textarea <% If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" class="input" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=Addr&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> rows="3" onfocus="this.select()" style="width: 100%" cols="1"><% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %></textarea>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							</td>
							<td width="16">
								<img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddress.U_<%=rSdk("AliasID")%>.value = ''" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
					<% ElseIf rSdk("TypeID") = "A" and rSdk("EditType") = "I" Then %>
						<table cellpadding="2" cellspacing="0" border="0">
							<tr>
								<td><img src="../pic.aspx?filename=<% If IsNull(FldVal) Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rSdk("AliasID")%>" border="1">
								<input type="hidden" name="U_<%=rSdk("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
								<td width="16" valign="bottom"><img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddress.U_<%=rSdk("AliasID")%>.value = '';document.frmAddress.imgU_<%=rSdk("AliasID")%>.src='../pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';" style="cursor: hand"></td>
							</tr>
							<tr>
								<td colspan="2" height="22">
								<p align="center">
								<input type="button" value="<%=getaddressesLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmAddress.U_<%=rSdk("AliasID")%>, document.frmAddress.imgU_<%=rSdk("AliasID")%>,180);"></td>
							</tr>
						</table>
						<% Else
						If Not IsNull(FldVal) Then 
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
							<input <% If readOnly Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" id="U_<%=rSdk("AliasID")%>" size="<%=fldSize%>" class="input" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click();"<% End If %> onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click()"<% End If %> <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=Addr&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> value="<% If Not IsNull(FldVal) Then %><%=FldVal%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %>" <% If rSdk("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> style="width: 100%" <% If isMaxSize Then %> onkeydown="return chkMax(event, this, <%=MaxSize%>);" maxlength="<%=MaxSize%>"<% End if %>>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
								</td>
								<td width="16">
									<img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddress.U_<%=rSdk("AliasID")%>.value = ''">
								</td>
							  </tr>
							</table>
							<% End If %>
						<% End If %>
			            </td>
			          </tr>
		<%	rSdk.movenext
			loop
			rg.movenext
			loop
			End If  %>
					<tr class="GeneralTbl">
						<td align="center" colspan="2">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td>
								<input type="submit" value="<%=getaddressesLngStr("DtxtApply")%>" name="btnApply" onclick="javascript:return valFrm();">
								<input type="submit" value="<% If Request("Op") = "add" Then %><%=getaddressesLngStr("DtxtAdd")%><% Else %><%=getaddressesLngStr("DtxtSave")%><% End If %>" name="btnSave" onclick="javascript:return valFrm();">&nbsp;
								</td>
								<% If Request("Op") <> "add" and not isUpdate Then %>
								<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="submit" value="<%=getaddressesLngStr("LtxtDelete")%>" name="btnDel" onclick="javascript:return confirm('<%=getaddressesLngStr("LtxtConfDel")%>'.replace('{0}', '<%=Replace(NewAddress, "'", "\'")%>'));"></td><% End IF %>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<input type="hidden" name="Op" value='<%=Request("Op")%>'>
			<input type="hidden" name="AdresType" value='<%=Request("AdresType")%>'>
		</form>
		<% End If %>
	</table>
	<table style="width: 100%">
		<tr>
			<% If Request("Op") = "" Then %><td><input type="button" value="<%=getaddressesLngStr("LtxtAddAddress")%>" name="btnAdd" onclick="javascript:document.FormTop.Op.value='add';document.FormTop.submit();;"></td><% End If %>
			<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="button" value="<%=getaddressesLngStr("DtxtClose")%>" name="btnClose" onclick="javascript:window.close();"></td>
		</tr>
	</table>
</div>
<script language="javascript">
function valFrm()
{
	if (document.frmAddress.NewAddress.value == '')
	{
		alert('<%=getaddressesLngStr("LtxtValNam")%>');
		document.frmAddress.NewAddress.focus();
		return false;
	}
	return true;
}
function changeCountry()
{
	$.post('addressesGetData.asp', { Country: document.frmAddress.Country.value, Type: 'S' }, function(data)
		{
			var cmd = document.frmAddress.State;
			var arrData = data.split('{S}');
			for (var i = cmd.length-1;i>=1;i--)
			{
				cmd.remove(i);
			}
			
			for (var i = 0;i<arrData.length;i++)
			{
				var arrCnt = arrData[i].split('{C}');
				cmd.options[i+1] = new Option(arrCnt[1], arrCnt[0]);
			}
		});
}


function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('<%=getaddressesLngStr("DtxtValFldMaxChar")%>'.replace('{0}', FSize));
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
							alert('<%=getaddressesLngStr("DtxtValNumVal")%>');
						}
						else if (parseInt(getNumericVB(Field.value)) < 1)
						{
							Field.value = '';
							alert('<%=getaddressesLngStr("DtxtValNumMinVal")%>'.replace('{0}', '1'));
						}
						else if (parseInt(getNumericVB(Field.value)) > 2147483647)
						{
							alert('<%=getaddressesLngStr("DtxtValNumMaxVal")%>'.replace('{0}', '2147483647'));
							Field.value = 2147483647;
						}
						else if (Field.value.indexOf('<%=GetFormatDec%>') > -1)
						{
							Field.value = '';
							alert('<%=getaddressesLngStr("DtxtValNumValWhole")%>');
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
					alert('<%=getaddressesLngStr("DtxtValNumVal")%>');
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

<% If Request("Op") <> "" Then %>
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
	loop %>
var objField;
function datePicker(page, w, h, s, r, o) 
{
objField = o
OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}
function setTimeStamp(Action, varDate) { 
objField.value = varDate }
<% End If %>
<% End If %>
</script>
</body>

</html>
