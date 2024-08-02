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
<!--#include file="lang/contacts.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl" <% End If %>="">

<%
set rs = Server.CreateObject("ADODB.RecordSet")
set rd = Server.CreateObject("ADODB.RecordSet")

isDefault = False
If Request("Op") <> "" and Request("Op") <> "add" Then
	cmd.CommandText = "DBOLKGetCrdCntData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("CrdRetVal")
	cmd("@LineNum") = CInt(Request("Op"))
	set rs = cmd.execute()
	EnableSDK = rs("EnableSDK") = "Y"
	NewName = rs("NewName")
	Title = rs("Title")
	Position = rs("Position")
	Address = rs("Address")
	Tel1 = rs("tel1")
	Tel2 = rs("tel2")
	Cellolar = rs("Cellolar")
	Fax = rs("Fax")
	EMail = rs("E_MailL")
	Pager = rs("Pager")
	Notes1 = rs("Notes1")
	Notes2 = rs("Notes2")
	Password = rs("Password")
	BirthPlace = rs("BirthPlace")
	BirthDate = rs("BirthDate")
	Gender = rs("Gender")
	Profession = rs("Profession")
	isDefault = rs("IsDefault") = "Y"
	isUpdate = rs("Command") = "U"
ElseIf Request("Op") = "add" Then
	cmd.CommandText = "DBOLKGetEnableSDK" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = "OCPR"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rs = cmd.execute()
	EnableSDK = rs(0) = "Y"
End If
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getcontactsLngStr("LtxtContacts")%></title>
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
			alert('<%=getcontactsLngStr("DtxtValNumVal")%>');
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
			<td><%=getcontactsLngStr("LtxtContacts")%></td>
		</tr>
		<form method="POST" action="contacts.asp" name="FormTop">
			<input type="hidden" name="pop" value="Y">
			<input type="hidden" name="AddPath" value="../">
			<% 
			cmd.CommandText = "DBOLKGetCrdCnts"
			cmd.Parameters.Refresh()
			cmd("@LogNum") = Session("CrdRetVal")
			set rd = cmd.execute()
			If Request("Op") <> "" Then %>
			<tr class="GeneralTlt">
				<td><%=getcontactsLngStr("DtxtContact")%>:
				<select size="1" name="Op" onchange="submit()">
				<option value=""><%=getcontactsLngStr("LtxtDoSel")%></option>
				<option <% If Request("Op") = "add" then response.write "selected "%>value="add">
				<%=getcontactsLngStr("LtxtAddContact")%></option>
				<% enableChkDef = False
				do while not rd.eof
				enableChkDef = True %>
				<option <% If CStr(Request("Op")) = CStr(rd("LineNum")) then response.write "selected "%>value='<%=rd("LineNum")%>'>
				<%=myHTMLEncode(rd("NewName"))%><% If rd("IsDefault") = "Y" Then %>&nbsp;(<%=getcontactsLngStr("DtxtDefault")%>)<% End If %></option>
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
						<td><%=getcontactsLngStr("DtxtName")%></td>
						<td><%=getcontactsLngStr("LtxtPosition")%></td>
						<td><%=getcontactsLngStr("DtxtPhone")%></td>
					</tr>
					<% enableChkDef = False
					do while not rd.eof
					enableChkDef = True %>
					<tr class="<% If rd("IsDefault") = "N" Then %>GeneralTbl<% Else %>CanastaTblExpense<% End If %>">
						<td width="15"><a href="#" onclick="javascript:document.FormTop.Op.value='<%=rd("LineNum")%>';submit();"><img border="0" src="../design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
						<td><% If rd("Title") <> "" Then %><%=rd("Title")%>&nbsp;<% End If %><%=rd("NewName")%></td>
						<td><%=rd("Position")%></td>
						<td><%=rd("Tel1")%></td>
					</tr>
					<% rd.movenext
					loop %>
				</table>
				</td>
			</tr>
			<input type="hidden" name="Op" value="">
			<% End If %>
		</form>
		<% If Request("Op") <> "" Then %>
		<form method="POST" action="submitContacts.asp" name="frmContact">
			<tr>
				<td>
				<table border="0" cellpadding="0" width="100%" id="table2">
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtName")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="NewName" size="50" maxlength="50" value="<%=myHTMLEncode(NewName)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtTitle")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Title" size="10" maxlength="10" value="<%=myHTMLEncode(Title)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtPosition")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Position" size="50" maxlength="90" value="<%=myHTMLEncode(Position)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtAddress")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Address" size="50" maxlength="90" value="<%=myHTMLEncode(Address)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtPhone")%>&nbsp;1</td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Tel1" size="20" maxlength="20" value="<%=myHTMLEncode(Tel1)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtPhone")%>&nbsp;2</td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Tel2" size="20" maxlength="20" value="<%=myHTMLEncode(Tel2)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtMobile")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Cellolar" size="20" maxlength="20" value="<%=myHTMLEncode(Cellolar)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtFax")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Fax" size="20" maxlength="20" value="<%=myHTMLEncode(Fax)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtEMail")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="EMail" size="50" maxlength="100" value="<%=myHTMLEncode(EMail)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtPager")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Pager" size="30" maxlength="30" value="<%=myHTMLEncode(Pager)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtObservations")%>&nbsp;1</td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Notes1" size="50" maxlength="100" value="<%=myHTMLEncode(Notes1)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtObservations")%>&nbsp;2</td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Notes2" size="50" maxlength="100" value="<%=myHTMLEncode(Notes2)%>"></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("DtxtPwd")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Password" size="8" maxlength="8" value="<%=myHTMLEncode(Password)%>"></td>
					</tr>
					<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then %>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtPOB")%></td>
						<td height="18" class="GeneralTbl">
						<select size="1" name="BirthPlace">
						<option value="0"><%=getcontactsLngStr("LtxtSel")%></option>
						<%
						set rd = Server.CreateObject("ADODB.RecordSet")
						cmd.CommandText = "DBOLKGetCountries" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						rd.open cmd, , 3, 1
						do while not rd.eof
						%>
						<option value="<%=rd("code")%>" <% If BirthPlace = rd("code") Then %>selected<% End If %>>
						<%=myHTMLEncode(rd.Fields("name"))%>
						</option>
						<% rd.movenext
						loop %>
						</select></td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtDOB")%></td>
						<td height="18" class="GeneralTbl">
						<table cellpadding="0" cellspacing="0" border="0">
			 			<tr>
			 				<td><input type="text" name="BirthDate" id="BirthDate" size="12" value="<%=FormatDate(BirthDate, False)%>" readonly onclick="javascript:btnBirthDate.click();"></td>
			 				<td>&nbsp;<img border="0" src="../images/cal.gif" id="btnBirthDate"></td>
			 			</tr>
			 		</table>
					</td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtGender")%></td>
						<td height="18" class="GeneralTbl">
						<select size="1" name="Gender">
						<option value=""><%=getcontactsLngStr("LtxtSel")%></option>
						<option value="M" <% If Gender = "M" Then %>selected<% End If %>><%=getcontactsLngStr("DtxtMale")%></option>
						<option value="F" <% If Gender = "F" Then %>selected<% End If %>><%=getcontactsLngStr("DtxtFemale")%></option>
						</select>
					</td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getcontactsLngStr("LtxtProfesion")%></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Profession" maxlength="50" size="50" value="<%=myHTMLEncode(Profession)%>"></td>
					</tr>
					<% End If %>
					<tr class="GeneralTblBold2">
						<td>&nbsp;</td>
						<td height="18" class="GeneralTbl">
						<input type="checkbox" name="SetDef" id="SetDef" <% If Not enableChkDef or isDefault Then %>checked disabled<% End IF %> value="Y" style="background: background-image; border: 0px solid"><label for="SetDef"><%=getcontactsLngStr("LtxtSetAsDef")%></label></td>
					</tr>
		<% 
		set rSdk = Server.CreateObject("ADODB.RecordSet")
		If EnableSDK Then
			set rg = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@TableID") = "OCPR"
			cmd("@UserType") = "V"
			cmd("@OP") = "O"
			set rg = cmd.execute()
		
			set rSdk = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@TableID") = "OCPR"
			cmd("@UserType") = "V"
			cmd("@OP") = "O"
			rSdk.open cmd, , 3, 1
			
			set rd = Server.CreateObject("ADODB.RecordSet")
			do while not rg.eof %>

			<tr class="GeneralTlt">
				<td colspan="2"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getcontactsLngStr("DtxtUDF")%><%
				Case Else
					Response.Write rg("GroupName")
				End Select %>
				</td>
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
			            	    	<img border="0" src="../images/<% If rSdk("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rSdk("AliasID")%>" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=Cnt&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',400,250,'yes', 'yes', document.frmContact.U_<%=rSdk("AliasID")%>)"<% End If %>>
			            	    </td>
			            	    <% End If %>
			            	</tr>
			              </table>
			            </td>
			            <td dir="ltr" bgcolor="#EAF5FF"><% If rSdk("DropDown") = "Y" or not IsNull(rSdk("RTable")) then 
							set rd = Server.CreateObject("ADODB.RecordSet")
							cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@TableID") = "OCPR"
							cmd("@FieldID") = rSdk("FieldID")
							rd.open cmd, , 3, 1 %><select size="1" name="U_<%=rSdk("AliasID")%>" class="input" style="width: 99%">
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
						<textarea <% If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" class="input" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=Cnt&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> rows="3" onfocus="this.select()" style="width: 100%" cols="1"><% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %></textarea>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							</td>
							<td width="16">
								<img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmContact.U_<%=rSdk("AliasID")%>.value = ''" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
					<% ElseIf rSdk("TypeID") = "A" and rSdk("EditType") = "I" Then %>
						<table cellpadding="2" cellspacing="0" border="0">
							<tr>
								<td><img src="../pic.aspx?filename=<% If IsNull(FldVal) Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rSdk("AliasID")%>" border="1">
								<input type="hidden" name="U_<%=rSdk("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
								<td width="16" valign="bottom"><img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmContact.U_<%=rSdk("AliasID")%>.value = '';document.frmContact.imgU_<%=rSdk("AliasID")%>.src='../pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';" style="cursor: hand"></td>
							</tr>
							<tr>
								<td colspan="2" height="22">
								<p align="center">
								<input type="button" value="<%=getcontactsLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmContact.U_<%=rSdk("AliasID")%>, document.frmContact.imgU_<%=rSdk("AliasID")%>,180);"></td>
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
							<input <% If readOnly Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" id="U_<%=rSdk("AliasID")%>" size="<%=fldSize%>" class="input" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click();"<% End If %> onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click()"<% End If %> <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=Cnt&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> value="<% If Not IsNull(FldVal) Then %><%=FldVal%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %>" <% If rSdk("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> style="width: 100%" <% If isMaxSize Then %> onkeydown="return chkMax(event, this, <%=MaxSize%>);" maxlength="<%=MaxSize%>"<% End if %>>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
								</td>
								<td width="16">
									<img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmContact.U_<%=rSdk("AliasID")%>.value = ''">
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
								<input type="submit" value="<%=getcontactsLngStr("DtxtApply")%>" name="btnApply" onclick="javascript:return valFrm();">
								<input type="submit" value="<% If Request("Op") = "add" Then %><%=getcontactsLngStr("DtxtAdd")%><% Else %><%=getcontactsLngStr("DtxtSave")%><% End If %>" name="btnSave" onclick="javascript:return valFrm();">&nbsp;
								</td>
								<% If Request("Op") <> "add" and not isUpdate Then %>
								<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="submit" value="<%=getcontactsLngStr("LtxtDelete")%>" name="btnDel" onclick="javascript:return confirm('<%=getcontactsLngStr("LtxtConfDel")%>'.replace('{0}', '<%=Replace(NewAddress, "'", "\'")%>'));"></td><% End IF %>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<input type="hidden" name="Op" value='<%=Request("Op")%>'>
		</form>
		<% End If %>
	</table>
	<table style="width: 100%">
		<tr>
			<% If Request("Op") = "" Then %><td><input type="button" value="<%=getcontactsLngStr("LtxtAddContact")%>" name="btnAdd" onclick="javascript:document.FormTop.Op.value='add';document.FormTop.submit();;"></td><% End If %>
			<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="button" value="<%=getcontactsLngStr("DtxtClose")%>" name="btnClose" onclick="javascript:window.close();"></td>
		</tr>
	</table>
</div>
<script type="text/javascript">
function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('<%=getcontactsLngStr("DtxtValFldMaxChar")%>'.replace('{0}', FSize));
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
							alert('<%=getcontactsLngStr("DtxtValNumVal")%>');
						}
						else if (parseInt(getNumericVB(Field.value)) < 1)
						{
							Field.value = '';
							alert('<%=getcontactsLngStr("DtxtValNumMinVal")%>'.replace('{0}', '1'));
						}
						else if (parseInt(getNumericVB(Field.value)) > 2147483647)
						{
							alert('<%=getcontactsLngStr("DtxtValNumMaxVal")%>'.replace('{0}', '2147483647'));
							Field.value = 2147483647;
						}
						else if (Field.value.indexOf('<%=GetFormatDec%>') > -1)
						{
							Field.value = '';
							alert('<%=getcontactsLngStr("DtxtValNumValWhole")%>');
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
					alert('<%=getcontactsLngStr("DtxtValNumVal")%>');
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
<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then %>
Calendar.setup({
    inputField     :    "BirthDate",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnBirthDate",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
<% End If %>
function valFrm()
{
	if (document.frmContact.NewName.value == '')
	{
		alert('<%=getcontactsLngStr("LtxtValNam")%>');
		document.frmContact.NewName.focus();
		return false;
	}
	return true;
}

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
