<% addLngPathStr = "activity/" %>
<!--#include file="lang/activityGeneral.asp" -->
<script language="javascript">var txtValNumVal = '<%=getactivityGeneralLngStr("DtxtValNumVal")%>';</script>
<script language="javascript" src="activity/addActivity.js"></script>

<head>
<style type="text/css">
.style1 {
				font-family: Verdana;
				font-size: xx-small;
}
.style2 {
				background-color: #75ACFF;
}
.style3 {
				font-family: Verdana;
				font-size: xx-small;
				background-color: #75ACFF;
}
.style4 {
				font-family: Verdana;
}
.style5 {
				font-size: xx-small;
}
</style>
</head>

<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetActivityGeneralData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ID") = Session("ActRetVal")
If Session("ActReadOnly") Then cmd("@ReadOnly") = "Y"
set rs = cmd.execute()

ReadOnly = Session("ActReadOnly")

ClgCode = rs("ClgCode")
Action = rs("Action")

If Request.Form.Count = 0 Then
	Status = rs("Status")
	Location = rs("Location")
	Recontact = FormatDate(rs("Recontact"), False)
	BeginTimeH = rs("BeginTimeH")
	BeginTimeM = rs("BeginTimeM")
	BeginTimeS = rs("BeginTimeS")
	endDate = FormatDate(rs("endDate"), False)
	ENDTimeH = rs("ENDTimeH")
	ENDTimeM = rs("ENDTimeM")
	ENDTimeS = rs("ENDTimeS")
	Duration = rs("Duration")
	DurType = rs("DurType")
	Reminder = rs("Reminder") = "Y"
	RemQty = rs("RemQty")
	RemType = rs("RemType")
	tentative = rs("tentative") = "Y"
	Inactive = rs("Inactive") = "Y"
	DocType = rs("DocType")
	DocNum = rs("DocNum")
	DocEntry = rs("DocEntry")
Else
	Status = Request("Status")
	Location = Request("Location")
	Recontact = Request("Recontact")
	BeginTimeH = Request("BeginTimeH")
	BeginTimeM = Request("BeginTimeM")
	BeginTimeS = Request("BeginTimeS")
	endDate = Request("endDate")
	ENDTimeH = Request("ENDTimeH")
	ENDTimeM = Request("ENDTimeM")
	ENDTimeS = Request("ENDTimeS")
	Duration = Request("Duration")
	DurType = Request("DurType")
	Reminder = Request("Reminder") = "Y"
	RemQty = Request("RemQty")
	RemType = Request("RemType")
	tentative = Request("tentative") = "Y"
	Inactive = Request("Inactive") = "Y"
	DocType = Request("DocType")
	DocNum = Request("DocNum")
	DocEntry = Request("DocEntry")

	If Request("Source") = "dur" and Duration = "" Then Duration = Request("DurationUndo")
	myType = Request("DurType")
	Select Case durType
	Case "M"
		myType = "n"
	Case "H"
		myType = "h"
	Case "D"
		myType = "d"
	End Select
	
	Select Case Request("Source")
		Case "beginT", "dur"
			calc = DateAdd(myType, Duration, getActBeginDateTime(Recontact, BeginTimeH, BeginTimeM, BeginTimeS))
			
			endDate = FormatDate(calc, False)
			ENDTimeH = Hour(calc)
			If ENDTimeH = 0 Then
				ENDTimeH = 12
				EndTimeS = "AM"
			ElseIf ENDTimeH > 12 Then
				ENDTimeH = ENDTimeH - 12
				EndTimeS = "PM"
			ElseIf ENDTimeH = 12 Then
				EndTimeS = "PM"
			Else
				EndTimeS = "AM"
			End If
			
			EndTimeM = Minute(calc)
		Case "endT"
			calc = DateAdd(myType, Duration*-1, getActBeginDateTime(endDate, ENDTimeH, ENDTimeM, ENDTimeS))
			
			Recontact = FormatDate(calc, False)
			BeginTimeH = Hour(calc)
			If BeginTimeH = 0 Then
				BeginTimeH = 12
				BeginTimeS = "AM"
			ElseIf BeginTimeH > 12 Then
				BeginTimeH = BeginTimeH - 12
				BeginTimeS = "PM"
			ElseIf BeginTimeH = 12 Then
				BeginTimeS = "PM"
			Else
				BeginTimeS = "AM"
			End If
			
			BeginTimeM = Minute(calc)
	End Select
	
	If Request("editVar") = "DocNum" and InStr(Request("DocNum"), "*") = 0 Then
		sql = "select top 1 DocEntry, DocNum from "
		Select Case Request("DocType")
			Case 23
				sql = sql & "OQUT"
			Case 17
				sql = sql & "ORDR"
			Case 15
				sql = sql & "ODLN"
			Case 16
				sql = sql & "ORDN"
			Case 13
				sql = sql & "OINV"
			Case 14
				sql = sql & "ORIN"
			Case 203
				sql = sql & "ODPI"
			Case 24
				sql = sql & "ORCT"
			Case 46
				sql = sql & "OVPM"
			Case 67
				sql = sql & "OWTR"
		End Select
		sql = sql & " where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DocNum like '" & Request("DocNum") & "%'"
		set rd = conn.execute(sql)
		If Not rd.Eof Then
			DocEntry = rd("DocEntry")
			DocNum = rd("DocNum")
		Else
			DocEntry = ""
			DocNum = ""
			ErrDocNum = True
		End If
	End If
End If

If BeginTimeH = 0 Then BeginTimeH = 12
If ENDTimeH = 0 Then ENDTimeH = 12

Function getActBeginDateTime(ByVal Value, ByVal bH, ByVal bM, ByVal bS)
	strDate = SaveSqlDate(Value) & " " 
	
	If bS = "PM" Then 
		If bH < 12 Then
			bH = bH + 12
		ElseIf bH = 24 Then 
			bH = 0
		End If
	ElseIf bS = "AM" Then
		If bH = 12 Then bH = 0
	End If
	
	strDate = strDate & Right("0" & bH, 2)
	
	strDate = strDate & ":"
	
	strDate = strDate & Right("0" & bM, 2)
	
	strDate = strDate & ":00"
	
	getActBeginDateTime = CDate(strDate)
End Function

If ErrDocNum Then %>
<script type="text/javascript">
<!--
alert('<%=getactivityGeneralLngStr("LtxtValDocNum")%>'.replace('{0}', '<%=Request("DocNum")%>'));
//-->
</script>
<% End IF %>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#9BC4FF">
	<tr>
		<td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		<table cellpadding="0" border="0">
			<tr>
				<td><img src="images/icon_activity_<% If Not IsNull(ClgCode) Then %>S<% Else %>O<% End If %>.gif"></td>
				<td><b><font face="Verdana" size="1"><%=getactivityGeneralLngStr("DtxtActivity")%>&nbsp;#<% If Not IsNull(ClgCode) Then Response.Write ClgCode Else Response.Write Session("ActRetVal") %>&nbsp;-&nbsp;<%=getactivityGeneralLngStr("LtxtGeneral")%></font></b></td>
			</tr>
		</table>
		</td>
	</tr>
	<% If Request("editVar") <> "DocNum" or Request("editVar") = "DocNum" and InStr(Request("DocNum"), "*") = 0 Then %>
	<form name="frmGeneral" method="post" action="activity/actSubmit.asp">
	<tr>
		<td width="100%">
		<!--#include file="activityMenu.asp"--></td>
	</tr>
	<input type="hidden" name="cmd" value="general">
	<input type="hidden" name="Source" value="">
	<input type="hidden" name="editVar" value="">
	<input type="hidden" name="System" value="Y">
	<input type="hidden" name="returnCmd" value="activityGeneral">
	<tr>
		<td>
		<table style="width: 240px;" cellspacing="0" cellpadding="0">
			<tr>
				<td>
					<table style="width: 100%">
					<% If Action = "T" Then %>
					<tr>
						<td class="style2" style="height: 22px"><span id="txtStatus" class="style1"><strong><%=getactivityGeneralLngStr("LtxtStatus")%></strong></span></td>
						<td style="height: 22px"><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="Status" id="Status" style="font-size: 10px; font-family: Verdana; height: 18px;">
						<% 
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetActStatus" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						set rd = cmd.execute()
						do while not rd.eof %>
						<option <% If CStr(rd("statusID")) = CStr(Status) Then %>selected<% End If %> value='<%=rd("statusID")%>'><%=myHTMLEncode(rd("name"))%></option>
						<% rd.movenext
						loop
						%></select></td>
					</tr>
					<% Else %>
					<input type="hidden" name="Status" value="<%=Status%>">
					<% End If %>
					<% rsdf.Filter = "FieldID = -12"
			        If Not rsdf.Eof Then %>
					<% If Action <> "E" then %>
					<tr>
						<td class="style3"><span id="txtLocation" class="style1"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityGeneralLngStr("DtxtLocation")%><% End If %></strong></span></td>
						<td><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="Location" id="Location" style='font-size: 10px; font-family: Verdana;'>
						<option></option>
						<% set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetActLocations" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						set rd = cmd.execute()
						do while not rd.eof %>
						<option <% If CStr(rd("Code")) = CStr(Location) Then %>selected<% End If %> value='<%=rd("Code")%>'><%=myHTMLEncode(rd("name"))%></option>
						<% rd.movenext
						loop
						%></select></td>
					</tr>
					<% Else %>
					<input type="hidden" name="Location" value="<%=Location%>">
					<% End If %>
					<% End If %>
					<% 
					rsdf.Filter = "FieldID = -14"
					If Not rsdf.eof Then %>
					<tr>
						<td colspan="2" class="style3"><span id="txtBeginTime"><strong><span class="style1"><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><% Select Case Action
						Case "T" %><%=getactivityGeneralLngStr("LtxtStartDate")%>
						<% Case "E" %><%=getactivityGeneralLngStr("LtxtTime")%>
						<% Case Else %><%=getactivityGeneralLngStr("LtxtStartTime")%>
						<% End Select %><% End If %></span></strong></span></td>
					</tr>
					<tr>
						<td colspan="2">
						<table cellpadding="0" cellspacing="2" border="0">
						<tr>
							<td><% If Not ReadOnly Then %><a href="#" onclick="javascript:getCal('Recontact');"><img border="0" src="images/cal.gif" id="btnBeginDate"></a><% Else %>&nbsp;<% End If %></td>
							<td><input class="input" type="text" name="Recontact" id="Recontact" readonly size="8" value='<%=Recontact%>' <% If Not ReadOnly Then %>onclick="btnBeginDate.click();"<% End If %> onchange="javascript:changeTime('beginT');"> </td>
							<td>&nbsp;</td>
							<td id="tdBeginTime" dir="ltr">
							<% If Action <> "T" Then %><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="BeginTimeH" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('beginT');"><% For i = 1 to 12 %>
							<option <% If i = CInt(BeginTimeH) Then %>selected<% End If %> value='<%=Right("0" & i, 2)%>'><%=Right("0" & i, 2)%></option>
							<% Next %></select><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="BeginTimeM" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('beginT');"><% For i = 0 to 60 %>
							<option <% If i = CInt(BeginTimeM) Then %>selected<% End If %> value='<%=Right("0" & i, 2)%>'><%=Right("0" & i, 2)%></option>
							<% Next %></select><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="BeginTimeS" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('beginT');">
							<option <% If BeginTimeS = "AM" Then %>selected<% End If %> value="AM">AM</option>
							<option <% If BeginTimeS = "PM" Then %>selected<% End If %> value="PM">PM</option>
							</select> 
							<% Else %>
							<input type="hidden" name="BeginTimeH" value="<%=BeginTimeH%>">
							<input type="hidden" name="BeginTimeM" value="<%=BeginTimeM%>">
							<input type="hidden" name="BeginTimeS" value="<%=BeginTimeS%>">
							<% End If %></td>
						</tr>
						</table>
						</td>
					</tr>
					<% 
					End If
					rsdf.Filter = "FieldID = -16"
					If Not rsdf.eof Then %>
					<% If Action <> "E" Then %>
					<tr>
						<td colspan="2" class="style2"><span id="txtENDTime" class="style1"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><% If Action <> "T" Then %><%=getactivityGeneralLngStr("LtxtEndTime")%><% Else %><%=getactivityGeneralLngStr("LtxtDueDate")%><% End If %><% End If %></strong></span></td>
					</tr>
					<tr>
						<td colspan="2">
						<table cellpadding="0" cellspacing="2" border="0" id="tblENDTime">
						<tr>
							<td><% If Not ReadOnly Then %><a href="#" onclick="javascript:getCal('endDate');"><img border="0" src="images/cal.gif" id="btnEndDate"></a><% Else %>&nbsp;<% End If %></td>
							<td id="tdENDDate"><input class="input" type="text" name="endDate" readonly id="endDate" size="8" value='<%=endDate%>' <% If Not ReadOnly Then %>onclick="btnEndDate.click();"<% End If %> onchange="javascript:changeTime('endT');"> </td>
							<td>&nbsp;</td>
							<td id="tdENDTime" dir="ltr">
							<% If Action <> "T" Then %><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="ENDTimeH" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('endT');"><% For i = 1 to 12 %>
							<option <% If i = CInt(ENDTimeH) Then %>selected<% End If %> value='<%=Right("0" & i, 2)%>'><%=Right("0" & i, 2)%></option>
							<% Next %></select><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="ENDTimeM" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('endT');"><% For i = 0 to 60 %>
							<option <% If i = CInt(ENDTimeM) Then %>selected<% End If %> value='<%=Right("0" & i, 2)%>'><%=Right("0" & i, 2)%></option>
							<% Next %></select><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="ENDTimeS" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('endT');">
							<option <% If ENDTimeS = "AM" Then %>selected<% End If %> value="AM">AM</option>
							<option <% If ENDTimeS = "PM" Then %>selected<% End If %> value="PM">PM</option>
							</select> 
							<% Else %>
							<input type="hidden" name="ENDTimeH" value="<%=ENDTimeH%>">
							<input type="hidden" name="ENDTimeM" value="<%=ENDTimeM%>">
							<input type="hidden" name="ENDTimeS" value="<%=ENDTimeS%>">
							<% End If %></td>
						</tr>
						</table>
						</td>
					</tr>
					<% End If %>
					<% End If %>
					<% 
					rsdf.Filter = "FieldID = -17"
					If Not rsdf.eof Then
					If Not (Action = "T" or Action = "E") Then %>
					<tr>
						<td class="style3"><span id="txtDuration" class="style1"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityGeneralLngStr("DtxtDuration")%><% End If %></strong></span></td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0" id="tblDuration">
							<tr>
								<td><input type="hidden" name="DurationUndo" value="<%=Duration%>"><input <% If ReadOnly Then %>readonly<% End If %> class="input" type="text" name="Duration" size="6" value='<%=Duration%>' maxlength="20" style="text-align: right" onchange="javascript:changeTime('dur');"> </td>
								<td>&nbsp;</td>
								<td><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="DurType" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeTime('dur');">
								<option <% If DurType = "M" Then %>selected<% End If %> value="M"><%=getactivityGeneralLngStr("LtxtMinutes")%></option>
								<option <% If DurType = "H" Then %>selected<% End If %> value="H"><%=getactivityGeneralLngStr("LtxtHours")%></option>
								<option <% If DurType = "D" Then %>selected<% End If %> value="D"><%=getactivityGeneralLngStr("LtxtDays")%></option>
								</select> </td>
							</tr>
						</table>
						</td>
					</tr>
					<% Else %>
					<input type="hidden" name="Duration" value="<%=Duration%>">
					<input type="hidden" name="DurType" value="<%=DurType%>">
					<% End If %>
					<% End If %>
					<% rsdf.Filter = "FieldID = -19"
		        	If Not rsdf.Eof Then %>
					<% If Action <> "T" Then %>
					<tr>
						<td class="style3"><span id="optReminder"><strong><span class="style5"><span class="style4"><input <% If ReadOnly Then %>disabled<% End If %> type="checkbox" name="Reminder" id="Reminder" <% If Reminder Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid"></span></span></strong><span class="style1"><label for="Reminder"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityGeneralLngStr("DtxtReminder")%><% End If %></strong></label></span></span></td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0" id="trReminder">
						<tr>
							<td><input type="hidden" name="RemQtyUndo" value="<%=RemQty%>"><input <% If ReadOnly Then %>readonly<% End If %> class="input" type="text" name="RemQty" size="6" value='<%=RemQty%>' maxlength="20" style="text-align: right" onchange="javascript:changeReminder();"> </td>
							<td>&nbsp;</td>
							<td><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="RemType" style="font-size: 10px; font-family: Verdana">
							<option <% If RemType = "M" Then %>selected<% End If %> value="M"><%=getactivityGeneralLngStr("LtxtMinutes")%></option>
							<option <% If RemType = "H" Then %>selected<% End If %> value="H"><%=getactivityGeneralLngStr("LtxtHours")%></option>
							</select></td>
						</tr>
						</table>
						</td>
					</tr>
					<% Else %>
					<input type="hidden" name="Reminder" value="<% If Reminder Then %>Y<% End If %>">
					<input type="hidden" name="RemQty" value="<%=RemQty%>">
					<input type="hidden" name="RemType" value="<%=RemType%>">
					<% End If %>
					<% End If %>
					<% rsdf.Filter = "FieldID = -22"
			        If Not rsdf.Eof Then %>
					<tr>
						<td colspan="2">
						<table style="width: 100%">
							<tr>
								<td class="style3" style="width: 50%"><% If Action = "M" Then %><span id="optTentative"><strong><span class="style4"><span class="style5"><input <% If ReadOnly Then %>disabled<% End If %> type="checkbox" name="tentative" id="tentative" <% If tentative Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid"></span></span><span class="style1"><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityGeneralLngStr("DtxtPosible")%><% End If %></span></strong></span><% Else %><input type="hidden" name="tentative" value="<% If tentative Then %>Y<% End If %>"><% End If %></td>
								<% rsdf.Filter = "FieldID = -23"
							        If Not rsdf.Eof Then %>
								<td class="style2" style="width: 50%"><span class="style4"><span class="style5"><strong><input <% If ReadOnly Then %>disabled<% End If %> type="checkbox" name="Inactive" id="Inactive" <% If Inactive Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid"></strong></span></span><span class="style1"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityGeneralLngStr("DtxtInactive")%><% End If %></strong></span></td><% End If %>
							</tr>
						</table>
						</td>
					</tr>
					<% End If %>
					<% rsdf.Filter = "FieldID = -25"
					If Not rsdf.Eof Then %>
					<tr>
						<td colspan="2" class="style3"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityGeneralLngStr("LtxtLinkDoc")%><% End If %></strong></td>
					</tr>
					<tr>
						<td class="style3"><strong><%=getactivityGeneralLngStr("LtxtDocType")%></strong></td>
						<td><select <% If ReadOnly Then %>disabled<% End If %> size="1" name="DocType" onchange="javascript:changeDocType();">
						<option value="-1"></option>
						<optgroup label="<%=getactivityGeneralLngStr("LtxtSale")%>">
						<option <% If DocType = 23 Then %>selected<% End If %> value="23"><%=txtQuotes%></option>
						<option <% If DocType = 17 Then %>selected<% End If %> value="17"><%=txtOrdrs%></option>
						<option <% If DocType = 15 Then %>selected<% End If %> value="15"><%=txtOdlns%></option>
						<option <% If DocType = 16 Then %>selected<% End If %> value="16"><%=txtOrnds%></option>
						<option <% If DocType = 13 Then %>selected<% End If %> value="13"><%=txtInvs%></option>
						<option <% If DocType = 14 Then %>selected<% End If %> value="14"><%=txtOrins%></option>
						<option <% If DocType = 203 Then %>selected<% End If %> value="203"><%=getactivityGeneralLngStr("LtxtInvDownPay")%></option>
						</optgroup>
						<optgroup label="<%=getactivityGeneralLngStr("LtxtBanks")%>">
						<option <% If DocType = 24 Then %>selected<% End If %> value="24"><%=txtRcts%></option>
						<option <% If DocType = 46 Then %>selected<% End If %> value="46"><%=txtOvpms%></option>
						</optgroup>
						<optgroup label="<%=getactivityGeneralLngStr("LtxtInventory")%>">
						<option <% If DocType = 67 Then %>selected<% End If %> value="67"><%=getactivityGeneralLngStr("LtxtInvTrans")%></option>
						</optgroup>
						</select></td>
					</tr>
					<tr>
						<td class="style3"><strong><%=getactivityGeneralLngStr("LtxtDocNum")%></strong></td>
						<td><input type="hidden" name="DocEntry" value="<%=DocEntry%>">
						<input <% If ReadOnly Then %>readonly<% End If %> type="text" name="DocNum" <% If DocType = -1 Then %>disabled<% End If %> size="20" value='<%=DocNum%>' onchange="javascript:changeDocNum();" style="text-align: right"></td>
					</tr>
					<% End If %>
					<% If Not IsNull(rs("parentType")) Then %>
					<tr>
						<%
						Select Case rs("parentType")
						Case 97
						source = getactivityGeneralLngStr("DtxtSalesOportunity")
						Case 191
						source = getactivityGeneralLngStr("DtxtServiceCall")
						End Select
						source = source & " #" & rs("parentID")
						%>
						<td class="style3"><strong><%=getactivityGeneralLngStr("DtxtSource")%></strong></td>
						<td><input type="text" name="Source" disabled size="30" value="<%=source%>"></td>
					</tr>
					<% End If %>
				</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<!--#include file="activityBottom.asp"-->
	</form>
	<% Else %>
	<tr>
		<td>
		<table border="0" width="490" id="table2" cellpadding="0">
			<form name="frmDocNum" action="operaciones.asp" method="post">
			<% 
			sql = "select DocEntry, DocNum, DocDate, Comments from "
			Select Case Request("DocType")
			Case 23
			sql = sql & "OQUT"
			Case 17
			sql = sql & "ORDR"
			Case 15
			sql = sql & "ODLN"
			Case 16
			sql = sql & "ORDN"
			Case 13
			sql = sql & "OINV"
			Case 14
			sql = sql & "ORIN"
			Case 203
			sql = sql & "ODPI"
			Case 24
			sql = sql & "ORCT"
			Case 46
			sql = sql & "OVPM"
			Case 67
			sql = sql & "OWTR"
			End Select
			sql = sql & " where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DocNum like '" & Replace(Request("DocNum"), "*", "%") & "' order by DocNum asc"
			set rd = conn.execute(sql)
			if not rd.eof then %>
			<tr class="style2">
				<td><%=getactivityGeneralLngStr("LtxtDocNum")%></td>
				<td><%=getactivityGeneralLngStr("DtxtDate")%></td>
				<td><%=getactivityGeneralLngStr("LtxtComments")%></td>
			</tr>
			<% 
			do while not rd.eof %>
			<tr class="CSpecialTbl">
				<td><a href="#" class="LinkCSpecial" onclick="setDoc(<%=rd("DocNum")%>,<%=rd("DocEntry")%>);"><%=rd("DocNum")%></a></td>
				<td><a href="#" class="LinkCSpecial" onclick="setDoc(<%=rd("DocNum")%>,<%=rd("DocEntry")%>);"><%=FormatDate(rd("DocDate"), False)%></a></td>
			<td><%=rd("Comments")%></td>
			</tr>
			<% rd.movenext
			loop
			else %>
			<tr class="CSpecialTbl">
				<td>
				<p align="center"><%=getactivityGeneralLngStr("DtxtNoData")%></td>
			</tr>
		<% End If %>
		<% For each itm in Request.Form %>
		<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
		<% Next %>
		</form>
		</table>
	</td>
	</tr>
	<script type="text/javascript">
	function setDoc(DocNum, DocEntry)
	{
	document.frmDocNum.DocNum.value = DocNum;
	document.frmDocNum.DocEntry.value = DocEntry;
	document.frmDocNum.editVar.value = '';
	document.frmDocNum.submit();
	}
	</script>
	<% End If %>
</table>
</center></div>


