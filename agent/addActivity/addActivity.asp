<% addLngPathStr = "addActivity/" %>
<!--#include file="lang/addActivity.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript">
var CalendarFormat = '<%=GetCalendarFormatString%>';
var DisplayFormat = '<%=myApp.DateFormat%>';
var SelDes = '<%=SelDes%>';
var dbName = '<%=Session("olkdb")%>';
var txtErrSaveData = '<%=getaddActivityLngStr("DtxtErrSaveData")%>';
</script>
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004


set rs = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SysID") = "OCLG"
cmd("@ObsID") = "TCLG"
set rs = cmd.execute()
If rs(0) = "Y" Then Response.Redirect "configErr.asp?errCmd=Card"

set rs = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKGetActData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@LogNum") = Session("ActRetVal")
set rs = cmd.execute()

ClgCode = rs("ClgCode")
EnableSDK = rs("EnableSDK") = "Y"

 %>
<form method="POST" action="agentActivitySubmit.asp" name="frmAddActivity">
<input type="hidden" name="isUpdate" value="<% If Not IsNull(ClgCode) Then %>True<% End If %>">
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><% If IsNull(ClgCode) Then %><%=getaddActivityLngStr("LttlNewActivity")%><% Else %><%=getaddActivityLngStr("LtxtEditActivity")%> #<%=ClgCode%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddActivityLngStr("DtxtActivity")%></td>
				<td>
				<select size="1" class="input" name="Action" onchange="javascript:changeAction(this.value);">
				<option <% If rs("Action") = "C" Then %>selected<% End If %> value="C">
				<%=getaddActivityLngStr("DtxtConv")%></option>
				<option <% If rs("Action") = "M" Then %>selected<% End If %> value="M">
				<%=getaddActivityLngStr("DtxtMeeting")%></option>
				<option <% If rs("Action") = "E" Then %>selected<% End If %> value="E">
				<%=getaddActivityLngStr("DtxtNote")%></option>
				<option <% If rs("Action") = "O" Then %>selected<% End If %> value="O">
				<%=getaddActivityLngStr("DtxtOther")%></option>
				<option <% If rs("Action") = "T" Then %>selected<% End If %> value="T">
				<%=getaddActivityLngStr("DtxtTask")%></option>
				</select>
				</td>
				<td class="GeneralTblBold2">
				<%=getaddActivityLngStr("DtxtClientCode")%></td>
				<td>
				<input class="inputDis" type="text" name="CardCode" readonly size="35" value="<%=myHTMLEncode(rs("CardCode"))%>"></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddActivityLngStr("DtxtType")%></td>
				<td>
				<% 
				
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetActTypes" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				rd.open cmd, , 3, 1 %>
				<select class="input" size="1" name="CntctType" style="font-size:10px; font-family:Verdana" onchange="javascript:changeType(this.value);">
		        <% do while not rd.eof %>
		        <option <% If rd("Code") = rs("CntctType") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
				<td class="GeneralTblBold2">
				<%=getaddActivityLngStr("DtxtName")%></td>
				<td>
				<input class="inputDis" type="text" name="CardName" readonly size="50" value="<%=myHTMLEncode(rs("CardName"))%>"></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddActivityLngStr("DtxtSubject")%></td>
				<td>
<% 
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetActivitySubjects" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@Type") = rs("CntctType")
				rd.open cmd, , 3, 1 %>
				<select class="input" size="1" name="CntctSbjct" id="CntctSbjct" style="font-size:10px; font-family:Verdana;" onchange="doProc('CntctSbjct', 'N', this.value);">
				<option value=""></option>
		        <% do while not rd.eof %>
		        <option <% If rd("Code") = rs("CntctSbjct") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
				<td class="GeneralTblBold2">
				<%=getaddActivityLngStr("DtxtContact")%></td>
				<td>
				<% 
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@CardCode") = Session("UserName")
				rd.open cmd, , 3, 1  %>
				<select class="input" size="1" name="CntctCode" style="font-size:10px; font-family:Verdana" onchange="javascript:changeContact(this.value);doProc('CntctCode', 'N', this.value);">
		        <% do while not rd.eof %>
		        <option <% If rd("CntctCode") = rs("CntctCode") Then %>selected<% End If %> value="<%=rd("cntctcode")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getaddActivityLngStr("LtxtAsignedTo")%></td>
				<td>
				<%
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetSysUsers" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				rd.open cmd, , 3, 1 %>
				<select class="input" size="1" name="AttendUser" style="font-size:10px; font-family:Verdana" onchange="doProc('AttendUser', 'N', this.value);">
		        <% do while not rd.eof %>
		        <option <% If rd("Code") = rs("AttendUser") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("Name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
				<td class="GeneralTblBold2">
				<%=getaddActivityLngStr("DtxtPhone")%></td>
				<td>
				<input class="input" type="text" name="Tel" size="35" value="<%=myHTMLEncode(rs("Tel"))%>" onkeydown="return chkMax(event, this, 20);" onchange="doProc('Tel', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%">&nbsp;</td>
				<td><font size="1" face="Verdana"><b>
                <input type="checkbox" style="background: background-image; border: 0px solid" name="personal" id="personal" value="Y"<% If rs("personal") = "Y" Then %> checked <% End If %> onchange="doProc('personal', 'S', GetYesNo(this.value));"></b></font><label for="personal"><%=getaddActivityLngStr("DtxtPersonal")%></label></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getaddActivityLngStr("LtxtGeneral")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getaddActivityLngStr("DtxtCommentaries")%></td>
				<td colspan="3" style="width: 82%" class="generalTbl">
				<input class="input" style="width: 100%; " type="text" name="Details" size="35" value='<%=myHTMLEncode(rs("Details"))%>' onkeydown="return chkMax(event, this, 60);" onchange="doProc('Details', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><span id="txtBeginTime"><% If rs("Action") <> "T" and rs("Action") <> "E" Then %><%=getaddActivityLngStr("LtxtStartTime")%><% ElseIf rs("Action") = "T" Then %><%=getaddActivityLngStr("LtxtStartDate")%><% ElseIf rs("Action") = "E" Then %><%=getaddActivityLngStr("LtxtTime")%><% End If %></span></td>
				<td width="25%" class="generalTbl" style="width: 50%">
				<table cellpadding="0" cellspacing="2" border="0">
					<tr class="generalTbl">
						<td><img border="0" src="images/cal.gif" id="btnBeginDate"></td>
						<td>
						<input class="input" readonly type="text" name="Recontact" id="Recontact" size="12" value="<%=FormatDate(rs("Recontact"), False)%>" onchange="javascript:changeTime('beginT');" onclick="btnBeginDate.click();">
						</td>
						<td>&nbsp;</td>
						<td id="tdBeginTime" <% If rs("Action") = "T" Then %>style="display: none"<% End If %> dir="ltr">
						<select class="input" size="1" name="BeginTimeH" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('beginT');">
						<% For i = 1 to 12 %>
						<option <% If i = CInt(rs("BeginTimeH")) Then %>selected<% End If %> value="<%=Right("0" & i, 2)%>"><%=Right("0" & i, 2)%></option>
						<% Next %>
				        </select><select class="input" size="1" name="BeginTimeM" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('beginT');">
				        <% For i = 0 to 59 %>
						<option <% If i = CInt(rs("BeginTimeM")) Then %>selected<% End If %> value="<%=Right("0" & i, 2)%>"><%=Right("0" & i, 2)%></option>
						<% Next %>
				        </select><select class="input" size="1" name="BeginTimeS" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('beginT');">
				        <option <% If rs("BeginTimeS") = "AM" Then %>selected<% End If %> value="AM">
						AM</option>
						<option <% If rs("BeginTimeS") = "PM" Then %>selected<% End If %> value="PM">
						PM</option>
				        </select>
						</td>
					</tr>
				</table>
				</td>
				<td width="6%"><%=getaddActivityLngStr("DtxtPriority")%></td>
				<td width="57%" class="generalTbl">
				<select class="input" size="1" name="Priority" style="font-size:10px; font-family:Verdana" onchange="doProc('Priority', 'S', this.value);">
				<option <% If rs("Priority") = "L" Then %>selected<% End If %> value="L"><%=getaddActivityLngStr("DtxtLow")%></option>
				<option <% If rs("Priority") = "N" Then %>selected<% End If %> value="N"><%=getaddActivityLngStr("DtxtNormal")%></option>
				<option <% If rs("Priority") = "H" Then %>selected<% End If %> value="H"><%=getaddActivityLngStr("DtxtHigh")%></option>
		        </select></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><span id="txtENDTime" <% If rs("Action") = "E" Then %>style="display: none"<% End If %>><% If rs("Action") <> "T" Then %><%=getaddActivityLngStr("LtxtEndTime")%><% Else %><%=getaddActivityLngStr("LtxtDueDate")%><% End If %></span></td>
				<td width="25%" class="generalTbl">
				<table cellpadding="0" cellspacing="2" border="0" id="tblENDTime" <% If rs("Action") = "E" Then %>style="display: none"<% End If %>>
					<tr class="generalTbl">
						<td><img border="0" src="images/cal.gif" id="btnEndDate"></td>
						<td id="tdENDDate"><input readonly class="input" type="text" name="endDate" id="endDate" size="12" value="<%=FormatDate(rs("endDate"), False)%>" onchange="javascript:changeTime('endT');" onclick="btnEndDate.click();">
						</td>
						<td>&nbsp;</td>
						<td id="tdENDTime" <% If rs("Action") = "T" Then %>style="display: none"<% End If %> dir="ltr">
						<select class="input" size="1" name="ENDTimeH" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('endT');">
						<% For i = 1 to 12 %>
						<option <% If i = CInt(rs("ENDTimeH")) Then %>selected<% End If %> value="<%=Right("0" & i, 2)%>"><%=Right("0" & i, 2)%></option>
						<% Next %>
				        </select><select class="input" size="1" name="ENDTimeM" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('endT');">
				        <% For i = 0 to 59 %>
						<option <% If i = CInt(rs("ENDTimeM")) Then %>selected<% End If %> value="<%=Right("0" & i, 2)%>"><%=Right("0" & i, 2)%></option>
						<% Next %>
				        </select><select class="input" size="1" name="ENDTimeS" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('endT');">
				        <option <% If rs("ENDTimeS") = "AM" Then %>selected<% End If %> value="AM">
						AM</option>
						<option <% If rs("ENDTimeS") = "PM" Then %>selected<% End If %> value="PM">
						PM</option>
				        </select>
						</td>
					</tr>
				</table>
				</td>
				<td width="6%"><span id="txtLocation" <% If rs("Action") = "E" Then %>style="display: none"<% End If %>><%=getaddActivityLngStr("DtxtLocation")%></span></td>
				<td width="57%" class="generalTbl">
				<% 
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetActLocations" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = cmd.execute() %>
				<select class="input" size="1" name="Location" id="Location" style="font-size:10px; font-family:Verdana;<% If rs("Action") = "E" Then %>display: none<% End If %>" onchange="doProc('Location', 'N', this.value);">
				<option></option>
		        <% do while not rd.eof %>
		        <option <% If rd("Code") = rs("Location") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><span id="txtDuration" <% If rs("Action") = "T" or rs("Action") = "E" Then %>style="display: none"<% End If %>><%=getaddActivityLngStr("DtxtDuration")%></span></td>
				<td width="25%" class="generalTbl">
				<table cellpadding="0" cellspacing="0" border="0" id="tblDuration" <% If rs("Action") = "T" or rs("Action") = "E" Then %>style="display: none"<% End If %>>
					<tr class="generalTbl">
						<td><input type="hidden" name="DurationUndo" value="<%=rs("Duration")%>"><input class="input" type="text" name="Duration" size="11" value='<%=rs("Duration")%>' onkeydown="return chkMax(event, this, 20);" style="text-align: right" onchange="javascript:changeTime('dur');">
						</td>
						<td>&nbsp;</td>
						<td>
						<select class="input" size="1" name="DurType" style="font-size:10px; font-family:Verdana" onchange="javascript:changeTime('dur');">
				        <option <% If rs("DurType") = "M" Then %>selected<% End If %> value="M">
						<%=getaddActivityLngStr("LtxtMinutes")%></option>
						<option <% If rs("DurType") = "H" Then %>selected<% End If %> value="H">
						<%=getaddActivityLngStr("LtxtHours")%></option>
						<option <% If rs("DurType") = "D" Then %>selected<% End If %> value="D">
						<%=getaddActivityLngStr("LtxtDays")%></option>
				        </select>
						</td>
					</tr>
				</table>
				</td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="generalTbl">
				<span id="optTentative" <% If rs("Action") <> "M" Then %>style="display: none"<% End If %>>
				<input type="checkbox" name="tentative" id="tentative" <% If rs("tentative") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid" onclick="doProc('tentative', 'S', GetYesNo(this.checked));"><label for="tentative"><%=getaddActivityLngStr("DtxtPosible")%></label></span></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><span id="txtStatus" <% If rs("Action") <> "T" Then %>style="display: none"<% End If %>><%=getaddActivityLngStr("LtxtStatus")%></span></td>
				<td width="25%" class="generalTbl">
				<% 
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetActStatus" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = cmd.execute() %>
				<select class="input" size="1" name="Status" id="Status" style="font-size:10px; font-family:Verdana" <% If rs("Action") <> "T" Then %>style="display: none"<% End If %> onchange="doProc('Status', 'N', this.value);">
		        <% do while not rd.eof %>
		        <option <% If rd("statusID") = rs("Status") Then %>selected<% End If %> value="<%=rd("statusID")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="generalTbl">
				<input type="checkbox" name="Inactive" id="Inactive" <% If rs("Inactive") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid" onclick="doProc('Inactive', 'S', GetYesNo(this.checked));"><label for="Inactive"><%=getaddActivityLngStr("DtxtInactive")%></label></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%">
				<span id="optReminder" <% If rs("Action") = "T" Then %>style="display: none"<% End If %>>
				<input type="checkbox" name="Reminder" id="Reminder" <% If rs("Reminder") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid" onclick="doProc('Reminder', 'S', GetYesNo(this.checked));"><label for="Reminder"><%=getaddActivityLngStr("DtxtReminder")%></label></span></td>
				<td width="25%" class="generalTbl">
				<table cellpadding="0" cellspacing="0" border="0" id="trReminder" <% If rs("Action") = "T" Then %>style="display: none"<% End If %>>
					<tr class="generalTbl">
						<td><input type="hidden" name="RemQtyUndo" value="<%=rs("RemQty")%>"><input class="input" type="text" name="RemQty" size="11" value='<%=rs("RemQty")%>' onkeydown="return chkMax(event, this, 20);" style="text-align: right" onchange="javascript:changeReminder();doProc('RemQty', 'N', this.value);">
						</td>
						<td>&nbsp;</td>
						<td><select class="input" size="1" name="RemType" style="font-size:10px; font-family:Verdana" onchange="doProc('RemType', 'S', this.value);">
				        <option <% If rs("RemType") = "M" Then %>selected<% End If %> value="M">
						<%=getaddActivityLngStr("LtxtMinutes")%></option>
						<option <% If rs("RemType") = "H" Then %>selected<% End If %> value="H">
						<%=getaddActivityLngStr("LtxtHours")%></option>
				        </select></td>
					</tr>
				</table>
				</td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="generalTbl"><input type="checkbox" name="Closed" id="Closed" value="Y" <% If rs("Closed") = "Y" Then %>checked<% End If %> style="background: background-image; border: 0px solid" onclick="doProc('Closed', 'S', GetYesNo(this.checked));"><label for="Closed"><%=getaddActivityLngStr("DtxtClosed")%></label></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2" id="trAddress1" <% If rs("Action") <> "M" Then %>style="display: none"<% End If %>>
		<td>
		<p align="center"><%=getaddActivityLngStr("DtxtAddress")%></td>
	</tr>
	<tr id="trAddress2" <% If rs("Action") <> "M" Then %>style="display: none"<% End If %>>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getaddActivityLngStr("DtxtCountry")%></td>
				<td width="25%" class="generalTbl">
				<% 
				set rd = Server.CreateObject("ADODB.RecordSet")
				cmd.CommandText = "DBOLKGetCountries" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = cmd.execute() %>
				<select class="input" size="1" name="Country" style="font-size:10px; font-family:Verdana" onchange="javascript:changeCountry(this.value);">
				<option></option>
		        <% do while not rd.eof %>
		        <option <% If rd("Code") = rs("Country") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
				<td width="6%"><%=getaddActivityLngStr("DtxtState")%></td>
				<td width="57%" class="generalTbl">
				<select class="input" size="1" name="State" id="State" style="font-size:10px; font-family:Verdana" onchange="doProc('State', 'S', this.value);">
				<option></option>
				<% 
				If rs("Country") <> "" Then
				set rd = Server.CreateObject("ADODB.RecordSet")

				cmd.CommandText = "DBOLKGetCountryStates" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@Code") = rs("Country")
				rd.open cmd, , 3, 1 %>
		        <% do while not rd.eof %>
		        <option <% If rd("Code") = rs("State") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        End If
		        %>
		        </select></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getaddActivityLngStr("DtxtCity")%></td>
				<td width="25%" class="generalTbl">
				<input class="input" type="text" name="city" size="35" value='<%=myHTMLEncode(rs("city"))%>' onkeydown="return chkMax(event, this, 100);" style="width: 222px" onchange="doProc('city', 'S', this.value);"></td>
				<td width="6%"><%=getaddActivityLngStr("DtxtStreet")%></td>
				<td width="57%" class="generalTbl">
				<input class="input" type="text" name="street" size="35" value='<%=myHTMLEncode(rs("street"))%>' onkeydown="return chkMax(event, this, 100);" style="width: 222px" onchange="doProc('street', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getaddActivityLngStr("DtxtRoom")%></td>
				<td width="25%" class="generalTbl">
				<input class="input" type="text" name="room" size="35" value='<%=myHTMLEncode(rs("room"))%>' onkeydown="return chkMax(event, this, 50);" style="width: 222px" onchange="doProc('room', 'S', this.value);"></td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="generalTbl">
				&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getaddActivityLngStr("LtxtContent")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTblBold2">
				<td align="center"><textarea rows="10" name="Notes" cols="20" class="input" style="width: 96%" onchange="doProc('Notes', 'S', this.value);"><%=myHTMLEncode(rs("Notes"))%></textarea></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getaddActivityLngStr("LtxtLinkDoc")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="generalTbl">
				<td class="GeneralTblBold2"><%=getaddActivityLngStr("LtxtDocType")%></td>
				<td>
				<select size="1" name="DocType" onchange="javascript:changeDocType();">
				<option value="-1"></option>
				<optgroup label="<%=getaddActivityLngStr("LtxtSale")%>">
					<option <% If rs("DocType") = 23 Then %>selected<% End If %> value="23"><%=txtQuotes%></option>
					<option <% If rs("DocType") = 17 Then %>selected<% End If %> value="17"><%=txtOrdrs%></option>
					<option <% If rs("DocType") = 15 Then %>selected<% End If %> value="15"><%=txtOdlns%></option>
					<option <% If rs("DocType") = 16 Then %>selected<% End If %> value="16"><%=txtOrnds%></option>
					<option <% If rs("DocType") = 13 Then %>selected<% End If %> value="13"><%=txtInvs%></option>
					<option <% If rs("DocType") = 14 Then %>selected<% End If %> value="14"><%=txtOrins%></option>
					<option <% If rs("DocType") = 203 Then %>selected<% End If %> value="203"><%=getaddActivityLngStr("LtxtInvDownPay")%></option>
				</optgroup>
				<optgroup label="<%=getaddActivityLngStr("LtxtBanks")%>">
					<option <% If rs("DocType") = 24 Then %>selected<% End If %> value="24"><%=txtRcts%></option>
					<option <% If rs("DocType") = 46 Then %>selected<% End If %> value="46"><%=txtOvpms%></option>
				</optgroup>
				<optgroup label="<%=getaddActivityLngStr("LtxtInventory")%>">
					<option <% If rs("DocType") = 67 Then %>selected<% End If %> value="67"><%=getaddActivityLngStr("LtxtInvTrans")%></option>
				</optgroup>
				</select></td>
				<td class="GeneralTblBold2"><%=getaddActivityLngStr("LtxtDocNum")%></td>
				<td>
				<input type="hidden" name="DocEntry" value="<%=rs("DocEntry")%>">
				<input type="text" name="DocNum" <% If rs("DocType") = -1 Then %>disabled<% End If %> size="20" value='<%=rs("DocNum")%>' onchange="javascript:changeDocNum();" style="text-align: right"></td>
			</tr>
			<% If Not IsNull(rs("parentType")) Then %>
			<tr class="generalTbl">
				<td class="GeneralTblBold2"><%=getaddActivityLngStr("DtxtSource")%></td>
				<% Select Case rs("parentType")
					Case 97
						source = getaddActivityLngStr("DtxtSalesOportunity")
					Case 191
						source = getaddActivityLngStr("DtxtServiceCall")
				End Select
				source = source & " #" & rs("parentID") %>
				<td><input type="text" id="txtSource" readonly value="<%=source%>" style="width: 100%;">
				</td>
				<td class="GeneralTblBold2">&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<% If EnableSDK Then
	
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCLG"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rg = cmd.execute()

	set rSdk = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCLG"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rSdk.open cmd, , 3, 1

	set rd = Server.CreateObject("ADODB.RecordSet")
	
	do while not rg.eof
	If CInt(rg("GroupID")) < 0 Then GroupID = "_1" Else GroupID = rg("GroupID")
	 %>
	<tr class="GeneralTblBold2">
		<td colspan="2">
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTblBold2" style="cursor: hand; " onclick="showHideSection(tdShowUDF<%=GroupID%>, trUDF<%=GroupID%>);">
				<td align="center"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getaddActivityLngStr("DtxtUDF")%><%
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
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			arrPos = Split("I,D", ",")
			For i = 0 to 1
			rSdk.Filter = "GroupID = " & rg("GroupID") & " and Pos = '" & arrPos(i) & "'"
			If not rSdk.eof then %>
				<td width="50%" valign="top">
			        <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
			        <% do while not rSdk.eof
			        ShowAddActivityUFD()
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
      loop %>
      <% End If %>
	<tr class="GeneralTbl" align="center">
		<td>
		<p align="center">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td>
					  <input type="button" value="<% If IsNull(ClgCode) Then %><%=getaddActivityLngStr("DtxtAdd")%><% Else %><%=getaddActivityLngStr("DtxtUpdate")%><% End If %>" name="btnAdd" onclick="if(valFrm()) { setActFlow(); doFlowAlert(); }"></td>
					<td>
						<% If IsNull(ClgCode) Then %>
					  <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					  <input type="button" value="<%=getaddActivityLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getaddActivityLngStr("LtxtConfCancel")%>'))window.location.href='activityCancel.asp?isUpdate=<%=JBool(Not IsNull(ClgCode))%>'"><% End If %></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="newActivitySubmit">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="doSubmit" value="Y">
<input type="hidden" name="Confirm" value="">
</form>
<script language="javascript">
var txtStartDate = '<%=getaddActivityLngStr("LtxtStartDate")%>';
var txtDueDate = '<%=getaddActivityLngStr("LtxtDueDate")%>';
var txtTime = '<%=getaddActivityLngStr("LtxtTime")%>';
var txtStartTime = '<%=getaddActivityLngStr("LtxtStartTime")%>';
var txtEndTime = '<%=getaddActivityLngStr("LtxtEndTime")%>';
var txtValNumVal = '<%=getaddActivityLngStr("DtxtValNumVal")%>';

function valFrm()
{
	<% If EnableSDK Then 
	cmd.CommandText = "DBOLKGetUDFNotNull" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = "V"
	cmd("@TableID") = "OCLG"
	cmd("@OP") = "O"
	set rd = cmd.execute()
	do while not rd.eof %>
	if (document.frmAddActivity.U_<%=rd("AliasID")%>.value == "")
	{
		alert('<%=getaddActivityLngStr("LtxtConfFld")%>'.replace('{0}', '<%=Replace(rd("Descr"), "'", "\'")%>'));
		showUDF(<%=rd("GroupID")%>);
		document.frmAddActivity.U_<%=rd("AliasID")%>.focus();
		return false;
	}
	<% rd.movenext
	loop 
	End If %>
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
	loop
End If %>

function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('<%=getaddActivityLngStr("DtxtValFldMaxChar")%>'.replace('{0}', FSize));
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
							alert('<%=getaddActivityLngStr("DtxtValNumVal")%>');
						}
						else if (parseInt(getNumericVB(Field.value)) < 1)
						{
							Field.value = '';
							alert('|D:txtValNumMinVal|'.replace('{0}', '1'));
						}
						else if (parseInt(getNumericVB(Field.value)) > 2147483647)
						{
							alert('<%=getaddActivityLngStr("DtxtValNumMaxVal")%>'.replace('{0}', '2147483647'));
							Field.value = 2147483647;
						}
						else if (Field.value.indexOf('<%=GetFormatDec%>') > -1)
						{
							Field.value = '';
							alert('<%=getaddActivityLngStr("DtxtValNumValWhole")%>');
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
					alert('<%=getaddActivityLngStr("DtxtValNumVal")%>');
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
<script language="javascript" src="addActivity/addActivity.js"></script>
<% Sub ShowAddActivityUFD()
	InsertID = rSdk("InsertID")
	FldVal = rs(InsertID)
	Select Case rSdk("TypeID")
		Case "B", "N"
			ProcType = "N"
		Case "M", "A"
			ProcType = "S"
		Case "D"
			ProcType = "D"
	End Select %>
				<tr class="generalTbl">
			            <td bgcolor="#EAF5FF" width="100" class="GeneralTblBold2">
			              <table border="0" cellpadding="0" cellspacing="0" width="100%">
			                <tr>
			            	  <td class="GeneralTblBold2">
			            	    <b><font size="1" face="Verdana"><%=rSdk("Descr")%><% If rSdk("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b>
			            	  </td>
			            	    <% If (rSdk("Query") = "Y" or rSdk("TypeID") = "D") and IsNull(rSdk("RTable")) Then %>
			            	    <td width="16" class="generalTbl">
			            	    	<img border="0" src="images/<% If rSdk("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rSdk("AliasID")%>" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Act&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',400,250,'yes', 'yes', document.frmAddActivity.U_<%=rSdk("AliasID")%>, '<%=ProcType%>')"<% End If %>>
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
							cmd("@TableID") = "OCLG"
							cmd("@FieldID") = rSdk("FieldID")
							rd.open cmd, , 3, 1
							 %>&nbsp;<select size="1" name="U_<%=rSdk("AliasID")%>" class="input" style="width: 99%" onchange="doProc(this.name, '<%=ProcType%>', this.value);">
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
						<textarea <% If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" class="input" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>);doProc(this.name, '<%=ProcType%>', this.value);" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Act&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> rows="3" onfocus="this.select()" style="width: 100%" cols="1"><% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %></textarea>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							</td>
							<td width="16">
								<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddActivity.U_<%=rSdk("AliasID")%>.value = '';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
					<% ElseIf rSdk("TypeID") = "A" and rSdk("EditType") = "I" Then %>
						<table cellpadding="0" cellspacing="2" border="0">
							<tr>
								<td><img src="pic.aspx?filename=<% If IsNull(rs(InsertID)) Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rSdk("AliasID")%>" border="1">
								<input type="hidden" name="U_<%=rSdk("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
								<td width="16" valign="bottom"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddActivity.U_<%=rSdk("AliasID")%>.value = '';document.frmAddActivity.imgU_<%=rSdk("AliasID")%>.src='pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand"></td>
							</tr>
							<tr>
								<td colspan="2" height="22">
								<p align="center">
								<input type="button" value="<%=getaddActivityLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.frmAddActivity.U_<%=rSdk("AliasID")%>, document.frmAddActivity.imgU_<%=rSdk("AliasID")%>,180);"></td>
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
							<input <% If readOnly Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" id="U_<%=rSdk("AliasID")%>" size="<%=fldSize%>" class="input" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>);doProc(this.name, '<%=ProcType%>', this.value);" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click()"<% End If %> <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Act&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this, '<%=ProcType%>')"<% End If %> value="<% If Not IsNull(FldVal) Then %><%=FldVal%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %>" <% If rSdk("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> style="width: 100%" <% If isMaxSize Then %> onkeydown="return chkMax(event, this, <%=MaxSize%>);"<% End if %>>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
								</td>
								<td width="16">
									<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddActivity.U_<%=rSdk("AliasID")%>.value = '';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');">
								</td>
							  </tr>
							</table>
							<% End If %><% End If %></td></tr><% End Sub %>