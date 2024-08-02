<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not myApp.EnableOCLG Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/agentActivitySubmit.asp" -->
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

If Session("ActRetVal") <> "" Then ActRetVal = Session("ActRetVal") Else ActRetVal = Session("ConfActRetVal")
If Request("doSubmit") = "Y" Then 
	set rs = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetLogStatus"
	cmd("@LogNum") = ActRetVal
	set rs = cmd.execute()
	If rs("Status") = "R" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandText = "DBOLKCheckObjectCheckSum" & Session("ID")
		cmd.CommandType = &H0004
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("ActRetVal")
		cmd.Execute()
	
		If cmd("@IsValid").Value = "N" Then
			Session("RetryRetVal") = Session("ActRetVal")
			Session("ActRetVal") = ""
			Response.Redirect "crcerror.asp"
		End If

		Session("NotifyAdd") = True
	
		cmd.CommandText = "DBOLKExecuteLog"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("ActRetVal")
		cmd("@SlpCode") = Session("vendid")
		cmd("@branchIndex") = Session("branch")
		cmd.execute()
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@sessiontype") = "A"
		cmd("@transtype") = "A"
		cmd("@object") = 33 
		cmd("@LogNum") = Session("ActRetVal")
		cmd("@CurrentSlpCode") = Session("vendid")
		cmd("@Branch") = Session("branch")
		cmd.execute()
		If goStatus <> "H" Then
			doSubmitActivityWait()
		Else
			ActRetVal = Session("ActRetVal")
			ShowNewActivity()
		End If
	Else
		doSubmitActivityWait()
	End If
Else
	ShowNewActivity()
End If

Sub doSubmitActivityWait()
	set mySubmit = new SubmitControl
	mySubmit.EnableRunInBackground = True
	mySubmit.LogNum = ActRetVal 
	mySubmit.LogNumID = "ActRetVal"
	If Request("isUpdate") = "True" Then
		mySubmit.TransactionOkMessage = getagentActivitySubmitLngStr("LtxtConfUpdAct")
	Else
		mySubmit.TransactionOkMessage = getagentActivitySubmitLngStr("LtxtConfAddAct")
		mySubmit.EndButtonDescription = getagentActivitySubmitLngStr("LtxtCreateNewAct")
		mySubmit.EndButtonFunction = "window.location.href='addActivity/goNewActivity.asp?AddPath=../';"
		mySubmit.SecondButtonDescription = getagentActivitySubmitLngStr("LtxtView") & " " & getagentActivitySubmitLngStr("DtxtActivity")
		mySubmit.SecondButtonFunction = "viewDetails('{0}');"
		mySubmit.RunInBackgroundRedir = "agentActivitySubmit.asp?RetVal=" & Session("ItmRetVal") & "&Confirm=Y&bg=Y"
	End If
	mySubmit.GenerateSubmit %>
<script type="text/javascript">
function Start(page) 
{
	OpenWin = this.open(page, 'objDetails', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, height=600,width=800');
	OpenWin.focus()
}

function viewDetails(actCode)
{
	Start('');
	doMyLink('addActivity/activityConfDetail.asp', 'DocType=33&DocEntry=' + actCode + '&pop=Y&AddPath=../', 'objDetails');
}
</script>
<% End Sub

Sub ShowNewActivity() 
cmd.CommandText = "DBOLKGetActDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@DocType") = -2
cmd("@DocEntry") = ActRetVal
set rs = cmd.execute()
EnableSDK = rs("EnableSDK")
%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><% If Request("bg") <> "Y" Then 
					If IsNull(rs("ClgCode")) Then
						Response.Write Replace(getagentActivitySubmitLngStr("LttlAddActivity"), "{0}", rs("ObjectCode"))
					Else
						Response.Write Replace(getagentActivitySubmitLngStr("LttlUpdActivity"), "{0}", rs("ClgCode"))
					End If
				Else 
					Response.Write getagentActivitySubmitLngStr("DtxtLogNum") & ActRetVal
				End If 
			%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentActivitySubmitLngStr("DtxtActivity")%></td>
				<td>
				<% Select Case rs("Action")
					Case "C" %><%=getagentActivitySubmitLngStr("DtxtConv")%>
				<%	Case "M" %><%=getagentActivitySubmitLngStr("DtxtMeeting")%>
				<% 	Case "E" %><%=getagentActivitySubmitLngStr("DtxtNote")%>
				<%	Case "O" %><%=getagentActivitySubmitLngStr("DtxtOther")%>
				<%	Case "T" %><%=getagentActivitySubmitLngStr("DtxtTask")%>
				<% End Select %>
				</td>
				<td class="GeneralTblBold2">
				<%=getagentActivitySubmitLngStr("DtxtClientCode")%></td>
				<td>
				<%=myHTMLEncode(rs("CardCode"))%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentActivitySubmitLngStr("DtxtType")%></td>
				<td>
				<%=rs("CntctType")%></td>
				<td class="GeneralTblBold2">
				<%=getagentActivitySubmitLngStr("DtxtName")%></td>
				<td>
				<%=myHTMLEncode(rs("CardName"))%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentActivitySubmitLngStr("DtxtSubject")%></td>
				<td>
				<%=rs("CntctSbjct")%></td>
				<td class="GeneralTblBold2">
				<%=getagentActivitySubmitLngStr("DtxtContact")%></td>
				<td>
				<%=rs("CntctName")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentActivitySubmitLngStr("LtxtAsignedTo")%></td>
				<td>
				<%=rs("AttendUser")%></td>
				<td class="GeneralTblBold2">
				<%=getagentActivitySubmitLngStr("DtxtPhone")%></td>
				<td>
				<%=myHTMLEncode(rs("Tel"))%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%">&nbsp;</td>
				<td><input type="checkbox" style="background: background-image; border: 0px solid" value="Y"<% If rs("personal") = "Y" Then %> checked<% End If %> disabled><%=getagentActivitySubmitLngStr("DtxtPersonal")%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getagentActivitySubmitLngStr("LtxtGeneral")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getagentActivitySubmitLngStr("DtxtCommentaries")%></td>
				<td colspan="3" style="width: 82%" class="GeneralTbl">
				<%=myHTMLEncode(rs("Details"))%></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><% If rs("Action") <> "T" and rs("Action") <> "E" Then %><%=getagentActivitySubmitLngStr("LtxtStartTime")%><% ElseIf rs("Action") = "T" Then %><%=getagentActivitySubmitLngStr("LtxtStartDate")%><% ElseIf rs("Action") = "E" Then %><%=getagentActivitySubmitLngStr("LtxtTime")%><% End If %></td>
				<td width="25%" class="GeneralTbl" style="width: 50%">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr class="GeneralTbl">
						<td>
						<%=FormatDate(rs("Recontact"), True)%>
						</td>
						<td>&nbsp;</td>
						<td dir="ltr">
						<% If rs("Action") <> "T" Then %><%=rs("BeginTime")%><% End If %>
						</td>
					</tr>
				</table>
				</td>
				<td width="6%"><%=getagentActivitySubmitLngStr("DtxtPriority")%></td>
				<td width="57%" class="GeneralTbl">
				<% Select Case rs("Priority")
					Case "L" %><%=getagentActivitySubmitLngStr("DtxtLow")%>
				<%	Case "N" %><%=getagentActivitySubmitLngStr("DtxtNormal")%>
				<%	Case "H" %><%=getagentActivitySubmitLngStr("DtxtHigh")%>
				<% End Select %>
				</td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><% If rs("Action") <> "E" Then %><% If rs("Action") <> "T" Then %><%=getagentActivitySubmitLngStr("LtxtEndTime")%><% Else %><%=getagentActivitySubmitLngStr("LtxtDueDate")%><% End If %><% End If %></td>
				<td width="25%" class="GeneralTbl">
				<% If rs("Action") <> "E" Then %>
				<table cellpadding="0" cellspacing="0" border="0" id="tblENDTime">
					<tr class="GeneralTbl">
						<td><%=FormatDate(rs("endDate"), True)%>
						</td>
						<td>&nbsp;</td>
						<td dir="ltr">
						 <% If rs("Action") <> "T" Then %><%=rs("ENDTime")%><% End If %>
						</td>
					</tr>
				</table>
				<% End If %>
				</td>
				<td width="6%"><% If rs("Action") <> "E" Then %><%=getagentActivitySubmitLngStr("DtxtLocation")%><% End If %></td>
				<td width="57%" class="GeneralTbl">
				<% If rs("Action") <> "E" Then %><%=rs("Location")%><% End If %></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><% If rs("Action") <> "T" and rs("Action") <> "E" Then %><%=getagentActivitySubmitLngStr("DtxtDuration")%><% End If %></td>
				<td width="25%" class="GeneralTbl">
				<% If rs("Action") <> "T" and rs("Action") <> "E" Then %>
				<%=rs("Duration")%>&nbsp;<% Select Case rs("DurType")
					Case "M" %><%=getagentActivitySubmitLngStr("LtxtMinutes")%>
				<% 	Case "H" %><%=getagentActivitySubmitLngStr("LtxtHours")%>
				<%	Case "D" %><%=getagentActivitySubmitLngStr("LtxtDays")%>
				<% End Select %>
				<% End If %>
				</td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="GeneralTbl">
				<% If rs("Action") = "M" Then %>
				<input type="checkbox" name="tentative" <% If rs("tentative") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid" disabled><%=getagentActivitySubmitLngStr("DtxtPosible")%><% End If %></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><% If rs("Action") = "T" Then %><%=getagentActivitySubmitLngStr("LtxtStatus")%><% End If %></td>
				<td width="25%" class="GeneralTbl"><% If rs("Action") = "T" Then %><%=rs("Status")%><% End If %></td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="GeneralTbl">
				<input type="checkbox" name="Inactive" <% If rs("Inactive") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid" disabled><%=getagentActivitySubmitLngStr("DtxtInactive")%></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%">
				<% If rs("Action") <> "T" Then %>
				<input type="checkbox" name="Reminder" <% If rs("Reminder") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid" disabled><%=getagentActivitySubmitLngStr("DtxtReminder")%><% End If %></td>
				<td width="25%" class="GeneralTbl">
				<% If rs("Action") <> "T" Then %><%=rs("RemQty")%>&nbsp;<% Select Case rs("RemType")
					Case "M" %><%=getagentActivitySubmitLngStr("LtxtMinutes")%>
				<%	Case "H" %><%=getagentActivitySubmitLngStr("LtxtHours")%>
				<% End Select %><% End If %>
				</td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="GeneralTbl"><input type="checkbox" name="Closed" value="Y" <% If rs("Closed") = "Y" Then %>checked<% End If %> disabled style="background: background-image; border: 0px solid"><%=getagentActivitySubmitLngStr("DtxtClosed")%></td>
			</tr>
			</table>
		</td>
	</tr>
	<% If rs("Action") = "M" Then %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getagentActivitySubmitLngStr("DtxtAddress")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getagentActivitySubmitLngStr("DtxtCountry")%></td>
				<td width="25%" class="GeneralTbl"><%=rs("Country")%></td>
				<td width="6%"><%=getagentActivitySubmitLngStr("DtxtState")%></td>
				<td width="57%" class="GeneralTbl"><%=rs("State")%></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getagentActivitySubmitLngStr("DtxtCity")%></td>
				<td width="25%" class="GeneralTbl">
				<%=myHTMLEncode(rs("city"))%></td>
				<td width="6%"><%=getagentActivitySubmitLngStr("DtxtStreet")%></td>
				<td width="57%" class="GeneralTbl">
				<%=myHTMLEncode(rs("street"))%></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getagentActivitySubmitLngStr("DtxtRoom")%></td>
				<td width="25%" class="GeneralTbl">
				<%=myHTMLEncode(rs("room"))%></td>
				<td width="6%">&nbsp;</td>
				<td width="57%" class="GeneralTbl">
				&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<% End If %>
	<% If rs("Notes") <> "" Then %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getagentActivitySubmitLngStr("LtxtContent")%></td>
	</tr>
	<tr>
		<td class="GeneralTblBold2">
		<%=myHTMLEncode(rs("Notes"))%>
		</td>
	</tr>
	<% End If %>
	<% If rs("DocType") <> -1 Then %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getagentActivitySubmitLngStr("LtxtLinkDoc")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td class="GeneralTblBold2"><%=getagentActivitySubmitLngStr("LtxtDocType")%></td>
				<td class="GeneralTbl" class="GeneralTbl">
				<% Select Case rs("DocType")
					Case 23 %><%=txtQuotes%>
				<%	Case 17 %><%=txtOrdrs%>
				<%	Case 15 %><%=txtOdlns%>
				<%	Case 16 %><%=txtOrnds%>
				<%	Case 13 %><%=txtInvs%>
				<%	Case 14 %><%=txtOrins%>
				<%	Case 203 %><%=getagentActivitySubmitLngStr("LtxtInvDownPay")%>
				<%	Case 24 %><%=txtRcts%>
				<%	Case 46 %><%=txtOvpms%>
				<%	Case 67 %><%=getagentActivitySubmitLngStr("LtxtInvTrans")%>
				<% End Select %></td>
				<td class="GeneralTblBold2"><%=getagentActivitySubmitLngStr("LtxtDocNum")%></td>
				<td class="GeneralTbl" class="GeneralTbl">
				<%=rs("DocNum")%></td>
			</tr>
		</table>
		</td>
	</tr>
	<% End If %>
	<% If EnableSDK = "Y" Then
	
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCLG"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rg = cmd.execute()


	set rc = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFReadCols" & Session("ID")
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCLG"
	cmd("@UserType") = userType
	cmd("@OP") = "O"
	rc.open cmd, , 3, 1

	set rd = Server.CreateObject("ADODB.RecordSet")
	do while not rg.eof
	 %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><% Select Case CInt(rg("GroupID"))
		Case -1 %><%=getagentActivitySubmitLngStr("DtxtUDF")%><%
		Case Else
			Response.Write rg("GroupName")
		End Select %></td>
	</tr>
      <tr>
        <td width="100%">
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
			<% 
			arrPos = Split("I,D", ",")
			For i = 0 to 1
			rc.Filter = "GroupID = " & rg("GroupID") & " and Pos = '" & arrPos(i) & "'"
			If not rc.eof then %>
				<td width="50%" valign="top">
			        <table border="0" cellpadding="0" width="100%">
			        <% do while not rc.eof %>
			          <tr class="GeneralTbl">
			            <td width="100" valign="top" class="GeneralTblBold2">
			              <%=rc("Descr")%>
			            </td>
			            <td dir="ltr">
			            <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rs("U_" & rc("AliasID"))%>"><% End If %>
			            <% If rc("TypeID") = "B" Then
			            If Not IsNull(rs("U_" & rc("AliasID"))) Then
			            	Select Case rc("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),myApp.MeasureDec)
			            	End Select
			            End If
			            ElseIf rc("TypeID") = "A" and rc("EditType") = "I" Then
			            If rs("U_" & rc("AliasID")) <> "" Then Picture = rs("U_" & rc("AliasID")) Else Picture = "pcard.gif" %>
			            <img src='pic.aspx?filename=<%=Picture%>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' border="0">
			            <% Else %>
			            <%=rs("U_" & rc("AliasID"))%>
			            <% End If %>
			            <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %></a><% End If %>
			            </td>
			          </tr>
			        <% rc.movenext
			        loop
			        rc.movefirst %>
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
</table>
<% End Sub %>
<!--#include file="agentBottom.asp"-->