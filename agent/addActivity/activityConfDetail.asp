<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<!--#include file="lang/activityConfDetail.asp" -->
<head>
<title><%=getactivityConfDetailLngStr("LttlActConfDetails")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<link rel="stylesheet" type="text/css" href="../design/0/style/stylenuevo.css">
</head>

<body>
<%
Dim myAut
set myAut = New clsAuthorization


hasAut = myAut.GetObjectProperty(33, "V")
If hasAut Then

set rs = Server.CreateObject("ADODB.recordset")

If Request("DocType") <> "" Then DocType = CInt(Request("DocType")) Else DocType = 33
DocEntry = CLng(Request("DocEntry"))

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetActDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@DocType") = DocType
cmd("@DocEntry") = DocEntry
set rs = cmd.execute()
EnableSDK = rs("EnableSDK") = "Y"
%>
<!--#include file="../loadAlterNames.asp"-->
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><%=getactivityConfDetailLngStr("LttlActConfDetails")%>&nbsp;<% If DocType = -2 Then %>(<%=getactivityConfDetailLngStr("DtxtLogNum")%>)&nbsp;<% End If %>#<%=Request("DocEntry")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getactivityConfDetailLngStr("DtxtActivity")%></td>
				<td>
				<% Select Case rs("Action")
					Case "C" %><%=getactivityConfDetailLngStr("DtxtConv")%>
				<%	Case "M" %><%=getactivityConfDetailLngStr("DtxtMeeting")%>
				<% 	Case "E" %><%=getactivityConfDetailLngStr("DtxtNote")%>
				<%	Case "O" %><%=getactivityConfDetailLngStr("DtxtOther")%>
				<%	Case "T" %><%=getactivityConfDetailLngStr("DtxtTask")%>
				<% End Select %>
				</td>
				<td class="GeneralTblBold2">
				<%=getactivityConfDetailLngStr("DtxtClientCode")%></td>
				<td>
				<%=rs("CardCode")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getactivityConfDetailLngStr("DtxtType")%></td>
				<td>
				<%=rs("CntctType")%></td>
				<td class="GeneralTblBold2">
				<%=getactivityConfDetailLngStr("DtxtName")%></td>
				<td>
				<%=rs("CardName")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getactivityConfDetailLngStr("DtxtSubject")%></td>
				<td>
				<%=rs("CntctSbjct")%></td>
				<td class="GeneralTblBold2">
				<%=getactivityConfDetailLngStr("DtxtContact")%></td>
				<td>
				<%=rs("CntctName")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getactivityConfDetailLngStr("LtxtAsignedTo")%></td>
				<td>
				<%=rs("AttendUser")%></td>
				<td class="GeneralTblBold2">
				<%=getactivityConfDetailLngStr("DtxtPhone")%></td>
				<td>
				<%=rs("Tel")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%">&nbsp;</td>
				<td>
				<input type="checkbox" style="background: background-image; border: 0px solid; background-color: #EDF5FE;" value="Y"<% If rs("personal") = "Y" Then %> checked<% End If %> disabled><%=getactivityConfDetailLngStr("DtxtPersonal")%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getactivityConfDetailLngStr("LtxtGeneral")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="10%" class="GeneralTblBold2"><%=getactivityConfDetailLngStr("DtxtCommentaries")%></td>
				<td colspan="3" style="width: 82%" class="GeneralTbl">
				<%=rs("Details")%></td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2"><% 
				Select Case rs("Action")
					Case "E" %><%=getactivityConfDetailLngStr("LtxtTime")%><%
					Case "T" %><%=getactivityConfDetailLngStr("LtxtStartDate")%><%
					Case Else %><%=getactivityConfDetailLngStr("LtxtStartTime")%><%
				End Select %></td>
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
				<td width="6%" class="GeneralTblBold2"><%=getactivityConfDetailLngStr("DtxtPriority")%></td>
				<td width="57%" class="GeneralTbl">
				<% Select Case rs("Priority")
					Case "L" %><%=getactivityConfDetailLngStr("DtxtLow")%>
				<%	Case "N" %><%=getactivityConfDetailLngStr("DtxtNormal")%>
				<%	Case "H" %><%=getactivityConfDetailLngStr("DtxtHigh")%>
				<% End Select %>
				</td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2"><% If rs("Action") <> "E" Then %><% If rs("Action") <> "T" Then %><%=getactivityConfDetailLngStr("LtxtEndTime")%><% Else %><%=getactivityConfDetailLngStr("LtxtDueDate")%><% End If %><% End If %></td>
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
				<td width="6%" class="GeneralTblBold2"><% If rs("Action") <> "E" Then %><%=getactivityConfDetailLngStr("DtxtLocation")%><% End If %></td>
				<td width="57%" class="GeneralTbl">
				<% If rs("Action") <> "E" Then %><%=rs("Location")%><% End If %></td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2"><% If rs("Action") <> "T" and rs("Action") <> "E" Then %><%=getactivityConfDetailLngStr("DtxtDuration")%><% End If %></td>
				<td width="25%" class="GeneralTbl">
				<% If rs("Action") <> "T" and rs("Action") <> "E" Then %>
				<%=rs("Duration")%>&nbsp;<% Select Case rs("DurType")
					Case "M" %><%=getactivityConfDetailLngStr("LtxtMinutes")%>
				<% 	Case "H" %><%=getactivityConfDetailLngStr("LtxtHours")%>
				<%	Case "D" %><%=getactivityConfDetailLngStr("LtxtDays")%>
				<% End Select %>
				<% End If %>
				</td>
				<td width="6%" class="GeneralTbl">&nbsp;</td>
				<td width="57%" class="GeneralTbl">
				<% If rs("Action") = "M" Then %><input type="checkbox" name="tentative" <% If rs("tentative") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid; background-color: #EDF5FE;" disabled><%=getactivityConfDetailLngStr("DtxtPosible")%><% End If %></td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2"><% If rs("Action") = "T" Then %><%=getactivityConfDetailLngStr("LtxtStatus")%><% End If %></td>
				<td width="25%" class="GeneralTbl"><% If rs("Action") = "T" Then %><%=rs("Status")%><% End If %></td>
				<td width="6%" class="GeneralTbl">&nbsp;</td>
				<td width="57%" class="GeneralTbl">
				<input type="checkbox" name="Inactive" <% If rs("Inactive") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid; background-color: #EDF5FE;" disabled><%=getactivityConfDetailLngStr("DtxtInactive")%></td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2">
				<% If rs("Action") <> "T" Then %>
				<input type="checkbox" name="Reminder" <% If rs("Reminder") = "Y" Then %>checked<% End If %> value="Y" style="background: background-image; border: 0px solid; background-color: #E1EEFD;" disabled><%=getactivityConfDetailLngStr("DtxtReminder")%><% End If %></td>
				<td width="25%" class="GeneralTbl">
				<% If rs("Action") <> "T" Then %><%=rs("RemQty")%>&nbsp;<% Select Case rs("RemType")
					Case "M" %><%=getactivityConfDetailLngStr("LtxtMinutes")%>
				<%	Case "H" %><%=getactivityConfDetailLngStr("LtxtHours")%>
				<% End Select %><% End If %>
				</td>
				<td width="6%" class="GeneralTbl">&nbsp;</td>
				<td width="57%" class="GeneralTbl">
				<input type="checkbox" name="Closed" value="Y" <% If rs("Closed") = "Y" Then %>checked<% End If %> disabled style="background: background-image; border: 0px solid; background-color: #EDF5FE;"><%=getactivityConfDetailLngStr("DtxtClosed")%></td>
			</tr>
			</table>
		</td>
	</tr>
	<% If rs("Action") = "M" Then %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getactivityConfDetailLngStr("DtxtAddress")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getactivityConfDetailLngStr("DtxtCountry")%></td>
				<td width="25%" class="GeneralTbl"><%=rs("Country")%></td>
				<td width="6%"><%=getactivityConfDetailLngStr("DtxtState")%></td>
				<td width="57%" class="GeneralTbl"><%=rs("State")%></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getactivityConfDetailLngStr("DtxtCity")%></td>
				<td width="25%" class="GeneralTbl">
				<%=rs("city")%></td>
				<td width="6%"><%=getactivityConfDetailLngStr("DtxtStreet")%></td>
				<td width="57%" class="GeneralTbl">
				<%=rs("street")%></td>
			</tr>
			<tr class="GeneralTblBold2">
				<td width="10%"><%=getactivityConfDetailLngStr("DtxtRoom")%></td>
				<td width="25%" class="GeneralTbl">
				<%=rs("room")%></td>
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
		<p align="center"><%=getactivityConfDetailLngStr("LtxtContent")%></td>
	</tr>
	<tr>
		<td class="GeneralTbl">
		<%=rs("Notes")%>
		</td>
	</tr>
	<% End If %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getactivityConfDetailLngStr("LtxtLinkDoc")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td class="GeneralTblBold2"><%=getactivityConfDetailLngStr("LtxtDocType")%></td>
				<td class="GeneralTbl">
				<% Select Case rs("DocType")
					Case 23 %><%=txtQuotes%>
				<%	Case 17 %><%=txtOrdrs%>
				<%	Case 15 %><%=txtOdlns%>
				<%	Case 16 %><%=txtOrnds%>
				<%	Case 13 %><%=txtInvs%>
				<%	Case 14 %><%=txtOrins%>
				<%	Case 203 %><%=getactivityConfDetailLngStr("LtxtInvDownPay")%>
				<%	Case 24 %><%=txtRcts%>
				<%	Case 46 %><%=txtOvpms%>
				<%	Case 67 %><%=getactivityConfDetailLngStr("LtxtInvTrans")%>
				<% End Select %></td>
				<td class="GeneralTblBold2"><%=getactivityConfDetailLngStr("LtxtDocNum")%></td>
				<td class="GeneralTbl">
				<%=rs("DocNum")%></td>
			</tr>
			<% If DocType = 33 Then
			If Not IsNull(rs("parentType")) Then
			Select Case rs("parentType")
				Case 97
					source = getactivityConfDetailLngStr("DtxtSalesOportunity")
				Case 191
					source = getactivityConfDetailLngStr("DtxtServiceCall")
			End Select
			source = source & " #" & rs("parentID") %>
			<tr>
				<td class="GeneralTblBold2"><%=getactivityConfDetailLngStr("DtxtSource")%></td>
				<td class="GeneralTbl"><%=source%></td>
				<td class="GeneralTblBold2"></td>
				<td class="GeneralTbl"></td>
			</tr>
			<% End If %>
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
		Case -1 %><%=getactivityConfDetailLngStr("DtxtUDF")%><%
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
			          <tr>
			            <td width="100" valign="top" class="GeneralTblBold2">
			            <%=rc("Descr")%>
			            </td>
			            <td dir="ltr" class="GeneralTbl">
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
			            If rs("U_" & rc("AliasID")) <> "" Then Picture = rs("U_" & rc("AliasID")) Else Picture = "n_a.jpg" %>
			            <img src="../pic.aspx?filename=<%=Picture%>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
			            <% Else %>
			            <% If Not IsNull(rs("U_" & rc("AliasID"))) Then %><%=rs("U_" & rc("AliasID"))%><% End If %>
			            <% End If %>
			            <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %></a><% End If %>
			            </td>
			          </tr>
			        <% rc.movenext
			        loop %>
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


<% Else %>
	<script type="text/javascript">
	alert('<%=getactivityConfDetailLngStr("DtxtNoAccessObj")%>'.replace('{0}', '<%=getactivityConfDetailLngStr("DtxtActivity")%>'));
	window.close();
	</script>
<% End If %>

</body>

</html>
