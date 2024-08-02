<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOCRD Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/agentClientSubmit.asp" -->
<% 
If Not myAut.GetCardProperty(Request("CardType"), "C") and Request("Confirm") <> "Y" or Request("isUpdate") = "True" Then goStatus = "C" Else goStatus = "H"
If Session("CrdRetVal") <> "" Then CardRetVal = Session("CrdRetVal") Else CardRetVal = Session("ConfCrdRetVal")
If Request("doSubmit") = "Y" Then
	set rs = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetLogStatus"
	cmd.Parameters.Refresh()
	cmd("@LogNum") = CardRetVal
	set rs = cmd.execute()
	If rs("Status") = "R" Then
		cmd.CommandText = "DBOLKCheckObjectCheckSum" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("CrdRetVal")
		cmd.Execute()
		
		If cmd("@IsValid").Value = "N" Then
			Session("RetryRetVal") = Session("CrdRetVal")
			Session("CrdRetVal") = ""
			Response.Redirect "crcerror.asp"
		End If
	
		Session("NotifyAdd") = True
		
		cmd.CommandText = "DBOLKExecuteLog"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("CrdRetVal")
		cmd("@SlpCode") = Session("vendid")
		cmd("@branchIndex") = Session("branch")
		If goStatus = "H" Then cmd("@Confirm") = "Y"
		cmd.execute()

		If goStatus = "H" Then 
			cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@UserType") = userType
			cmd("@ExecAt") = "C1"
			cmd("@ObjectEntry") = Session("CrdRetVal")
			cmd("@AgentID") = Session("vendid")
			cmd("@LanID") = Session("LanID")
			cmd("@branch") = Session("branch")
			cmd("@SetLogNumConf") = "Y"
			cmd.Execute()
		End If
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@sessiontype") = "A"
		cmd("@transtype") = "A"
		cmd("@object") = 2 
		cmd("@LogNum") = Session("CrdRetVal")
		cmd("@CurrentSlpCode") = Session("vendid")
		cmd("@Branch") = Session("branch")
		cmd.execute()
	End If
	doCardNextOp()
Else
	ShowNewCard()
End If

Sub doCardNextOp()
	If goStatus <> "H" Then
		doSubmitCardWait()
	Else
		CardRetVal = Session("CrdRetVal")
		ShowNewCard()
	End If
End Sub

Sub doSubmitCardWait()
	set mySubmit = new SubmitControl
	mySubmit.EnableRunInBackground = True
	mySubmit.LogNum = CardRetVal 
	mySubmit.LogNumID = "CrdRetVal"
	If Request("isUpdate") = "True" Then
		mySubmit.TransactionOkMessage = getagentClientSubmitLngStr("LtxtConfUpdCrd")
		mySubmit.EndButtonDescription = getagentClientSubmitLngStr("LtxtBPOp")
		mySubmit.EndButtonFunction = "goOp('" & Replace(Session("UserName"), "'", "\'") & "');"
	Else
		mySubmit.TransactionOkMessage = getagentClientSubmitLngStr("LtxtConfAddCrd")
		mySubmit.EndButtonDescription = getagentClientSubmitLngStr("LtxtCreateNewBP")
		mySubmit.EndButtonFunction = "window.location.href='addCard/goNewCard.asp?AddPath=';"
		mySubmit.SecondButtonDescription = getagentClientSubmitLngStr("LtxtView") & " " & getagentClientSubmitLngStr("DtxtBP")
		mySubmit.SecondButtonFunction = "viewDetails('{0}');"
		mySubmit.ThirdButtonDescription = getagentClientSubmitLngStr("LtxtBPOp")
		mySubmit.ThirdButtonFunction = "goOp('{0}');"
		mySubmit.RunInBackgroundRedir = "adminClientSubmit.asp?RetVal=" & Session("CrdRetVal") & "&Confirm=Y&bg=Y"
	End If
	mySubmit.GenerateSubmit %>
<script type="text/javascript">
function Start(page) 
{
	OpenWin = this.open(page, 'objDetails', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, height=600,width=800');
	OpenWin.focus()
}

function viewDetails(cardCode)
{
	Start('');
	doMyLink('addCard/crdConfDetailOpen.asp', 'CardCode=' + cardCode + '&pop=Y&AddPath=', 'objDetails');
}
</script>
<% End Sub

Sub ShowNewCard()

Confirm = goStatus = "H" or Request("bg") = "Y"
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCrdDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
If Not Confirm Then
	cmd("@CrdType") = 2
	cmd("@LogNum") = CardRetVal
	cmd("@SlpCode") = Session("vendid")
Else
	cmd("@CrdType") = -2
	cmd("@LogNum") = CardRetVal
End If
set rs = cmd.execute()
EnableSDK = rs("EnableSDK") %>
<p align="center">
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><% If Request("isUpdate") <> "True" Then %><%=getagentClientSubmitLngStr("LttlAddClient")%><% Else %><%=getagentClientSubmitLngStr("LttlUpdClient")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentClientSubmitLngStr("DtxtCode")%></td>
				<td width="46%"><%=rs("CardCode")%> - <%=rs("CardType")%></td>
				<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then %>
				<td class="GeneralTblBold2" width="5%"><%=getagentClientSubmitLngStr("DtxtType")%>:</td>
				<td class="GeneralTbl" width="37%"><%=Server.HTMLEncode(rs("ClientType"))%>&nbsp;</td>
				<% End If %>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentClientSubmitLngStr("DtxtName")%></td>
				<td colspan="3"><%=rs("CardName")%>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentClientSubmitLngStr("DtxtGroup")%></td>
				<td colspan="3"><%=rs("GroupName")%>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><% 
				Select Case myApp.LawsSet 
					Case "PA", "IL", "US", "CA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "ZA" %><%=getagentClientSubmitLngStr("DtxtLicTradNum")%><% 
					Case "MX", "CR", "GT" %>RFC<% 
					Case "GB" %><%=getagentClientSubmitLngStr("DtxtVatNum")%><%
					Case "CL" %>RUT<% 
				End Select %></td>
				<td colspan="3"><%=rs("LicTradNum")%>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getagentClientSubmitLngStr("DtxtAgent")%></td>
				<td colspan="3"><%=rs("SlpName")%>&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getagentClientSubmitLngStr("DtxtAddData")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="10%" class="GeneralTblBold2"><%=getagentClientSubmitLngStr("LtxtPhone1")%></td>
				<td width="25%" class="GeneralTbl"><%=rs("Phone1")%>&nbsp;</td>
				<td width="6%" class="GeneralTblBold2"><nobr><%=getagentClientSubmitLngStr("LtxtMobile")%></nobr></td>
				<td width="57%" class="GeneralTbl"><%=rs("Cellular")%>&nbsp;</td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2"><%=getagentClientSubmitLngStr("LtxtPhone2")%></td>
				<td width="25%" class="GeneralTbl"><%=rs("Phone2")%>&nbsp;</td>
				<td width="6%" class="GeneralTblBold2"><%=getagentClientSubmitLngStr("LtxtFax")%></td>
				<td width="57%" class="GeneralTbl"><%=rs("Fax")%>&nbsp;</td>
			</tr>
			<tr>
				<td width="10%" class="GeneralTblBold2"><%=getagentClientSubmitLngStr("DtxtEMail")%></td>
				<td colspan="3" class="GeneralTbl"><%=rs("E_Mail")%>&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td valign="top" width="25%">
				<table border="0" cellpadding="0" width="100%">
					<tr class="generalTblBold2">
						<td><%=getagentClientSubmitLngStr("DtxtImage")%></td>
					</tr>
					<tr>
						<td height="180" class="GeneralTbl">
						<p align="center">
                <font face="Verdana" size="1">
                <img id="ItemImg" src='pic.aspx?filename=<% If Not IsNull(rs("Picture")) and Trim(rs("Picture")) <> "" Then Response.Write rs("Picture") Else Response.Write "pcard.gif" %>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' border="1" name="ItemImg"></font></td>
					</tr>
				</table>
				</td>
				<td valign="top" width="75%">
				<table border="0" cellpadding="0" width="100%">
					<tr class="generalTblBold2">
						<td><%=getagentClientSubmitLngStr("DtxtObservations")%></td>
					</tr>
					<tr>
						<td class="GeneralTbl" height="180" valign="top"><%=rs("Notes")%>&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<% If EnableSDK = "Y" Then
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCRD"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rg = cmd.execute()

	set rc = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFReadCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OCRD"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rc.open cmd, , 3, 1

	set rd = Server.CreateObject("ADODB.RecordSet")
	do while not rg.eof
	 %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><% Select Case CInt(rg("GroupID"))
		Case -1 %><%=getagentClientSubmitLngStr("DtxtUDF")%><%
		Case Else
			Response.Write rg("GroupName")
		End Select %></td>
	</tr>
      <tr>
        <td width="100%">
        <table border="0" cellpadding="0" cellspacing="2" width="100%">
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
			            If rs("U_" & rc("AliasID")) <> "" Then Picture = rs("U_" & rc("AliasID")) Else Picture = "n_a.jpg" %>
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