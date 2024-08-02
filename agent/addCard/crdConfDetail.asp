<!--#include file="lang/crdConfDetail.asp" -->
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = "&H0004"
cmd.CommandText = "DBOLKGetCrdDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid")
If Request("CardCode") <> "" or strScriptName = "activeclient.asp" Then
	cmd("@CrdType") = 2
	If Request("CardCode") <> "" Then
		cmd("@CardCode") = Request("CardCode")
	Else
		cmd("@CardCode") = Session("UserName")
	End If
	cmd("@SlpCode") = Session("vendid")
Else
	cmd("@CrdType") = -2
	cmd("@LogNum") = CardRetVal
End If
set rs = cmd.execute()

CardCode = rs("CardCode")

If rs.Eof Then Response.Redirect "../nodata.asp"
EnableSDK = rs("EnableSDK")

Select Case rs("CardType")
	Case "C"
		CardTypeDesc = txtClient
	Case "L"
		CardTypeDesc = getCrdConfDetailLngStr("DtxtLead")
	Case "S"
		CardTypeDesc = getcrdConfDetailLngStr("DtxtSupplier")
End Select

If strScriptName = "activeclient.asp" Then
	Select Case rs("CardType")
		Case "L"
			hasAut = myAut.HasAuthorization(75)
		Case "C"
			hasAut = myAut.HasAuthorization(23)
		Case "S"
			hasAut = myAut.HasAuthorization(74)
	End Select
Else
	hasAut = myAut.GetCardProperty(rs("CardType"), "V") or userType = "C"
End If


If hasAut Then %>
<table border="0" cellpadding="0" width="100%" id="table12">
	<tr class="GeneralTlt">
		<td><% If Request("CardCode") = "" Then %><%=getcrdConfDetailLngStr("LttlClientConfDetails")%><% Else %><%=getcrdConfDetailLngStr("LttlClientDetails")%><% End If %>&nbsp;	<% If Request("DocEntry") <> "" Then %>#<%=Request("DocEntry")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table13">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getcrdConfDetailLngStr("DtxtCode")%></td>
				<td width="46%"><%=rs("CardCode")%> - <%=CardTypeDesc%></td>
				<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then %>
				<td class="GeneralTblBold2" width="5%"><%=getcrdConfDetailLngStr("DtxtType")%>:</td>
				<td class="GeneralTblBold2" width="13%"><% Select Case rs("ClientType")
					Case "C"
						Response.Write getcrdConfDetailLngStr("DtxtCmp")
					Case "I"
						Response.Write getcrdConfDetailLngStr("LtxtNatPer")
				End Select %>&nbsp;</td>
				<% End If %>
				<% If Request("CardCode") <> "" or strScriptName = "activeclient.asp" Then %>
				<td class="GeneralTblBold2" width="12%"><%=getcrdConfDetailLngStr("DtxtBalance")%></td>
				<td width="12%" align="right">
				<% If (rs("CardType") = "S" and myAut.HasAuthorization(95) or (rs("CardType") = "C" or rs("CardType") = "L") and myAut.HasAuthorization(94)) and (myAut.HasAuthorization(174) or not myAut.HasAuthorization(174) and rs("IsSlpAssign") = "Y") and ClientShowBalance Then %>
				<nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(rs("Balance"),myApp.SumDec)%></nobr><% Else %>****<% End If %></td>
				<% End If %>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getcrdConfDetailLngStr("DtxtName")%></td>
				<td colspan="5"><% If Not IsNull(rs("CardName")) Then %><%=rs("CardName")%><% End If %>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getcrdConfDetailLngStr("DtxtGroup")%></td>
				<td colspan="5"><% If Not IsNull(rs("GroupName")) Then %><%=rs("GroupName")%><% End If %>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><% 
				Select Case myApp.LawsSet 
					Case "PA", "IL", "US", "CA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "ZA" %><%=getcrdConfDetailLngStr("DtxtLicTradNum")%><% 
					Case "MX", "CR", "GT" %>RFC<% 
					Case "GB" %>|D:txtVatNum|<%
					Case "CL" %>RUT<% 
				End Select %></td>
				<td colspan="5"><%=rs("LicTradNum")%>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="12%"><%=getcrdConfDetailLngStr("DtxtAgent")%></td>
				<td colspan="5"><%=rs("SlpName")%>&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getcrdConfDetailLngStr("DtxtAddData")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table16">
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getcrdConfDetailLngStr("LtxtPhone1")%></td>
				<td width="25%"><%=rs("Phone1")%>&nbsp;</td>
				<td width="10%" class="GeneralTblBold2"><%=getcrdConfDetailLngStr("LtxtMobile")%></td>
				<td width="52%"><%=rs("Cellular")%>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getcrdConfDetailLngStr("LtxtPhone2")%></td>
				<td width="25%"><%=rs("Phone2")%>&nbsp;</td>
				<td width="10%" class="GeneralTblBold2"><%=getcrdConfDetailLngStr("LtxtFax")%></td>
				<td width="52%"><%=rs("Fax")%>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td width="10%" class="GeneralTblBold2"><%=getcrdConfDetailLngStr("DtxtEMail")%></td>
				<td colspan="3"><%=rs("E_Mail")%>&nbsp;</td>
			</tr>
			<%
			If Request("CardCode") <> "" or strScriptName = "activeclient.asp" Then
			set rx = Server.CreateObject("ADODB.RecordSet")
			cmd.CommandText = "DBOLKGetCardRepRead" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@UserType") = userType
			cmd("@OP") = "O"
			rx.open cmd, , 3, 1
			If Not rx.eof Then
				rx.movefirst
				lastRow = rx("rowOrder") %>
				<tr class="GeneralTbl">
				<% do while not rx.eof
				If lastRow <> rx("rowOrder") Then
					Response.Write "</tr><tr class=""GeneralTbl"">"
					lastRow = rx("rowOrder")
				End If %>
				<td class="GeneralTblBold2"><%=rx("rowName")%></td>
				<td <% If rx("colCount") = 1 Then %>colspan="3"<% End If %>><%=rs("CardRep" & rx("rowIndex"))%>&nbsp;</td>
				<% rx.movenext
				loop %>
			</tr>
			<% End If %>
		<% End If %>			
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<% If strScriptName <> "activeclient.asp" Then %>
				<td valign="top" width="25%">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td class="GeneralTblBold2"><%=getcrdConfDetailLngStr("DtxtImage")%></td>
					</tr>
					<tr>
						<td class="GeneralTbl" height="100">
						<p align="center">
                <font face="Verdana" size="1">
                <img id="ItemImg" src="<% If Request("cmd") = "" Then %>../<% End If %>pic.aspx?filename=<% If Not IsNull(rs("Picture")) and Trim(rs("Picture")) <> "" Then Response.Write rs("Picture") Else Response.Write "pcard.gif" %>&MaxSize=223&dbName=<%=Session("olkdb")%>" border="1" name="ItemImg"></font></td>
					</tr>
				</table>
				</td>
				<% End If %>
				<td valign="top" <% If strScriptName <> "activeclient.asp" Then %>width="75%"<% End If %>>
				<table border="0" cellpadding="0" width="100%" id="table19">
					<tr class="generalTblBold2">
						<td><%=getcrdConfDetailLngStr("DtxtObservations")%></td>
					</tr>
					<tr>
						<td class="GeneralTbl" height="100" valign="top"><% If Not IsNull(rs("Notes")) Then %><%=rs("Notes")%><% End If %>&nbsp;</td>
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
		Case -1 %><%=getcrdConfDetailLngStr("DtxtUDF")%><%
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
									Response.Write FormatNumber(CDbl(rs("U_" & rc("AliasID"))),MeasureDec)
			            	End Select
			            End If
			            ElseIf rc("TypeID") = "A" and rc("EditType") = "I" Then
			            If rs("U_" & rc("AliasID")) <> "" Then Picture = rs("U_" & rc("AliasID")) Else Picture = "n_a.gif" %>
			            <img src="<% If Request("cmd") = "" Then %>../<% End If %>pic.aspx?filename=<%=Picture%>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
			            <% Else %>
			            <% If Not IsNull(rs("U_" & rc("AliasID"))) Then %><%=rs("U_" & rc("AliasID"))%><% End If %>
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
	<% Else %>
	<script type="text/javascript">
	alert('<%=getcrdConfDetailLngStr("LtxtNoAccessCard")%>'.replace('{0}', '<%=CardTypeDesc%>'));
	window.close();
	</script>
	<% End If %>
<%
If strScriptName <> "activeclient.asp" Then
	set rs = nothing
	set rd = nothing
	conn.close
End If %>