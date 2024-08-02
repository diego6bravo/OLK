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

<!--#include file="lang/soConfDetail.asp" -->
<head>
<title><%=getsoConfDetailLngStr("LttlSOConfDetails")%></title>
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
set rd = Server.CreateObject("ADODB.recordset")

If Request("DocType") <> "" Then DocType = CInt(Request("DocType")) Else DocType = 33
DocEntry = CLng(Request("DocEntry"))

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetSODetails" & Session("ID")
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
		<td><%=getsoConfDetailLngStr("LttlSOConfDetails")%>&nbsp;<% If DocType = -2 Then %>(<%=getsoConfDetailLngStr("DtxtLogNum")%>)&nbsp;<% End If %>#<%=Request("DocEntry")%></td>
	</tr>
<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("DtxtClientCode")%></td>
				<td><%=myHTMLEncode(rs("CardCode"))%></td>
				<td class="GeneralTblBold2">
				<%=getsoConfDetailLngStr("LtxtOpporName")%></td>
				<td><%=myHTMLEncode(rs("Name"))%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("DtxtName")%></td>
				<td><%=myHTMLEncode(rs("CardName"))%></td>
				<td class="GeneralTblBold2">
				<%=getsoConfDetailLngStr("LtxtOpporId")%></td>
				<td><%=myHTMLEncode(rs("OpprId"))%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("DtxtContact")%></td>
				<td><%=rs("CprCode")%></td>
				<td class="GeneralTblBold2">
				<%=getsoConfDetailLngStr("LtxtStatus")%></td>
				<td><% Select Case rs("Status")
					Case "O" %><%=getsoConfDetailLngStr("LtxtOpen")%><%
					Case "W" %><%=getsoConfDetailLngStr("LtxtWon")%><%
					Case "L" %><%=getsoConfDetailLngStr("LtxtLost")%><%
				End Select %></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtInvTotal")%></td>
				<td><%=rs("Currency")%><%=FormatNumber(CDbl(rs("TotalInvoice")), myApp.SumDec)%></td>
				<td class="GeneralTblBold2">
				<%=getsoConfDetailLngStr("LtxtStartDate")%></td>
				<td><%=rs("OpenDate")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("DtxtTerritory")%></td>
				<td><%=rs("TerritoryDesc")%></td>
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtEndDate")%></td>
				<td><%=rs("CloseDate")%></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("DtxtAgent")%></td>
				<td><%=rs("SlpName")%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtOwner")%></td>
				<td><%=rs("Owner")%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getsoConfDetailLngStr("LtxtPotential")%></td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
			<table cellpadding="0" cellspacing="0" style="width: 100%">
				<tr>
					<td style="width: 55%" valign="top">
					<table style="width: 100%" cellpadding="0">
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtClosePlan")%></td>
							<td>
							<%=rs("PredDateQty")%>&nbsp;<% Select Case rs("DifType")
																Case "M" %><%=getsoConfDetailLngStr("DtxtMonths")%><% 
																Case "W" %><%=getsoConfDetailLngStr("DtxtWeeks")%><% 
																Case "D" %><%=getsoConfDetailLngStr("DtxtDays")%><%
															End Select %></td>
						</tr>
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtCloseDate")%></td>
							<td><%=rs("PredDate")%></td>
						</tr>
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtMaxSumLoc")%></td>
							<td><% If Not IsNull(rs("MaxSumLoc")) Then Response.Write FormatNumber(CDbl(rs("MaxSumLoc")), myApp.SumDec)%></td>
						</tr>
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtWtSumLoc")%></td>
							<td><% If Not IsNull(rs("WtSumLoc")) Then Response.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)%></td>
						</tr>
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtPrcntProf")%></td>
							<td><% If Not IsNull(rs("PrcntProf")) Then Response.Write FormatNumber(CDbl(rs("PrcntProf")), myApp.PercentDec)%></td>
						</tr>
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtSumProfL")%></td>
							<td><% If Not IsNull(rs("SumProfL")) Then Response.Write FormatNumber(CDbl(rs("SumProfL")), myApp.SumDec)%></td>
						</tr>
						<tr class="GeneralTbl">
							<td class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtIntRate")%></td>
							<td><%=rs("IntRate")%>
					        </td>
						</tr>
					</table>
					</td>
					<td style="width: 45%" valign="top">
					<table cellpadding="0" style="width: 100%">
						<tr class="GeneralTblBold2">
							<td colspan="3"><%=getsoConfDetailLngStr("LtxtIntRange")%></td>
						</tr>
						<tr class="GeneralTblBold2">
							<td style="width: 40px; text-align: center;">#</td>
							<td><%=getsoConfDetailLngStr("DtxtDescription")%></td>
							<td style="width: 100px; text-align:center;"></td>
						</tr>
						<%						
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetSODetailsIntRange" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						cmd("@DocType") = DocType
						cmd("@DocEntry") = DocEntry
						rd.open cmd, , 3, 1
						do while not rd.eof %>
						<tr class="GeneralTbl">
							<td style="width: 40px; text-align: right;"><%=rd("Line")%></td>
							<td><%=rd("Name")%></td>
							<td style="width: 100px; text-align:center;"><% If rd("Prmry") = "Y" Then %><%=getsoConfDetailLngStr("LtxtPrimary")%><% End If %></td>
						</tr>
						<% rd.movenext
						loop %>
					</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td><p align="center"><%=getsoConfDetailLngStr("LtxtGeneral")%></td>
	</tr>
    <tr>
        <td width="100%">
					<table style="width: 100%">
						<tr class="GeneralTbl">
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtChnCode")%></td>
							<td style="width: 30%">
							<%=myHTMLEncode(rs("ChnCrdCode"))%></td>
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoConfDetailLngStr("DtxtProject")%></td>
							<td style="width: 30%"><%=rs("Project")%></td>
						</tr>
						<tr class="GeneralTbl">
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtChnName")%></td>
							<td style="width: 30%">
							<%=myHTMLEncode(rs("ChnCrdName"))%></td>
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtInfSource")%></td>
							<td style="width: 30%"><%=rs("Source")%></td>
						</tr>
						<tr class="GeneralTbl">
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtChnCnt")%></td>
							<td style="width: 30%">
							<%=rs("ChnCrdCon")%></td>
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtIndustry")%></td>
							<td style="width: 30%"><%=rs("Industry")%></td>
						</tr>
						<tr class="GeneralTbl">
							<td style="width: 20%; vertical-align: top; padding-top: 2px;" class="GeneralTblBold2"><%=getsoConfDetailLngStr("LtxtRemarks")%></td>
							<td colspan="3">
							<%=rs("Memo")%></td>
						</tr>
					</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td><p align="center"><%=getsoConfDetailLngStr("LtxtStages")%></td>
	</tr>
    <tr>
        <td width="100%">
				<table cellpadding="0" style="width: 100%" id="tblStages">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoConfDetailLngStr("LtxtStage")%></td>
						<td><%=getsoConfDetailLngStr("LtxtStartDate")%></td>
						<td><%=getsoConfDetailLngStr("LtxtEndDate")%></td>
						<td><%=getsoConfDetailLngStr("DtxtAgent")%></td>
						<td>%</td>
						<td><%=getsoConfDetailLngStr("LtxtMaxSumLoc")%></td>
						<td><%=getsoConfDetailLngStr("LtxtWtSumLoc")%></td>
						<td><%=getsoConfDetailLngStr("LtxtDocType")%></td>
						<td class="style1"><%=getsoConfDetailLngStr("LtxtDocNum")%></td>
						<td class="style1"><%=getsoConfDetailLngStr("LtxtOwner")%></td>
					</tr>
					<%
										
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetSODetailsStages" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@DocType") = DocType
					cmd("@DocEntry") = DocEntry
					rd.close
					rd.open cmd, , 3, 1
					do while not rd.eof %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><%=rd("LineDesc")%></td>
						<td><%=rd("StepId")%></td>
						<td><%=rd("OpenDate")%></td>
						<td><%=rd("CloseDate")%></td>
						<td><%=rd("SlpName")%></td>
						<td><%=FormatNumber(CDbl(rd("ClosePrcnt")), myApp.PercentDec)%></td>
						<td><% If Not IsNull(rd("MaxSumLoc")) Then Response.Write FormatNumber(CDbl(rd("MaxSumLoc")), myApp.SumDec)%></td>
						<td><% If Not IsNull(rd("WtSumLoc")) Then Response.Write FormatNumber(CDbl(rd("WtSumLoc")), myApp.SumDec)%></td>
						<td><% Select Case rd("ObjType") 
							Case 23
								Response.Write txtQuote
							Case 17
								Response.Write txtOrdr
							Case 15
								Response.Write txtOdln
							Case 13
								Response.Write txtInv
						End Select %></td>
						<td><%=rd("DocNumber")%></td>
						<td class="style1"><%=rd("Owner")%></td>
					</tr>
					<% rd.movenext
					loop %>
				</table>
		</td>
	</tr>
	<% 
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetSODetailsBP" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@DocEntry") = DocEntry
	cmd("@DocType") = DocType
	rd.close
	rd.open cmd, , 3, 1
	If Not rd.Eof Then %>
	<tr class="GeneralTblBold2">
		<td><p align="center"><%=getsoConfDetailLngStr("DtxtBPS")%></td>
	</tr>
    <tr>
        <td width="100%">
				<table cellpadding="0" style="width: 100%" id="tblBP">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoConfDetailLngStr("DtxtName")%></td>
						<td><%=getsoConfDetailLngStr("LtxtRelationship")%></td>
						<td style="width: 120px"><%=getsoConfDetailLngStr("LtxtRelBP")%></td>
						<td><%=getsoConfDetailLngStr("LtxtRemarks")%></td>
					</tr>
					<%
					do while not rd.eof %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right; "><%=rd("Line")%></td>
						<td><%=rd("ParterId")%></td>
						<td><%=rd("OrlCode")%></td>
						<td style="width: 120px; height: 24px;"><%=myHTMLEncode(rd("RelatCard"))%></td>
						<td><%=myHTMLEncode(rd("Memo"))%></td>
					</tr>
					<% rd.movenext
					loop %>
				</table>
		</td>
	</tr>
	<% End If
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetSODetailsCompetition" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@DocType") = DocType
	cmd("@DocEntry") = DocEntry
	rd.close
	rd.open cmd, , 3, 1
	If Not rd.Eof Then %>
	<tr class="GeneralTblBold2">
		<td><p align="center"><%=getsoConfDetailLngStr("LtxtCompetition")%></td>
	</tr>
    <tr>
        <td width="100%">
				<table cellpadding="0" style="width: 100%" id="tblComp">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoConfDetailLngStr("DtxtName")%></td>
						<td><%=getsoConfDetailLngStr("LtxtThreat")%></td>
						<td><%=getsoConfDetailLngStr("LtxtRemarks")%></td>
						<td style="width: 80px"></td>
					</tr>
					<%
					do while not rd.eof %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><%=rd("Line")%></td>
						<td><%=rd("CompetId")%></td>
						<td><% Select Case rd("ThreatLevl")
						Case 1 %>|D:txtLow|<%
						Case 2 %>|D:txtMedium|<%
						Case 3 %>|D:txtHigh|<%
						End Select %></td>
						<td><%=myHTMLEncode(rd("Memo"))%></td>
						<td style="text-align: center;">
						<% If rd("Won") = "Y" Then %><%=getsoConfDetailLngStr("LtxtWon")%><% End If %></td>
					</tr>
					<% rd.movenext
					loop %>
				</table>
		</td>
	</tr>
	<% End If
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetSODetailsReasons" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@DocType") = DocType
	cmd("@DocEntry") = DocEntry
	rd.close
	rd.open cmd, , 3, 1
	If Not rd.Eof Then %>
	<tr class="GeneralTblBold2">
		<td><p align="center"><%=getsoConfDetailLngStr("LtxtReason")%></td>
	</tr>
    <tr>
        <td width="100%">
				<table cellpadding="0" style="width: 100%" id="tblReason">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoConfDetailLngStr("DtxtDescription")%></td>
					</tr>
					<%
					do while not rd.eof %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><%=rd("Line")%></td>
						<td><%=rd("ReasonId")%></td>
					</tr>
					<% rd.movenext
					loop %>
				</table>
		</td>
	</tr>
	<% End If
	If EnableSDK Then
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OOPR"
	cmd("@UserType") = userType
	cmd("@OP") = "O"
	set rg = cmd.execute()
	
	set rc = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFReadCols" & Session("ID")
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OOPR"
	cmd("@UserType") = userType
	cmd("@OP") = "O"
	rc.open cmd, , 3, 1
	
	set rd = Server.CreateObject("ADODB.RecordSet")
	
	do while not rg.eof
	 %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><% Select Case CInt(rg("GroupID"))
		Case -1 %><%=getsoConfDetailLngStr("DtxtUDF")%><%
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
			            <img src='../../AGECLI/pic.aspx?filename=<%=Picture%>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' border="0">
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
	alert('<%=getsoConfDetailLngStr("DtxtNoAccessObj")%>'.replace('{0}', '<%=getsoConfDetailLngStr("DtxtSO")%>'));
	window.close();
	</script>
<% End If %>

</body>

</html>
