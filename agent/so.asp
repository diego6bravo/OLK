<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->

<head>
<style type="text/css">
.style1 {
	direction: ltr;
}
</style>
</head>

<% If Not myApp.EnableOOPR Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" 
set rd = Server.CreateObject("ADODB.RecordSet") %>
<!--#include file="lang/so.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript">
var CalendarFormat = '<%=GetCalendarFormatString%>';
var DisplayFormat = '<%=myApp.DateFormat%>';
var SelDes = '<%=SelDes%>';
var dbName = '<%=Session("olkdb")%>';
</script>
<!--#include file="topGetValueSelect.inc"-->
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004


set rs = Server.CreateObject("ADODB.RecordSet")
set rg = Server.CreateObject("ADODB.RecordSet")
set rSdk = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SysID") = "OORP"
cmd("@ObsID") = "TORP"
set rs = cmd.execute()
If rs(0) = "Y" Then Response.Redirect "configErr.asp?errCmd=SO"

set rs = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKGetSOData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@LogNum") = Session("SORetVal")
set rs = cmd.execute()

OpprId = rs("OpprId")
EnableSDK = rs("EnableSDK") = "Y"
MaxStageNum = rs("MaxStageNum")

If EnableSDK Then
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OOPR"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rg.open cmd, , 3, 1


	cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OOPR"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rSdk.open cmd, , 3, 1
End If
 %>
<form method="POST" action="SOSubmit.asp" name="frmAddSO">
<input type="hidden" name="isUpdate" value="<% If Not IsNull(OpprId) Then %>True<% End If %>">
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><% If IsNull(OpprId) Then %><%=getsoLngStr("LttlNewSO")%><% Else %><%=getsoLngStr("LtxtEditSO")%> #<%=OpprId%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("DtxtClientCode")%></td>
				<td>
				<input class="inputDis" type="text" name="CardCode" size="20" value='<%=myHTMLEncode(rs("CardCode"))%>' style="width: 222px" readonly></td>
				<td class="GeneralTblBold2">
				<%=getsoLngStr("LtxtOpporName")%></td>
				<td>
				<input class="input" type="text" name="Name" size="35" value='<%=myHTMLEncode(rs("Name"))%>' maxlength="100" style="width: 222px" onfocus="this.select();" onchange="doProc('Name', 'S', this.value);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("DtxtName")%></td>
				<td>
				<input class="inputDis" type="text" name="CardName" size="20" value='<%=myHTMLEncode(rs("CardName"))%>' style="width: 222px" readonly></td>
				<td class="GeneralTblBold2">
				<%=getsoLngStr("LtxtOpporId")%></td>
				<td>
				<input class="inputDis" type="text" name="OpprId" size="20" value='<%=myHTMLEncode(rs("OpprId"))%>' style="width: 222px" readonly></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("DtxtContact")%></td>
				<td>
				<select class="input" size="1" name="CprCode" style="font-size:10px; font-family:Verdana; width: 222px" onchange="doProc('CprCode', 'N', this.value);">
				<option></option>
		        <% 
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@CardCode") = rs("CardCode")
				set rd = cmd.execute()
		        do while not rd.eof %>
		        <option <% If rd("CntctCode") = rs("CprCode") Then %>selected<% End If %> value="<%=rd("CntctCode")%>"><%=myHTMLEncode(rd("name"))%></option>
		        <% rd.movenext
		        loop
		        %>
		        </select></td>
				<td class="GeneralTblBold2">
				<%=getsoLngStr("LtxtStatus")%></td>
				<td>
				<select name="Status" class="input" size="1" onchange="doProc('Status', 'S', this.value);">
				<option value="O"><%=getsoLngStr("LtxtOpen")%></option>
				<option <% If rs("Status") = "W" Then %>selected<% End If %> value="W"><%=getsoLngStr("LtxtWon")%></option>
				<option <% If rs("Status") = "L" Then %>selected<% End If %> value="L"><%=getsoLngStr("LtxtLost")%></option>
				</select>
				</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("LtxtInvTotal")%></td>
				<td>
				<input class="inputDis" type="text" name="TotalInvoice" size="20" value='<%=rs("Currency")%><%=FormatNumber(CDbl(rs("TotalInvoice")), myApp.SumDec)%>' style="width: 222px; text-align: right;" readonly></td>
				<td class="GeneralTblBold2">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="GeneralTblBold2">
						<td><%=getsoLngStr("LtxtStartDate")%></td>
						<td align="right" width="16"><img border="0" src="images/cal.gif" id="btnOpenDate"></td>
					</tr>
				</table></td>
				<td>
				<input class="input" type="text" name="OpenDate" size="12" value='<%=rs("OpenDate")%>' readonly onclick="btnOpenDate.click()" onchange="SetCloseData(1);"></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("DtxtTerritory")%></td>
				<td>
				<input type="hidden" id="Territory" size="35" value='<%=rs("Territory")%>'>
				<input class="input" type="text" id="TerritoryDesc" size="35" value='<%=rs("TerritoryDesc")%>' onfocus="this.select();" onchange="fetchValue(0, null, document.getElementById('Territory'), this);" style="width: 222px"></td>
				<td class="GeneralTblBold2"><%=getsoLngStr("LtxtEndDate")%></td>
				<td>
				<input class="inputDis" type="text" name="CloseDate" size="12" value='<%=rs("CloseDate")%>' readonly></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("DtxtAgent")%></td>
				<td>
				<% 
						set ra = Server.CreateObject("ADODB.RecordSet")
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetAgents" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						If myAut.HasAuthorization(96) Then %>
						<select class="input" size="1" name="SlpCode" style="width: 222px" onchange="doProc('SlpCode', 'N', this.value);">
                      <% ra.open cmd, , 3, 1
                    	do while not ra.eof %>
                      <option value="<%=ra("SlpCode")%>" <% If rs("SlpCode") = ra("SLPCode") Then %>selected<% End If %>><%=myHTMLEncode(ra("SlpName"))%></option>
                      <% ra.movenext
                      loop %>
                      </select><% Else
                      cmd("@Filter") = rs("SlpCode")
                      set ra = cmd.execute()
                      %><%=ra("SlpName")%><input type="hidden" name="SlpCode" value="<%=rs("SlpCode")%>"><% End If %></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"><%=getsoLngStr("LtxtOwner")%></td>
				<td>
				<select class="input" size="1" name="Owner" style="width: 222px" onchange="doProc('Owner', 'N', this.value);">
				<option></option>
		        <% 
				set ro = Server.CreateObject("ADODB.RecordSet")
		        set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetEmployees" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				ro.open cmd, , 3, 1
		        do while not ro.eof %>
		        <option <% If ro("Code") = rs("Owner") Then %>selected<% End If %> value="<%=ro("Code")%>"><%=myHTMLEncode(ro("name"))%></option>
		        <% ro.movenext
		        loop
		        %>
		        </select></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<div id="itemDetTabs">
				<ul>
					<li><a href="#itemDetTabs-1" style="font-size: xx-small;"><%=getsoLngStr("LtxtPotential")%></a></li>
					<li><a href="#itemDetTabs-2" style="font-size: xx-small;"><%=getsoLngStr("LtxtGeneral")%></a></li>
					<li><a href="#itemDetTabs-3" style="font-size: xx-small;"><%=getsoLngStr("LtxtStages")%></a></li>
					<li><a href="#itemDetTabs-4" style="font-size: xx-small;"><%=getsoLngStr("DtxtBPS")%></a></li>
					<li><a href="#itemDetTabs-5" style="font-size: xx-small;"><%=getsoLngStr("LtxtCompetition")%></a></li>
					<li><a href="#itemDetTabs-6" style="font-size: xx-small;"><%=getsoLngStr("LtxtReason")%></a></li><% 
					If EnableSDK Then
						do while not rg.eof
						If CInt(rg("GroupID")) < 0 Then GroupID = "_1" Else GroupID = rg("GroupID")
						 %><li><a href="#itemDetTabs-<%=rg.bookmark+6%>" style="font-size: xx-small;"><%=rg("GroupName")%></a></li><%
						 rg.movenext
						 loop
						 rg.movefirst
					End If %>
				</ul>
				<div id="itemDetTabs-1" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
					<table cellpadding="0" cellspacing="0" style="width: 100%">
						<tr>
							<td style="width: 55%" valign="top">
							<table style="width: 100%" cellpadding="0">
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><%=getsoLngStr("LtxtClosePlan")%></td>
									<td>
									<input class="input" type="text" name="PredDateQty" size="35" onfocus="this.select();" value='<%=rs("PredDateQty")%>' style="width: 62px; font-size: 8pt; text-align: right;" onchange="SetCloseData(2);"><select class="input" size="1" name="DifType" style="width: 158px; font-size: 8pt;" onchange="SetCloseData(3);">
									<option <% If rs("DifType") = "M" Then %>selected<% End If %> value="M"><%=getsoLngStr("DtxtMonths")%></option>
									<option <% If rs("DifType") = "W" Then %>selected<% End If %> value="W"><%=getsoLngStr("DtxtWeeks")%></option>
									<option <% If rs("DifType") = "D" Then %>selected<% End If %> value="D"><%=getsoLngStr("DtxtDays")%></option>
							        </select></td>
								</tr>
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr class="GeneralTblBold2">
										<td><%=getsoLngStr("LtxtCloseDate")%></td>
										<td align="right" width="16"><img border="0" src="images/cal.gif" id="btnPredDate"></td>
									</tr>
								</table></td>
									<td>
									<input class="input" type="text" readonly name="txtPredDate" size="12" value='<%=rs("PredDate")%>' style="font-size: 8pt;" onchange="SetCloseData(4);" onclick="btnPredDate.click();"></td>
								</tr>
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><%=getsoLngStr("LtxtMaxSumLoc")%></td>
									<td>
									<input class="input" type="text" id="MaxSumLoc" onchange="SetSumData(1);" onfocus="this.select();" onkeydown="return valKeyNumDec(event);"  size="35" value='<% If Not IsNull(rs("MaxSumLoc")) Then Response.Write FormatNumber(CDbl(rs("MaxSumLoc")), myApp.SumDec)%>' style="width: 222px; font-size: 8pt; text-align: right;"></td>
								</tr>
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><%=getsoLngStr("LtxtWtSumLoc")%></td>
									<td>
									<input class="input" type="text" id="WtSumLoc" onchange="SetSumData(2);" onfocus="this.select();" onkeydown="return valKeyNumDec(event);"  size="35" value='<% If Not IsNull(rs("WtSumLoc")) Then Response.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)%>' style="width: 222px; font-size: 8pt; text-align: right;"></td>
								</tr>
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><%=getsoLngStr("LtxtPrcntProf")%></td>
									<td>
									<input class="input" type="text" id="PrcntProf" onchange="SetSumData(4);" onfocus="this.select();" onkeydown="return valKeyNumDec(event);"  size="35" value='<% If Not IsNull(rs("PrcntProf")) Then Response.Write FormatNumber(CDbl(rs("PrcntProf")), myApp.PercentDec)%>' style="width: 222px; font-size: 8pt; text-align: right;"></td>
								</tr>
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><%=getsoLngStr("LtxtSumProfL")%></td>
									<td>
									<input class="input" type="text" id="SumProfL" onchange="SetSumData(3);" onfocus="this.select();" onkeydown="return valKeyNumDec(event);"  size="35" value='<% If Not IsNull(rs("SumProfL")) Then Response.Write FormatNumber(CDbl(rs("SumProfL")), myApp.SumDec)%>' style="width: 222px; font-size: 8pt; text-align: right;"></td>
								</tr>
								<tr class="GeneralTbl">
									<td class="GeneralTblBold2"><%=getsoLngStr("LtxtIntRate")%></td>
									<td>
									<select class="input" size="1" name="IntRate" style="width: 222px; font-size: 8pt;" onchange="doProc('IntRate', 'N', this.value);">
									<option></option>
							        <% 
							        set cmd = Server.CreateObject("ADODB.Command")
									cmd.ActiveConnection = connCommon
									cmd.CommandType = &H0004
									cmd.CommandText = "DBOLKGetIntRate" & Session("ID")
									cmd.Parameters.Refresh()
									cmd("@LanID") = Session("LanID")
									set rd = cmd.execute()
							        do while not rd.eof %>
							        <option <% If rd("Code") = rs("IntRate") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
							        <% rd.movenext
							        loop
							        %>
							        </select>
							        </td>
								</tr>
							</table>
							</td>
							<td style="width: 45%" valign="top">
							<div style="height: 252px; overflow: auto;">
							<table cellpadding="0" style="width: 100%" id="tblIntRange">
								<tr class="GeneralTblBold2">
									<td colspan="3" style="height: 21px"><%=getsoLngStr("LtxtIntRange")%></td>
								</tr>
								<tr class="GeneralTblBold2">
									<td style="width: 40px; text-align: center;">#</td>
									<td><%=getsoLngStr("DtxtDescription")%></td>
									<td style="width: 100px; text-align:center;"><%=getsoLngStr("LtxtPrimary")%></td>
								</tr>
								<%
								set ri = Server.CreateObject("ADODB.RecordSet")
								set cmd = Server.CreateObject("ADODB.Command")
								cmd.ActiveConnection = connCommon
								cmd.CommandType = &H0004
								cmd.CommandText = "DBOLKGetIntRange" & Session("ID")
								cmd.Parameters.Refresh()
								cmd("@LanID") = Session("LanID")
								ri.open cmd, , 3, 1
								
								set cmd = Server.CreateObject("ADODB.Command")
								cmd.ActiveConnection = connCommon
								cmd.CommandType = &H0004
								cmd.CommandText = "DBOLKGetSOIntRangeData" & Session("ID")
								cmd.Parameters.Refresh()
								cmd("@LanID") = Session("LanID")
								cmd("@LogNum") = Session("SORetVal")
								rd.close
								rd.open cmd, , 3, 1
								do while not rd.eof %>
								<tr class="GeneralTbl" id="intRangeNum<%=rd("Line")%>">
									<td style="width: 40px; text-align: right;"><%=rd.bookmark%></td>
									<td>
									<select class="input" size="1" name="IntRange<%=rd("Line")%>" style="width: 100%;" onchange="doProcLine(4, <%=rd("Line")%>, 'IntId', 'N', this.value);">
									<option></option>
									<% ri.movefirst
									do while not ri.eof %>
									<option <% If rd("IntId") = ri("Code") Then %>selected<% End If %> value="<%=ri("Code")%>"><%=ri("Name")%></option>
									<% ri.movenext
									loop %>
							        </select></td>
									<td style="width: 100px; text-align:center;"><input onclick="doProcLine(4, <%=rd("Line")%>, 'Prmry', 'S', 'Y');" type="radio" name="IntRangePrim" class="noborder" <% If rd("Prmry") = "Y" Then %>checked<% End If %> value="<%=rd("Line")%>"></td>
								</tr>
								<% rd.movenext
								loop %>
								<tr class="GeneralTbl">
									<td style="width: 40px; text-align: right;">&nbsp;<span id="intRangeCountText"><%=rd.RecordCount+1%></span><input type="hidden" id="intRangeCount" value="<%=rd.RecordCount+1%>"></td>
									<td>
									<select class="input" size="1" id="IntRangeNew" style="width: 100%;" onchange="doNewInt(this.value);">
									<option></option>
									<% If ri.recordcount > 0 Then ri.movefirst
									do while not ri.eof %>
									<option value="<%=ri("Code")%>"><%=ri("Name")%></option>
									<% ri.movenext
									loop %>
							        </select></td>
									<td style="width: 100px;">&nbsp;</td>
								</tr>
							</table>
							</div>
							</td>
						</tr>
					</table>
				</div>
				<div id="itemDetTabs-2" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
					<table style="width: 100%">
						<tr class="GeneralTbl">
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoLngStr("LtxtChnCode")%></td>
							<td style="width: 30%">
							<input class="input" type="text" id="ChnCrdCode" size="20" value='<%=myHTMLEncode(rs("ChnCrdCode"))%>' onfocus="this.select();" onchange="fetchValue(1, null, this, document.getElementById('ChnCrdName'));" style="width: 222px; font-size: 8pt;"></td>
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoLngStr("DtxtProject")%></td>
							<td style="width: 30%">
							<select class="input" size="1" name="PrjCode" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProc('PrjCode', 'S', this.value);">
							<option></option>
					        <% set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							cmd.CommandText = "DBOLKGetProjects" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							set rd = cmd.execute()
							do while not rd.eof %>
							<option value="<%=rd("PrjCode")%>" <% If rs("PrjCode") = rd("PrjCode") Then %>selected<% End If %>><%=myHTMLEncode(rd("PrjName"))%></option>
							<% rd.movenext
							loop %>
					        </select></td>
						</tr>
						<tr class="GeneralTbl">
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoLngStr("LtxtChnName")%></td>
							<td style="width: 30%">
							<input class="inputDis" type="text" id="ChnCrdName" size="20" value='<%=myHTMLEncode(rs("ChnCrdName"))%>' style="width: 222px; font-size: 8pt;"></td>
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoLngStr("LtxtInfSource")%></td>
							<td style="width: 30%">
							<select class="input" size="1" name="Source" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProc('Source', 'N', this.value);">
							<option></option>
					        <% 
					        set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							cmd.CommandText = "DBOLKGetInfSource" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							set rd = cmd.execute()
					        do while not rd.eof %>
					        <option <% If rd("Code") = rs("Source") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
					        <% rd.movenext
					        loop
					        %>
					        </select></td>
						</tr>
						<tr class="GeneralTbl">
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoLngStr("LtxtChnCnt")%></td>
							<td style="width: 30%">
							<select class="input" size="1" id="ChnCrdCon" style="font-size:10px; font-family:Verdana" onchange="if(!isUpdating)doProc('ChnCrdCon', 'N', this.value);">
							<option></option>
					        <% 
					        If Not IsNull(rs("ChnCrdCode")) Then 
							set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@CardCode") = rs("ChnCrdCode")
							set rd = cmd.execute()
					        do while not rd.eof %>
					        <option <% If rd("CntctCode") = rs("ChnCrdCon") Then %>selected<% End If %> value="<%=rd("CntctCode")%>"><%=myHTMLEncode(rd("name"))%></option>
					        <% rd.movenext
					        loop
					        End If
					        %>
					        </select></td>
							<td style="width: 20%" class="GeneralTblBold2"><%=getsoLngStr("LtxtIndustry")%></td>
							<td style="width: 30%">
							<select class="input" size="1" name="Industry" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProc('Industry', 'N', this.value);">
							<option></option>
					        <% 
					        set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							cmd.CommandText = "DBOLKGetIndustry" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							set rd = cmd.execute()
					        do while not rd.eof %>
					        <option <% If rd("Code") = rs("Industry") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=myHTMLEncode(rd("name"))%></option>
					        <% rd.movenext
					        loop
					        %>
					        </select></td>
						</tr>
						<tr class="GeneralTbl">
							<td style="width: 20%; vertical-align: top; padding-top: 2px;" class="GeneralTblBold2"><%=getsoLngStr("LtxtRemarks")%></td>
							<td colspan="3">
							<textarea cols="20" name="Memo" rows="12" style="width: 100%; font-size: 8pt;" onchange="doProc('Memo', 'S', this.value);"><%=rs("Memo")%></textarea></td>
						</tr>
					</table>
				</div>
				<div id="itemDetTabs-3" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
				<div style="height: 252px; overflow: auto;">
				<table cellpadding="0" style="width: 100%" id="tblStages">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoLngStr("LtxtStage")%></td>
						<td><%=getsoLngStr("LtxtStartDate")%></td>
						<td><%=getsoLngStr("LtxtEndDate")%></td>
						<td><%=getsoLngStr("DtxtAgent")%></td>
						<td>%</td>
						<td><%=getsoLngStr("LtxtMaxSumLoc")%></td>
						<td><%=getsoLngStr("LtxtWtSumLoc")%></td>
						<td><%=getsoLngStr("LtxtDocType")%></td>
						<td class="style1"><%=getsoLngStr("LtxtDocNum")%></td>
						<td class="style1"><%=getsoLngStr("LtxtOwner")%></td>
					</tr>
					<%
					
					set rx = Server.CreateObject("ADODB.RecordSet")
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetStages" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					rx.open cmd, , 3, 1
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetSOStagesData" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = Session("SORetVal")
					rd.close
					rd.open cmd, , 3, 1
					do while not rd.eof
					readOnly = rd("Line") <> MaxStageNum %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><%=rd("LineDesc")%></td>
						<td>
						<select <% If not readOnly Then %>class="input"<% Else %>class="inputDis" disabled<% End If %> size="1" id='StageStepID<%=rd("Line")%>' onchange="doProcStepID(<%=rd("Line")%>, this.value);">
                      <% rx.movefirst
                    	do while not rx.eof %>
                      <option value="<%=rx("Code")%>" <% If rd("StepId") = rx("Code") Then %>selected<% End If %>><%=myHTMLEncode(rx("Name"))%></option>
                      <% rx.movenext
                      loop %>
                      </select></td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="GeneralTblBold2">
								<td align="right" width="16"><img <% If readOnly Then %>disabled<% End IF %> border="0" src="images/cal.gif" id="btnStageOpenDate<%=rd("Line")%>"></td>
								<td><input <% If not readOnly Then %>class="input"<% Else %>class="inputDis"<% End If %> type="text" onclick="btnStageOpenDate<%=rd("Line")%>.click();" id="StageOpenDate<%=rd("Line")%>" size="12" value='<%=rd("OpenDate")%>' readonly onchange="doProcLine(1, <%=rd("Line")%>, 'OpenDate', 'D', this.value);"></td>
							</tr>
						</table>
						</td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="GeneralTblBold2">
								<td align="right" width="16"><img <% If readOnly Then %>disabled<% End IF %> border="0" src="images/cal.gif" id="btnStageCloseDate<%=rd("Line")%>"></td>
								<td><input <% If not readOnly Then %>class="input"<% Else %>class="inputDis"<% End If %> type="text" onclick="btnStageCloseDate<%=rd("Line")%>.click();" id="StageCloseDate<%=rd("Line")%>" size="12" value='<%=rd("CloseDate")%>' readonly onchange="doProcLine(1, <%=rd("Line")%>, 'CloseDate', 'D', this.value);"></td>
							</tr>
						</table>
						</td>
						<td>
						<select <% If not readOnly Then %>class="input"<% Else %>class="inputDis" disabled<% End If %> size="1" id="StageSlpCode<%=rd("Line")%>" onchange="doProcLine(1, <%=rd("Line")%>, 'SlpCode', 'I', this.value);">
                      <% ra.movefirst
                    	do while not ra.eof %>
                      <option value="<%=ra("SlpCode")%>" <% If rd("SlpCode") = ra("SlpCode") Then %>selected<% End If %>><%=myHTMLEncode(ra("SlpName"))%></option>
                      <% ra.movenext
                      loop %>
                      </select></td>
						<td>
						<input onkeydown="return valKeyNumDec(event);" <% If not readOnly Then %>class="input" onfocus="this.select();"<% Else %>class="inputDis" readonly<% End If %> type="text" id='StageClosePrcnt<%=rd("Line")%>' size="8" value='<%=FormatNumber(CDbl(rd("ClosePrcnt")), myApp.PercentDec)%>' style="text-align: right;" onchange="doProcStepClosePer(<%=rd("Line")%>, this.value);"></td>
						<td>
						<input onkeydown="return valKeyNumDec(event);" <% If not readOnly Then %>class="input" onfocus="this.select();"<% Else %>class="inputDis" readonly<% End If %> type="text" id="StageMaxSumLoc<%=rd("Line")%>" size="20" value='<% If Not IsNull(rd("MaxSumLoc")) Then Response.Write FormatNumber(CDbl(rd("MaxSumLoc")), myApp.SumDec)%>' style="font-size: 8pt; text-align: right;" onchange="doProcStepMaxSum(<%=rd("Line")%>, this.value);"></td>
						<td>
						<input onkeydown="return valKeyNumDec(event);" <% If not readOnly Then %>class="input" onfocus="this.select();"<% Else %>class="inputDis" readonly<% End If %> type="text" id='StageWtSumLoc<%=rd("Line")%>' size="20" value='<% If Not IsNull(rd("WtSumLoc")) Then Response.Write FormatNumber(CDbl(rd("WtSumLoc")), myApp.SumDec)%>' style="font-size: 8pt; text-align: right;" onchange="doProcStepWtSum(<%=rd("Line")%>, this.value);"></td>
						<td>
						<select <% If not readOnly Then %>class="input"<% Else %>class="inputDis" disabled<% End If %> size="1" id='StageObjType<%=rd("Line")%>' onchange="doChStageObjType(<%=rd("Line")%>,this.value);">
						<option></option>
						<option <% If rd("ObjType") = 23 Then %>selected<% End If %> value="23"><%=txtQuote%></option>
						<option <% If rd("ObjType") = 17 Then %>selected<% End If %> value="17"><%=txtOrdr%></option>
						<option <% If rd("ObjType") = 15 Then %>selected<% End If %> value="15"><%=txtOdln%></option>
						<option <% If rd("ObjType") = 13 Then %>selected<% End If %> value="13"><%=txtInv%></option>
				        </select></td>
						<td>
						<input onkeydown="return valKeyNumSearch(event);" onfocus="this.select();" <% If not readOnly and rd("ObjType") <> "" Then %>class="input"<% Else %>class="inputDis" readonly<% End If %> type="text" id='StageDocNumber<%=rd("Line")%>' size="20" value='<%=rd("DocNumber")%>' onchange="fetchValue(2, <%=rd("Line")%>, this, null);" style="font-size: 8pt; text-align: right;"></td>
						<td class="style1">
						<select <% If not readOnly Then %>class="input"<% Else %>class="inputDis" disabled<% End If %> size="1" id="StageOwner<%=rd("Line")%>" onchange="doProcLine(1, <%=rd("Line")%>, 'Owner', 'I', this.value);">
						<option></option>
				        <% 
						If ro.recordcount > 0 Then ro.movefirst
				        do while not ro.eof %>
				        <option <% If ro("Code") = rd("Owner") Then %>selected<% End If %> value="<%=ro("Code")%>"><%=myHTMLEncode(ro("name"))%></option>
				        <% ro.movenext
				        loop
				        %>
				        </select></td>
					</tr>
					<% rd.movenext
					loop %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><span id="StageCountText"><%=rd.recordcount+1%></span><input type="hidden" id="StageCount" value="<%=rd.recordcount+1%>"><input type="hidden" id="MaxStageNum" value="<%=MaxStageNum%>"></td>
						<td>
						<select class="input" size="1" id='NewStageStepID' onchange="doNewStage(this.value);">
						<option></option>
                      <% rx.movefirst
                    	do while not rx.eof %>
                      <option value="<%=rx("Code")%>"><%=myHTMLEncode(rx("Name"))%></option>
                      <% rx.movenext
                      loop %>
                      </select></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>
						<select style="display: none;" size="1" id="NewStageSlpCode">
                      <% ra.movefirst
                    	do while not ra.eof %>
                      <option value="<%=ra("SlpCode")%>"><%=myHTMLEncode(ra("SlpName"))%></option>
                      <% ra.movenext
                      loop %>
                      </select></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>
						<select style="display: none;" size="1" id='NewStageObjType'>
						<option></option>
						<option value="23"><%=txtQuote%></option>
						<option value="17"><%=txtOrdr%></option>
						<option value="15"><%=txtOdln%></option>
						<option value="13"><%=txtInv%></option>
				        </select></td>
						<td>&nbsp;</td>
						<td class="style1">
						<select style="display: none;" size="1" id="NewStageOwner">
				        <% 
						If ro.recordcount > 0 Then ro.movefirst
				        do while not ro.eof %>
				        <option value="<%=ro("Code")%>"><%=myHTMLEncode(ro("name"))%></option>
				        <% ro.movenext
				        loop
				        %>
				        </select></td>
					</tr>
				</table>
				</div>
				</div>
				<div id="itemDetTabs-4" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
				<div style="height: 252px; overflow: auto;">
				<table cellpadding="0" style="width: 100%" id="tblBP">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoLngStr("DtxtName")%></td>
						<td><%=getsoLngStr("LtxtRelationship")%></td>
						<td style="width: 120px"><%=getsoLngStr("LtxtRelBP")%></td>
						<td><%=getsoLngStr("LtxtRemarks")%></td>
					</tr>
					<%
					ri.close
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetPartners" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					ri.open cmd, , 3, 1
					
					set rx = Server.CreateObject("ADODB.RecordSet")
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetRelationship" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					rx.open cmd, , 3, 1
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetSOBPData" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = Session("SORetVal")
					rd.close
					rd.open cmd, , 3, 1
					do while not rd.eof %>
					<tr class="GeneralTbl" id="bpNum<%=rd("Line")%>">
						<td style="width: 40px; text-align: right; "><%=rd.bookmark%></td>
						<td style="height: 24px">
							<select class="input" size="1" name="BPPartnerId<%=rd("Line")%>" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProcBP(<%=rd("Line")%>, this.value);">
							<option></option>
					        <% ri.movefirst
					        do while not ri.eof %>
					        <option value="<%=ri("Code")%>" <% If ri("Code") = rd("ParterId") Then %>selected<% End If %>><%=ri("Name")%></option>
					        <% ri.movenext
					        loop %>
					        </select></td>
						<td style="height: 24px">
							<select class="input" size="1" id='BPOrlCode<%=rd("Line")%>' style="width: 222px; font-size:10px; font-family:Verdana;">
							<option></option>
					        <% rx.movefirst
					        do while not rx.eof %>
					        <option value="<%=rx("Code")%>" <% If rx("Code") = rd("OrlCode") Then %>selected<% End If %>><%=rx("Name")%></option>
					        <% rx.movenext
					        loop %>
					        </select></td>
						<td style="width: 120px; "><input type="text" class="inputDis" id="RelatCard<%=rd("Line")%>" maxlength="15" value="<%=myHTMLEncode(rd("RelatCard"))%>" readonly style="width: 100%;"></td>
						<td style="height: 24px"><input type="text" class="input" maxlength="50" value="<%=myHTMLEncode(rd("Memo"))%>" id="BPMemo<%=rd("Line")%>" style="width: 100%;" onchange="doProcLine(2, <%=rd("Line")%>, 'Memo', 'S', this.value);"></td>
					</tr>
					<% rd.movenext
					loop %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><span id="NewBPCountText"><%=rd.recordcount+1%></span><input type="hidden" id="NewBPCount" value="<%=rd.recordcount+1%>"></td>
						<td>
							<select class="input" size="1" id="NewBP" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doNewBP(this.value);">
							<option></option>
					        <% If ri.recordcount > 0 Then ri.movefirst
					        do while not ri.eof %>
					        <option value="<%=ri("Code")%>"><%=ri("Name")%></option>
					        <% ri.movenext
					        loop %>
					        </select></td>
						<td>
							<select size="1" id='NewBPOrlCode' style="display: none;">
					        <% If rx.recordcount > 0 Then rx.movefirst
					        do while not rx.eof %>
					        <option value="<%=rx("Code")%>"><%=rx("Name")%></option>
					        <% rx.movenext
					        loop %>
					        </select></td>
						<td style="width: 120px">&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
				</table>
				</div>
				</div>
				<div id="itemDetTabs-5" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
				<div style="height: 252px; overflow: auto;">
				<table cellpadding="0" style="width: 100%" id="tblComp">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoLngStr("DtxtName")%></td>
						<td><%=getsoLngStr("LtxtThreat")%></td>
						<td><%=getsoLngStr("LtxtRemarks")%></td>
						<td style="width: 80px"><%=getsoLngStr("LtxtWon")%></td>
					</tr>
					<%
					ri.close
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetCompetitors" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					ri.open cmd, , 3, 1
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetSOCompetitionData" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = Session("SORetVal")
					rd.close
					rd.open cmd, , 3, 1
					do while not rd.eof %>
					<tr class="GeneralTbl" id="compNum<%=rd("Line")%>">
						<td style="width: 40px; text-align: right;"><%=rd.bookmark%></td>
						<td>
							<select class="input" size="1" name="CompPartnerId<%=rd("Line")%>" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProcCompetition(<%=rd("Line")%>, this.value);">
							<option></option>
					        <% ri.movefirst
					        do while not ri.eof %>
					        <option value="<%=ri("Code")%>" <% If ri("Code") = rd("CompetId") Then %>selected<% End If %>><%=ri("Name")%></option>
					        <% ri.movenext
					        loop %>
					        </select></td>
						<td><% Select Case rd("ThreatLevl")
						Case 1
							ThreatDesc = getsoLngStr("DtxtLow")
						Case 2
							ThreatDesc = getsoLngStr("DtxtMedium")
						Case 3
							ThreatDesc = getsoLngStr("DtxtHigh")
						End Select %><input type="text" class="inputDis" maxlength="15" id="Threat<%=rd("Line")%>" value="<%=ThreatDesc%>" readonly style="width: 100%;"></td>
						<td><input type="text" id="CompMemo<%=rd("Line")%>" class="input" maxlength="50" value="<%=myHTMLEncode(rd("Memo"))%>" style="width: 100%;" onchange="doProcLine(3, <%=rd("Line")%>, 'Memo', 'S', this.value);"></td>
						<td style="text-align: center;">
						<input id="CompChkWon<%=rd("Line")%>" type="checkbox" <% If rd("Won") = "Y" Then %>checked<% End If %> class="noborder" value="Y" onclick="doProcLine(3, <%=rd("Line")%>, 'Won', 'S', (this.checked ? 'Y' : 'N'));"></td>
					</tr>
					<% rd.movenext
					loop %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><span id="CompNewCountText"><%=rd.recordcount + 1%></span><input type="hidden" id="CompNewCount" value="<%=rd.recordcount + 1%>"></td>
						<td>
							<select class="input" size="1" id="CompNew" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doCompNew(this.value);">
							<option></option>
					        <% If ri.recordcount > 0 Then ri.movefirst
					        do while not ri.eof %>
					        <option value="<%=ri("Code")%>"><%=ri("Name")%></option>
					        <% ri.movenext
					        loop %>
					        </select></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td style="text-align: center;">&nbsp;</td>
					</tr>
				</table>
				</div>
				</div>
				<div id="itemDetTabs-6" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
				<div style="height: 252px; overflow: auto;">
				<table cellpadding="0" style="width: 100%" id="tblReason">
					<tr class="GeneralTblBold2">
						<td style="width: 40px; text-align: center;">#</td>
						<td><%=getsoLngStr("DtxtDescription")%></td>
					</tr>
					<%
					ri.close
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetReasons" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					ri.open cmd, , 3, 1
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetSOReasonsData" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = Session("SORetVal")
					rd.close
					rd.open cmd, , 3, 1
					do while not rd.eof %>
					<tr class="GeneralTbl" id="reasonNum<%=rd("Line")%>">
						<td style="width: 40px; text-align: right;"><%=rd.bookmark%></td>
						<td>
						<select class="input" size="1" name="Reason<%=rd("Line")%>" style="width: 100%; font-size:10px; font-family:Verdana;" onchange="doProcLine(5, <%=rd("Line")%>, 'ReasonId', 'N', this.value);">
						<option></option>
				        <% ri.movefirst
				        do while not ri.eof %>
				        <option value="<%=ri("Code")%>" <% If ri("Code") = rd("ReasonId") Then %>selected<% End If %>><%=ri("Name")%></option>
				        <% ri.movenext
				        loop %>
				        </select></td>
					</tr>
					<% rd.movenext
					loop %>
					<tr class="GeneralTbl">
						<td style="width: 40px; text-align: right;"><span id="ReasonNewCountText"><%=rd.recordcount + 1%></span><input type="hidden" id="ReasonNewCount" value="<%=rd.recordcount + 1%>"></td>
						<td>
							<select class="input" size="1" id="ReasonNew" style="width: 100%; font-size:10px; font-family:Verdana;" onchange="doNewReason(this.value);">
							<option></option>
					        <% If ri.recordcount > 0 Then ri.movefirst
					        do while not ri.eof %>
					        <option value="<%=ri("Code")%>"><%=ri("Name")%></option>
					        <% ri.movenext
					        loop %>
					        </select></td>
					</tr>
				</table>
				</div>
				</div>
				<% 
				If EnableSDK Then
					do while not rg.eof
					If CInt(rg("GroupID")) < 0 Then GroupID = "_1" Else GroupID = rg("GroupID")
					 %>
					<div id="itemDetTabs-<%=rg.bookmark+6%>" style="height: 260px; background-color: #FFFFFF; overflow: auto;">
					<div style="height: 252px; overflow: auto;">
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
					</div>
					</div><%
					 rg.movenext
					 loop
					 rg.movefirst
				End If %>
			</div>
		</td>
	</tr>
	<tr class="GeneralTbl" align="center">
		<td>
		<p align="center">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td>
					  <input type="button" value="<% If IsNull(OpprId) Then %><%=getsoLngStr("DtxtAdd")%><% Else %><%=getsoLngStr("DtxtUpdate")%><% End If %>" name="btnAdd" onclick="if(valFrm()) { setActFlow(); doFlowAlert(); }"></td>
					<td>
						<% If IsNull(OpprId) Then %>
					  <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					  <input type="button" value="<%=getsoLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getsoLngStr("LtxtConfCancel")%>'))window.location.href='soCancel.asp?isUpdate=<%=JBool(Not IsNull(OpprId))%>'"><% End If %></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="newSOSubmit">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="doSubmit" value="Y">
<input type="hidden" name="Confirm" value="">
</form>
<script language="javascript">
var txtStartDate = '<%=getsoLngStr("LtxtStartDate")%>';
var txtDueDate = '|L:txtDueDate|';
var txtTime = '|L:txtTime|';
var txtStartTime = '|L:txtStartTime|';
var txtEndTime = '|L:txtEndTime|';
var txtValNumVal = '|D:txtValNumVal|';

function valFrm()
{
	<% If EnableSDK Then 
	cmd.CommandText = "DBOLKGetUDFNotNull" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = "V"
	cmd("@TableID") = "OOPR"
	cmd("@OP") = "O"
	set rd = cmd.execute()
	do while not rd.eof %>
	if (document.frmAddSO.U_<%=rd("AliasID")%>.value == "")
	{
		alert('|L:txtConfFld|'.replace('{0}', '<%=Replace(rd("Descr"), "'", "\'")%>'));
		showUDF(<%=rd("GroupID")%>);
		document.frmAddSO.U_<%=rd("AliasID")%>.focus();
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

<% If MaxStageNum <> "" Then %>
Calendar.setup({
    inputField     :    "StageOpenDate<%=MaxStageNum%>",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnStageOpenDate<%=MaxStageNum%>",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
Calendar.setup({
    inputField     :    "StageCloseDate<%=MaxStageNum%>",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnStageCloseDate<%=MaxStageNum%>",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
<% End If %>
var txtLow = '<%=getsoLngStr("DtxtLow")%>';
var txtMedium = '<%=getsoLngStr("DtxtMedium")%>';
var txtHigh= '<%=getsoLngStr("DtxtHigh")%>';

function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('|D:txtValFldMaxChar|'.replace('{0}', FSize));
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
							alert('|D:txtValNumVal|');
						}
						else if (parseInt(getNumericVB(Field.value)) < 1)
						{
							Field.value = '';
							alert('|D:txtValNumMinVal|'.replace('{0}', '1'));
						}
						else if (parseInt(getNumericVB(Field.value)) > 2147483647)
						{
							alert('|D:txtValNumMaxVal|'.replace('{0}', '2147483647'));
							Field.value = 2147483647;
						}
						else if (Field.value.indexOf('<%=GetFormatDec%>') > -1)
						{
							Field.value = '';
							alert('|D:txtValNumValWhole|');
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
					alert('|D:txtValNumVal|');
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
<script language="javascript" src="addSO/addSO.js.asp"></script>
<script language="javascript" src="addSO/addSO.js"></script>
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
			            	    	<img border="0" src="images/<% If rSdk("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rSdk("AliasID")%>" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Act&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',400,250,'yes', 'yes', document.frmAddSO.U_<%=rSdk("AliasID")%>, '<%=ProcType%>')"<% End If %>>
			            	    </td>
			            	    <% End If %>
			            	</tr>
			              </table>
			            </td>
			            <td dir="ltr" bgcolor="#EAF5FF"><% If rSdk("DropDown") = "Y" or not IsNull(rSdk("RTable")) then
			            	set rd = Server.CreateObject("ADODB.RecordSet")
			            	If rSdk("DropDown") = "Y" Then 
								cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
								cmd.Parameters.Refresh()
								cmd("@LanID") = Session("LanID")
								cmd("@TableID") = "OOPR"
								cmd("@FieldID") = rSdk("FieldID")
								rd.open cmd, , 3, 1
							  Else
							  	sql = "select Code, Name from [@" & rSdk("RTable") & "] order by 2"
							  	set rd = conn.execute(sql)
							  End If
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
								<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddSO.U_<%=rSdk("AliasID")%>.value = '';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
					<% ElseIf rSdk("TypeID") = "A" and rSdk("EditType") = "I" Then %>
						<table cellpadding="0" cellspacing="2" border="0">
							<tr>
								<td><img src="pic.aspx?filename=<% If IsNull(rs(InsertID)) Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rSdk("AliasID")%>" border="1">
								<input type="hidden" name="U_<%=rSdk("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
								<td width="16" valign="bottom"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.frmAddSO.U_<%=rSdk("AliasID")%>.value = '';document.frmAddSO.imgU_<%=rSdk("AliasID")%>.src='pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand"></td>
							</tr>
							<tr>
								<td colspan="2" height="22">
								<p align="center">
								<input type="button" value="|D:txtAddImg|" name="B1" onclick="javascript:getImg(document.frmAddSO.U_<%=rSdk("AliasID")%>, document.frmAddSO.imgU_<%=rSdk("AliasID")%>,180);"></td>
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
									<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddSO.U_<%=rSdk("AliasID")%>.value = '';doProc('U_<%=rSdk("AliasID")%>', '<%=ProcType%>', '');">
								</td>
							  </tr>
							</table>
							<% End If %><% End If %></td></tr><% End Sub %>
<!--#include file="agentBottom.asp"-->