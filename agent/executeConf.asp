<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (Session("useraccess") = "P" or Session("HasActionConfAut")) Then Response.Redirect "unauthorized.asp" %>
<!--#include file="lang/executeConf.asp" -->
<html dir="ltr">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>
.trProcStatus
{
	text-align: center; font-family: Verdana; font-size: xx-small; vertical-align: middle;
}
</style>
</head>
<script language="javascript">
var txtTransactionPool = '<%=getexecuteConfLngStr("LtxtTransactionPool")%>';
var txtConfirmed = '<%=getexecuteConfLngStr("LtxtConfirmed")%>';
var txtRejected = '<%=getexecuteConfLngStr("DtxtRejected")%>';
var txtCanceled = '<%=getexecuteConfLngStr("LtxtCanceled")%>';
var txtProcesing = '<%=getexecuteConfLngStr("DtxtProcesing")%>';
var txtError = '<%=getexecuteConfLngStr("DtxtError")%>';
var txtRetry = '<%=getexecuteConfLngStr("DtxtRetry")%>';
var txtNote = '<%=getexecuteConfLngStr("DtxtNote")%>';
var msgAllreadyProc = '<%=getexecuteConfLngStr("LmsgAllreadyProc")%>';
var SumDec = <%=myApp.SumDec%>;
</script>
<SCRIPT LANGUAGE="JavaScript" src="executeConf.js"></SCRIPT>
<%
execType = Request("Type")
Select Case execType
	Case "A"
		strFldID = "ID"
		strFormFldID = "ID"
	Case "C", "I", "R"
		strFldID = "LogNum"
		strFormFldID = "LogNum"
	Case "D"
		strFldID = "ObjectEntry"
		strFormFldID = "LogNum"
End Select

set rv = Server.CreateObject("ADODB.RecordSet")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetExecConfRepRead" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@ExecType") = execType
set rc = Server.CreateObject("ADODB.RecordSet")
rc.open cmd, , 3, 1

cmd.CommandText = "DBOLKGetExecConfRepLinkVars" & Session("ID")
cmd("@ExecType") = execType
set rcVars = Server.CreateObject("ADODB.RecordSet")
rcVars.open cmd, , 3, 1

rcColSpan = rc.recordcount
rc.Filter = "LinkObject <> null"
rcColSpan = rcColSpan + rc.recordcount
rc.Filter = ""

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetExecConf" & execType & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid") 

Select Case execType
	Case "A"
		ColSpan = 12 + rcColSpan
		
		cmd("@UserAccess") = Session("UserAccess")
	Case "C"
		ColSpan = 13 + rcColSpan
		
		cmd("@UserAccess") = Session("UserAccess")
	Case "D"
		ColSpan = 14 + rcColSpan
		
		cmd("@UserAccess") = Session("UserAccess") 
	Case "I"
		ColSpan = 11 + rcColSpan 
	Case "R"
		ColSpan = 14 + rcColSpan 
End Select
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1
%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td>&nbsp;<% Select Case execType
			Case "A" %><%=getexecuteConfLngStr("LtxtConfActions")%>
		<%	Case "C" %><%=getexecuteConfLngStr("LtxtBPConf")%>
		<% 	Case "I" %><%=getexecuteConfLngStr("LtxtItemConf")%>
		<%	Case "R" %><%=getexecuteConfLngStr("LtxtPayConf")%>
		<%	Case "D" %><%=getexecuteConfLngStr("LtxtDocConf")%>
			<% End Select %></td>
	</tr>
	<tr class="GeneralTbl">
		<td><b><%=getexecuteConfLngStr("DtxtExecTime")%>:</b> <%=FormatDate(Now(), True)%>&nbsp;<%=FormatTime(Now())%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTblBold2">
				<td width="13"></td>
				<td align="center">#</td>
				<td align="center"><%=getexecuteConfLngStr("LtxtReqBy")%></td>
				<% If execType <> "I" Then %>
				<td></td>
				<td align="center"><%=getexecuteConfLngStr("DtxtCode")%></td>
				<td align="center"><%=getexecuteConfLngStr("DtxtBP")%></td>
				<% If execType = "D" or execType = "R" Then %>
				<td></td>
				<td align="center"><%=getexecuteConfLngStr("DtxtBalance")%></td>
				<% End If %>
				<% End If %>
				<td align="center" style="width: 20px">&nbsp;</td>
				<td align="center"><%=getexecuteConfLngStr("DtxtDescription")%></td>
				<td align="center"><%=getexecuteConfLngStr("LtxtID")%></td>
				<td align="center"><%=getexecuteConfLngStr("DtxtDate")%></td>
				<% If execType = "A" Then %><td align="center"><%=getexecuteConfLngStr("DtxtType")%></td><% End If %>
				<% Select Case execType
				Case "C" %>
				<td align="center"><%=getexecuteConfLngStr("DtxtGroup")%></td>
				<td align="center"><%=getexecuteConfLngStr("DtxtCountry")%></td>
				<% Case "I" %>
				<td align="center"><%=txtAlterGrp%></td>
				<td align="center"><%=txtAlterFrm%></td>
				<% Case "R", "D" %>
				<td align="center"><%=getexecuteConfLngStr("DtxtTotal")%></td>
				<% End Select %>
				<% If Not rc.Eof Then
				do while not rc.eof %>
				<% If Not IsNull(rc("LinkObject")) Then %><td align="center"></td><% End If %>
				<td align="center"><%=rc("Name")%></td>
				<% rc.movenext
				loop
				rc.movefirst
				End If %>
				<td align="center"><%=getexecuteConfLngStr("DtxtState")%></td>
			</tr>
			<% If Not rs.Eof Then 
			do while not rs.eof %>
			<tr class="<% If rs("UserType") = "V" Then %>GeneralTbl<% Else %>CanastaTblExpense<% End If %>" id="act<%=rs("ID")%>" style="<% Select Case rs("Status") 
				Case "P" %>background-color: #CCFF99;<% Case "E" %>background-color: #FFD2A6;<% End Select%> ">
				<td width="13"><% If rs("HasDetails") = "Y" Then %><a href="javascript:viewFlowLog(<%=rs("ID")%>);">
				<img src="images/log_details.gif" alt="<%=getexecuteConfLngStr("LtxtViewLog")%>" border="0" style="height: 14px"></a><% End If %></td>
				<td align="right"><%=rs("ID")%></td>
				<td><% 
				Select Case rs("UserType")
					Case "V"
						Response.Write rs("SlpName")
					Case "C"
						Response.Write rs("CardName")
				End Select 
				Select Case execType
					Case "A"
						ObjectCode = rs("ObjectCode")
						Select Case rs("ObjectCode")
							Case 2, 4
								ObjEntry = rs("ActionObjCode")
							Case Else
								ObjEntry = rs("ObjectEntry")
						End Select 
					Case "C"
						ObjectCode = 2
						ObjEntry = rs("ObjectEntry")
					Case "I"
						ObjectCode = 4
						ObjEntry = rs("ObjectEntry")
					Case "R"
						ObjectCode = 24
						ObjEntry = rs("ObjectEntry")
					Case "D"
						ObjectCode = rs("Object")
						ObjEntry = rs("ObjectEntry")
				End Select %>&nbsp;</td>
				<% If execType <> "I" Then %>
				<td style="width: 20px">
				<% If execType <> "C" and rs("CardCode") <> "" and not IsNull(rs("CardCode")) Then %>
				<a href="javascript:goBP('<%=rs("CardCode")%>');">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a><% End If %></td>
				<td><%=rs("CardCode")%></td>
				<td><%=rs("CardName")%></td>
				<% If execType = "D" or execType = "R" Then %>
				<td style="width: 20px">
						<a href="javascript:goCXC('<%=JScriptHRefEncode(rs("CardCode"))%>');">
						<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td align="right"><%=rs("Currency")%>&nbsp;<%=FormatNumber(CDbl(rs("Balance")), myApp.SumDec)%></td>
				<% End If %>
				<% End If %>
				<td style="width: 20px">
						<a href="javascript:goDetail('<%=rs("ExecAt")%>', <%=ObjectCode%>, '<%=ObjEntry%>');">
						<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td><% Select Case ObjectCode
					Case 2
						ObjDesc = getexecuteConfLngStr("DtxtBP")
					Case 4
						ObjDesc = getexecuteConfLngStr("DtxtItem")
					Case 33
						ObjDesc = getexecuteConfLngStr("DtxtActivity")
					Case 23
						ObjDesc = txtQuote
					Case 13
						Select Case rs("ReserveInvoice")
							Case "Y"
								ObjDesc = txtInvRes
							Case Else
								ObjDesc = txtInv
						End Select
					Case 17
						ObjDesc = txtOrdr
					Case 24
						ObjDesc = txtRct
					Case 15
						ObjDesc = txtOdln
					Case 16
						ObjDesc = txtOrdn
					Case 14
						ObjDesc = txtOrin
					Case 22
						ObjDesc = txtOpor
					Case 20
						ObjDesc = txtOpdn
					Case 21
						ObjDesc = txtOrpd
					Case 18
						ObjDesc = txtOpch
					Case 19
						ObjDesc = txtOrpc
					Case 46
						ObjDesc = txtOvpm
					Case 48
						ObjDesc = txtInv & " / " & txtRct
					Case 191
						ObjDesc = getexecuteConfLngStr("DtxtServiceCall")
					Case 203
						ObjDesc = txtODPIReq
					Case 204
						ObjDesc = txtODPIInv
					Case Else
						ObjDesc = "N/A"
				End Select
				Response.Write ObjDesc 
				%><input type="hidden" id="ObjDesc<%=rs("ID")%>" value="<%=myHTMLEncode(ObjDesc)%>"></td>
				
				<% 
				Select Case execType
					Case "A"
						Select Case ObjectCode
							Case 2, 4
								ObjDispCode = rs("ActionObjCode")
							Case 33
								ObjDispCode = rs("ObjectEntry")
							Case Else
								ObjDispCode = rs("DocNum")
						End Select
					Case "C"
						ObjDispCode = rs("CardCode")
					Case "R", "D"
						ObjDispCode = rs("ObjectEntry")
					Case "I"
						ObjDispCode = rs("ItemCode")
				End Select %>
				<td <% If IsNumeric(ObjDispCode) Then %>align="right"<% End If %>><%=ObjDispCode%></td>
				<td align="center"><nobr><%=FormatDate(rs("RequestDate"), True)%>&nbsp;<%=FormatTime(rs("RequestDate"))%></nobr></td>
				<% If execType = "A" Then %><td><% Select Case rs("ExecAt")
					Case "O0" %><%=getexecuteConfLngStr("LtxtAprovOrder")%>
				<%	Case "O1" %><%=getexecuteConfLngStr("LtxtConvQuoteOrder")%>
				<%	Case "O7" %><%=getexecuteConfLngStr("LtxtConvOrderInv")%>
				<%	Case "O2" %><%=getexecuteConfLngStr("LtxtCloseObj")%>
				<%	Case "O3" %><%=getexecuteConfLngStr("LtxtCancelObj")%>
				<%	Case "O4" %><%=getexecuteConfLngStr("LtxtRemObj")%>
				<% End Select %></td><% End If %>
				<% Select Case execType
				Case "C" %>
				<td align="center"><%=rs("Group")%></td>
				<td align="center"><%=rs("Country")%></td>
				<% Case "I" %>
				<td align="center"><%=rs("Group")%></td>
				<td align="center"><%=rs("FirmName")%></td>
				<% Case "R", "D" %>
				<td align="right"><%=rs("DocCur")%>&nbsp;<%=FormatNumber(CDbl(rs("DocTotal")), myApp.SumDec)%></td>
				<% End Select %>
				<% If Not rc.Eof Then
				do while not rc.eof
				colID = rc("ID") %>
				<% If Not IsNull(rc("LinkObject")) Then %><td align="center" style="width: 20px"><a href="javascript:go<% Select Case rc("LinkType")
				Case "R" %>Rep<% Case "F" %>Form<% End Select %><%=rc("LinkObject")%>(<%=execConfRepVars()%>);">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td><% End If %>
				<td align="center"><%=rs("ExecRep" & colID)%></td>
				<% rc.movenext
				loop
				rc.movefirst
				End If %>
				<td style="height: 20px; text-align: center;">
				<% If rs("WaitForConf") = "Y" Then %><%=getexecuteConfLngStr("LtxtWaitForConf")%><% Else %>
				<select size="1" id="status<%=rs("ID")%>" onchange="showProcess(this.value, <%=rs("ID")%>);" <% If rs("Status") = "P" Then %>style="display: none;" <% End If %>>
				<option value="1"><%=getexecuteConfLngStr("DtxtWaiting")%></option>
				<option value="2"><%=getexecuteConfLngStr("LtxtConfirmed")%></option>
				<option value="3"><%=getexecuteConfLngStr("DtxtRejected")%></option>
				<option value="4"><%=getexecuteConfLngStr("LtxtCanceled")%></option>
				</select>
				<span id="txtStatus<%=rs("ID")%>"<% If rs("Status") <> "P" Then %> style="display: none;"<% End If %>><%=getexecuteConfLngStr("DtxtWaiting")%>...</span>
				<% End If %></td>
			</tr>
			<% If execType = "D" or execType = "R" Then
			If Not IsNull(rs("Comments")) and rs("Comments") <> "" Then %>
			<tr class="CanastaTblExpense">
				<td colspan="<%=ColSpan%>">
				<%=getexecuteConfLngStr("DtxtObservations")%>: <%=rs("Comments")%></td>
			</tr>
			<% End If
			End If %>
			<tr id="trProc<%=rs("ID")%>" style="<% If rs("Status") = "O" Then %>display: none;<% End If %><% If rs("Status") <> "O" Then %>background-color: #FFFFCC;<% End If %>">
				<td colspan="<%=ColSpan%>" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<table style="width: 100%;">
								<tr id="trNote<%=rs("ID")%>" class="GeneralTblBold2" style="<% If rs("Status") <> "O" Then %>display: none;<% End If %><% If rs("Status") = "E" Then %>background-color: #FFFFCC;<% End If %>">
												<td><textarea id="txtNote<%=rs("ID")%>" rows="10" class="input" style="width: 100%;"></textarea></td>
								</tr>
								<tr id="trSubmit<%=rs("ID")%>" class="GeneralTblBold2" style="<% If rs("Status") <> "O" Then %>display: none;<% End If %><% If rs("Status") = "E" Then %>background-color: #FFFFCC;<% End If %>">
												<td style="text-align: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>; padding-right: 2px; padding-left: 2px;">
												<input type="button" id="btnSubmit<%=rs("ID")%>" value="<%=getexecuteConfLngStr("DtxtAccept")%>" onclick="executeProcess(<%=rs("ID")%>);"></td>
								</tr>
								<tr id="trProcStatus<%=rs("ID")%>" class="trProcStatus" <% If rs("Status") = "O" Then %>style="display: none;"<% End If %>>
									<td>
									<img border="0" src="images/cargando_gif.gif" id="imgProc<%=rs("id")%>" <% If rs("Status") = "E" Then %>style="display: none;"<% End If %>><div id="txtStatusProc<%=rs("ID")%>"><% Select Case rs("Status") 
										Case "P" 
											Response.Write Replace(getexecuteConfLngStr("LtxtTransactionPool"), "{0}", rs("PoolNumber"))
										Case "E" %><b><%=getexecuteConfLngStr("DtxtError")%>: <%=rs("ErrMessage")%></b>
									<%	End Select %></div></td>
								</tr>
				</table>
				</td>
			</tr><input type="hidden" id="dbStatus<%=rs("ID")%>" value="<%=rs("Status")%>">
			<% If rs("Status") = "P" Then %><script language="javascript">setTimeout('checkProcess(<%=rs("ID")%>);', 2000);</script><% End If %>
			<% rs.movenext
			loop
			Else %>
			<tr class="GeneralTblBold2">
				<td colspan="<%=ColSpan%>">
				<p align="center"><%=getexecuteConfLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
</table>
<%
rc.Filter = "LinkActive = 'Y' and LinkObject <> null"
If Not rc.Eof Then %>
<script type="text/javascript">
<% do while not rc.eof
Select Case rc("LinkType")
Case "R"
	rsIndex = rc("LinkObject")
	strLinkVars = ""
	strLinkAtt = ""
	If rc("VerfyVars") = "Y" Then
		rcVars.Filter = "ID = " & rc("ID")
		do while not rcVars.eof
			If strLinkVars <> "" Then 
				strLinkVars = strLinkVars & ", "
			End If
			strLinkVars = strLinkVars & rcVars("VarVar")
			strLinkAtt = strLinkAtt & "var" & rcVars("varIndex") & "=' + " & rcVars("VarVar") & "+ '&"
		rcVars.movenext
		loop
	End If %>
	function goRep<%=rsIndex%>(<%=strLinkVars%>)
	{
		Start('');
		var strRepLink = 'rsIndex=<%=rsIndex%>&<%=strLinkAtt%>itemSmallRep=Y&pop=Y&AddPath=';
		doMyLink('viewReportPrint.asp', strRepLink, 'objDetails');
	}
<% Case "F" %>
function goForm<%=rc("LinkObject")%>(id)
{
	var strRepLink = 'pop=Y&AddPath=&SecID=<%=rc("LinkObject")%>&<%=Replace(rc("LinkLink"), "'", "\'")%>'.replace('{@<%=strFormFldID%>}', id);
	Start('sectionPopup.asp?' + strRepLink);
}
<% End Select
rc.movenext
loop %>
</script>
<% 
End If
set myFlowView = New FlowViewControl
myFlowView.GenerateFlowView

Function execConfRepVars
	strRepLnkVars = ""
	Select Case rc("LinkType")
		Case "R"
			If rc("VerfyVars") = "Y" Then
				rcVars.Filter = "ID = " & rc("ID")
				do while not rcVars.eof
					If strRepLnkVars <> "" Then strRepLnkVars = strRepLnkVars & ", "
					Select Case rcVars("By") 
						Case "V"
							If rcVars("varDataType") <> "datetime" Then
								strVal = rcVars("Value")
							Else
								strVal = FormatDate(rcVars("ValueDate"), False)
							End If
						Case "F"
							Select Case rcVars("Value")
								Case "@ID"
									strVal = rs(strFldID)
								Case "@LogNum"
									strVal = rs(strFldID)
							End Select
						Case "Q"
							sql = ""
							Select Case execType
								Case "A"
									sql = sql & "declare @ID int set @ID = " & rs(strFldID) & " "
								Case "C", "I", "R", "D"
									sql = sql & "declare @LogNum int set @LogNum = " & rs(strFldID) & " "
							End Select
							sql = sql & " select (" & rcVars("Value") & ") "
							set rv = conn.execute(sql)
							If Not rv.Eof Then strVal = rv(0) Else strVal = ""
					End Select
					strRepLnkVars = strRepLnkVars & "'" & Replace(strVal, "'", "\'") & "'"
				rcVars.movenext
				loop
			End If
		Case "F"
			strRepLnkVars = rs(strFldID)
	End Select
	execConfRepVars = strRepLnkVars 
End Function
%>
<!--#include file="agentBottom.asp"-->