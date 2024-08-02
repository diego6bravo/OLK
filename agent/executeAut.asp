<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (Session("useraccess") = "P" or Session("HasActionConfAut")) Then Response.Redirect "unauthorized.asp" %>
<!--#include file="lang/executeAut.asp" -->
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
var SumDec = <%=myApp.SumDec%>;
var txtConfAut = '<%=getexecuteAutLngStr("LtxtConfAut")%>';
var txtConfReject = '<%=getexecuteAutLngStr("LtxtConfReject")%>';
var txtOtherProc = '<%=getexecuteAutLngStr("LtxtOtherProc")%>';
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
cmd.CommandText = "DBOLKGetExecAut" & execType & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid") 

Select Case execType
	Case "A"
		ColSpan = 12 + rcColSpan + 1
		
		cmd("@UserAccess") = Session("UserAccess")
	Case "C"
		ColSpan = 13 + rcColSpan + 1
		
		cmd("@UserAccess") = Session("UserAccess")
	Case "D"
		ColSpan = 14 + rcColSpan + 1
		
		cmd("@UserAccess") = Session("UserAccess") 
	Case "I"
		ColSpan = 11 + rcColSpan + 1
	Case "R"
		ColSpan = 14 + rcColSpan + 1
End Select
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1
%>
<table border="0" cellpadding="0" width="100%">
	<form name="frmAut">
	<tr class="GeneralTlt">
		<td>&nbsp;<% Select Case execType
			Case "A" %><%=getexecuteAutLngStr("LtxtActionsAut")%>
		<%	Case "C" %><%=getexecuteAutLngStr("LtxtBPAut")%>
		<% 	Case "I" %><%=getexecuteAutLngStr("LtxtItemAut")%>
		<%	Case "R" %><%=getexecuteAutLngStr("LtxtPayAut")%>
		<%	Case "D" %><%=getexecuteAutLngStr("LtxtDocAut")%>
			<% End Select %></td>
	</tr>
	<tr class="GeneralTbl">
		<td><b><%=getexecuteAutLngStr("DtxtExecTime")%>:</b> <%=FormatDate(Now(), True)%>&nbsp;<%=FormatTime(Now())%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTblBold2">
				<td width="13"></td>
				<td align="center">#</td>
				<td align="center"><%=getexecuteAutLngStr("LtxtReqBy")%></td>
				<% If execType <> "I" Then %>
				<td></td>
				<td align="center"><%=getexecuteAutLngStr("DtxtCode")%></td>
				<td align="center"><%=getexecuteAutLngStr("DtxtBP")%></td>
				<% If execType = "D" or execType = "R" Then %>
				<td></td>
				<td align="center"><%=getexecuteAutLngStr("DtxtBalance")%></td>
				<% End If %>
				<% End If %>
				<td align="center" style="width: 20px">&nbsp;</td>
				<td align="center"><%=getexecuteAutLngStr("DtxtDescription")%></td>
				<td align="center"><%=getexecuteAutLngStr("LtxtID")%></td>
				<td align="center"><%=getexecuteAutLngStr("DtxtDate")%></td>
				<% If execType = "A" Then %><td align="center"><%=getexecuteAutLngStr("DtxtType")%></td><% End If %>
				<% Select Case execType
				Case "C" %>
				<td align="center">|D:txtGroup|</td>
				<td align="center">|D:txtCountry|</td>
				<% Case "I" %>
				<td align="center"><%=txtAlterGrp%></td>
				<td align="center"><%=txtAlterFrm%></td>
				<% Case "R", "D" %>
				<td align="center"><%=getexecuteAutLngStr("DtxtTotal")%></td>
				<% End Select %>
				<% If Not rc.Eof Then
				do while not rc.eof %>
				<% If Not IsNull(rc("LinkObject")) Then %><td align="center"></td><% End If %>
				<td align="center"><%=rc("Name")%></td>
				<% rc.movenext
				loop
				rc.movefirst
				End If %>
				<td align="center"><%=getexecuteAutLngStr("DtxtFlow")%></td>
				<td align="center"><%=getexecuteAutLngStr("DtxtState")%></td>
			</tr>
			<% If Not rs.Eof Then
			LastID = -1
			LastFlowID = -1
			LastLineID = -1 
			do while not rs.eof
			AutID = rs("ID") & "_" & rs("FlowID") & "_" & rs("LineID") %>
			<tr class="<% If rs("UserType") = "V" Then %>GeneralTbl<% Else %>CanastaTblExpense<% End If %>" id="act<%=rs("ID")%>" style="<% Select Case rs("Status") 
				Case "P" %>background-color: #CCFF99;<% Case "E" %>background-color: #FFD2A6;<% End Select%> ">
				<% If LastID <> rs("ID") Then %>
				<td width="13"><a href="javascript:viewFlowLog(<%=rs("ID")%>);">
				<img src="images/log_details.gif" alt="|L:txtViewLog|" border="0" style="height: 14px"></a></td>
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
						ObjDesc = getexecuteAutLngStr("DtxtBP")
					Case 4
						ObjDesc = getexecuteAutLngStr("DtxtItem")
					Case 33
						ObjDesc = getexecuteAutLngStr("DtxtActivity")
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
						ObjDesc = getexecuteAutLngStr("DtxtServiceCall")
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
					Case "O0" %><%=getexecuteAutLngStr("LtxtAprovOrder")%>
				<%	Case "O1" %><%=getexecuteAutLngStr("LtxtConvQuoteOrder")%>
				<%	Case "O7" %><%=getexecuteAutLngStr("LtxtConvOrderInv")%>
				<%	Case "O2" %><%=getexecuteAutLngStr("LtxtCloseObj")%>
				<%	Case "O3" %><%=getexecuteAutLngStr("LtxtCancelObj")%>
				<%	Case "O4" %><%=getexecuteAutLngStr("LtxtRemObj")%>
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
				<% Else %>
				<td colspan="<%=ColSpan-2%>"></td>
				<% End If %>
				<td><%=rs("FlowName")%></td>
				<td style="height: 20px; text-align: center;">
				<input type="hidden" name="AutID" value="<%=rs("ID")%>">
				<input type="hidden" name="FlowID" value="<%=rs("FlowID")%>">
				<input type="hidden" name="LineID" value="<%=rs("LineID")%>">
				<input type="button" name="btnApprove" id="btnApprove<%=AutID%>" value="<%=getexecuteAutLngStr("DtxtAuthorize")%>" onclick="executeAut(<%=rs("ID")%>,<%=rs("FlowID")%>,<%=rs("LineID")%>, 'A');">
				<input type="button" name="btnReject" id="btnReject<%=AutID%>" value="<%=getexecuteAutLngStr("DtxtReject")%>" onclick="executeAut(<%=rs("ID")%>,<%=rs("FlowID")%>,<%=rs("LineID")%>, 'R');">
				</td>
			</tr>
			<% If execType = "D" or execType = "R" Then
			If Not IsNull(rs("Comments")) and rs("Comments") <> "" Then %>
			<tr class="CanastaTblExpense">
				<td colspan="<%=ColSpan%>">
				|D:txtObservations|: <%=rs("Comments")%></td>
			</tr>
			<% End If
			End If %>
			<% LastID = rs("ID")
			LastFlowID = rs("FlowID")
			LastLineID = rs("LineID")
			rs.movenext
			loop
			Else %>
			<tr class="GeneralTblBold2">
				<td colspan="<%=ColSpan%>">
				<p align="center"><%=getexecuteAutLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	</form>
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