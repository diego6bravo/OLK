<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not myAut.HasAuthorization(39) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/extPollView.asp" -->
<head>
<style>
a 
{
	font-size: x-small;
}
</style>
</head>
<% 
set rd = server.createobject("ADODB.RecordSet")
sql = "select IsNull(T1.AlterName, T0.Name) Name, IsNull(T1.AlterDescription, T0.Description) Description, StartDate, EndDate, Filter " & _
"from OLKADPoll T0 " & _
"left outer join OLKADPollAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID " & _
"where T0.AdPollID = " & request("AdPollID")
set rs = conn.execute(sql)
qcName = rs("Name")
qcDesc = rs("Description")
qcDate = FormatDate(rs("StartDate"), True)
qcEndDate = FormatDate(rs("EndDate"), True)

If rs("Filter") <> "" Then
	qcFilter = " and " & rs("Filter")
Else
	qcFilter = ""
End If

sql = "select Count('') TotalCard, Sum(Answer) TotalVote from " & _
"( " & _
"	select Case When exists(select '' from OLKADPollAnswers where AdPollID = " & Request("AdPollID") & " and CardCode = OCRD.CardCode) Then 1 Else 0 End Answer " & _
"	from OCRD  " & _
"	where CardType in ('C', 'L') " & qcFilter & " " & _
") X0 "
set rs = conn.execute(sql)
TotalCardCode = rs("TotalCard")
TotalVote = rs("TotalVote")
TotalPend = TotalCardCode - TotalVote
%>
<table border="0" cellpadding="0" cellspacing="2" width="100%">
	<tr class="GeneralTlt">
		<td><a class="GeneralTlt" href="extpollman.asp"><%=getextPollViewLngStr("LtxtExtPolls")%></a> - <%=getextPollViewLngStr("LtxtViewRes")%></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewLngStr("DtxtName")%></td>
		<td colspan="2" class="GeneralTbl"><%=qcName%>&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewLngStr("LtxtStartDate")%></td>
		<td width="250" class="GeneralTbl"><%=qcDate%>&nbsp;</td>
		<td class="GeneralTblBold2" valign="top"><%=getextPollViewLngStr("DtxtDescription")%></td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewLngStr("LtxtEndDate")%></td>
		<td width="250" class="GeneralTbl"><%=qcEndDate%>&nbsp;</td>
		<td rowspan="4" class="GeneralTbl" valign="top"><% If Not IsNull(qcDesc) Then %><%=Replace(qcDesc,VbNewLine,"<br>")%><% End If %>&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewLngStr("LtxtAdvance")%></td>
		<td width="250" class="GeneralTbl"><%=TotalVote%>/<%=TotalCardCode%> (<%=FormatNumber(TotalVote*100/TotalCardCode, 2)%>%)&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewLngStr("DtxtDate")%></td>
		<td width="250" class="GeneralTbl"><%=FormatDate(Now(), false)%> - <%=FormatTime(Now())%>&nbsp;</td>
	</tr>
</table>
<% sql = "select T0.LineID, IsNull(T1.AlterQuestion, T0.Question) Question, T0.Type, Count('') TotalVotes " & _
"from OLKADPollLines T0  " & _
"left outer join OLKADPollLinesAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID and T1.LineID = T0.LineID  " & _
"left outer join OLKADPollAnswers T2 on T2.AdPollID = T0.AdPollID and T2.LineID = T0.LineID " & _
"where T0.AdPollID = " & Request("AdPollID") & " " & _
"Group By T0.LineID, T1.AlterQuestion, T0.Question, T0.Type "
rs.close
rs.open sql, conn, 3, 1 

set rd = Server.CreateObject("ADODB.RecordSet")
sql = 	"select X0.LineID, X0.Ordr, Convert(nvarchar(50),X0.ChoiceID) ChoiceID, IsNull(X2.AlterChoice, X0.Choice) Choice, X0.Color, " & _
		"(select Count('') from OLKADPollAnswers where AdPollID = X0.AdPollID and LineID = X0.LineID and Answer = X0.ChoiceID) Votes " & _
		"from OLKADPollLinesChoices X0 " & _
		"left outer join OLKADPollLinesChoicesAlterNames X2 on X2.LanID = " & Session("LanID") & " and X2.AdPollID = X0.AdPollID and X2.LineID = X0.LineID and X2.ChoiceID = X0.ChoiceID " & _
		"where X0.AdPollID = " & Request("AdPollID") & " " & _
		"union " & _
		"select X0.LineID, X1.Ordr, X1.ChoiceID, X1.Choice, X1.Color, " & _
		"(select Count('') from OLKADPollAnswers where AdPollID = X0.AdPollID and LineID = X0.LineID and Answer = X1.ChoiceID) Votes " & _
		"from OLKADPollLines X0 " & _
		"cross join ( " & _
		"	select 'Y' ChoiceID, N'" & getextPollViewLngStr("DtxtNo") & "' Choice, '#FF9797' Color, 2 Ordr  " & _
		"	union  " & _
		"	select 'N' ChoiceID, N'" & getextPollViewLngStr("DtxtYes") & "' choice, '#00CC00' Color, 1 Ordr) X1 " & _
		"where X0.AdPollID = " & Request("AdPollID") & " and X0.Type = 'B' " & _
		"union " & _
		"select X0.LineID, X1.Ordr, X1.ChoiceID, X1.Choice, X1.Color, " & _
		"(select Count('') from OLKADPollAnswers where AdPollID = X0.AdPollID and LineID = X0.LineID and Answer = X1.ChoiceID) Votes " & _
		"from OLKADPollLines X0 " & _
		"cross join ( " & _
		"	select '1' ChoiceID, N'" & getextPollViewLngStr("LtxtTerrible") & "' Choice, '#FF9797' Color, 1 Ordr  " & _
		"	union  " & _
		"	select '2' ChoiceID, N'" & getextPollViewLngStr("LtxtBad") & "' choice, '#FFC082' Color, 2 Ordr " & _
		"	union  " & _
		"	select '3' ChoiceID, N'" & getextPollViewLngStr("LtxtRegular") & "' choice, '#FFFF99' Color, 3 Ordr " & _
		"	union  " & _
		"	select '4' ChoiceID, N'" & getextPollViewLngStr("LtxtGood") & "' choice, '#33CCFF' Color, 4 Ordr " & _
		"	union  " & _
		"	select '5' ChoiceID, N'" & getextPollViewLngStr("LtxtExcelent") & "' choice, '#00CC00' Color, 5 Ordr) X1 " & _
		"where X0.AdPollID = " & Request("AdPollID") & " and X0.Type = 'R' " & _
		"order by 1, 2 "
rd.open sql, conn, 3, 1
varx = 0 %>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<% 
		do while not rs.eof 
		rd.Filter = "LineID = " & rs("LineID")
		varx = varx + 1 
		varMax = 0
		totalVotes = rs("TotalVotes")
		%>
		<td width="50%" valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTblBold2">
					<td align="center" height="30">
					<%=rs("Question")%></td>
				</tr>
				<% 
				do while not rd.eof
					If rd("Votes") > varMax Then varMax = rd("Votes")
				rd.movenext
				loop
				If rs("Type") <> "M" Then
				%>
				<tr>
					<td style="border-left-style:solid; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-style:none; border-top-width:medium; border-bottom-style:solid; border-bottom-width:1px">
					<div align="center">
					<table border="0" cellpadding="0" cellspacing="0" id="table4" height="190">
						<tr>
							<%
							rd.movefirst
							do while not rd.eof
								If varMax > 0 Then
									vars = rd("Votes")*85/varMax
									varP = rd("Votes")*100/totalVotes
								Else
									vars = 0
								End If %>
							<td width="60" align="center" valign="bottom">
							<table width="90%" height="100%">
								<tr>
									<td valign="bottom"><font size="1">
									<p align="center"><%=FormatNumber(varP,2)%>%</font>
									</td>
								</tr>
								<% If vars > 0 Then %>
								<tr>
									<td bgcolor="<%=rd("Color")%>" height="<%=vars%>%"></td>
								</tr>
								<% End If %>
								</table>

							</td>
							<%
							rd.movenext
							loop %>
						</tr>
						<tr>
							<% rd.movefirst
							do while not rd.eof %>
							<td width="60" valign="bottom" align="center" height="10" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top: 1px solid #C0C0C0; border-bottom-width: 1px" class="GeneralTbl">
							<a href="javascript:doMyLink('extPollViewDetails.asp', 'AdPollID=<%=Request("AdPollID")%>&amp;LineID=<%=rs("LineID")%>&amp;col=<%=rd("ChoiceID")%>', '_self');"><%=rd("Votes")%><br><%=rd("Choice")%></a></td>
							<% rd.movenext
							loop %>
						</tr>
					</table>
					</div>
					</td>
				</tr>
				<% Else %>
				<tr>
					<td style="border-left-style:solid; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-style:none; border-top-width:medium; border-bottom-style:solid; border-bottom-width:1px">
					<div align="center"><br>
					<table border="0" cellpadding="0" cellspacing="0" width="300">
						<% rd.movefirst
						do while not rd.eof
						If varMax > 0 Then
							vars = rd("Votes")*85/varMax
							varP = rd("Votes")*100/totalVotes
						Else
							vars = 0
						End If %>
						<tr>
							<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top: 1px solid #C0C0C0; border-bottom-width: 1px" class="GeneralTbl">
							<a href="javascript:doMyLink('extPollViewDetails.asp', 'AdPollID=<%=Request("AdPollID")%>&amp;LineID=<%=rs("LineID")%>&amp;col=<%=rd("ChoiceID")%>', '_self');"><%=rd("Votes")%> - <%=rd("Choice")%></a></td>
						</tr>
						<tr>
							<td>
							<table width="100%">
								<tr>
									<% If vars > 0 Then %><td bgcolor="<%=rd("Color")%>" width="<%=vars%>%"></td><% End If %>
									<td><font size="1" face="Verdana"><%=FormatNumber(varP, 2)%>%</font></td>
								</tr>
							</table>
							</td>
						</tr>
						<% rd.movenext 
						loop %>
					</table>
					<br>
					</div>
					</td>
				</tr>
				<% End If %>
			</table>
		</td>
	<% 
	If varx = 2 then
		varx = 0
		response.write "</tr><tr>"
	end if
	rs.movenext
	loop %>
	</tr>
</table>

<% If Request("Refresh") <> "" Then %>
<% If Request("Refresh") > 0 Then %>
<script language="javascript">
setTimeout("reloadRep()", <%=Request("Refresh")*60000%>);
function reloadRep()
{
	document.frmReload.btnReload.click();
}
</script>
<% End If %><% End If %>
<!--#include file="agentBottom.asp"-->