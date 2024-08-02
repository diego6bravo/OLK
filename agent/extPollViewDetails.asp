<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not myAut.HasAuthorization(39) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/extPollViewDetails.asp" -->
<% 
set rd = server.createobject("ADODB.RecordSet")
sql = "select IsNull(T1.AlterName, T0.Name) Name, IsNull(T1.AlterDescription, T0.Description) Description, StartDate, EndDate, Filter, " & _
"IsNull(T3.AlterQuestion, T2.Question) Question, T2.Type, " & _
"(select Count('') from OLKADPollAnswers where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & ") Total, " & _
"(select Count('') from OLKADPollAnswers where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & " and Answer = '" & Request("col") & "') TotalAct " & _
" from OLKADPoll T0 " & _
"left outer join OLKADPollAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID " & _
"inner join OLKADPollLines T2 on T2.AdPollID = T0.AdPollID and T2.LineID = " & Request("LineID") & " " & _
"left outer join OLKADPollLinesAlterNames T3 on T3.LanID = " & Session("LanID") & " and T3.AdPollID = T0.AdPollID and T3.LineID = T2.LineID " & _
"where T0.AdPollID = " & request("AdPollID")
set rs = conn.execute(sql)
qcName = rs("Name")
qcDesc = rs("Description")
qcDate = FormatDate(rs("StartDate"), True)
qcEndDate = FormatDate(rs("EndDate"), True)
qcQuestion = rs("Question")
qcType = rs("Type")
varTotal = rs("Total")
varATotal = rs("TotalAct")

If rs("Filter") <> "" Then
	qcFilter = " and " & rs("Filter")
Else
	qcFilter = ""
End If
sql = "select Count('') from OCRD where CardType in ('C', 'L') " & qcFilter & " and not exists(select 'a' from OLKADPollAnswers where AdPollID = " & Request("AdPollID") & " and CardCode = OCRD.CardCode)"
set rs = conn.execute(sql)
pendientes = rs(0)

sql = "select T0.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName, (select name from OCPR where CntctCode = T1.CntctPrsn) CntctPrsn, " & _
	"AnswerDate, T1.Notes " & _
	"from OCRD T0 " & _
	"inner join OLKADPollAnswers T1 on T1.CardCode = T0.CardCode and AdPollID = " & Request("AdPollID") & " and LineID = " & _
	Request("LineID") & " and Answer = '" & Request("col") & "' "
rs.close
rs.open sql, conn, 3, 1
%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><a class="GeneralTlt" href="extpollman.asp"><%=getextPollViewDetailsLngStr("LtxtExtPolls")%></a> - <a class="GeneralTlt" href="javascript:doMyLink('extPollview.asp', 'AdPollID=<%=Request("AdPollID")%>', '_self');"><%=getextPollViewDetailsLngStr("LtxtViewRes")%></a> - <%=getextPollViewDetailsLngStr("LtxtDetails")%></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewDetailsLngStr("DtxtName")%></td>
		<td colspan="2" class="GeneralTbl"><%=qcName%>&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewDetailsLngStr("LtxtStartDate")%></td>
		<td width="250" class="GeneralTbl"><%=qcDate%>&nbsp;</td>
		<td class="GeneralTblBold2"><%=getextPollViewDetailsLngStr("DtxtDescription")%></td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewDetailsLngStr("LtxtEndDate")%></td>
		<td width="250" class="GeneralTbl"><%=qcEndDate%>&nbsp;</td>
		<td rowspan="4" class="GeneralTbl" valign="top"><% If Not IsNull(qcDesc) Then %><%=Replace(qcDesc,VbNewLine,"<br>")%><% End If %>&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewDetailsLngStr("LtxtPending")%></td>
		<td width="250" class="GeneralTbl"><%=pendientes%>&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollViewDetailsLngStr("DtxtDate")%></td>
		<td width="250" class="GeneralTbl"><%=FormatDate(Now(), false)%> - <%=FormatTime(Now())%>&nbsp;</td>
	</tr>
</table><br>
<table border="0" cellpadding="0" width="100%" >
	<tr>
		<td class="GeneralTblBold2"><%=Replace(getextPollViewDetailsLngStr("LttlQuestionDetails"), "{0}", Request("lineIndex")+1)%>: <%=qcQuestion%><br>
		<% myText = getextPollViewDetailsLngStr("LtxtPplAnswered")
		myText = Replace(myText, "{0}",varATotal)
		Select Case qcType
			Case "B"
				Select Case Request("col")
					Case "Y"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("DtxtYes"))
					Case "N"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("DtxtNo"))
				End Select
			Case "R"
				Select Case Request("col")
					Case "1"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("LtxtTerrible"))
					Case "2"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("LtxtBad"))
					Case "3"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("LtxtRegular"))
					Case "4"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("LtxtGood"))
					Case "5"
						myText = Replace(myText, "{1}",getextPollViewDetailsLngStr("LtxtExcelent"))
				End Select
			Case "M"
				set rd = Server.CreateObject("ADODB.RecordSet")
				sql = "select IsNull(T1.AlterChoice, T0.Choice) " & _
					"from OLKADPollLinesChoices T0 " & _
					"left outer join OLKADPollLinesChoicesAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID and T1.LineID = T0.LineID and T1.ChoiceID = T0.ChoiceID " & _
					"where T0.AdPollID = " & Request("AdPollID") & " and T0.LineID = " & Request("LineID") & " and T0.ChoiceID = " & Request("col")
				set rd = conn.execute(sql)
				myText = Replace(myText, "{1}",rd(0))
		End Select
		myText = Replace(myText, "{2}", varTotal) %> 
		<%=myText%>
		</td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTblBold2">
		<td width="20">&nbsp;</td>
		<td><%=getextPollViewDetailsLngStr("DtxtNote")%></td>
		<td><%=getextPollViewDetailsLngStr("DtxtName")%></td>
		<td><%=getextPollViewDetailsLngStr("DtxtContact")%></td>
		<td><%=getextPollViewDetailsLngStr("DtxtDate")%></td>
		<td><%=getextPollViewDetailsLngStr("DtxtNote")%></td>
	</tr>
	<% do while not rs.eof %>
	<tr class="GeneralTbl">
		<td width="20">
		<a href="javascript:doMyLink('extPollExec.asp', 'AdPollID=<%=Request("AdPollID")%>&CardCode=<%=rs("CardCode")%>&LineID=<%=Request("LineID")%>&col=<%=Request("col")%>', '_self');">
		<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
		<td><%=rs("CardCode")%>&nbsp;</td>
		<td><%=rs("CardName")%>&nbsp;</td>
		<td><%=rs("CntctPrsn")%>&nbsp;</td>
		<td><%=FormatDate(rs("AnswerDate"), True)%>&nbsp;</td>
		<td><%=rs("Notes")%>&nbsp;</td>
	</tr>
	<% rs.movenext
	loop %>
</table>
<!--#include file="agentBottom.asp"-->