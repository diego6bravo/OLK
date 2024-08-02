<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myAut.HasAuthorization(39) Then Response.Redirect "unauthorized.asp" %>
<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<% addLngPathStr = "" %>
<!--#include file="lang/extPollList.asp" -->
<% 
sql = "select T0.AdPollID, T0.Filter " & _
"from OLKADPoll T0 " & _
"inner join OLKADPollAgents T1 on T1.AdPollID = T0.AdPollID and (T1.SlpCode = " & Session("vendid") & " or T1.SlpCode = -2) " & _
"where DateDiff(day,StartDate,getdate()) >= 0 and DateDiff(day,getdate(),EndDate) >= 0 and Status = 'A'"
set rs = conn.execute(sql) 

If Not rs.Eof Then
	sql = "SELECT T0.AdPollID, IsNull(T1.AlterName, T0.Name) Name, StartDate, EndDate, Case T0.ADPollID "
	
	do while not rs.eof
		If rs("Filter") <> "" Then qcFilter = " and " & rs("Filter") else qcFilter = ""
		sql = sql & "When " & rs("ADPollID") & " Then (select count('') from OCRD where CardType in ('C', 'L') " & qcFilter & ") "
	rs.movenext
	loop
	
	sql = sql & "End As 'TotalLL', Case T0.ADPollID "
	
	rs.movefirst
	do while not rs.eof
		If rs("Filter") <> "" Then qcFilter = " and " & rs("Filter") else qcFilter = ""
		sql = sql & "When " & rs("ADPollID") & " Then (select count('') from OCRD where CardType in ('C', 'L') " & qcFilter & " and not exists(select 'S' from OLKADPollAnswers where ADPollID = T0.ADPollID and CardCode = OCRD.CardCode)) "	
	rs.movenext
	loop
	sql = sql & " End As 'TotalP' from OLKADPoll T0 " & _
	"left outer join OLKADPollAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID =  T0.AdPollID " & _
	"where DateDiff(day,StartDate,getdate()) >= 0 and DateDiff(day,getdate(),EndDate) >= 0 and Status = 'A'"
	set rs = conn.execute(sql)
End If
%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><%=getextPollListLngStr("LtxtExtPolls")%></td>
	</tr>
</table>
<table border="0" cellpadding="0"  width="100%">
	<tr class="GeneralTblBold2">
		<td width="16">&nbsp;</td>
		<td class="style1"><%=getextPollListLngStr("LtxtPoll")%></td>
		<td class="style1"><%=getextPollListLngStr("LtxtStartDate")%></td>
		<td class="style1"><%=getextPollListLngStr("LtxtEndDate")%></td>
		<td class="style1"><%=getextPollListLngStr("LtxtPending")%></td>
		<td class="style1"><%=getextPollListLngStr("DtxtTotal")%></td>
	</tr>
	<% If Not rs.Eof Then %>
	<% do while not rs.eof %>
	<tr class="GeneralTbl">
		<td width="16">
		<a href="javascript:doMyLink('extPollOpen.asp', 'ADPollID=<%=rs("AdPollID")%>', '_self');"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
		<td><%=rs("Name")%></td>
		<td><%=FormatDate(rs("StartDate"), True)%></td>
		<td><%=FormatDate(rs("EndDate"), True)%></td>
		<td class="style1"><%=rs("TotalP")%></td>
		<td class="style1"><%=rs("TotalLL")%></td>
	</tr>
	<% rs.movenext
	loop %>
	<% Else %>
	<tr class="GeneralTbl">
		<td colspan="6" align="center"><%=getextPollListLngStr("DtxtNoData")%></td>
	</tr>
	<% End If %>
</table>
<!--#include file="agentBottom.asp"-->