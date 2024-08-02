<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>

<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<%
If Not myAut.HasAuthorization(39) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/extPollMan.asp" -->
<!--#include file="genman/adminTradForm.asp"-->
<% 

sql = "select AdPollID, Filter from OLKADPoll where Status <> 'D'"
rs.close
rs.open sql, conn, 3, 1

If Not rs.Eof Then

	sql = "SELECT AdPollID, Name, StartDate, EndDate, Case AdPollID "
	
	do while not rs.eof
		If rs("Filter") <> "" Then qcFilter = " and " & rs("Filter") Else qcFilter = ""
		sql = sql & "When " & rs("AdPollID") & " Then (select count('') from OCRD where CardType in ('C', 'L') and DateDiff(day,CreateDate,T0.EndDate) >= 0 " & qcFilter & ") "
	rs.movenext
	loop
	
	sql = sql & "End As 'TotalCard', Case AdPollID "
	
	if rs.recordcount > 0 then rs.movefirst
	do while not rs.eof
		If rs("Filter") <> "" Then qcFilter = " and " & rs("Filter") Else qcFilter = ""
		sql = sql & "When " & rs("AdPollID") & " Then (select count('') from OCRD X0 where CardType in ('C', 'L') and DateDiff(day,CreateDate,T0.EndDate) >= 0 " & qcFilter & " and exists(select 'S' from OLKADPollAnswers where AdPollID = T0.AdPollID and CardCode = X0.CardCode)) "	
	rs.movenext
	loop
	
	sql = sql & " End As 'TotalVote', Status from OLKADPoll T0 where Status <> 'R'"
	
	If rs.recordcount > 0 then
		rs.close
		rs.open sql, conn, 3, 1
	End If
End If
%>
<script type="text/javascript">
function viewPoll(AdPollID)
{
	doMyLink('extPollView.asp', 'AdPollID=' + AdPollID, '_self');
}
function valFrm()
{
	var found = false;
	if (document.frmPolls.delID)
	{
		if (document.frmPolls.delID.length)
		{
			for (var i = 0;i<document.frmPolls.delID.length;i++)
			{
				if (document.frmPolls.delID[i].checked)
				{
					found = true;
					break;
				}
			}
		}
		else
		{
			found = document.frmPolls.delID.checked;
		}
	}
	
	if (!found)
	{
		alert('<%=getextPollManLngStr("LtxtValSelPoll")%>');
		return false;
	}
	else
	{
		return confirm('<%=getextPollManLngStr("LtxtConfDelPoll")%>');
	}
}
</script>
<form method="post" name="frmPolls" action="genman/extPollSubmit.asp" onsubmit="return valFrm();">
<input type="hidden" name="cmd" value="del">
<table border="0" cellpadding="0" cellspacing="2" width="100%">
	<tr class="GeneralTlt">
		<td><%=getextPollManLngStr("LtxtExtPolls")%></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr class="FirmTlt3">
		<td style="width: 16px">&nbsp;</td>
		<td style="width: 16px">&nbsp;</td>
		<td align="center"><%=getextPollManLngStr("LtxtPoll")%></td>
		<td align="center"><%=getextPollManLngStr("LtxtStartDate")%></td>
		<td align="center"><%=getextPollManLngStr("LtxtEndDate")%></td>
		<td align="center"><%=getextPollManLngStr("LtxtAdvance")%></td>
		<td align="center"><%=getextPollManLngStr("DtxtActive")%></td>
	</tr>
	<% If Not rs.Eof Then %>
	<% do while not rs.eof %>
	<tr class="GeneralTbl">
		<td style="width: 16px" class="style1">
		<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("AdPollID")%>);">
		<input type="checkbox" name="delID" id="delID<%=rs("AdPollID")%>" value="<%=rs("AdPollID")%>" style="display: none;"></td>
		<td style="width: 16px"><a href='extPollEdit.asp?AdPollID=<%=rs("AdPollID")%>'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTbl">
				<td>
				<%=rs("Name")%>
				</td>
				<td width="16"><a href="javascript:doFldTrad('ADPoll', 'AdPollID', '<%=rs("AdPollID")%>', 'AlterName', 'T', null);"><img src="images/trad.gif" alt="<%=getextPollManLngStr("DtxtTranslate")%>" border="0"></a></td>
			</tr>
		</table>
		</td>
		<td class="style1"><%=FormatDate(rs("StartDate"), True)%></td>
		<td class="style1"><%=FormatDate(rs("EndDate"), True)%></td>
		<td class="style1">
		<table cellpadding="0" cellspacing="2" border="0">
			<tr class="GeneralTbl">
				<td><img border="0" src="images/eye_icon.gif" dir="ltr" width="15" height="13" alt="<%=getextPollManLngStr("LtxtViewResults")%>" style="cursor: hand" onclick="javascript:viewPoll(<%=rs("AdPollID")%>);"></td>
				<td><%=rs("TotalVote")%>/<%=rs("TotalCard")%> (<% If rs("TotalCard") > 0 Then %><%=FormatNumber(rs("TotalVote")*100/rs("TotalCard"),2)%><% Else %><%=FormatNumber(0, 2)%><% End If %>%)</td>
			</tr>
		</table>
		</td>
		<td align="center"><% Select Case rs("Status")
			Case "A" %><%=getextPollManLngStr("DtxtYes")%>
			<% Case "N" %><%=getextPollManLngStr("DtxtNo")%>
			<% End Select %></td>
	</tr>
	<% rs.movenext
	loop %>
	<% Else %>
	<tr>
		<td colspan="7" align="center">
		<%=getextPollManLngStr("DtxtNoData")%>
		</td>
	</tr>
	<% End If %>
	<tr>
		<td colspan="7">
		<table cellpadding="0" border="0" width="100%">
			<tr>
				<td width="1"><input type="submit" value="<%=getextPollManLngStr("DtxtDelete")%>" name="btnDel"></td>
				<td>&nbsp;</td>
				<td width="1"><input type="button" value="<%=getextPollManLngStr("LtxtNewExtPoll")%>" name="btnNew" onclick="javascript:window.location.href='extPollEdit.asp';"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<script type="text/javascript">
<!--
function doCheckDel(Img, LogNum)
{
	if (!document.getElementById('delID' + LogNum).checked)
	{
		document.getElementById('delID' + LogNum).checked = true;
		Img.src = 'images/checkbox_on.jpg';
	}
	else
	{
		document.getElementById('delID' + LogNum).checked = false;
		Img.src = 'images/checkbox_off.jpg';
	}
}
//-->
</script>
<!--#include file="agentBottom.asp"-->