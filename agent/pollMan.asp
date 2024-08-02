<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myAut.HasAuthorization(38) Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/pollMan.asp" -->
<!--#include file="genman/adminTradForm.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<% 
conn.execute("use [" & Session("OLKDB") & "]")
           set rs = Server.CreateObject("ADODB.recordset")
           sql = "select pollIndex, pollName, pollStatus, pollDate, " & _
            	 "(select count('A') from olkpolldb where pollIndex = olkpoll.pollIndex) Votes, " & _
            	 "Case pollIndex When (select max(pollIndex) from olkPoll where pollStatus = 'O' " & _
            	 "and pollDate < getdate()) then 'Y' End Verfy from olkpoll where pollStatus <> 'D' " & _
            	 "order by Convert(int,pollDate) desc"
           set rs = conn.execute(sql)
           %>
<script language="javascript">
function viewPoll(pollIndex) {
	page = 'genman/pollView.asp?pollIndex='+pollIndex+'&pop=Y';
	OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=250,height=250");
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
		alert('<%=getpollManLngStr("LtxtValSelPoll")%>');
		return false;
	}
	else
	{
		return confirm('<%=getpollManLngStr("LtxtConfDelPoll")%>');
	}
}
</script>
<form method="post" name="frmPolls" action="genman/pollSubmit.asp" onsubmit="return valFrm();">
<input type="hidden" name="pollCmd" value="delPoll">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td class="GeneralTlt">
			<%=getpollManLngStr("LttlPolls")%>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="FirmTlt3">
				<td align="center" style="width: 16px">
				&nbsp;</td>
				<td align="center" style="width: 16px">
				&nbsp;</td>
				<td style="width: 100px" align="center">
				<nobr><%=getpollManLngStr("DtxtDateOfPub")%>&nbsp;</nobr></td>
				<td align="center"><%=getpollManLngStr("DtxtName")%></td>
				<td style="width: 100px" align="center"><%=getpollManLngStr("LtxtVotes")%></td>
				<td style="width: 100px" align="center"><%=getpollManLngStr("DtxtActive")%></td>
			</tr>
			<% do while not rs.eof %>
			<tr class="GeneralTbl">
				<td style="width: 16px" class="style1">
				<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("pollIndex")%>);">
				<input type="checkbox" name="delID" id="delID<%=rs("pollIndex")%>" value="<%=rs("pollIndex")%>" style="display: none;"></td>
				<td style="width: 16px">
				<a href="javascript:doMyLink('pollEdit.asp', 'pollIndex=<%=rs("pollIndex")%>', '_self');">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td style="width: 100px"><%=FormatDate(rs("pollDate"), True)%></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td valign="middle" class="GeneralTbl"><%=rs("pollName")%>&nbsp;<% If rs("Verfy") = "Y" Then %>(<%=getpollManLngStr("LtxtOnline")%>)<%end if%>
						</td>
						<td width="16">
						<a href="javascript:doFldTrad('Poll', 'pollIndex', <%=rs("pollIndex")%>, 'alterPollTitle', 'T', null);"><img src="images/trad.gif" alt="<%=getpollManLngStr("DtxtTranslate")%>" border="0">
						</a></td>
					</tr>
				</table>
				</td>
				<td align="center" style="width: 100px">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><img border="0" src="images/eye_icon.gif" dir="ltr" width="15" height="13" alt="<%=getpollManLngStr("LtxtViewResults")%>" style="cursor: hand" onclick="javascript:viewPoll(<%=rs("pollIndex")%>);"></td>
						<td class="GeneralTbl">&nbsp;<%=rs("votes")%>&nbsp;</td>
					</tr>
				</table>
				</td>
				<td style="width: 100px; text-align: center;"><% 
				Select Case rs("pollStatus")
				Case "O"
					Response.Write getpollManLngStr("DtxtYes") '"Activado"
				Case "C"
					Response.Write getpollManLngStr("DtxtNo") '"Desactivado"
				End Select %>&nbsp;</td>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getpollManLngStr("DtxtDelete")%>" name="btnDelete"></td>
				<td>&nbsp;</td>
				<td width="77">
				<input type="button" value="<%=getpollManLngStr("LtxtNewPoll")%>" name="B1" onclick="javascript:doMyLink('pollEdit.asp', '', '_self');"></td>
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
</script><!--#include file="agentBottom.asp"-->