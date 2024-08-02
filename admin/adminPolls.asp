<!--#include file="top.asp" -->
<!--#include file="lang/adminPolls.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	border-bottom-style: none;
	border-bottom-width: medium;
}
.style2 {
	border-left-width: 1px;
	border-right-width: 1px;
	border-top: medium none #C0C0C0;
	border-bottom-width: 1px;
}
.style3 {
	border-left-width: 1px;
	border-right-width: 1px;
	border-top: medium none #C0C0C0;
	border-bottom-width: 1px;
	text-align: center;
}
.style4 {
	border-bottom-style: none;
	border-bottom-width: medium;
	text-align: center;
	color: #31659C;
}
.style5 {
	text-align: center;
	font-weight: normal;
	color: #31659C;
}
.style6 {
	border-bottom-style: none;
	border-bottom-width: medium;
	font-weight: bold;
	text-align: center;
}
.style7 {
	border-bottom-style: none;
	border-bottom-width: medium;
	text-align: center;
	font-size: xx-small;
	color: #31659C;
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
	page = 'pollView.asp?pollIndex='+pollIndex+'&pop=Y';
	OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=250,height=250");
}
</script>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font color="#31659C" face="Verdana" size="1"><%=getadminPollsLngStr("LttlPolls")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminPollsLngStr("LttlPollsNote")%></font></td>
	</tr>
	<form method="post" action="adminSubmit.asp" name="frmPolls">
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table6" style="font-family: Verdana; font-size: 10px; width: 100%;">
			<tr>
				<td align="center" bgcolor="#E2F3FC" class="style1" style="width: 16px">
				&nbsp;</td>
				<td bgcolor="#E2F3FC" class="style7" style="width: 100px">
				<strong><%=getadminPollsLngStr("DtxtDateOfPub")%></strong></td>
				<td bgcolor="#E2F3FC" class="style6">
				<p class="style5">
				<strong><%=getadminPollsLngStr("DtxtName")%></strong></td>
				<td bgcolor="#E2F3FC" class="style4" style="width: 100px">
				<strong><%=getadminPollsLngStr("LtxtVotes")%></strong></td>
				<td bgcolor="#E2F3FC" class="style4" style="width: 100px">
				<strong><%=getadminPollsLngStr("DtxtActive")%></strong></td>
				<td align="center" bgcolor="#E2F3FC" class="style1" style="width: 16px">
				&nbsp;</td>
			</tr>
			<% do while not rs.eof %>
			<input type="hidden" name="pollIndex" value="<%=rs("pollIndex")%>">
			<tr>
				<td class="style2" bgcolor="#F3FBFE" style="width: 16px">
				<a href="adminPollEdit.asp?pollIndex=<%=rs("pollIndex")%>">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td class="style2" bgcolor="#F3FBFE" style="width: 100px"><font color="#4783C5" size="1"><%=FormatDate(rs("pollDate"), False)%></font></td>
				<td class="style2" bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td valign="middle"><font color="#4783C5" size="1"><%=rs("pollName")%>&nbsp;<% If rs("Verfy") = "Y" Then %>(<%=getadminPollsLngStr("LtxtOnline")%>)<%end if%></font>
						</td>
						<td width="16">
						<a href="javascript:doFldTrad('Poll', 'pollIndex', <%=rs("pollIndex")%>, 'alterPollTitle', 'T', null);"><img src="images/trad.gif" alt="<%=getadminPollsLngStr("DtxtTranslate")%>" border="0">
						</a></td>
					</tr>
				</table>
				</td>
				<td class="style2" bgcolor="#F3FBFE" align="center" style="width: 100px">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><img border="0" src="images/eye_icon.gif" dir="ltr" width="15" height="13" alt="<%=getadminPollsLngStr("LtxtViewResults")%>" style="cursor: hand" onclick="javascript:viewPoll(<%=rs("pollIndex")%>);"></td>
						<td><font color="#4783C5" size="1">&nbsp;<%=rs("votes")%>&nbsp;</font></td>
					</tr>
				</table>
				</td>
				<td class="style3" bgcolor="#F3FBFE" style="width: 100px">
				<input type="checkbox" class="noborder" name="Status<%=rs("pollIndex")%>" value="O" <% If rs("pollStatus") = "O" Then %>checked<% End If %>></td>
				<td class="style2" bgcolor="#F3FBFE" style="width: 16px">
				<a href="javascript:if(confirm('<%=getadminPollsLngStr("LtxtConfDelPoll")%>'.replace('{0}', '<%=Replace(rs("pollName"), "'", "\'")%>')))window.location.href='pollSubmit.asp?pollCmd=D&pollIndex=<%=rs("pollIndex")%>'">
				<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminPollsLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminPollsLngStr("DtxtNew")%>" name="B1" class="OlkBtn" onclick="javascript:window.location.href='adminPollEdit.asp'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="admPolls">
	</form>
</table>
<!--#include file="bottom.asp" -->