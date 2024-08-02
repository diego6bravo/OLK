<!--#include file="top.asp" -->
<!--#include file="lang/adminPollEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}
function setTimeStamp(vardate) {
document.frmPoll.pollDate.value = vardate
}
</script>
<% 
conn.execute("use [" & Session("OLKDB") & "]")
set rs = Server.CreateObject("ADODB.recordset")
If Request("pollIndex") <> "" Then
	sql = "select pollIndex, pollName, pollTitle, pollDate, pollStatus,  " & _
			"(select count('A') from olkpolldb where pollIndex = T0.pollIndex) Votes, " & _
			"IsNull((select Max(lineOrder) from OLKPollLines where pollIndex = T0.pollIndex)+1, 1) lineOrder " & _
			"from olkPoll T0 " & _
			"where pollIndex = " & Request("pollIndex")
	set rs = conn.execute(sql)
	pollCmd = "updatePoll"
	pollName = rs("pollName")
	pollTitle = rs("pollTitle")
	pollDate = rs("pollDate")
	pollStatus = rs("pollStatus")
	lineOrder = rs("lineOrder")
	Votes = rs("Votes")
	Status = rs("pollStatus")
Else
	pollCmd = "addPoll"
	pollName = ""
	pollTitle = ""
End If
%>
<script language="javascript">
function valFrm()
{
	if (document.frmPoll.pollName.value == '')
	{
		alert("<%=getadminPollEditLngStr("LtxtValPollNam")%>");
		document.frmPoll.pollName.focus();
		return false;
	}
	else if (document.frmPoll.pollTitle.value == '')
	{
		alert("<%=getadminPollEditLngStr("LtxtValPollTtl")%>");
		document.frmPoll.pollTitle.focus();
		return false;
	}
	else if (document.frmPoll.pollDate.value == '') 
	{
		alert("<%=getadminPollEditLngStr("LtxtValPollDat")%>");
		document.frmPoll.pollDate.focus();
		return false;
	}
	return true;
}
function changepic(lineID,img_src) {
document['img' + lineID].src="Poll/colo"+img_src+".gif";
document.getElementById('cIndex' + lineID).value = img_src;
}
</script>
<script language="javascript" src="js_up_down.js"></script>
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
}
.style2 {
	background-color: #F3FBFE;
}
.style3 {
	background-color: #E2F3FC;
	text-align: center;
	color: #31659C;
}
.style4 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style5 {
	color: #31659C;
}
.style6 {
	background-color: #E2F3FC;
	text-align: center;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<form method="POST" action="pollSubmit.asp" name="frmPoll" onsubmit="javascript:return valFrm();">
<% If Request("pollIndex") = "" Then %>
<input type="hidden" name="pollTitleTrad">
<% End If %>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><% If Request("pollIndex") <> "" Then %><%=getadminPollEditLngStr("LttlEditPoll")%><% Else %><%=getadminPollEditLngStr("LttlAddPoll")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminPollEditLngStr("LttlPollNote")%></font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="616" id="table6">
			<tr>
				<td width="141" class="style4">
				<p align="justify" class="style5">
				<font face="Verdana" size="1"><strong><%=getadminPollEditLngStr("DtxtName")%></strong></font></td>
				<td class="style2">
				<p align="justify">
				<input class="input" type="text" name="pollName" size="50" value="<%=Server.HTMLEncode(pollName)%>" onkeydown="return chkMax(event, this, 50);">
				</td>
				<td class="style2">
				<% If Request("pollIndex") <> "" Then %>
				<p align="right">
				<input type="checkbox" name="pollStatus" value="O" <% If Status = "O" Then %>checked<% End If %> id="chkActive" class="noborder"><font color="#4783C5" face="Verdana" size="1"><label for="chkActive"><%=getadminPollEditLngStr("DtxtActive")%></label></font><% End If %></td>
			</tr>
			<tr>
				<td width="141" class="style4">
				<font face="Verdana" size="1"><strong><%=getadminPollEditLngStr("DtxtTitle")%></strong></font></td>
				<td colspan="2" class="style2">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input class="input" type="text" name="pollTitle" style="width: 100%; " size="70" value="<%=Server.HTMLEncode(pollTitle)%>" onkeydown="return chkMax(event, this, 100);"></td>
						<td style="width: 16px"><a href="javascript:doFldTrad('Poll', 'pollIndex', '<%=Request("pollIndex")%>', 'alterPollTitle', 'T', <% If Request("pollIndex") <> "" Then %>null<% Else %>document.frmPoll.pollTitleTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="141" class="style4">
				<font face="Verdana" size="1"><strong><%=getadminPollEditLngStr("DtxtDateOfPub")%></strong></font></td>
				<td colspan="2" class="style2">				
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><img border="0" src="images/cal.gif" id="btnPollDate" width="16" height="16">&nbsp;</td>
						<td>
						<input readonly class="input" type="text" name="pollDate" id="pollDate" size="11" value="<%=FormatDate(pollDate, False)%>" onclick="btnPollDate.click()"></td>
					</tr>
				</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<% If Request("pollIndex") <> "" Then
	   rs.close
	   sql = "select * from olkPollLines where pollIndex = " & Request("pollIndex") & _
	   " order by lineOrder asc"
	   rs.open sql, conn, 3, 1
	%>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminPollEditLngStr("LttlPollLines")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminPollEditLngStr("LttlPollLinesNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8">
			<tr>
				<td colspan="2" style="font-size: 4px">&nbsp;<table border="0" id="table10">
					<tr>
						<td width="206" class="style6">
		<font face="Verdana" size="1"><strong><%=getadminPollEditLngStr("DtxtDescription")%></strong></font></td>
						<td class="style3">
		<font face="Verdana" size="1"><strong><%=getadminPollEditLngStr("DtxtOrder")%></strong></font></td>
						<td class="style3">
		<font face="Verdana" size="1"><strong><%=getadminPollEditLngStr("DtxtColor")%></strong></font></td>
						<td width="101" class="style1">
		&nbsp;</td>
						<td class="style1" style="width: 16px">
						</td>
					</tr>
				<% do while not rs.eof %>
					<tr>
						<td style="font-family: Verdana; font-size: 10px; color: #4783C5" width="206" class="style2">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name="opt<%=rs("PollLineNum")%>" size="30" value="<%=Server.HTMLEncode(rs("LineText"))%>" onfocus="this.select()">
								</td>
								<td width="16"><a href="javascript:doFldTrad('PollLines', 'pollIndex,pollLineNum', '<%=Request("pollIndex")%>,<%=rs("pollLineNum")%>', 'AlterLineText', 'T', null);"><img src="images/trad.gif" alt="<%=getadminPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</td>
						<td class="style2">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name='order<%=rs("PollLineNum")%>' id="order<%=rs("PollLineNum")%>" size="4" value='<%=rs("lineOrder")%>' onfocus="this.select()"></td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnorder<%=rs("PollLineNum")%>Up"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnorder<%=rs("PollLineNum")%>Down"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						<script language="javascript">NumUDAttach('frmPoll', 'order<%=rs("PollLineNum")%>', 'btnorder<%=rs("PollLineNum")%>Up', 'btnorder<%=rs("PollLineNum")%>Down');</script>
						</td>
						<td class="style2">
						<img border="0" id="img<%=rs("pollLineNum")%>" name="img<%=rs("pollLineNum")%>" src="Poll/colo<%=rs("colorIndex")%>.gif" width="50" height="12"><input type="hidden" name="cIndex<%=rs("pollLineNum")%>" id="cIndex<%=rs("pollLineNum")%>" value="<%=rs("colorIndex")%>"></td>
						<td width="101" class="style2">
						<img border="0" src="Poll/colo7.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','7')"><img border="0" src="Poll/colo0.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','0')"><img border="0" src="Poll/colo1.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','1')"><img border="0" src="Poll/colo5.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','5')"><img border="0" src="Poll/colo4.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','4')"><img border="0" src="Poll/colo3.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','3')"><img border="0" src="Poll/colo6.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','6')"><img border="0" src="Poll/colo2.gif" width="12" height="12" onclick="changepic('<%=rs("pollLineNum")%>','2')"></td>
						<td style="font-family: Verdana; font-size: 10px; color: #4783C5; width: 16px;" class="style2"><% if rs.recordcount > 1 then %>
						<p align="center">
						<% If Votes = 0 Then %>
						<a href="javascript:if(confirm('<%=getadminPollEditLngStr("LtxtConfDelLine")%>'))window.location.href='pollSubmit.asp?pollCmd=del&pollIndex=<%=Request("pollIndex")%>&pollLineNum=<%=rs("pollLineNum")%>'">
						<img border="0" src="images/remove.gif" width="16" height="16"></a><% end if %><% End If %></td>
					</tr>
					<% rs.movenext
					loop %>
					<tr>
						<td style="font-family: Verdana; font-size: 10px; color: #4783C5" width="206" class="style2">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td>
								<input type="hidden" name="optNewTrad">
								<input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name="optNew" size="30" onfocus="this.select()">
								</td>
								<td width="16"><a href="javascript:doFldTrad('PollLines', 'pollIndex,pollLineNum', '', 'AlterLineText', 'T', frmPoll.optNewTrad);"><img src="images/trad.gif" alt="<%=getadminPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</td>
						<td class="style2">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name='orderNew' id="orderNew" size="4" value='<%=lineOrder%>' onfocus="this.select()"></td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnorderNewUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnorderNewDown"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						<script language="javascript">NumUDAttach('frmPoll', 'orderNew', 'btnorderNewUp', 'btnorderNewDown');</script>
						</td>
						<td class="style2">
						<%
						randomize()
						randomNumber=Int(7 * rnd()) %>
						<img border="0" id="imgNew" name="imgNew" src="Poll/colo<%=randomNumber%>.gif" width="50" height="12"><input type="hidden" name="cIndexNew" id="cIndexNew" value="<%=randomNumber%>"></td>
						<td width="101" class="style2">
						<img border="0" src="Poll/colo7.gif" width="12" height="12" onclick="changepic('New','7')"><img border="0" src="Poll/colo0.gif" width="12" height="12" onclick="changepic('New','0')"><img border="0" src="Poll/colo1.gif" width="12" height="12" onclick="changepic('New','1')"><img border="0" src="Poll/colo5.gif" width="12" height="12" onclick="changepic('New','5')"><img border="0" src="Poll/colo4.gif" width="12" height="12" onclick="changepic('New','4')"><img border="0" src="Poll/colo3.gif" width="12" height="12" onclick="changepic('New','3')"><img border="0" src="Poll/colo6.gif" width="12" height="12" onclick="changepic('New','6')"><img border="0" src="Poll/colo2.gif" width="12" height="12" onclick="changepic('New','2')"></td>
						<td style="font-family: Verdana; font-size: 10px; color: #4783C5; width: 16px;" class="style2">
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td style="font-size: 4px">&nbsp;</td>
				<td style="font-size: 4px">&nbsp;</td>
			</tr>
		</table>

		</td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminPollEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminPollEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminPollEditLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminPollEditLngStr("DtxtConfCancel")%>'))window.location.href='adminPolls.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	</table>
<input type="hidden" name="pollIndex" value="<%=Request("pollIndex")%>">
<input type="hidden" name="pollCmd" value="<%=pollCmd%>">
</form>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "pollDate",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnPollDate",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
<!--#include file="bottom.asp" -->