<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not myAut.HasAuthorization(39) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/extPollEdit.asp" -->
<!--#include file="genman/adminTradForm.asp"-->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript" src="js_up_down.js"></script>
<script language="javascript">
function valFrm() 
{
	if (document.frmNewQc.Name.value == '') 
	{
		alert("<%=getextPollEditLngStr("LtxtValPollNam")%>");
		document.frmNewQc.Name.focus();
		return false;
	}
	else if (document.frmNewQc.valQuery.value == 'Y' && document.frmNewQc.Filter.value != '')
	{
		alert('<%=getextPollEditLngStr("LtxtVerfyFilter")%>');
		document.frmNewQc.btnVerfyFilter.focus();
		return false;
	}
	else if (document.frmNewQc.Status.checked && document.frmNewQc.Agents.value == '')
	{
		alert('<%=getextPollEditLngStr("LtxtValSelAgent")%>');
		document.frmNewQc.txtAgents.focus();
		return false;
	}
	return true;
}
function agentsTo(agentsCode, agentsDesc)
{
	document.frmNewQc.Agents.value = agentsCode;
	document.frmNewQc.txtAgents.value = agentsDesc;
}

</script>
<% 
If Request("AdPollID") <> "" Then
	sql = "select Name, Description, StartDate, EndDate, Filter, Status, (select top 1 SlpCode from OLKADPollAgents where AdPollID = T0.AdPollID) SlpCode " & _
	"from OLKADPoll T0 where AdPollID = " & Request("AdPollID")
	set rs = conn.execute(sql)
	Name = rs("Name")
	Description = rs("Description")
	StartDate = FormatDate(rs("StartDate"), False)
	EndDate = FormatDate(rs("EndDate"), False)
	qFilter = rs("Filter")
	Status = rs("Status")
Else
	Status = "D"
	StartDate = FormatDate(Now(), False)
	EndDate = FormatDate(Now()+7, False)
End If

If Request("AdPollID") <> "" Then

	If rs("SlpCode") <> -2 Then
		sql = "SELECT T0.SlpCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, T0.SlpName) SlpName " & _
			  "FROM OSLP T0 " & _
			  "inner join OLKAgentsAccess T1 on T1.SlpCode = T0.SlpCode and T1.Access <> 'D' " & _
			  "inner join OLKADPollAgents T2 on T2.SlpCode = T0.SlpCode " & _
			  "WHERE T2.AdPollID = " & Request("AdPollID") & " and T0.SlpCode <> -1 " & _
			  "ORDER BY SlpName "
	
		set rd = Server.CreateObject("ADODB.RecordSet")
		set rd = conn.execute(sql)
		
		agentsCode = ""
		agentsDesc = ""
		
		do while not rd.eof
			If agentsCode <> "" Then agentsCode = agentsCode & ", "
			If agentsDesc <> "" Then agentsDesc = agentsDesc & ", "
			
			agentsCode = agentsCode & rd("SlpCode")
			agentsDesc = agentsDesc & rd("SlpName")
		rd.movenext
		loop
	ElseIf rs("SlpCode") = -2 Then
		agentsCode = "-2"
		agentsDesc = getextPollEditLngStr("DtxtAll")
	End If
	 
End If
%>

<form method="POST" action="genman/extPollSubmit.asp" onsubmit="return valFrm()" name="frmNewQc">
<% If Request("AdPollID") = "" Then %>
<input type="hidden" name="NameTrad" id="NameTrad" value="">
<input type="hidden" name="DescriptionTrad" id="DescriptionTrad" value="">
<% End If %>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td class="GeneralTlt"><% If Request("AdPollID") <> "" Then %><%=getextPollEditLngStr("LttlEditExtPoll")%><% Else %><%=getextPollEditLngStr("LttlAddExtPoll")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
		<tr>
			<td width="100" class="CanastaTblResaltada"><%=getextPollEditLngStr("DtxtName")%><font color="#FF0000">*</font></td>
			<td class="CanastaTbl">
			<table cellpadding="0" cellspacing="2" border="0">
				<tr class="CanastaTbl">
					<td width="384">
					<table border="0" cellpadding="0" width="100%">
						<tr>
							<td>
							<input class="input" size="20" style="width: 100%;" value="<%=Name%>" name="Name" onkeydown="return chkMax(event, this, 50);">
							</td>
							<td width="16"><a href="javascript:doFldTrad('ADPoll', 'AdPollID', '<%=Request("AdPollID")%>', 'AlterName', 'T', <% If Request("AdPollID") = "" Then %>document.frmNewQc.NameTrad<% Else %>null<% End If %>);"><img src="images/trad.gif" alt="<%=getextPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>
					<td><input type="checkbox" name="Status" class="noborder" id="Status" <% If Status = "A" Then %>checked<% End If %> value="A"></td>
					<td><label for="Status"><%=getextPollEditLngStr("DtxtActive")%></label></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100" valign="top" class="CanastaTblResaltada"><%=getextPollEditLngStr("DtxtDescription")%></td>
			<td class="CanastaTbl">
			<table border="0" cellpadding="0">
				<tr>
					<td>
					<textarea rows="6" name="Description" cols="70"><%=Description%></textarea>
					</td>
					<td width="16" valign="bottom"><a href="javascript:doFldTrad('ADPoll', 'AdPollID', '<%=Request("AdPollID")%>', 'AlterDescription', 'M', <% If Request("AdPollID") = "" Then %>document.frmNewQc.DescriptionTrad<% Else %>null<% End If %>);"><img src="images/trad.gif" alt="<%=getextPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100" class="CanastaTblResaltada">
			<a class="LinkNoticiasMas" href="javascript:Pic('genman/extAgentsUsers.asp?agentsusers='+document.frmNewQc.Agents.value+'&pop=Y&AddPath=../',320,480,'yes','no')">
			<font size="1"><% If 1 = 2 Then %>Agentes<% Else %><%=txtAgents%><% End If %></font></a></td>
			<td class="CanastaTbl">	
			<input type="text" name="txtAgents" readonly size="80" onclick="javascript:Pic('genman/extAgentsUsers.asp?agentsusers='+document.frmNewQc.Agents.value+'&pop=Y&AddPath=../',320,480,'yes','no')" value="<%=agentsDesc%>">
			<input type="hidden" name="Agents" value="<%=agentsCode%>"></td>
		</tr>
		<tr>
			<td width="100" class="CanastaTblResaltada"><%=getextPollEditLngStr("LtxtStartDate")%></td>
			<td class="CanastaTbl">	
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td><img border="0" src="images/cal.gif" id="btnStartDate" width="16" height="16">&nbsp;</td>
					<td>
					<input readonly class="input" type="text" name="StartDate" id="StartDate" size="11" value="<%=StartDate%>" onclick="btnStartDate.click()"></td>
				</tr>
			</table></td>
		</tr>
		<tr>
			<td width="100" class="CanastaTblResaltada"><%=getextPollEditLngStr("LtxtEndDate")%></td>
			<td class="CanastaTbl">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td><img border="0" src="images/cal.gif" id="btnEndDate" width="16" height="16">&nbsp;</td>
					<td>
					<input readonly class="input" type="text" name="EndDate" id="EndDate" size="11" value="<%=EndDate%>" onclick="btnEndDate.click()"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100" class="CanastaTblResaltada" valign="top"><%=getextPollEditLngStr("LtxtFilter")%><br>
						(<em>from OCRD where ...</em>)</td>
			<td class="CanastaTbl">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td class="style1"><textarea dir="ltr" rows="10" style="width: 100%" name="Filter" cols="87" class="input" onkeypress="javascript:document.frmNewQc.btnVerfyFilter.src='images/btnValidate.gif';document.frmNewQc.btnVerfyFilter.style.cursor = 'hand';document.frmNewQc.valQuery.value='Y';"><%=myHTMLEncode(qFilter)%></textarea>
					</td>
					<td width="24" valign="bottom" class="style1">
					<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getextPollEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmNewQc.valQuery.value == 'Y')VerfyQuery();">
					<input type="hidden" name="valQuery" value="N"></td>
				</tr>
			</table>
			</td>
		</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="1">
				<input type="submit" value="<%=getextPollEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="1">
				<input type="submit" value="<%=getextPollEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td>&nbsp;</td>
				<td width="1">
				<input type="button" value="<%=getextPollEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getextPollEditLngStr("DtxtConfCancel")%>'))window.location.href='extpollman.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="AdPollID" value="<%=Request("AdPollID")%>">
<input type="hidden" name="cmd" value="editQC">
</form>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "StartDate",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnStartDate",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    Calendar.setup({
        inputField     :    "EndDate",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnEndDate",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });

function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frmNewQc.Filter.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	//document.form2.btnVerfy.disabled = true;
	document.frmNewQc.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmNewQc.btnVerfyFilter.style.cursor = '';
	document.frmNewQc.valQuery.value='N';
}
</script>
<form name="frmVerfyQuery" action="genman/verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="ocrdFilter">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>

<% If Request("AdPollID") <> "" Then %>
<script language="javascript">
function valFrmAddQ() 
{
	if (document.frmAddQuestion.Question.value == '') 
	{
		alert("<%=getextPollEditLngStr("LtxtValQuestion")%>");
		document.frmAddQuestion.Question.focus();
		return false;
	}
	return true;
}
</script>
<% If Request("editIndex") <> "" Then
	sql = "select Question, Type, MandatoryNote, Ordr from OLKADPollLines where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("editIndex")
	set rs = conn.execute(sql)
	Question = rs("Question")
	qType = rs("Type")
	Ordr = rs("Ordr") 
	MandatoryNote = rs("MandatoryNote")
Else
	sql = "select IsNull(Max(Ordr)+1, 1) Ordr from OLKADPollLines where AdPollID = " & Request("AdPollID")
	set rs = conn.execute(sql)
	Ordr = rs("Ordr")
End If %>
<form method="POST" action="genman/extPollSubmit.asp" name="frmAddQuestion" onsubmit="return valFrmAddQ()">
<% If Request("editIndex") = "" Then %>
<input type="hidden" name="QuestionTrad" value="">
<% End If %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tblEditLine">
	<tr>
		<td class="GeneralTlt"><% If Request("editIndex") <> "" Then %><%=getextPollEditLngStr("LttlEditLine")%><% Else %><%=getextPollEditLngStr("LtxtAddLine")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
		<tr class="CanastaTblResaltada">
			<td align="center"><%=getextPollEditLngStr("LtxtQuestion")%></td>
			<td style="width: 140px">
			<p align="center"><%=getextPollEditLngStr("DtxtType")%></td>
			<td style="width: 60px">
			<p align="center"><%=getextPollEditLngStr("DtxtNote")%></td>
			<td style="width: 60px" align="center">
			<%=getextPollEditLngStr("DtxtOrder")%></td>
		</tr>
		<tr class="CanastaTbl">
			<td>
			<table border="0" cellpadding="0" id="table13" width="100%">
				<tr>
					<td>
					<input type="text" name="Question" style="width: 100%; " size="94" value="<%=Question%>">
					</td>
					<td width="16"><a href="javascript:doFldTrad('ADPollLines', 'AdPollID,LineID', '<%=Request("AdPollID")%>,<%=Request("editIndex")%>', 'AlterQuestion', 'T', <% If Request("editIndex") = "" Then %>document.frmAddQuestion.QuestionTrad<% Else %>null<% End If %>);"><img src="images/trad.gif" alt="<%=getextPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
				</tr>
			</table>
			</td>
			<td style="width: 140px">
			<p align="center">
			<select size="1" name="Type">
			<option <% If qType = "R" Then %>selected<% End If %> value="R"><%=getextPollEditLngStr("LtxtRange")%></option>
			<option <% If qType = "B" Then %>selected<% End If %> value="B"><%=getextPollEditLngStr("DtxtYes")%>/<%=getextPollEditLngStr("DtxtNo")%></option>
			<option <% If qType = "M" Then %>selected<% End If %> value="M"><%=getextPollEditLngStr("LtxtMultipleChoice")%></option>
			</select></td>
			<td style="width: 60px" align="center">
			<input type="checkbox" name="MandatoryNote" <% If MandatoryNote = "Y" Then %>checked<% End If %> class="noborder" value="Y">
			</td>
			<td style="width: 60px">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name='Ordr' id="Ordr" size="4" value='<%=Ordr%>' onfocus="this.select()"></td>
					<td valign="middle">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><img src="images/img_nud_up.gif" id="btnOrdrUp"></td>
						</tr>
						<tr>
							<td><img src="images/spacer.gif"></td>
						</tr>
						<tr>
							<td><img src="images/img_nud_down.gif" id="btnOrdrDown"></td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			<script language="javascript">NumUDAttach('frmAddQuestion', 'Ordr', 'btnOrdrUp', 'btnOrdrDown');</script>
			</td>
		</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="1">
				<input type="submit" value="<%=getextPollEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="1">
				<input type="submit" value="<% If Request("editIndex") = "" Then %><%=getextPollEditLngStr("DtxtAdd")%><% Else %><%=getextPollEditLngStr("DtxtSave")%><% End If %>" name="btnSave"></td>
				<td>&nbsp;</td>
				<td width="1">
				<input type="button" value="<%=getextPollEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getextPollEditLngStr("DtxtConfCancel")%>'))window.location.href='extPollEdit.asp?AdPollID=<%=Request("AdPollID")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="editQuestion">
<input type="hidden" name="AdPollID" value="<%=Request("AdPollID")%>">
<input type="hidden" name="LineID" value="<%=Request("editIndex")%>">
</form>
<% sql = "select Question, [Type], LineID, MandatoryNote, Ordr, " & _
"Case When Exists(select top 1 '' from OLKADPollAnswers where AdPollID = T0.AdPollID and LineID = T0.LineID) Then 'Y' Else 'N' End Verfy " & _
"from OLKADPollLines T0 " & _
"where AdPollID = " & Request("AdPollID") & " " & _
"order by ordr"
rs.close
rs.open sql, conn, 3, 1
If not rs.eof then %>
<form name="frmQuestions" method="post" action="genman/extPollSubmit.asp">
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tblQuestions">
	<tr>
		<td class="GeneralTlt"><%=getextPollEditLngStr("LtxtPollLines")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
		<tr class="CanastaTblResaltada">
			<td style="width: 15px">&nbsp;</td>
			<td><%=getextPollEditLngStr("LtxtQuestion")%></td>
			<td align="center" style="width: 130px"><%=getextPollEditLngStr("DtxtType")%></td>
			<td align="center" style="width: 80px"><%=getextPollEditLngStr("DtxtOptions")%></td>
			<td align="center" style="width: 60px"><%=getextPollEditLngStr("DtxtNote")%></td>
			<td align="center" style="width: 60px"><%=getextPollEditLngStr("DtxtOrder")%></td>
			<td style="width: 15px">&nbsp;</td>
		</tr>
		<% do while not rs.eof
		If CStr(Request("editChoice")) = CStr(rs("LineID")) Then editDesc = rs("Question") %>
		<tr class="CanastaTbl">
			<td style="width: 15px"><a href="javascript:doMyLink('extPollEdit.asp', 'AdPollID=<%=Request("AdPollID")%>&editIndex=<%=rs("LineID")%>', '_self');"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
			<td>
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr class="GeneralTbl">
					<td>
					<%=rs("Question")%>
					</td>
					<td width="16"><a href="javascript:doFldTrad('ADPollLines', 'AdPollID,LineID', '<%=Request("AdPollID")%>, <%=rs("LineID")%>', 'AlterQuestion', 'T', null);"><img src="images/trad.gif" alt="<%=getextPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
				</tr>
			</table>
			</td>
			<td align="center" style="width: 130px"><% Select Case rs("Type")
				Case "R" %><%=getextPollEditLngStr("LtxtRange")%>
				<% Case "B"%><%=getextPollEditLngStr("DtxtYes")%>/<%=getextPollEditLngStr("DtxtNo")%>
				<% Case "M" %><%=getextPollEditLngStr("LtxtMultipleChoice")%>
				<% End Select %></td>
			<td align="center" style="width: 80px"><% If rs("Type") = "M" Then %><a href='extPollEdit.asp?AdPollID=<%=Request("AdPollID")%>&amp;editChoice=<%=rs("LineID")%>&amp;#editChoice'><img src="images/options.gif" alt="<%=getextPollEditLngStr("DtxtEdit")%>" border="0"></a><% Else %>&nbsp;<% End If %></td>
			<td align="center" style="width: 60px"><% Select Case rs("MandatoryNote")
			Case "Y" %><%=getextPollEditLngStr("DtxtYes")%>
			<% Case "N"%><%=getextPollEditLngStr("DtxtNo")%>
			<% End Select%>&nbsp;</td>
			<td style="width: 60px">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name='Ordr<%=rs("LineID")%>' id="Ordr<%=rs("LineID")%>" size="4" value='<%=rs("Ordr")%>' onfocus="this.select()"></td>
					<td valign="middle">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><img src="images/img_nud_up.gif" id="btnOrdr<%=rs("LineID")%>Up"></td>
						</tr>
						<tr>
							<td><img src="images/spacer.gif"></td>
						</tr>
						<tr>
							<td><img src="images/img_nud_down.gif" id="btnOrdr<%=rs("LineID")%>Down"></td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			<script language="javascript">NumUDAttach('frmQuestions', 'Ordr<%=rs("LineID")%>', 'btnOrdr<%=rs("LineID")%>Up', 'btnOrdr<%=rs("LineID")%>Down');</script>
			</td>
			<td style="width: 15px">
			<% If rs("Verfy") = "N" Then %><a href="javascript:delLine(<%=rs("LineID")%>);"><img border="0" src="images/<%=Session("rtl")%>remove.gif"></a><% End If %>
			</td>
		</tr>
		<% rs.movenext
		loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<input type="submit" value="<%=getextPollEditLngStr("DtxtSave")%>" name="btnSave">
		</td>
	</tr>
</table>
<input type="hidden" name="AdPollID" value="<%=Request("AdPollID")%>">
<input type="hidden" name="cmd" value="updateQuestions">
</form>
<script type="text/javascript">
<!--
function delLine(LineID)
{
	if(confirm('<%=getextPollEditLngStr("LtxtConfDelQuestion")%>'))window.location.href='genman/extPollSubmit.asp?cmd=delQuestion&AdPollID=<%=Request("AdPollID")%>&LineID=' + LineID;
}
//-->
</script>
<% If Request("editChoice") <> "" Then %>
<% 
MaxOrdr = 0
sql = "select T0.ChoiceID, T0.Choice, T0.Color, T0.Ordr, " & _
"Case When Exists(select top 1 '' from OLKADPollAnswers where AdPollID = T0.AdPollID and LineID = T0.LineID and Answer = Convert(nvarchar(50),T0.ChoiceID)) Then 'Y' Else 'N' End Verfy " & _
"from OLKADPollLinesChoices T0 " & _
"where T0.AdPollID = " & Request("AdPollID") & " and T0.LineID = " & Request("editChoice") & " " & _
"order by T0.ordr"
rs.close
rs.open sql, conn, 3, 1 %>
<script type="text/javascript">
<!--
function valChoicesFrm()
{
	if (document.frmChoices.ChoiceNew.value != '' && document.frmChoices.ForeColorNew.value == '')
	{
		alert('<%=getextPollEditLngStr("LtxtValNewChoiceColor")%>');
		document.frmChoices.ForeColorNew.click();
		return false;
	}
	return true;
}
//-->
</script>

<form name="frmChoices" method="post" action="genman/extPollSubmit.asp" onsubmit="javascript:return valChoicesFrm();">
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="editChoice">
	<tr>
		<td class="GeneralTlt"><%=getextPollEditLngStr("LtxtMultipleChoice")%> - <%=editDesc%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
		<tr class="CanastaTblResaltada">
			<td style="width: 120px"><%=getextPollEditLngStr("DtxtColor")%></td>
			<td><%=getextPollEditLngStr("DtxtOption")%></td>
			<td style="width: 60px"><%=getextPollEditLngStr("DtxtOrder")%></td>
			<td style="width: 15px">&nbsp;</td>
		</tr>
		<% do while not rs.eof
		If rs("Ordr") > MaxOrdr Then MaxOrdr = rs("Ordr") %>
		<tr class="CanastaTbl">
			<td style="width: 120px">
			<table cellpadding="0" cellpadding="2" border="0">
				<tr>
					<td style="width: 74px">
					<table cellpadding="0" cellspacing="0" border="0" style="border: 1px solid">
						<tr>
							<td style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
								<input type="text" readonly name="ForeColor<%=rs("ChoiceID")%>" size="9" style="cursor: default" maxlength="7" value="<%=rs("Color")%>" onclick="showColorPicker(btnChangeForeColor<%=rs("ChoiceID")%>,this,ForeColor<%=rs("ChoiceID")%>Sample)"></td>
							<td style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"><img src="images_picker/select_arrow_small.gif" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" id="btnChangeForeColor<%=rs("ChoiceID")%>" onclick="showColorPicker(this,document.frmChoices.ForeColor<%=rs("ChoiceID")%>,ForeColor<%=rs("ChoiceID")%>Sample)"></td>
						</tr>
					</table>
					</td>
					<td width="46" bgcolor="<%=rs("Color")%>" style="border: 1px solid; font-size: 10px" id="ForeColor<%=rs("ChoiceID")%>Sample">&nbsp;</td>
				</tr>
			</table>
			</td>
			<td>
			<table cellpadding="0" cellspacing="0" border="0" style="width: 100%; ">
				<tr class="GeneralTbl">
					<td>
					<input type="text" name="Choice<%=rs("ChoiceID")%>" maxlength="256" style="width: 100%; " size="94" value="<%=rs("Choice")%>">
					</td>
					<td width="16"><a href="javascript:doFldTrad('ADPollLinesChoices', 'AdPollID,LineID,ChoiceID', '<%=Request("AdPollID")%>, <%=Request("editChoice")%>, <%=rs("ChoiceID")%>', 'AlterChoice', 'T', null);"><img src="images/trad.gif" alt="<%=getextPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
				</tr>
			</table>
			</td>
			<td style="width: 60px">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name='choiceOrdr<%=rs("ChoiceID")%>' id="choiceOrdr<%=rs("ChoiceID")%>" size="4" value='<%=rs("Ordr")%>' onfocus="this.select()"></td>
					<td valign="middle">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><img src="images/img_nud_up.gif" id="btnChoiceOrdr<%=rs("ChoiceID")%>Up"></td>
						</tr>
						<tr>
							<td><img src="images/spacer.gif"></td>
						</tr>
						<tr>
							<td><img src="images/img_nud_down.gif" id="btnChoiceOrdr<%=rs("ChoiceID")%>Down"></td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			<script language="javascript">NumUDAttach('frmChoices', 'choiceOrdr<%=rs("ChoiceID")%>', 'btnChoiceOrdr<%=rs("ChoiceID")%>Up', 'btnChoiceOrdr<%=rs("ChoiceID")%>Down');</script>
			</td>
			<td style="width: 15px">
			<% If rs("Verfy") = "N" Then %><a href="javascript:delChoice(<%=rs("ChoiceID")%>);"><img border="0" src="images/<%=Session("rtl")%>remove.gif"></a><% Else %>&nbsp;<% End If %>
			</td>
		</tr>
		<% rs.movenext
		loop %>
		<tr class="CanastaTbl">
			<td style="width: 120px">
			<table cellpadding="0" cellpadding="2" border="0">
				<tr>
					<td style="width: 74px">
					<table cellpadding="0" cellspacing="0" border="0" style="border: 1px solid">
						<tr>
							<td style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
								<input type="text" readonly name="ForeColorNew" size="9" style="cursor: default" maxlength="7" value="" onclick="showColorPicker(btnChangeForeColorNew,this,ForeColorNewSample)"></td>
							<td style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"><img src="images_picker/select_arrow_small.gif" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" id="btnChangeForeColorNew" onclick="showColorPicker(this,document.frmChoices.ForeColorNew,ForeColorNewSample)"></td>
						</tr>
					</table>
					</td>
					<td width="46" style="border: 1px solid; font-size: 10px" id="ForeColorNewSample">&nbsp;</td>
				</tr>
			</table>
			</td>
			<td>
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr class="GeneralTbl">
					<td>
					<input type="text" name="ChoiceNew" style="width: 100%; " maxlength="256" size="94" value="">
					<input type="hidden" name="ChoiceNewTrad">
					</td>
					<td width="16"><a href="javascript:doFldTrad('ADPollLinesChoices', 'AdPollID,LineID,ChoiceID', '', 'AlterChoice', 'T', document.frmChoices.ChoiceNewTrad);"><img src="images/trad.gif" alt="<%=getextPollEditLngStr("DtxtTranslate")%>" border="0"></a></td>
				</tr>
			</table>
			</td>
			<td style="width: 60px">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td><input class="input" onkeydown="return chkMax(event, this, 254);" type="text" name='choiceOrdrNew' id="choiceOrdrNew" size="4" value='<%=MaxOrdr+1%>' onfocus="this.select()"></td>
					<td valign="middle">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><img src="images/img_nud_up.gif" id="btnChoiceOrdrNewUp"></td>
						</tr>
						<tr>
							<td><img src="images/spacer.gif"></td>
						</tr>
						<tr>
							<td><img src="images/img_nud_down.gif" id="btnChoiceOrdrNewDown"></td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			<script language="javascript">NumUDAttach('frmChoices', 'choiceOrdrNew', 'btnChoiceOrdrNewUp', 'btnChoiceOrdrNewDown');</script>
			</td>
			<td style="width: 15px">
			&nbsp;
			</td>
		</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="1">
				<input type="submit" value="<%=getextPollEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="1">
				<input type="submit" value="<%=getextPollEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td>&nbsp;</td>
				<td width="1">
				<input type="button" value="<%=getextPollEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getextPollEditLngStr("DtxtConfCancel")%>'))window.location.href='extPollEdit.asp?AdPollID=<%=Request("AdPollID")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="AdPollID" value="<%=Request("AdPollID")%>">
<input type="hidden" name="editChoice" value="<%=Request("editChoice")%>">
<input type="hidden" name="cmd" value="updateChoice">
</form>
<script type="text/javascript">
<!--
function delChoice(ChoiceID)
{
	if(confirm('<%=getextPollEditLngStr("LtxtConfDelChoice")%>'))window.location.href='genman/extPollSubmit.asp?cmd=delChoice&AdPollID=<%=Request("AdPollID")%>&LineID=<%=Request("editChoice")%>&ChoiceID=' + ChoiceID;
}
//-->
</script>
<% End If %>

<% End If %><% End If %>
<!--#include file="agentBottom.asp"-->