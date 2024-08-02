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

<% addLngPathStr = "" %>
<!--#include file="lang/extPollExec.asp" -->
<body>
<% 
sql = "select IsNull(T1.AlterName, T0.Name) Name, IsNull(T1.AlterDescription, T0.Description) Description, StartDate, EndDate, Filter, " & _
"Case When Exists(select '' from OLKADPollAnswers where AdPollID = T0.AdPollID and CardCode = N'" & Request("CardCode") & "') Then 'Y' Else 'N' End Verfy, " & _
"(select top 1 CntctPrsn from OLKADPollAnswers where AdPollID = T0.AdPollID and CardCode = N'" & Request("CardCode") & "') SelectedContact " & _
"from OLKADPoll T0 " & _
"left outer join OLKADPollAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID " & _
"where T0.AdPollID = " & request("AdPollID")
set rs = conn.execute(sql)
qcName = rs("Name")
qcDesc = rs("Description")
qcDate = FormatDate(rs("StartDate"), True)
qcEndDate = FormatDate(rs("EndDate"), True)
qcReadOnly = rs("Verfy") = "Y"
SelectedContact = rs("SelectedContact")
sql = "select CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) CardName, Phone1, Phone2, CntctPrsn from OCRD where CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "'"
rs.close
set rs = conn.execute(sql)
CntctPrsn = rs("CntctPrsn")

set rd = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.LineID, T0.ChoiceID, IsNull(T1.AlterChoice, Choice) Choice " & _
	"from OLKADPollLinesChoices T0 " & _
	"left outer join OLKADPollLinesChoicesAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID and T1.LineID = T0.LineID and T1.ChoiceID = T0.ChoiceID " & _
	"where T0.AdPollID = " & Request("AdPollID") & " " & _
	"order by T0.LineID, T0.Ordr"
rd.open sql, conn, 3, 1
%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><% If Request("LineID") = "" Then %><%=getextPollExecLngStr("LtxtExtPolls")%><% Else %>
		<a class="GeneralTlt" href="extPollExec.asp?cmd=extpollman"><%=getextPollExecLngStr("LtxtExtPolls")%></a> - <a class="GeneralTlt" href="javascript:doMyLink('extPollview.asp', 'AdPollID=<%=Request("AdPollID")%>', '_self');"><%=getextPollExecLngStr("LtxtViewRes")%></a> - <a class="GeneralTlt" href="javascript:doMyLink('extPollview.asp', 'AdPollID=<%=Request("AdPollID")%>&LineID=<%=Request("LineID")%>&col=<%=Request("col")%>', '_self');"><%=getextPollExecLngStr("LtxtDetails")%></a> - <%=getextPollExecLngStr("LtxtClientPollSum")%><% End If %></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollExecLngStr("DtxtName")%></td>
		<td colspan="2"class="GeneralTbl"><%=qcName%>&nbsp;</td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollExecLngStr("LtxtStartDate")%></td>
		<td width="250"class="GeneralTbl"><%=qcDate%>&nbsp;</td>
		<td class="GeneralTblBold2"><%=getextPollExecLngStr("DtxtDescription")%></td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2" valign="top" style="padding-top: 1px;"><%=getextPollExecLngStr("LtxtEndDate")%></td>
		<td width="250" class="GeneralTbl" valign="top" style="padding-top: 1px;"><%=qcEndDate%>&nbsp;</td>
		<td rowspan="2" valign="top"class="GeneralTbl"><% If Not IsNull(qcDesc) Then %><%=Replace(qcDesc,VbNewLine,"<br>")%><% End If %>&nbsp;</td>
	</tr>
	</table>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTblBold2">
		<td class="style1"><%=getextPollExecLngStr("DtxtCode")%></td>
		<td class="style1"><%=getextPollExecLngStr("DtxtName")%></td>
		<td class="style1"><%=getextPollExecLngStr("DtxtPhone")%>&nbsp;1</td>
		<td class="style1"><%=getextPollExecLngStr("DtxtPhone")%>&nbsp;2</td>
		<td class="style1"><%=getextPollExecLngStr("DtxtContact")%></td>
	</tr>
	<% do while not rs.eof %>
	<tr class="GeneralTbl">
		<td><%=rs("CardCode")%></td>
		<td><%=rs("CardName")%></td>
		<td><%=rs("Phone1")%></td>
		<td><%=rs("Phone2")%></td>
		<td><%=rs("CntctPrsn")%></td>
	</tr>
	<% rs.movenext
	loop %>
</table>
<form method="post" name="frmEnc" action="poll/extPollExecSubmit.asp" onsubmit="return valFrm()">
<% sql = "select CntctCode, Name, Position, Tel1, Tel2, Cellolar, Fax, E_MailL from ocpr where CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "'"
rs.close
rs.open sql, conn, 3, 1 %>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTblBold2">
		<td width="1"></td>
		<td class="style1"><%=getextPollExecLngStr("DtxtContact")%></td>
		<td class="style1"><%=getextPollExecLngStr("LtxtPosition")%></td>
		<td class="style1"><%=getextPollExecLngStr("DtxtPhone")%>&nbsp;1</td>
		<td class="style1"><%=getextPollExecLngStr("DtxtPhone")%>&nbsp;2</td>
		<td class="style1"><%=getextPollExecLngStr("LtxtMobile")%></td>
		<td class="style1"><%=getextPollExecLngStr("DtxtFax")%></td>
		<td class="style1"><%=getextPollExecLngStr("DtxtEMail")%></td>
	</tr>
	<% do while not rs.eof %>
	<label for="cn<%=rs(0)%>">
	<tr class="GeneralTbl">
		<td width="1"><input <% If qcReadOnly Then %> disabled <% If CStr(SelectedContact) = CStr(rs(0)) Then %>checked<% End If %><% End If %> type="radio" class="noborder" id="cn<%=rs(0)%>" value="<%=rs(0)%>" name="CntctCode"></td>
		<td><%=rs(1)%>&nbsp;</td>
		<td><%=rs(2)%>&nbsp;</td>
		<td><%=rs(3)%>&nbsp;</td>
		<td><%=rs(4)%>&nbsp;</td>
		<td><%=rs(5)%>&nbsp;</td>
		<td><%=rs(6)%>&nbsp;</td>
		<td><%=rs(7)%>&nbsp;</td>
	</tr></label>
	<% rs.movenext
	loop %>
</table>
<% sql = "select T0.LineID, IsNull(T1.AlterQuestion, T0.Question) Question, T0.Type, T0.MandatoryNote, T2.Answer, T2.Notes " & _
"from OLKADPollLines T0 " & _
"left outer join OLKADPollLinesAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID and T1.LineID = T0.LineID " & _
"left outer join OLKADPollAnswers T2 on T2.AdPollID = T0.AdPollID and T2.CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' and T2.LineID = T0.LineID " & _
"where T0.AdPollID = " & Request("AdPollID") & " " & _
"order by Ordr asc"
rs.close
rs.open sql, conn, 3, 1 %>
<table border="0" cellpadding="0" cellspacing="2" width="100%">
	<% do while not rs.eof %>
	<tr class="GeneralTblBold2">
		<td><%=rs.bookmark%> - <%=rs("Question")%></td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5" style="font-family: Verdana; font-size: 10px">
			<tr class="GeneralTbl">
				<td width="200">
				<% Select Case rs("Type")
					Case "R" %>
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "1" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>1" value="1"><label for="qc<%=rs("LineID")%>1">1</label> 
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "2" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>2" value="2"><label for="qc<%=rs("LineID")%>2">2</label> 
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "3" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>3" value="3"><label for="qc<%=rs("LineID")%>3">3</label> 
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "4" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>4" value="4"><label for="qc<%=rs("LineID")%>4">4</label> 
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "5" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>5" value="5"><label for="qc<%=rs("LineID")%>5">5</label>
				<% Case "B" %>
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "Y" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>Y" value="Y"><label for="qc<%=rs("LineID")%>Y"><%=getextPollExecLngStr("DtxtYes")%></label> 
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "N" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%>N" value="N"><label for="qc<%=rs("LineID")%>N"><%=getextPollExecLngStr("DtxtNo")%></label>
				<% End Select %>
				</td>
				<td><% If rs("Type") = "M" Then
				rd.Filter = "LineID = " & rs("LineID")
				do while not rd.eof %>
				<input type="radio" class="noborder" <% If qcReadOnly Then %> disabled <% If rs("Answer") = "Y" Then %> checked <% End If %><% End If %> name="qc<%=rs("LineID")%>" id="qc<%=rs("LineID")%><%=rd("ChoiceID")%>" value="<%=rd("ChoiceID")%>"><label for="qc<%=rs("LineID")%><%=rd("ChoiceID")%>"><%=rd("Choice")%></label><br>
				<% rd.movenext
				loop 
				End If %><%=getextPollExecLngStr("DtxtNote")%>:
				<input type="text" name="N<%=rs("LineID")%>" <% If qcReadOnly Then %>readonly<% End If %> size="84" style="width: 100%;" value="<% If qcReadOnly Then %><%=rs("Notes")%><% End If %>"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="LineID" value="<%=rs("LineID")%>">
	<input type="hidden" name="Type<%=rs("LineID")%>" value="<%=rs("Type")%>">
	<input type="hidden" name="MandatoryNote<%=rs("LineID")%>" value="<%=rs("MandatoryNote")%>">
	<input type="hidden" name="LineN<%=rs("LineID")%>" value="<%=rs.bookmark%>">
	<% rs.movenext
	loop %>
</table>
<% If Not qcReadOnly Then %>
<table cellpadding="0" cellspacing="2" border="0" width="100%">
	<tr>
		<td><input type="submit" name="btnSave" value="<%=getextPollExecLngStr("DtxtSave")%>"></td>
		<td width="1"><input type="button" name="btnCancel" value="<%=getextPollExecLngStr("DtxtCancel")%>" onclick="javascript:if(confirm('<%=getextPollExecLngStr("DtxtConfCancel")%>'))doMyLink('extPollOpen.asp', 'ADPollID=<%=Request("ADPollID")%>', '_self');"></td>
	</tr>
</table>
<% End If %>
<input type="hidden" name="AdPollID" value="<%=Request("AdPollID")%>">
<input type="hidden" name="CardCode" value="<%=Request("CardCode")%>">
</form>
<script language="javascript">
function valFrm() 
{

	if (!chkRdGrp(document.frmEnc.CntctCode)) 
	{
		alert("<%=getextPollExecLngStr("LtxtValSelContact")%>");
		if (document.frmEnc.CntctCode.length)
			document.frmEnc.CntctCode[0].focus();
		else
			document.frmEnc.CntctCode.focus();
		return false;
	}
	
	var lineID = document.frmEnc.LineID;
	if (lineID.length)
	{
		for (var i = 0;i<lineID.length;i++)
		{
			var qc = document.getElementById('qc' + lineID[i].value);
			var note = document.getElementById('N' + lineID[i].value);
			var noteNotNull = document.getElementById('MandatoryNote' + lineID[i].value);
			var lineNum = document.getElementById('LineN' + lineID[i].value).value;

			if (!chkValues(qc, note, noteNotNull, lineNum))
			{
				return false;
			}			
		}
	}
	else
	{
		var qc = document.getElementById('qc' + lineID.value);
		var note = document.getElementById('N' + lineID.value);
		var noteNotNull = document.getElementById('MandatoryNote' + lineID.value);
		var lineNum = document.getElementById('LineN' + lineID.value).value;
		if (!chkValues(qc, note, noteNotNull, lineNum))
		{
			return false;
		}
	}
	
	return confirm('<%=getextPollExecLngStr("LtxtConfEndPoll")%>');
}

function chkValues(qc, note, noteNotNull, lineNum)
{
	if (!chkRdGrp(qc))
	{
		alert('<%=getextPollExecLngStr("LtxtValAnswer")%>'.replace('{0}', lineNum));
		return false;
	}
	else if (noteNotNull.value == 'Y' && note.value.length < 5)
	{
		alert('<%=getextPollExecLngStr("LtxtValNote")%>'.replace('{0}', lineNum));
		note.focus();
		return false;
	}
	return true;
}

function chkRdGrp(fld) 
{
	var retVal = false;
	if (fld)
	{
		if (fld.length)
		{
			for (var i = 0;i<fld.length;i++) 
			{
				if (fld[i].checked) 
				{ 
					retVal = true;
					break;
				}
			}
		}
		else
			retVal = fld.checked;
	}
	else
		retVal = true;
	return retVal;
}

</script>
<!--#include file="agentBottom.asp"-->