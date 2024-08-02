<% addLngPathStr = "messages/" %>
<!--#include file="lang/messageNew.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	direction: ltr;
}
</style>
</head>

<% 
Session("sendmsg") = False
conn.execute("delete olkMsgTemp where slpcode = " & Session("VendId")) 
If Request("ClientsTo") <> "" Then
sql = "insert olkmsgtemp(SlpCode, CardCode) Values(" & Session("vendid") & ", N'" & saveHTMLDecode(Request("ClientsTo"), False) & "')"
conn.execute(sql)
End If
sql = "select (select Status from OLKSections where SecType = 'S' and SecID = -4) EnableClientMsg"
set rs = conn.execute(sql)
If rs(0) = "A" Then optCMsg = True Else optCMsg = False
rs.close %>
<script type="text/javascript">
function sapTo(sapUsers, sapUsersName)
{
	document.frmMessage.SapTo.value = sapUsers;
	document.frmMessage.SapToName.value = sapUsersName;
}

function agentsTo(agentsUsers)
{
	document.frmMessage.AgentsTo.value = agentsUsers;
}

function clientsTo(clientsUsers)
{
	document.frmMessage.ClientsTo.value = clientsUsers;
}

header = "Bodegas"
function valFrmMsg()
{
	if (document.frmMessage.SapTo.value == '' && document.frmMessage.AgentsTo.value == '' && document.frmMessage.ClientsTo.value == '')
	{
		alert ('<%=getmessageNewLngStr("LtxtValMsgToUser")%>');
		return false;
	}
	else if (document.frmMessage.Subject.value == '')
	{
		alert ('<%=getmessageNewLngStr("LtxtValMsgSubject")%>');
		document.frmMessage.Subject.focus();
		return false;
	}
	return true;
}
</script>
<form method="POST" action="newMessageSubmit.asp" name="frmMessage" onsubmit="return valFrmMsg();">
	<div align="center">
	<table border="0" cellpadding="0" width="550" id="table1">
		<tr class="GeneralTlt">
			<td colspan="2">
			<%=getmessageNewLngStr("LttlSendMsg")%></td>
		</tr>
		<tr class="GeneralTbl">
			<td width="97">
			<a id="lnkSAPUsers" class="LinkNoticiasMas" href="javascript:Pic('messages/sapUsers.asp?sapusers='+document.frmMessage.SapTo.value+'&pop=Y&AddPath=../',320,480,'yes','no')">
			<%=getmessageNewLngStr("DtxtSAP")%></a></td>
			<td width="447"><input type="hidden" name="SapTo" value=""><input readonly type="text" name="SapToName" size="88" onclick="javascript:Pic('messages/sapUsers.asp?sapusers='+document.frmMessage.SapTo.value+'&pop=Y&AddPath=../',320,480,'yes','no')"></td>
		</tr>
		<tr class="GeneralTbl" >
			<td  width="97">
			<a class="LinkNoticiasMas" href="javascript:Pic('messages/agentsUsers.asp?agentsusers='+document.frmMessage.AgentsTo.value+'&pop=Y&AddPath=../',320,480,'yes','no')">
			<% If 1 = 2 Then %>Agentes<% Else %><%=txtAgents%><% End If %></a></td>
			<td width="447"><input readonly type="text" name="AgentsTo" size="88" onclick="javascript:Pic('messages/agentsUsers.asp?agentsusers='+document.frmMessage.AgentsTo.value+'&pop=Y&AddPath=../',320,480,'yes','no')"></td>
		</tr>
		<% If optCMsg Then %>
		<tr class="GeneralTbl" >
			<td  width="97">
			<a class="LinkNoticiasMas" href="javascript:Pic('messages/clientSearch.asp?agentsusers='+document.frmMessage.AgentsTo.value+'&pop=Y&AddPath=../',500,290,'no','no')">
			<% If 1 = 2 Then %>Clientes<% Else %><%=txtClients%><% End If %></a></td>
			<td width="447"><input type="text" readonly name="ClientsTo" size="88" onclick="javascript:Pic('messages/clientSearch.asp?agentsusers='+document.frmMessage.AgentsTo.value+'&pop=Y&AddPath=../',500,290,'no','no')" value="<%=Request("ClientsTo")%>"></td>
		</tr>
		<% End If %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold" width="544" colspan="2">
			<table border="0" cellpadding="0" width="100%" id="table2">
				<tr class="GeneralTbl">
					<td class="GeneralTblBold2" width="73"><%=getmessageNewLngStr("DtxtSubject")%></td>
					<td><input type="text" name="Subject" size="76" onkeydown="return chkMax(event, this, 80);"></td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="checkbox" name="Urgente" value="Y" style="background:background-image;border: 0px solid" id="fp1"><label for="fp1"><%=getmessageNewLngStr("LtxtUrgent")%></label></td>
				</tr>
				<tr class="GeneralTbl">
					<td class="style1" colspan="3"><%=getmessageNewLngStr("DtxtMessage")%></td>
				</tr>
				<tr class="GeneralTbl">
					<td class="GeneralTblBold2" colspan="3">
					<textarea rows="21" name="Message" cols="87"></textarea></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold" colspan="2">
			<p align="center">
			<input type="submit" value="<%=getmessageNewLngStr("DtxtSend")%>" name="B1"> -
			<input type="reset" value="<%=getmessageNewLngStr("DtxtClear")%>" name="B2"></td>
		</tr>
		</table>
    	</div>
    	<input type="hidden" name="cmd" value="messagePost">
    	<input type="hidden" name="redir" value="newMessage">
    </form>