<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C"
	user = Session("UserName")
	MainDoc = "clientes" %><!--#include file="clientTop.asp"-->
<% 
If (Session("UserName") = "-Anon-") Then Response.Redirect "default.asp"
Case "V"
	user = Session("vendid")
	MainDoc = "ventas" %><!--#include file="agentTop.asp"-->
<%
End Select
addLngPathStr = "" %>
<!--#include file="lang/messageReturn.asp" -->
<script language="javascript">
function getAgents() {
OpenWin = this.open('', "GetAgents", "scrollbars=no,resizable=no, width=320,height=480");
doMyLink('messages/agentsUsers.asp', 'agentsusers='+document.frmMsg.AgentsTo.value+'&rParent=Y&pop=Y&AddPath=../', 'GetAgents');
OpenWin.focus();
}
</script>
<%
Session("sendmsg") = False
set rs = Server.CreateObject("ADODB.recordset")

If userType = "V" Then 
	user = Session("vendid")
ElseIf userType = "C" Then
	user = Session("UserName")
End If 
If Request("cmd") <> "newMsg" Then 

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetMessageReturn" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@User") = user
	cmd("@OlkLog") = Request("olklog")
	set rs = cmd.execute()
    msgDate = rs("Date")
End If

If Request("cmd") <> "newMsg" Then Title = Replace(getmessageReturnLngStr("LttlRespMsg"), "{0}", Request("OlkLog")) Else Title = getmessageReturnLngStr("LttlNewMsg")
      %>
<script type="text/javascript">
function agentsTo(agentsUsers)
{
	document.frmMsg.AgentsTo.value = agentsUsers;
}
</script>
<form method="POST" name="frmMsg" action="messagePostCL.asp">
	<table border="0" cellpadding="0" width="100%">
		<% If tblCustTtl = "" Then %>
		<tr class="TablasTituloSec">
			<td colspan="3" id="tdMyTtl">
			<p align="center"><%=Title%>&nbsp;
			</td>
		</tr>
		<% Else %>
		<tr class="TablasTituloSec">
			<td colspan="3">
			<table cellpadding="0" cellspacing="0" width="100%" border="0">
			<tr>
				<td>
				<%=Replace(Replace(tblCustTtl, "{txtTitle}", Title), "{AddPath}", "../")%>
				</td>
			</tr>
			</table>
			</td>
		</tr>
		<% End If %>
		<tr class="CanastaTblResaltada">
			<td width="49%" colspan="2">
			<p align="center"><%=getmessageReturnLngStr("DtxtDate")%></p>
			</td>
			<td width="50%">
			<p align="center"><%=getmessageReturnLngStr("DtxtTo")%></p>
			</td>
		</tr>
		<tr class="CanastaTbl">
			<td width="12%">&nbsp;<img border="0" src="images/mail_icon_urgent.gif" width="13" height="12">
			<input type="checkbox" name="C1" value="ON" style="background:background-image; border-width:0px"></td>
			<td>
			<p align="center"><%=FormatDate(msgDate, True)%></p>
			</td>
			<td>
			<p align="center"><% If Request("cmd") <> "newMsg" Then %><%=rs("olkUToName")%><% Else %>
			<input type="text" name="AgentsTo" style="width: 100%" readonly onclick="javascript:getAgents();"><% End If %></p>
			</td>
		</tr>
		<tr class="CanastaTblResaltada">
			<td colspan="3">
			<%=getmessageReturnLngStr("DtxtSubject")%>:
			</td>
		</tr>
		<tr class="CanastaTbl">
			<td colspan="3">
			<p align="center"><font size="1" face="Verdana">
			<input type="text" name="Subject" style="width: 100%" value="<% If Request("cmd") = "reply" Then %>RE: <%=Left(rs("olkSubject"),80)%><% End If %>" onfocus="this.select()" onkeydown="return chkMax(event, this, 80);"></font></p>
			</td>
		</tr>
		<tr class="CanastaTblResaltada">
			<td colspan="3"><%=getmessageReturnLngStr("DtxtMessage")%>:</td>
		</tr>
		<tr>
			<td colspan="3">
			<table border="0" cellpadding="0" width="100%" cellspacing="1">
				<tr class="CanastaTbl">
					<td>
					<p align="center">
					<textarea rows="15" name="Message" class="input" style="width: 100%" cols="20"><% 
		         If Request("cmd") = "reply" Then
		         	response.write VbCrLf
		         	ArrVal = Split(rs("olkmsg"),VbCrLf)
		         	For i = 0 to UBound(ArrVal)
		         		response.write VbCrLf & ">" & ArrVal(i)
		         	next
		         	End If %></textarea></p>
					</td>
				</tr>
				<tr class="CanastaTbl">
					<td>
					<p align="center">
					<input type="button" name="btnSend" value="<%=getmessageReturnLngStr("DtxtSend")%>" class="btnSend" onclick="vbscript:<% If Request("cmd") = "newMsg" Then %>if document.frmMsg.AgentsTo.value = '' Then Alert('<%=getmessageReturnLngStr("LtxtSelAgent")%>'.replace('{0}', '<%=txtAgent%>')) else <% End If %>document.frmMsg.submit">
					&nbsp;
					<input type="button" name="btnCancel" value="<%=getmessageReturnLngStr("DtxtCancel")%>" class="btnClose" onclick="javascript:history.go(-1);"></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	<% If Request("cmd") <> "newMsg" Then %>
	<input type="hidden" name="AgentsTo" value="<% If rs("olkufromType") = "V" Then Response.Write myHTMLEncode(rs("olkUToName")) %>">
	<input type="hidden" name="ClientsTo" value="<% If rs("olkufromType") = "C" Then Response.Write myHTMLEncode(rs("olkufrom")) %>">
	<% End If %> <input type="hidden" name="return" value="Y">
	<input type="hidden" name="addPath" value="<%=Request("addPath")%>">
	<input type="hidden" name="pop" value="Y">
</form>
<% If Request("cmd") = "reply" Then
set rs = nothing 
End If %> <% If setCustTtl and userType = "C" Then %>
<script language="javascript" src="setTltBg.js.asp?custTtlBgL=<%=custTtlBgL%>&amp;custTtlBgM=<%=custTtlBgM%>&amp;AddPath=../"></script>
<script language="javascript">setTtlBg(false);</script>
<% End If %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>