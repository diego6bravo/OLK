<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<!--#include file="lang/crcerror.asp" -->
<%
LogNum = Session("RetryRetVal")

sql = "select Object from R3_ObsCommon..TLOG where LogNum = " & LogNum
set rs = conn.execute(sql)
obj = rs("Object")


Select Case obj
	Case 2
		sql = "select CardCode from R3_ObsCommon..TCRD where LogNum = " & LogNum
		set rs = conn.execute(sql)
		CardCode = rs("CardCode")
		action = "addCard/goEditCard.asp"
	Case 33
		sql = "select ClgCode, CardCode from R3_ObsCommon..TCLG where LogNum = " & LogNum
		set rs = conn.execute(sql)
		ClgCode = rs("ClgCode")
		CardCode = rs("CardCode")
		action = "addActivity/goEditActivity.asp"
	Case 97
		sql = "select OpprId, CardCode from R3_ObsCommon..TOPR where LogNum = " & LogNum
		set rs = conn.execute(sql)
		OpprId = rs("OpprId")
		CardCode = rs("CardCode")
		action = "addSO/goEditSO.asp"
End Select

%>
<div align="center">
	<table border="0" cellspacing="0" cellpadding="0" width="435">
		<tr>
			<td height="182" background="images/error_olCOnfig.gif" valign="top">
			<table border="0" cellspacing="0" width="100%">
				<form name="frmReOpen" action="<%=action%>" method="post">
				<input type="hidden" name="CardCode" value="<%=myHTMLEncode(CardCode)%>">
				<input type="hidden" name="ClgCode" value="<%=ClgCode%>">
				<input type="hidden" name="ID" value="<%=OpprId%>">
				<tr>
					<td height="24">
					<p align="center"><b>
					<font face="Verdana" size="2" color="#0066CC">
					<% Select Case obj
					Case 2 %><%=getcrcerrorLngStr("DtxtClient")%> "<%=CardCode%>"
				<%	Case 33 %><%=getcrcerrorLngStr("DtxtActivity")%> #<%=ClgCode%>
				<%	Case 97 %><%=getcrcerrorLngStr("DtxtSO")%> #<%=OpprId%><%
					End Select %></font></b></td>
				</tr>
				<tr>
					<td height="24">
					<p align="center">
					<font face="Verdana" size="1" color="#0066CC">
					<%=getcrcerrorLngStr("LtxtCRCError")%></font></td>
				</tr>
				<tr>
					<td height="90">&nbsp;</td>
				</tr>
				<tr>
					<td align="center">
					<input type="submit" class="input" name="btnReOpen" value="<%=getcrcerrorLngStr("DtxtEditNewData")%>">
					</td>
				</tr>
				</form>
			</table>
			</td>
		</tr>
	</table>
</div>
<!--#include file="agentBottom.asp"-->