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
		action = "goCrdEdit.asp"
	Case 33
		sql = "select ClgCode, CardCode from R3_ObsCommon..TCLG where LogNum = " & LogNum
		set rs = conn.execute(sql)
		ClgCode = rs("ClgCode")
		CardCode = rs("CardCode")
		action = "goActEdit.asp"
End Select
%>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
				<form name="frmReOpen" action="<%=action%>" method="post">
				<input type="hidden" name="CardCode" value="<%=myHTMLEncode(CardCode)%>">
				<input type="hidden" name="ClgCode" value="<%=ClgCode%>">
				<tr>
					<td width="100%" bgcolor="#75ACFF">
					<p align="center"><b><font face="Verdana" size="1"><% Select Case obj
					Case 2 %><%=getcrcerrorLngStr("DtxtClient")%> "<%=CardCode%>"
				<%	Case 33 %><%=getcrcerrorLngStr("DtxtActivity")%> #<%=ClgCode%><%
					End Select %></font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" bgcolor="#75ACFF">
					<p align="center"><font face="Verdana" size="1"><%=getcrcerrorLngStr("LtxtCRCError")%></font></p>
					</td>
				</tr>
				<tr>
					<td width="100%" align="center"><img src="images/errorIcon.gif"> </td>
				</tr>
				<tr>
					<td align="center">
					<input type="submit" name="btnReOpen" value="<%=getcrcerrorLngStr("DtxtEditNewData")%>">
					</td>
				</tr>
				<tr>
					<td align="center">
					<input type="button" name="btnGoHome" value="<%=getcrcerrorLngStr("DtxtHome")%>" onclick="window.location.href='?cmd=home';">
					</td>
				</tr>
				</form>
			</table>
			</td>
		</tr>
	</table>
</center></div>