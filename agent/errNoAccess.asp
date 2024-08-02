<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<!--#include file="lang/errNoAccess.asp" -->
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td>
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
				<tr class="TablasTituloSec">
					<td><%=geterrNoAccessLngStr("LtxtAccessDenied")%>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td class="CanastaTblResaltada" style="width: 100px">
					<nobr><%=geterrNoAccessLngStr("LtxtSection")%>&nbsp;</nobr></td>
					<td class="CanastaTbl">
					<% Select Case Request("type")
						Case "form" %><%=geterrNoAccessLngStr("DtxtForm")%>
					<% End Select %></td>
				</tr>
				<tr>
					<td class="CanastaTblResaltada" valign="top" style="width: 100px">
					<nobr><% Select Case Request("type")
						Case "form" %><%=geterrNoAccessLngStr("DtxtForm")%>
					<% End Select %>&nbsp;</nobr></td>
					<td class="CanastaTbl">
					<% Select Case Request("type")
						Case "form"
						sql = "select IsNull(AlterSecName, SecName) SecName from OLKSections T0 " & _
							"left outer join OLKSectionsAlterNames T1 on T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
							"where T0.SecID = " & Request("secID")
						set rs = conn.execute(sql)
						 %><%=rs("SecName")%>
					<% End Select %>&nbsp;</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	</center>
</div>

<!--#include file="agentBottom.asp"-->