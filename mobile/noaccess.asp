<!--#include file="lang/noaccess.asp" -->
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
		<tr>
			<td>
			<img src="images/spacer.gif" width="100%" height="1" border="0" alt=""></td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
				<tr>
					<td width="100%">
					<p align='<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>'>
					<b><font face="Verdana" size="1"><%=getnoaccessLngStr("LtxtAccessDenied")%> </font>
					</b></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
			<table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" id="AutoNumber2">
				<tr>
					<td bgcolor="#66A4FF">
					<font face="verdana" color="#000000" size="1"><%=getnoaccessLngStr("LtxtSection")%></font></td>
					<td bgcolor="#9BC4FF">
					<font face="verdana" color="#000000" size="1">
					<% Select Case Request("type")
						Case "form" %><%=getnoaccessLngStr("DtxtForm")%>
					<% End Select %></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF" valign="top">
					<font face="verdana" color="#000000" size="1"><% Select Case Request("type")
						Case "form" %><%=getnoaccessLngStr("DtxtForm")%>
					<% End Select %></font></td>
					<td bgcolor="#9BC4FF">
					<font face="verdana" color="#000000" size="1"><% Select Case Request("type")
						Case "form"
						sql = "select IsNull(AlterSecName, SecName) SecName from OLKSections T0 " & _
							"left outer join OLKSectionsAlterNames T1 on T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
							"where T0.SecID = " & Request("secID")
						set rs = conn.execute(sql)
						 %><%=rs("SecName")%>
					<% End Select %>&nbsp;</font></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	</center>
</div>
