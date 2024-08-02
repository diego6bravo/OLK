<!--#include file="top.asp" -->
<!--#include file="lang/adminUsersLng.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style2 {
	color: #31659C;
}
</style>
</head>

<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<% conn.execute("use [" & Session("OLKDB") & "]") %>
<form method="POST" name="frmAlerts" action="adminSubmit.asp">
	<input type="hidden" name="submitCmd" value="adminUsersLic">
	<table border="0" cellpadding="0" width="100%">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF" size="1">&nbsp;</font><font size="1" face="Verdana" color="#31659C"><%=getadminUsersLngLngStr("LttlUsersLang")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5"><%=getadminUsersLngLngStr("LttlUsersLangNote")%>
			</font></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="100" bgcolor="#E1F3FD" valign="top" class="style1">
					<font face="Verdana" size="1"><strong><%=getadminUsersLngLngStr("LtxtOlkAgents")%></strong></font></td>
					<td>
					<table border="0" cellpadding="0">
						<tr>
							<td bgcolor="#E1F3FD" style="width: 300px" class="style1">
							<p align="center" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminUsersLngLngStr("DtxtUser")%></strong></font></p>
							</td>
							<td align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminUsersLngLngStr("LtxtLang")%></strong></font></td>
						</tr>
						<%
						oCount = 0
						sql = "select T1.SlpCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, T1.SlpName) SlpName, T2.AlertLang " & _
						"from OLKAgentsAccess T0 " & _
						"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
						"left outer join OLKAlertsTo T2 on T2.ToType = 'O' and T2.ToID = T1.SlpCode and T2.AlertType = 'S' and T2.AlertID = 2 " & _
						"where T0.Access <> 'D' " & _
						"order by T1.SlpName"
						set rd = Server.CreateObject("ADODB.RecordSet")
						set rd = conn.execute(sql)
						do while not rd.eof
						oCount = oCount + 1 %>
						<tr>
							<td bgcolor="#F3FBFE" style="width: 300px">
							<font face="Verdana" size="1" color="#4783C5"><%=rd("SlpName")%></font>&nbsp;
							</td>
							<td bgcolor="#F3FBFE">
							<p>
							<select name='cmbLngO<%=rd("SlpCode")%>' class="input">
							<option value=""><%=getadminUsersLngLngStr("LtxtNatLng")%></option>
							<% For i = 0 to UBound(myLanIndex) %>
							<option <% If rd("AlertLang") = myLanIndex(i)(0) Then %>selected<% End If %>="" value="<%=myLanIndex(i)(0)%>">
							<%=myLanIndex(i)(1)%></option>
							<% Next %></select></p>
							</td>
						</tr>
						<% 
						Alter = Not Alter
						rd.movenext
						loop %>
					</table>
					</td>
				</tr>
				<tr <% If AlertID = 6 Then %>style="display: none" <% End If %>="">
					<td width="100" bgcolor="#E1F3FD" valign="top" class="style1">
					<font face="Verdana" size="1"><strong><span class="style2"><%=getadminUsersLngLngStr("LtxtSapUsers")%></span></strong></font><span class="style2"><font face="Verdana" size="1"><strong>
					</strong>
					</font></span></td>
					<td >
					<table border="0" cellpadding="0" id="table8">
						<tr>
							<td bgcolor="#E1F3FD" style="width: 300px" class="style1">
							<p align="center" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminUsersLngLngStr("DtxtUser")%></strong></font></p>
							</td>
							<td bgcolor="#E1F3FD">
							<p align="center" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminUsersLngLngStr("LtxtLang")%></strong></font></p>
							</td>
						</tr>
						<%
					sCount = 0
					sql = "select T0.USERID, IsNull(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_Name', INTERNAL_K, U_Name), USER_CODE) UName, T2.AlertLang " & _
					"from OUSR T0 " & _
					"left outer join OLKAlertsTo T2 on T2.ToType = 'S' and T2.ToID = T0.USERID and T2.AlertType = 'S' and T2.AlertID = 2 " & _
					"where T0.Groups <> 99 and USER_CODE not like 'B1i%' " & _
					"order by IsNull(U_NAME, USER_CODE)"
					set rd = conn.execute(sql)
					do while not rd.eof
					sCount = sCount + 1 %>
						<tr>
							<td bgcolor="#F3FBFE" style="width: 300px">
							<font face="Verdana" size="1" color="#4783C5"><%=rd("UName")%></font>
							</td>
							<td bgcolor="#F3FBFE">
							<p>
							<select name='cmbLngS<%=rd("UserID")%>' class="input">
							<option value=""><%=getadminUsersLngLngStr("LtxtNatLng")%></option>
							<% For i = 0 to UBound(myLanIndex) %>
							<option <% If rd("AlertLang") = myLanIndex(i)(0) Then %>selected<% End If %>="" value="<%=myLanIndex(i)(0)%>">
							<%=myLanIndex(i)(1)%></option>
							<% Next %></select></p>
							</td>
						</tr>
						<% Alter = Not Alter
						rd.movenext
						loop %>
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
					<td width="77">
					<input type="submit" value="<%=getadminUsersLngLngStr("DtxtSave")%>" name="B1" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size: 10px; width: 75; height: 23; font-weight: bold"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</form>
<!--#include file="bottom.asp" -->