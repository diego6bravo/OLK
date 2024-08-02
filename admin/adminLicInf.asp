<!--#include file="top.asp" -->
<!--#include file="lang/adminLicInf.asp" -->

<head>
<%
ClientLic = "NO"
TotalTrans = 0
UsedTrans = 0
AvlTrans = 0

AgentLic = "NO"
AgentLicTotal = 0
AgentLicUse = 0
AgentLicDisp = 0

MobileLic = "NO"
MobileLicTotal = 0
MobileLicUse = 0
MobileLicDisp = 0

AMLic = "NO"
AMLicTotal = 0
AMLicUse = 0
AMLicDisp = 0

set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")

	with olic
        .LicenceServer = licip
        .LicencePort = licport
        
        isAlive = oLic.IsAlive
        
		if isAlive then 
			strStatus = getadminLicInfLngStr("DtxtActive")
			
			ClientLic = oLic.HasLicence(50)
			TotalTrans = oLic.GetTotalTrans(50)
			AvlTrans = oLic.GetAvlTrans(50)
			UsedTrans = TotalTrans-AvlTrans
			
			AgentLic = oLic.HasLicence(51)
			If AgentLic = "YES" Then
				retVal = oLic.LicAvlUsers(51)
				AgentLicTotal = Split(retVal,"|")(2)
				AgentLicUse = Split(retVal,"|")(3)
				AgentLicDisp = Split(retVal,"|")(4)
			End If
						
			MobileLic = oLic.HasLicence(52)
			If MobileLic = "YES" Then
				retVal = oLic.LicAvlUsers(52)
				MobileLicTotal = Split(retVal,"|")(2)
				MobileLicUse = Split(retVal,"|")(3)
				MobileLicDisp = Split(retVal,"|")(4)
			End If
			
			AMLic = oLic.HasLicence(53)
			If AMLic = "YES" Then
				retVal = oLic.LicAvlUsers(53)
				AMLicTotal = Split(retVal,"|")(2)
				AMLicUse = Split(retVal,"|")(3)
				AMLicDisp = Split(retVal,"|")(4)
			End If
		else
			strStatus = getadminLicInfLngStr("DtxtNotActive")
		end if 
	end with 


%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<table border="0" cellpadding="0" width="98%">
	<tr>
		<td height="15"></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">&nbsp;<%=getadminLicInfLngStr("LttlLicInf")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" color="#4783C5" size="1"><%=getadminLicInfLngStr("LttlLicInfNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="center">
			<table border="0" cellpadding="0" style="width: 480">
				<tr bgcolor="#E1F3FD">
					<td class="style1" style="height: 18px;"><font size="1" face="Verdana"><strong><%=getadminLicInfLngStr("DtxtState2")%></strong></font></td>
					<td style="height: 18px;"><font face="Verdana" color="#4783C5" size="1"><%=strStatus%></font></td>
				</tr>
				<tr>
					<td colspan="2" style="height: 18px;">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2" style="height: 18px;">
					<table style="width: 100%">
						<tr>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD">&nbsp;</td>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><strong><%=getadminLicInfLngStr("LtxtLicense")%></strong></td>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><strong><%=getadminLicInfLngStr("DtxtTotal")%></strong></td>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><strong><%=getadminLicInfLngStr("LtxtUsed")%></strong></td>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><strong><%=getadminLicInfLngStr("DtxtAvl")%></strong></td>
						</tr>
						<tr>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><font face="Verdana" size="1">
							<strong><%=getadminLicInfLngStr("DtxtClient")%></strong></font></td>
							<td align="center"><font face="Verdana" color="#4783C5" size="1"><% Select Case ClientLic
					Case "YES" %><%=getadminLicInfLngStr("DtxtYes")%><%
					Case "NO" %><%=getadminLicInfLngStr("DtxtNo")%><%
					Case "EXP" %><%=getadminLicInfLngStr("LtxtExpired")%><%
					End Select %></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><%=TotalTrans%></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><%=UsedTrans%></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><%=AvlTrans%></font></td>
						</tr>
						<tr>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><font face="Verdana" size="1">
							<strong><%=getadminLicInfLngStr("DtxtAgent")%></strong></font></td>
							<td align="center"><font face="Verdana" color="#4783C5" size="1"><% Select Case AgentLic
					Case "YES" %><%=getadminLicInfLngStr("DtxtYes")%><%
					Case "NO" %><%=getadminLicInfLngStr("DtxtNo")%><%
					Case "EXP" %><%=getadminLicInfLngStr("LtxtExpired")%><%
					End Select %></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtAgentTotal"><%=AgentLicTotal%></span></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtAgentUse"><%=AgentLicUse%></span></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtAgentDisp"><%=AgentLicDisp%></span></font></td>
						</tr>
						<tr>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><font face="Verdana" size="1">
							<strong><%=getadminLicInfLngStr("DtxtPocket")%></strong></font></td>
							<td align="center"><font face="Verdana" color="#4783C5" size="1"><% Select Case MobileLic
					Case "YES" %><%=getadminLicInfLngStr("DtxtYes")%><%
					Case "NO" %><%=getadminLicInfLngStr("DtxtNo")%><%
					Case "EXP" %><%=getadminLicInfLngStr("LtxtExpired")%><%
					End Select %></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtMobileTotal"><%=MobileLicTotal%></span></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtMobileUse"><%=MobileLicUse%></span></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtMobileDisp"><%=MobileLicDisp%></span></font></td>
						</tr>
						<tr>
							<td class="style1" style="height: 18px;" bgcolor="#E1F3FD"><font face="Verdana" size="1">
							<strong><%=getadminLicInfLngStr("DtxtAgent")%> / <%=getadminLicInfLngStr("DtxtPocket")%></strong></font></td>
							<td align="center"><font face="Verdana" color="#4783C5" size="1"><% Select Case AMLic
					Case "YES" %><%=getadminLicInfLngStr("DtxtYes")%><%
					Case "NO" %><%=getadminLicInfLngStr("DtxtNo")%><%
					Case "EXP" %><%=getadminLicInfLngStr("LtxtExpired")%><%
					End Select %></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtAMTotal"><%=AMLicTotal%></span></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtAMUse"><%=AMLicUse%></span></font></td>
							<td align="right"><font face="Verdana" color="#4783C5" size="1"><span id="txtAMDisp"><%=AMLicDisp%></span></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td colspan="2" style="height: 18px;">
					<p align="center">
					<input type="button" value="<%=getadminLicInfLngStr("DtxtRefresh")%>" name="B1" class="OlkBtn" onclick="window.location.reload();"></td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	<% If Session("ID") <> "" or myApp.SingleSignOn Then %>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">&nbsp;<%=getadminLicInfLngStr("LtxtUserLicAsign")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" color="#4783C5" size="1"><%=getadminLicInfLngStr("LtxtUserLicAsignDesc")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<form method="post" action="adminSubmit.asp" name="frmAdminLic">
		<div align="center">
			<table border="0" cellpadding="0" width="500">
				<tr bgcolor="#E1F3FD">
					<td class="style1" style="height: 18px;" align="center"><font face="Verdana" size="1">
					<strong><%=getadminLicInfLngStr("DtxtUser")%></strong></font></td>
					<% If Not myApp.SingleSignOn Then %><td class="style1" style="height: 18px;" align="center"><font face="Verdana" size="1">
					<strong><%=getadminLicInfLngStr("DtxtName")%></strong></font></td><% End If %>
					<td class="style1" style="height: 18px;" align="center"><font face="Verdana" size="1">
					<strong><%=getadminLicInfLngStr("DtxtAgent")%></strong></font></td>
					<td class="style1" style="height: 18px;" align="center"><font face="Verdana" size="1">
					<strong><%=getadminLicInfLngStr("DtxtPocket")%></strong></font></td>
					<td class="style1" style="height: 18px;" align="center"><font face="Verdana" size="1">
					<strong><%=getadminLicInfLngStr("DtxtAgent")%> / <%=getadminLicInfLngStr("DtxtPocket")%></strong></font></td>
				</tr>
				<% 
				set rs = Server.CreateObject("ADODB.RecordSet")
				If Not myApp.SingleSignOn Then
					sql = "select T0.UserName, T1.SlpName, T0.SlpCode from OLKAgentsAccess T0 inner join OSLP T1 on T1.SlpCode = T0.SlpCode where T0.Access <> 'D'"
					rs.open sql, conn, 3, 1
				Else
					sql = "select T0.UserName from OLKAgentsAccess T0 where Status = 'Y'"
					rs.open sql, connCommon, 3, 1
				End If
				do while not rs.eof
				If IsAlive Then
					ChkAgent = oLic.ConfHasLic(51, 0, rs(0))
					ChkMobile = oLic.ConfHasLic(52, 0, rs(0))
					ChkAM = oLic.ConfHasLic(53, 0, rs(0))
				End If %>
				<tr bgcolor="#E1F3FD">
					<td><font face="Verdana" color="#4783C5" size="1"><%=Server.HTMLEncode(rs(0))%></font><input type="hidden" name="ID" value="<%=rs.bookmark%>"><input type="hidden" name="UserName<%=rs.bookmark%>" value="<%=myHTMLEncode(rs(0))%>"></td>
					<% If Not myApp.SingleSignOn Then %><td><font face="Verdana" color="#4783C5" size="1"><%=Server.HTMLEncode(rs(1))%></font></td><% End If %>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<p align="center"><input class="OptionButton" style="background:background-image" <% If ChkAgent Then %>checked<% End If %> <% If Not IsAlive or AgentLic <> "YES" Then %>disabled<% End If %> type="checkbox" name="ChkAgent<%=rs.bookmark%>" value="Y" onclick="doCheck(this, event, 'A');"></td>					
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<p align="center"><input class="OptionButton" style="background:background-image" <% If ChkMobile Then %>checked<% End If %> <% If Not IsAlive or MobileLic <> "YES" Then %>disabled<% End If %> type="checkbox" name="ChkMobile<%=rs.bookmark%>" value="Y" onclick="doCheck(this, event, 'M');"></td>					
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<p align="center"><input class="OptionButton" style="background:background-image" <% If ChkAM Then %>checked<% End If %> <% If Not IsAlive or AMLic <> "YES" Then %>disabled<% End If %> type="checkbox" name="ChkAM<%=rs.bookmark%>" value="Y" onclick="doCheck(this, event, 'AM');"></td>
				</tr>
				<% rs.movenext
				loop
				rs.close %>
				<tr bgcolor="#E1F3FD">
					<td colspan="<% If Not myApp.SingleSignOn Then %>5<% Else %>4<% End If %>">
					<p align="center">
					<input type="submit" value="<%=getadminLicInfLngStr("DtxtSave")%>" name="btnSave" <% If Not IsAlive Then %>disabled<% End If %> ></td>
				</tr>
			</table>
		</div>
		<input type="hidden" name="submitCmd" value="setLic">
		</form>
		</td>
	</tr>
	<% End If %>
</table>
<% If Session("ID") <> "" Then %>
<script language="javascript">
function doCheck(o, e, t)
{

	var txtTotal;
	var txtDisp;
	var txtUse;
	
	switch (t)
	{
		case 'A':
			txtTotal = txtAgentTotal;
			txtDisp = txtAgentDisp;
			txtUse = txtAgentUse;
			break;
		case 'M':
			txtTotal = txtMobileTotal;
			txtDisp = txtMobileDisp;
			txtUse = txtMobileUse;
			break;
		case 'AM':
			txtTotal = txtAMTotal;
			txtDisp = txtAMDisp;
			txtUse = txtAMUse;
			break;
	}
	var disp = parseInt(txtDisp.innerHTML);
	if (o.checked && disp == 0)
	{
		alert('<%=getadminLicInfLngStr("LtxtNoMoreLic")%>');
		o.checked = false;
	}
	else
	{
		var total = parseInt(txtTotal.innerHTML);
		
		if (!o.checked) disp++; else disp--;
		txtDisp.innerHTML = disp;
		txtUse.innerHTML = total-disp;
	}
}
</script>
<% End If %>
<!--#include file="bottom.asp" -->