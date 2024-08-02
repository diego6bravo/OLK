<!--#include file="top.asp" -->
<!--#include file="lang/adminAlerts.asp" -->

<head>
<% 

If Request("AlertID") = "" Then AlertID = 12 Else AlertID = CInt(Request("AlertID"))

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKAdminGetAlertData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@AlertID") = AlertID
cmd("@EnableBranchs") = GetYN(myApp.EnableBranchs)
set rs = cmd.execute()

%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E1F3FD;
}
.style2 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style3 {
	color: #31659C;
}
</style>
</head>

<form method="post" action="adminAlerts.asp" name="frmGo">
	<input type="hidden" name="AlertID" value="">
</form>
<script language="javascript">
function goID(ID)
{
	if (document.frmAlerts.B1.disabled || 
	    (!document.frmAlerts.B1.disabled /*&& confirm('<%=getadminAlertsLngStr("LtxtConfChangeAlr")%>')*/)
	    )
	{	
		document.frmGo.AlertID.value = ID;
		document.frmGo.submit();
	}
	else
		document.frmAlerts.AlertID.value = <%=AlertID%>;
}
var uField;
function setBranchs(o, t, i) {
uField = o
OpenWin = this.open('adminAlertsBranchs.asp?pop=Y&AlertType=S&AlertID=<%=Request("AlertID")%>&ToType='+t+'&ToID='+i+'&branchIndex='+o.value, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no,width=300,height=400");
}
function acceptBranchs(b)
{
	uField.value = b;
}
</script>
<form method="POST" name="frmAlerts" action="adminSubmit.asp" onsubmit="return valFrm();">
	<input type="hidden" name="submitCmd" value="adminAlert">
	<table border="0" cellpadding="0" width="100%" id="table3">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF" size="1">&nbsp;</font><font size="1" face="Verdana" color="#31659C"><%=getadminAlertsLngStr("LttlAlertDef")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5"><%=getadminAlertsLngStr("LttlAlertDefNot")%>
			</font></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table4">
				<tr>
					<td width="100" bgcolor="#E1F3FD" class="style2">
					<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("LtxtAlert")%></strong></font></td>
					<td  align="center">
					<table border="0" cellpadding="0" width="100%" id="table6">
						<tr>
							<td class="style1">
							<p align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>">
							<%
							formatAlert1 = Replace(getadminAlertsLngStr("LformatAlert1"), "{0}", getadminAlertsLngStr("DtxtNew"))
							formatAlert2 = Replace(getadminAlertsLngStr("LformatAlert2"), "{0}", getadminAlertsLngStr("DtxtNew2"))
							%>
							<select size="1" name="AlertID" onchange="javascript:goID(this.value);" class="input">
							<optgroup label="<%=getadminAlertsLngStr("LtxtGeneral")%>">
								<option <% If 12 = AlertID Then %>selected<% End If %> value="12"><%=Replace(formatAlert1, "{1}", getadminAlertsLngStr("DtxtActivity"))%></option>							
								<option <% If 5 = AlertID Then %>selected<% End If %> value="5"><%=Replace(formatAlert1, "{1}", getadminAlertsLngStr("DtxtClient"))%></option>
								<option <% If 4 = AlertID Then %>selected<% End If %> value="4"><%=Replace(formatAlert1, "{1}", getadminAlertsLngStr("DtxtItem"))%></option>
							</optgroup>
							<optgroup label="<%=getadminAlertsLngStr("LtxtPur")%>">
								<option <% If 11 = AlertID Then %>selected<% End If %> value="11"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtPurOrder"))%></option>
							</optgroup>
							<optgroup label="<%=getadminAlertsLngStr("LtxtSale")%>">
								<option <% If 0 = AlertID Then %>selected<% End If %> value="0"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtQuote"))%></option>
								<option <% If 2 = AlertID Then %>selected<% End If %> value="2"><%
									If Session("rtl") = "" Then
										Response.Write Replace(formatAlert1, "{1}", getadminAlertsLngStr("DtxtSalesOrder"))
									Else
										Response.Write Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtSalesOrder"))
									End If %></option>
								<option <% If 7 = AlertID Then %>selected<% End If %> value="7"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtDelivery"))%></option>
								<option <% If 9 = AlertID Then %>selected<% End If %> value="9"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtARDownPayReq"))%></option>
								<option <% If 10 = AlertID Then %>selected<% End If %> value="10"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtARDownPayInv"))%></option>
								<option <% If 1 = AlertID Then %>selected<% End If %> value="1"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtInvoice"))%></option>
								<option <% If 8 = AlertID Then %>selected<% End If %> value="8"><%=Replace(formatAlert2, "{1}", getadminAlertsLngStr("DtxtInvoice") & " (" & getadminAlertsLngStr("DtxtReserved") & ")" )%></option>
							</optgroup>
							<optgroup label="<%=getadminAlertsLngStr("LtxtBanks")%>">
								<option <% If 3 = AlertID Then %>selected<% End If %> value="3"><%=Replace(formatAlert1, "{1}", getadminAlertsLngStr("DtxtReceipt"))%></option>
							</optgroup>
							<optgroup label="<%=getadminAlertsLngStr("DtxtOLK")%>">
								<option <% If 6 = AlertID Then %>selected<% End If %> value="6"><%=getadminAlertsLngStr("LtxtOffers")%></option>
							</optgroup>
							</select></p>
							</td>
							<td width="10" class="style1">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td width="10" class="style1">
									<input <% If rs("AsignedSLP") = "N" Then %>disabled<% End If %> type="checkbox" id="asigned" name="asigned" <% If rs("asigned") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
									<td class="style1" style="width: 150px">
									<nobr>
									<label for="asigned">
									<font face="Verdana" size="1" color="#4783C5"><%=getadminAlertsLngStr("LtxtAsignAgent")%></font></label></nobr></td>
								</tr>
							</table></td>
							<td width="10" class="style1" style="width: 50px">
							<table cellpadding="0" cellspacing="0" style="width: 100%">
								<tr>
									<td width="10" class="style1">
									<input type="checkbox" name="Status" <% If rs("Status") = "A" Then %>checked<% End If %> id="Status" value="A" class="noborder"></td>
									<td width="40" class="style1">
									<font face="Verdana" size="1" color="#4783C5">
									<label for="Status"><%=getadminAlertsLngStr("DtxtActive")%></label></font></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#E1F3FD" valign="top" class="style2">
					<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("LtxtOlkAgents")%></strong></font></td>
					<td>
					<table border="0" cellpadding="0">
						<tr>
							<td bgcolor="#E1F3FD" style="width: 300px" class="style2">
							<p align="center" class="style3">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("LtxtNotifyTo")%></strong></font></p>
							</td>
							<td style="width: 2px;">&nbsp;</td>
							<% If rs("branch") = "Y" Then %>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtBranches")%></strong></font></td>
							<% End If %>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtInternal")%></strong></font></td>
							<% If 1 = 2 Then %>
							<td width="120" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtEMail")%></strong></font></td>
							<% End If %>
							<% If rs("draft") = "Y" Then %>
							<td style="width: 2px;">&nbsp;</td>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtRegular")%></strong></font></td>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtDraft")%></strong></font></td><% End If %>
						</tr>
						<%
					oCount = 0
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKAdminGetAlertTo" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@AlertID") = AlertID
					cmd("@ToType") = "O"
					cmd("@LanID") = Session("LanID")
					
					set rd = Server.CreateObject("ADODB.RecordSet")
					set rd = cmd.execute()
					
					do while not rd.eof
					oCount = oCount + 1 %>
						<tr>
							<td bgcolor="#F3FBFE" style="width: 300px">
							<font face="Verdana" size="1" color="#4783C5"><%=rd("SlpName")%></font>&nbsp;
							</td>
							<td style="width: 2px;">&nbsp;</td>
							<% If rs("branch") = "Y" Then %>
							<td bgcolor="#F3FBFE">
							<p align="center">
							<input type="button" value="..." name="B2" style="border: 1px solid #68A6C0; color: #3F7B96; font-family: Verdana; font-size: 10px; font-weight: bold; padding: 0; background-color: #D9F0FD" onclick="javascript:setBranchs(document.frmAlerts.branchO<%=rd("SlpCode")%>,'O', <%=rd("SlpCode")%>);"></p>
							</td><input type="hidden" name="branchO<%=rd("SlpCode")%>" value="<%=rd("selectedBranchs")%>">
							<% End If %>
							<td align="center" bgcolor="#F3FBFE">
							<input type="checkbox" id="chkIntrnl" name="chkIntrnlO<%=rd("SlpCode")%>" <% If rd("SendIntrnl") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<td align="center" style="display: none;" bgcolor="#F3FBFE">
							<input <% If Not rd("VerfyMail") Then %>disabled<% End If %> type="checkbox" id="chkMail" name="chkMailO<%=rd("SlpCode")%>" <% If rd("SendEMail") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<% If rs("draft") = "Y" Then %>
							<td style="width: 2px;">&nbsp;</td>
							<td align="center" bgcolor="#F3FBFE">
							<input type="checkbox" id="chkReg" name="chkRegO<%=rd("SlpCode")%>" <% If rd("AlertReg") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<td align="center" bgcolor="#F3FBFE">
							<input type="checkbox" id="chkDraft" name="chkDraftO<%=rd("SlpCode")%>" <% If rd("AlertDraft") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td><% End If %>
						</tr>
						<% 
						rd.movenext
						loop %>
					</table>
					</td>
				</tr>
				<tr <% If AlertID = 6 Then %>style="display: none" <% End If %>>
					<td width="100" bgcolor="#E1F3FD" valign="top" class="style2">
					<font face="Verdana" size="1"><strong><span class="style3"><%=getadminAlertsLngStr("LtxtSapUsers")%></span></strong></font><span class="style3"><font face="Verdana" size="1"><strong>
					</strong>
					</font></span></td>
					<td>
					<table border="0" cellpadding="0">
						<tr>
							<td bgcolor="#E1F3FD" style="width: 300px" class="style2">
							<p align="center" class="style3">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("LtxtNotifyTo")%></strong></font></p>
							</td>
							<td style="width: 2px;">&nbsp;</td>
							<% If rs("branch") = "Y" Then %>
							<td bgcolor="#E1F3FD" width="60" class="style2">
							<p align="center" class="style3">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtBranches")%></strong></font></p>
							</td>
							<% End If %>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtInternal")%></strong></font></td>
							<td width="120" align="center" bgcolor="#E1F3FD" class="style3">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtEMail")%></strong></font></td>
							<td bgcolor="#E1F3FD" width="60">
							<p align="center" class="style3">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("LtxtSMS")%></strong></font></p>
							</td>
							<td bgcolor="#E1F3FD" width="60">
							<p align="center" class="style3">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("LtxtFax")%></strong></font></p>
							</td>
							<% If rs("draft") = "Y" Then %>
							<td style="width: 2px;">&nbsp;</td>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtRegular")%></strong></font></td>
							<td width="60" align="center" bgcolor="#E1F3FD" class="style2">
							<font face="Verdana" size="1"><strong><%=getadminAlertsLngStr("DtxtDraft")%></strong></font></td><% End If %>
						</tr>
						<%
					sCount = 0
					
					cmd("@ToType") = "S"
					set rd = cmd.execute()
					do while not rd.eof
					sCount = sCount + 1 %>
						<tr>
							<td bgcolor="#F3FBFE" style="width: 300px">
							<font face="Verdana" size="1" color="#4783C5"><%=rd("UName")%></font>
							</td>
							<td style="width: 2px;">&nbsp;</td>
							<% If rs("branch") = "Y" Then %>
							<td bgcolor="#F3FBFE">
							<p align="center">
							<input <% If rs("branch") = "N" Then %>disabled<% End If %> type="button" value="..." name="B3" style="border: 1px solid #68A6C0; color: #3F7B96; font-family: Verdana; font-size: 10px; font-weight: bold; padding: 0; background-color: #D9F0FD" onclick="javascript:setBranchs(document.frmAlerts.branchS<%=rd("UserID")%>, 'S', <%=rd("UserID")%>);"></p>
							</td>
							<input type="hidden" name="branchS<%=rd("UserID")%>" value="<%=rd("selectedBranchs")%>">
							<% End If %>
							<td align="center" bgcolor="#F3FBFE">
							<input type="checkbox" <% If rd("SendIntrnl") = "Y" Then %>checked<% End If %> id="chkIntrnl" name="chkIntrnlS<%=rd("UserID")%>" value="Y" <% If AlertID = 6 Then %>disabled<% End If %> class="noborder"></td>
							<td align="center" bgcolor="#F3FBFE">
							<input <% If Not rd("VerfyMail") Then %>disabled<% End If %> type="checkbox" id="chkMail" name="chkMailS<%=rd("UserID")%>" <% If rd("SendEMail") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<td bgcolor="#F3FBFE">
							<p align="center">
							<input <% If Not rd("VerfySMS") Then %>disabled<% End If %> type="checkbox" id="chkSMS" name="chkSMSS<%=rd("UserID")%>" <% If rd("SendSMS") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></p>
							</td>
							<td bgcolor="#F3FBFE">
							<p align="center">
							<input <% If Not rd("VerfyFAX") Then %>disabled<% End If %> type="checkbox" id="chkFax" name="chkFaxS<%=rd("UserID")%>" <% If rd("SendFax") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></p>
							</td>
							<% If rs("draft") = "Y" Then %>
							<td style="width: 2px;">&nbsp;</td>
							<td align="center" bgcolor="#F3FBFE">
							<input type="checkbox" id="chkReg" name="chkRegS<%=rd("USERID")%>" <% If rd("AlertReg") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td>
							<td align="center" bgcolor="#F3FBFE">
							<input type="checkbox" id="chkDraft" name="chkDraftS<%=rd("USERID")%>" <% If rd("AlertDraft") = "Y" Then %>checked<% End If %> value="Y" class="noborder"></td><% End If %>
						</tr>
						<% rd.movenext
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
					<input type="submit" value="<%=getadminAlertsLngStr("DtxtSave")%>" name="B1" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size: 10px; width: 75; height: 23; font-weight: bold"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</form>
<script language="javascript">
<!--
function valFrm()
{
	if (document.frmAlerts.Status.checked)
	{
		f = false;
		<% If oCount + sCount > 1 Then %>
		chkIntrnl = document.frmAlerts.chkIntrnl;
		chkMail = document.frmAlerts.chkMail;
		for (var i = 0;i<chkIntrnl.length;i++)
		{
			if (chkIntrnl[i].checked) { f = true; break; }
			else if (chkMail[i].checked) { f = true; break; }
		}
		<% Else %>
		if (document.frmAlerts.chkIntrnl.checked) { f = true; break; }
		else if (document.frmAlerts.chkMail.checked) { f = true; break; }
		<% End If %>
		<% If sCount > 1 Then %>
		chkSMS = document.frmAlerts.chkSMS;
		chkFax = document.frmAlerts.chkFax;
		for (var i = 0;i<chkSMS.length;i++)
		{
			if (chkSMS[i].checked) { f = true; break; }
			else if (chkFax[i].checked) { f = true; break; }
		}
		<% Else %>
		if (document.frmAlerts.chkSMS.checked) { f = true; break; }
		else if (document.frmAlerts.chkFax.checked) { f = true; break; }
		<% End If %>
		if (!f && (!document.frmAlerts.asigned.checked && !document.frmAlerts.asigned.disabled || document.frmAlerts.asigned.disabled))
		{
			alert('<%=getadminAlertsLngStr("LtxtSelDest")%>');
			return false;
		}
	}
	return true;
}
//-->
</script>
<!--#include file="bottom.asp" -->