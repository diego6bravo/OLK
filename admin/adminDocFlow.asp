<!--#include file="top.asp" -->
<!--#include file="lang/adminDocFlow.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<% conn.execute("use [" & Session("OLKDB") & "]") %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="js_up_down.js"></script>
<script language="javascript">
function delFlow(FlowID)
{
	if(confirm('<%=getadminDocFlowLngStr("LtxtConfDelFlow")%>'))
		window.location.href='adminSubmit.asp?submitCmd=delDocFlow&FlowID=' + FlowID;
}
var txtConfDel = '<%=getadminDocFlowLngStr("LtxtConfDel")%>';
var maxOrdr = 0;
var txtAsignedSLP = '<%=getadminDocFlowLngStr("LtxtAsignedSLP")%>';
var txtAlertFilter = '<%=getadminDocFlowLngStr("LtxtAlertFilter")%>';
</script>
<style type="text/css">
.style1 {
	background-color: #F3FBFE;
}
.style2 {
	color: #31659C;
}
</style>
</head>
<% sql = "select T0.ID, T1.Name + ' - ' + T0.Name Name " & _  
		"from OLKOps T0  " & _  
		"inner join OLKOpsGrps T1 on T1.ID = T0.GroupID " & _  
		"where T0.Status <> 'D' " 
set ro = Server.CreateObject("ADODB.RecordSet")
ro.open sql, conn, 3, 1
%>
<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("NewFlow") <> "Y" and Request("FlowID") = "" Then %>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminDocFlowLngStr("LttlDocFlow")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminDocFlowLngStr("LttlDocFlowNote")%></font></td>
	</tr>
	<form method="post" action="adminSubmit.asp" name="frmDocFlow">
	<tr>
		<td width="100%">
		<table border="0" cellpadding="0" width="100%" id="table6">
			<tr>
				<td colspan="2">
				<table border="0" cellpadding="0" id="table12" style="width: 100%">
					<tr>
						<td align="center" bgcolor="#E2F3FC" width="10">&nbsp;</td>
						<td align="center" bgcolor="#E2F3FC"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminDocFlowLngStr("DtxtType")%></b>&nbsp;</font></td>
						<td align="center" bgcolor="#E2F3FC"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminDocFlowLngStr("LtxtTitle")%></b>&nbsp;</font></td>
						<td align="center" bgcolor="#E2F3FC"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminDocFlowLngStr("LtxtExecution")%></b>&nbsp;</font></td>
						<td align="center" bgcolor="#E2F3FC"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminDocFlowLngStr("DtxtOrder")%></b>&nbsp;</font></td>
						<td align="center" bgcolor="#E2F3FC"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminDocFlowLngStr("DtxtActive")%></b>&nbsp;</font></td>
						<td align="center" bgcolor="#E2F3FC" style="width: 16px"><font size="1">&nbsp;</font></td>
					</tr>

					<% 
					sql = "select FlowID, Name,  " & _
					"Case Type When 1 Then N'" & getadminDocFlowLngStr("DtxtConfirm") & "' When 0 Then N'" & getadminDocFlowLngStr("DtxtError") & "' When 2 Then N'" & getadminDocFlowLngStr("DtxtFlow") & "' End Type, [Order], " & _
					"Active, " & _
					"Case ExecAt When 'D1' Then N'" & getadminDocFlowLngStr("DtxtComDocs") & " - " & getadminDocFlowLngStr("LtxtDocCreation") & "' " & _
					"When 'D2' Then N'" & getadminDocFlowLngStr("DtxtComDocs") & " - " & Replace(getadminDocFlowLngStr("LtxtAddItem"), "'", "''") & "' " & _
					"When 'D3' Then N'" & getadminDocFlowLngStr("DtxtComDocs") & " - " & getadminDocFlowLngStr("DtxtAdd") & "/" & getadminDocFlowLngStr("LtxtDocConf") & "' " & _
					"When 'R1' Then N'" & getadminDocFlowLngStr("DtxtReceipts") & " - " & getadminDocFlowLngStr("LtxtCreation") & "' " & _
					"When 'R2' Then N'" & getadminDocFlowLngStr("DtxtReceipts") & " - " & getadminDocFlowLngStr("DtxtAdd") & "/" & getadminDocFlowLngStr("LtxtRcpConf") & "' " & _
					"When 'A1' Then N'" & getadminDocFlowLngStr("DtxtItem") & " - " & getadminDocFlowLngStr("DtxtAdd") & "/" & Replace(getadminDocFlowLngStr("LtxtItmConf"), "'", "''") & "' " & _
					"When 'C1' Then N'" & getadminDocFlowLngStr("DtxtClient") & " - " & getadminDocFlowLngStr("DtxtAdd") & "/" & getadminDocFlowLngStr("LtxtClientConf") & "' " & _
					"When 'C2' Then N'" & getadminDocFlowLngStr("DtxtClient") & " - " & getadminDocFlowLngStr("DtxtAdd") & "/" & getadminDocFlowLngStr("LtxtActivityConf") & "' " & _
					"When 'C3' Then N'" & getadminDocFlowLngStr("DtxtClient") & " - " & getadminDocFlowLngStr("DtxtAdd") & "/" & getadminDocFlowLngStr("LtxtSOConf") & "' " & _
					"When 'O1' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtConvQuoteOrder") & "'  " & _
					"When 'O10' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtConvOrderDel") & "'  " & _
					"When 'O7' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtConvOrderInv") & "'  " & _
					"When 'O11' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtAprovDraft") & "'  " & _
					"When 'O0' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtAprovOrder") & "'  " & _
					"When 'O8' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtAprovPurOrdr") & "'  " & _
					"When 'O2' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtCloseObj") & "' " & _
					"When 'O3' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtCancelObj") & "'  " & _
					"When 'O4' Then N'" & getadminDocFlowLngStr("LtxtAction") & " - " & getadminDocFlowLngStr("LtxtRemObj") & "' "
					
					do while not ro.eof
						sql = sql & "When 'OP" & ro("ID") & "' Then N'" & ro("Name") & "' "
					ro.movenext
					loop
					
					sql = sql & "End ExecAt " & _
					"from OLKUAF T0 " & _
					"where Active <> 'D' " & _
					"order by T0.Type, ExecAt,  [Order]"
					set rs = conn.execute(sql)
					do while not rs.eof
					FlowID = Replace(rs("FlowID"), "-", "_") %>
					<input type="hidden" name="FlowID" value="<%=rs("FlowID")%>">
					<tr>
						<td width="10" bgcolor="#F3FBFE">
						<a href="adminDocFlow.asp?FlowID=<%=rs("FlowID")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
						<td align="center" bgcolor="#F3FBFE"><font size="1" face="Verdana" color="#4783C5"><%=rs("Type")%></font></td>
						<td bgcolor="#F3FBFE">
						
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><font size="1" face="Verdana" color="#4783C5"><%=rs("Name")%></font>
								</td>
								<td width="16"><a href="javascript:doFldTrad('UAF', 'FlowID', <%=rs("FlowID")%>, 'AlterName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminDocFlowLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						
						</td>
						<td bgcolor="#F3FBFE"><font size="1" face="Verdana" color="#4783C5"><%=rs("ExecAt")%></font></td>
						<td bgcolor="#F3FBFE" align="center">
						<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><input type="text" name="Order<%=rs("FlowID")%>" id="Order<%=FlowID%>" size="5" style="text-align:right" class="input" value="<%=rs("Order")%>" onfocus="this.select()" onchange="chkThis(this, 0, 0)" onkeydown="return chkMax(event, this, 6);"></td>
									<td valign="middle">
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><img src="images/img_nud_up.gif" id="btnOrder<%=FlowID%>Up"></td>
										</tr>
										<tr>
											<td><img src="images/spacer.gif"></td>
										</tr>
										<tr>
											<td><img src="images/img_nud_down.gif" id="btnOrder<%=FlowID%>Down"></td>
										</tr>
									</table>
									</td>
								</tr>
							</table>
						<script language="javascript">NumUDAttach('frmDocFlow', 'Order<%=FlowID%>', 'btnOrder<%=FlowID%>Up', 'btnOrder<%=FlowID%>Down');</script>
						</td>
						<td align="center" bgcolor="#F3FBFE">
						<input type="checkbox" class="noborder" name="Active<%=rs("FlowID")%>" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %>></td>
						<td bgcolor="#F3FBFE" style="width: 16px"><% If rs("FlowID") > -1 Then %><a href="javascript:delFlow(<%=rs("FlowID")%>);"><img border="0" src="images/remove.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
					</tr>
					<% rs.movenext
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
				<input type="submit" value="<%=getadminDocFlowLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminDocFlowLngStr("DtxtNew")%>" name="B2" class="OlkBtn" onclick="javascript:window.location.href='adminDocFlow.asp?NewFlow=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="admDocFlow">
	</form>
	<% Else %>
	<form method="POST" action="adminsubmit.asp" name="frmAddEditFlow" onsubmit="javascript:return valFrm2()">
	<% If Request("NewFlow") = "Y" Then %>
	<input type="hidden" name="FlowNameTrad">
	<input type="hidden" name="NoteTextTrad">
	<input type="hidden" name="FlowQueryDef">
	<input type="hidden" name="NoteQueryDef">
	<input type="hidden" name="LineQueryDef">
	<% End If %>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("FlowID") <> "" Then %><%=getadminDocFlowLngStr("LtxtEditFlow")%><% Else %><%=getadminDocFlowLngStr("LtxtAddFlow")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminDocFlowLngStr("LtxtFlowNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table20">
		<% ExecAt = "  "
		If Request("FlowID") <> "" Then
			sql = "select T0.*, " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 23) Then 'Y' Else 'N' End 'OQUT', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 17) Then 'Y' Else 'N' End 'ORDR', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 13) Then 'Y' Else 'N' End 'OINV', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = -13) Then 'Y' Else 'N' End 'OINVR', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 48) Then 'Y' Else 'N' End 'OIR', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 15) Then 'Y' Else 'N' End 'ODLN', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 203) Then 'Y' Else 'N' End 'ODPIReq', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 204) Then 'Y' Else 'N' End 'ODPIInv', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 22) Then 'Y' Else 'N' End 'OPOR', " & _
			"Case When Exists(select 'A' from OLKUAF2 where FlowID = T0.FlowID and ObjectCode = 540000006) Then 'Y' Else 'N' End 'OPQT', " & _			
			"Case When Exists(select 'A' from OLKUAF1 where FlowID = T0.FlowID and SlpCode = -999) Then 'Y' Else 'N' End AllAgents, " & _
			"Case When Exists(select top 1 '' from OLKUAFControl1 where FlowID = T0.FlowID) Then 'Y' Else 'N' End LockExec " & _
			"from OLKUAF T0 " & _
			"where FlowID = " & Request("FlowID")
			set rs = conn.execute(sql) 
			FlowName = rs("Name")
			FlowType = rs("Type")
			Order = rs("Order")
			Active = rs("Active")
			FlowQuery = rs("Query")
			LineQuery = rs("LineQuery")
			NoteBuilder = rs("NoteBuilder")
			NoteQuery = rs("NoteQuery")
			NoteText = rs("NoteText")
			OQUT = rs("OQUT")
			ORDR = rs("ORDR")
			OINV = rs("OINV")
			OIR = rs("OIR")
			ODLN = rs("ODLN")
			ODPIReq = rs("ODPIReq")
			ODPIInv = rs("ODPIInv")
			OINVR = rs("OINVR")
			OPOR = rs("OPOR")
			OPQT = rs("OPQT")
			ExecAt = rs("ExecAt")
			ApplyToClient = rs("ApplyToClient")
			AllAgents = rs("AllAgents")
			Draft = rs("Draft")
			Authorize = rs("Authorize")
			LockExec = rs("LockExec")
			
			If AllAgents = "N" Then
				sql = "select SlpCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) SlpName from OSLP T0 where exists(select 'A' from OLKUAF1 where FlowID = " & Request("FlowID") & " and SlpCode = T0.SlpCode)"
				set rs = conn.execute(sql)
				SlpCode = ""
				Agents = ""
				do while not rs.eof
					If SlpCode <> "" Then
						SlpCode = SlpCode & ", "
						Agents = Agents & ", "
					End If
					SlpCode = SlpCode & rs("SlpCode")
					Agents = Agents & rs("SlpName")
				rs.movenext
				loop
			Else
				SlpCode = "-999"
				Agents = getadminDocFlowLngStr("DtxtAll")
				If Agents = "" Then Agents = "Todos"
			End If
		Else
			LockExec = False
			Order = 1
			NoteText = ""
			FlowName = ""
		End If
		%>
			<tr>
				<td colspan="3">
					<table cellpadding="0" width="100%" border="0">
						<tr>
							<td align="center" bgcolor="#E2F3FC" width="100"><b>
							<font face="Verdana" size="1" color="#31659C"><%=getadminDocFlowLngStr("DtxtType")%></font></b></td>
							<td align="center" bgcolor="#E2F3FC" width="100"><b>
							<font face="Verdana" size="1" color="#31659C"><%=getadminDocFlowLngStr("DtxtOrder")%></font></b></td>
							<td align="center" bgcolor="#E2F3FC" width="350"><b>
							<font size="1" face="Verdana" color="#31659C"><%=getadminDocFlowLngStr("LtxtTitle")%></font></b></td>
							<td align="center" bgcolor="#E2F3FC" style="width: 100px">
							<b><font face="Verdana" size="1" color="#31659C">
							<%=getadminDocFlowLngStr("DtxtActive")%></font></b></td>
							<td align="center" bgcolor="#E2F3FC">
							&nbsp;</td>
						</tr>
						<tr>
							<td valign="top" width="100" class="style1">
							<select size="1" name="FlowType" class="input" onchange="changeFlow(this.value);">
							<option <% If FlowType = 1 Then %>selected<% End If %> value="1"><%=getadminDocFlowLngStr("DtxtConfirm")%></option>
							<option <% If FlowType = 0 Then %>selected<% End If %> value="0"><%=getadminDocFlowLngStr("DtxtError")%></option>
							<option <% If FlowType = 2 Then %>selected<% End If %> value="2"><%=getadminDocFlowLngStr("DtxtFlow")%></option>
							</select></td>
							<td valign="top" align="center" width="100" class="style1">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><input type="text" name="Order" id="Order" size="5" style="text-align:right" class="input" value="<%=Order%>" onfocus="this.select()" onchange="chkThis(this, 0, 0)" onkeydown="return chkMax(event, this, 6);"></td>
										<td valign="middle">
										<table cellpadding="0" cellspacing="0" border="0">
											<tr>
												<td><img src="images/img_nud_up.gif" id="btnOrderUp"></td>
											</tr>
											<tr>
												<td><img src="images/spacer.gif"></td>
											</tr>
											<tr>
												<td><img src="images/img_nud_down.gif" id="btnOrderDown"></td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
							<script language="javascript">NumUDAttach('frmAddEditFlow', 'Order', 'btnOrderUp', 'btnOrderDown');</script></td>
							<td valign="top" width="370" class="style1">
							<p>
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td><input name="FlowName" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(FlowName)%>" size="50" onkeydown="return chkMax(event, this, 100);"></td>
									<td style="width: 16px"><a href="javascript:doFldTrad('UAF', 'FlowID', '<%=Request("FlowID")%>', 'AlterName', 'T', <% If Request("NewFlow") <> "Y" Then %>null<% Else %>document.frmAddEditFlow.FlowNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminDocFlowLngStr("DtxtTranslate")%>" border="0"></a></td>
								</tr>
							</table>
							</td>
							<td valign="top" align="center" class="style1" style="width: 100px">
							<input type="checkbox" name="FlowActive" value="Y" <% If Active = "Y" Then %>checked<% End If %> class="noborder"></td>
							<td valign="top" align="center" class="style1">
							&nbsp;</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td valign="top" colspan="3">
				<table border="0" width="100%" cellpadding="0">
					<tr>
						<td bgcolor="#E2F3FC" style="width: 100px">
						<a hrefx="#" href="javascript:selectAgents()">
						<b><font size="1" face="Verdana" color="#31659C"><%=getadminDocFlowLngStr("DtxtAgents")%></font></b></a></td>
						<td class="style1">
						<input type="hidden" name="SlpCode" value="<%=SlpCode%>">
						<input name="Agents" readonly onclick="javascript:selectAgents()"  class="input" value="<%=myHTMLEncode(Agents)%>" size="80">
					</td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 100px">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminDocFlowLngStr("DtxtClients")%></font></b></td>
						<td bgcolor="#F3FBFE">
						<input <% If Left(ExecAt,1) <> "D" and ExecAt <> "C1" and Left(ExecAt, 1) <> "O" Then %>disabled<% End If %> type="checkbox" <% If ApplyToClient = "Y" Then %>checked<% End If %> name="ApplyToClient" value="Y" id="ApplyToClient" onclick="changeApplyToClient()" class="noborder"><label for="ApplyToClient"><font face="Verdana" size="1" color="#31659C"><%=getadminDocFlowLngStr("DtxtApply")%></font></label></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 100px">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminDocFlowLngStr("LtxtExecution")%></font></b></td>
						<td class="style1">
						<select size="1" name="ExecAt" class="input" <% If LockExec = "Y" Then %>disabled<% End If %> onchange="javascript:changeDoc(this.value);">
						<option></option>
						<option <% If ExecAt = "O1" Then %>selected<% End If %> value="O1"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtConvQuoteOrder")%></option>
						<option <% If ExecAt = "O10" Then %>selected<% End If %> value="O10"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtConvOrderDel")%></option>
						<option <% If ExecAt = "O7" Then %>selected<% End If %> value="O7"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtConvOrderInv")%></option>
						<option <% If ExecAt = "O11" Then %>selected<% End If %> value="O11"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtAprovDraft")%></option>
						<option <% If ExecAt = "O8" Then %>selected<% End If %> value="O8"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtAprovPurOrdr")%></option>
						<option <% If ExecAt = "O0" Then %>selected<% End If %> value="O0"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtAprovOrder")%></option>
						<option <% If ExecAt = "O2" Then %>selected<% End If %> value="O2"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtCloseObj")%></option>
						<option <% If ExecAt = "O3" Then %>selected<% End If %> value="O3"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtCancelObj")%></option>
						<option <% If ExecAt = "O4" Then %>selected<% End If %> value="O4"><%=getadminDocFlowLngStr("LtxtAction")%> - <%=getadminDocFlowLngStr("LtxtRemObj")%></option>
						<option <% If ExecAt = "D1" Then %>selected<% End If %> value="D1"><%=getadminDocFlowLngStr("DtxtComDocs")%> - <%=getadminDocFlowLngStr("LtxtDocCreation")%></option>
						<option <% If ExecAt = "D2" Then %>selected<% End If %> value="D2"><%=getadminDocFlowLngStr("DtxtComDocs")%> - <%=getadminDocFlowLngStr("LtxtAddItem")%></option>
						<option <% If ExecAt = "D3" Then %>selected<% End If %> value="D3"><%=getadminDocFlowLngStr("DtxtComDocs")%> - <%=getadminDocFlowLngStr("DtxtAdd")%>/<%=getadminDocFlowLngStr("LtxtDocConf")%></option>
						<option <% If ExecAt = "R1" Then %>selected<% End If %> value="R1"><%=getadminDocFlowLngStr("DtxtReceipts")%> - <%=getadminDocFlowLngStr("LtxtCreation")%></option>
						<option <% If ExecAt = "R2" Then %>selected<% End If %> value="R2"><%=getadminDocFlowLngStr("DtxtReceipts")%> - <%=getadminDocFlowLngStr("DtxtAdd")%>/<%=getadminDocFlowLngStr("LtxtRcpConf")%></option>
						<option <% If ExecAt = "A1" Then %>selected<% End If %> value="A1"><%=getadminDocFlowLngStr("DtxtItem")%> - <%=getadminDocFlowLngStr("DtxtAdd")%>/<%=getadminDocFlowLngStr("LtxtItmConf")%></option>
						<option <% If ExecAt = "C1" Then %>selected<% End If %> value="C1"><%=getadminDocFlowLngStr("DtxtClient")%> - <%=getadminDocFlowLngStr("DtxtAdd")%>/<%=getadminDocFlowLngStr("LtxtClientConf")%></option>
						<option <% If ExecAt = "C2" Then %>selected<% End If %> value="C2"><%=getadminDocFlowLngStr("DtxtClient")%> - <%=getadminDocFlowLngStr("DtxtAdd")%>/<%=getadminDocFlowLngStr("LtxtActivityConf")%></option>
						<option <% If ExecAt = "C3" Then %>selected<% End If %> value="C3"><%=getadminDocFlowLngStr("DtxtClient")%> - <%=getadminDocFlowLngStr("DtxtAdd")%>/<%=getadminDocFlowLngStr("LtxtSOConf")%></option>
						<% do while not ro.eof %><option <% If ExecAt = "OP" & ro("ID") Then %>selected<% End If %> value="<%="OP" & ro("ID")%>"><%=getadminDocFlowLngStr("DtxtOp")%> - <%=ro("Name")%></option><% 
						ro.movenext
						loop %>
						</select></td>
					</tr>
					<% If LockExec = "Y" Then %>
					<tr>
						<td bgcolor="#FFD2A6" colspan="2" align="center">
						<b><font face="Verdana" size="1" color="#666666">
						<%=getadminDocFlowLngStr("LtxtLockedFlow")%></font></b>
					</tr>
					<% End If %>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 100px; vertical-align: top;">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminDocFlowLngStr("DtxtDocs")%></font></b></td>
						<td class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If OQUT = "Y" Then %>checked<% End If %> name="ObjectCode" value="23" id="ObjectCode1" class="noborder"><label for="ObjectCode1"><%=getadminDocFlowLngStr("DtxtQuote")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If ORDR = "Y" Then %>checked<% End If %> name="ObjectCode" value="17" id="ObjectCode2" class="noborder"><label for="ObjectCode2"><%=getadminDocFlowLngStr("DtxtSalesOrder")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If ODLN = "Y" Then %>checked<% End If %> name="ObjectCode" value="15" id="ObjectCode5" class="noborder"><label for="ObjectCode5"><%=getadminDocFlowLngStr("DtxtDelivery")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If ODPIReq = "Y" Then %>checked<% End If %> name="ObjectCode" value="203" id="ObjectCode7" class="noborder"><label for="ObjectCode7"><%=getadminDocFlowLngStr("DtxtARDownPayReq")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If ODPIInv = "Y" Then %>checked<% End If %> name="ObjectCode" value="204" id="ObjectCode8" class="noborder"><label for="ObjectCode8"><%=getadminDocFlowLngStr("DtxtARDownPayInv")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If OINV = "Y" Then %>checked<% End If %> name="ObjectCode" value="13" id="ObjectCode3" class="noborder"><label for="ObjectCode3"><%=getadminDocFlowLngStr("DtxtInvoice")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If OINVR = "Y" Then %>checked<% End If %> name="ObjectCode" value="-13" id="ObjectCode6" class="noborder"><label for="ObjectCode6"><%=getadminDocFlowLngStr("DtxtInvoice")%> (<%=getadminDocFlowLngStr("DtxtReservada")%>)</label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If OIR = "Y" Then %>checked<% End If %> name="ObjectCode" value="48" id="ObjectCode4" class="noborder"><label for="ObjectCode4"><%=getadminDocFlowLngStr("DtxtInvoice")%>/<%=getadminDocFlowLngStr("DtxtReceipt")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If OPQT = "Y" Then %>checked<% End If %> name="ObjectCode" value="540000006" id="ObjectCode10" class="noborder"><label for="ObjectCode10"><%=getadminDocFlowLngStr("DtxtPurQuote")%></label><br>
						<input <% If Left(ExecAt,1) <> "D" Then %>disabled<% End If %> type="checkbox" <% If OPOR = "Y" Then %>checked<% End If %> name="ObjectCode" value="22" id="ObjectCode9" class="noborder"><label for="ObjectCode9"><%=getadminDocFlowLngStr("DtxtPurOrder")%></label>
						</font></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 100px" rowspan="2" valign="top">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminDocFlowLngStr("DtxtAdvanced")%></font></b></td>
						<td class="style1">
						<font face="Verdana" size="1" color="#4783C5">
							<input type="checkbox" name="FlowDraft" id="FlowDraft" value="Y" <% If Draft = "Y" Then %>checked<% End If %> class="noborder" <% If Request("NewFlow") = "Y" or Request("FlowID") <> "" and (FlowType <> 1 or FlowType = 1 and ExecAt <> "D3" and ExecAt <> "R2") Then %>disabled<% End If %>><label for="FlowDraft"><%=getadminDocFlowLngStr("LtxtCreateAsDraft")%></label></font></td>
					</tr>
					<tr>
						<td class="style1"><font face="Verdana" size="1" color="#4783C5">
							<input type="checkbox" name="FlowAuthorize" id="FlowAuthorize" value="Y" <% If Authorize = "Y" Then %>checked<% End If %> class="noborder" <% If Request("NewFlow") = "Y" or Request("FlowID") <> "" and (FlowType <> 1 or FlowType = 1 and ExecAt <> "D3") Then %>disabled<% End If %>><label for="FlowAuthorize"><%=getadminDocFlowLngStr("LtxtReqQutOrdr")%></label></font></td>
					</tr>
					<tr>
						<td colspan="2">
						<table border="0" cellpadding="0" width="100%">
							<tr>
								<td colspan="3" style="">
								<table border="0" cellpadding="0" width="300">
									<tr>
										<td bgcolor="#D9F5FF" align="center" style="border: 1px solid #31659C" onclick="javascript:showTab('tabQry');" id="btntabQry" onmouseover="javascript:if(document.getElementById('tabQry').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('tabQry').style.display=='none')this.bgColor='#BFEEFE';">
								<b><font color="#31659C" face="Verdana" size="1"><%=getadminDocFlowLngStr("DtxtQuery")%></font></b></td>
										<td bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand" onclick="javascript:showTab('tabMsg');" id="btntabMsg" onmouseover="javascript:if(document.getElementById('tabMsg').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('tabMsg').style.display=='none')this.bgColor='#BFEEFE';">
								<b>
								<font color="#31659C" face="Verdana" size="1"><%=getadminDocFlowLngStr("LtxtMsg")%></font></b></td>
										<td bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand" onclick="javascript:showTab('tabLines');" id="btntabLines" onmouseover="javascript:if(document.getElementById('tabLines').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('tabLines').style.display=='none')this.bgColor='#BFEEFE';"><b>
										<font color="#31659C" face="Verdana" size="1">
										<%=getadminDocFlowLngStr("DtxtLines")%></font></b></td>
										<td bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand" onclick="javascript:showTab('tabAutGrp');" id="btntabAutGrp" onmouseover="javascript:if(document.getElementById('tabAutGrp').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('tabAutGrp').style.display=='none')this.bgColor='#BFEEFE';"><b>
										<font color="#31659C" face="Verdana" size="1">
										<%=getadminDocFlowLngStr("LtxtAutGrp")%></font></b></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr id="tabAutGrp" style="display: none;">
								<td colspan="3" bgcolor="#F3FBFE">
								<div style="width: 100%; background-color: #E2F3FC; ">
								<font face="Verdana" size="1">
								<strong><span class="style2"><%=getadminDocFlowLngStr("LtxtAutGrp")%></span></strong></font>
								</div>
								<table id="tblAutGrp" border="0">
									<tr>
										<td bgcolor="#E2F3FC" width="200">
										<font face="Verdana" size="1" color="#31659C"><strong><%=getadminDocFlowLngStr("DtxtGroup")%></strong></font></td>
										<td bgcolor="#E2F3FC" width="200">&nbsp;</td>
										<td bgcolor="#E2F3FC" width="200">&nbsp;</td>
										<td class="TblRepTlt" align="center">
										<font face="Verdana" size="1" color="#31659C"><strong><%=getadminDocFlowLngStr("DtxtOrder")%></strong></font></td>
										<td class="TblRepTlt" width="1"></td>
									</tr>
									<% 
									If Request("FlowID") <> "" Then
										set rd = Server.CreateObject("ADODB.RecordSet")
										set cmd = Server.CreateObject("ADODB.Command")
										cmd.ActiveConnection = connCommon
										cmd.CommandType = &H0004
										cmd.CommandText = "DBOLKGetUAFAutGrp" & Session("ID")
										cmd.Parameters.Refresh()
										cmd("@FlowID") = Request("FlowID")
										rd.open cmd, , 3, 1
										do while not rd.eof
										GrpID = rd("GrpID")
										 %>
										<tr>
											<td class="TblRepNrm"><script type="text/javascript">maxOrdr=<%=rd("Ordr")%>+1;</script>
											<input type="hidden" name="GrpID" value="<%=GrpID%>"><%=rd("GrpName")%></td>
											<td class="TblRepNrm"><input type="checkbox" name="AsignedSLP<%=GrpID%>" id="AsignedSLP<%=GrpID%>" value="Y" class="noborder" <% If rd("AsignedSLP") = "Y" Then %>checked<% End If %>><label for="AsignedSLP<%=GrpID%>"><%=getadminDocFlowLngStr("LtxtAsignedSLP")%></label></td>
											<td class="TblRepNrm"><input type="checkbox" name="GrpQuery<%=GrpID%>" id="GrpQuery<%=GrpID%>" value="Y" class="noborder" <% If Not IsNull(rd("Query")) Then %>checked<% End If %> onclick="document.getElementById('trQryGrpID<%=GrpID%>').style.display=this.checked?'':'none';"><label for="GrpQuery<%=GrpID%>"><%=getadminDocFlowLngStr("LtxtAlertFilter")%></label></td>
											<td align="center"><table cellpadding="0" border="0">
													<tr>
														<td class="TblRepNrm"><input type="text" name="Order<%=GrpID%>" id="Order<%=GrpID%>" size="5" style="text-align:right" class="input" value="<%=rd("Ordr")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);"></td>
														<td valign="middle">
														<table cellpadding="0" cellspacing="0" border="0">
															<tr>
																<td><img src="images/img_nud_up.gif" id="btnOrder<%=GrpID%>Up"></td>
															</tr>
															<tr>
																<td><img src="images/spacer.gif"></td>
															</tr>
															<tr>
																<td><img src="images/img_nud_down.gif" id="btnOrder<%=GrpID%>Down"></td>
															</tr>
														</table>
														</td>
													</tr>
												</table>
											<script language="javascript">NumUDAttach('frmAddEditFlow', 'Order<%=GrpID%>', 'btnOrder<%=GrpID%>Up', 'btnOrder<%=GrpID%>Down');</script>
											</td>
											<td class="TblRepNrm"><img border="0" src="images/remove.gif" width="16" height="16" onclick="delGrp(this, <%=GrpID%>);"></td>
										</tr>
										<tr id="trQryGrpID<%=GrpID%>" <% If IsNull(rd("Query")) Then %>style="display: none;"<% End If %>>
											<td colspan="5">
												<font face="Verdana" size="1" color="#4783C5">where SlpCode in (...)</font>
												<table cellpadding="0" cellspacing="0" border="0" width="100%">
													<tr>
														<td rowspan="2">
															<textarea dir="ltr" rows="10" style="width: 100%" name="GrpValue<%=GrpID%>Query" id="GrpValue<%=GrpID%>Query" cols="100" class="input" onkeypress="javascript:document.frmAddEditFlow.btnVerfyGrpValue<%=GrpID%>Query.src='images/btnValidate.gif';document.frmAddEditFlow.btnVerfyGrpValue<%=GrpID%>Query.style.cursor = 'hand';document.frmAddEditFlow.valGrpValue<%=GrpID%>Query.value='Y';"><%=myHTMLEncode(rd("Query"))%></textarea>
														</td>
														<td valign="top" width="1">
															<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteGrpFilter" alt="<%=getadminDocFlowLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(11, 'GrpValue<%=GrpID%>Query', '', null);">
														</td>
													</tr>
													<tr>
														<td valign="bottom" width="1">
															<img src="images/btnValidateDis.gif" id="btnVerfyGrpValue<%=GrpID%>Query" alt="<%=getadminDocFlowLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmAddEditFlow.valGrpValue<%=GrpID%>Query.value == 'Y')VerfyQuery('GrpValue<%=GrpID%>');">
															<input type="hidden" name="valGrpValue<%=GrpID%>Query" id="valGrpValue<%=GrpID%>Query" value="N">
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<% rd.movenext
										loop
									End If %>
										<tr>
											<td bgcolor="#E2F3FC" class="TblRepNrm">
										<font face="Verdana" size="1" color="#31659C"><strong><%=getadminDocFlowLngStr("DtxtAdd")%></strong></font></td>
											<td class="TblRepNrm"><select name="AddGrp" id="AddGrp" onchange="doAddGrp(this);">
											<option></option>
											<% set rd = Server.CreateObject("ADODB.RecordSet")
											set cmd = Server.CreateObject("ADODB.Command")
											cmd.ActiveConnection = connCommon
											cmd.CommandType = &H0004
											cmd.CommandText = "DBOLKGetUAFAutGrpList" & Session("ID")
											cmd.Parameters.Refresh()
											If Request("FlowID") <> "" Then cmd("@FlowID") = Request("FlowID") Else cmd("@FlowID") = -1
											rd.open cmd, , 3, 1
											do while not rd.eof %>
											<option value="<%=rd("GrpID")%>"><%=rd("GrpName")%></option>
											<% rd.movenext
											loop %>
											</select><input type="hidden" name="delID" value=""></td>
										</tr>
									</table>								
								</td>
							</tr>
							<tr id="tabQry" style="">
								<td colspan="3" bgcolor="#F3FBFE">
								<div style="width: 100%; background-color: #E2F3FC; ">
								<font size="1" face="Verdana" color="#31659C">
								(<%=getadminDocFlowLngStr("LtxtMustStart")%> <b>Select 'TRUE'</b>)
								</font></div>
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="15" style="width: 100%" name="FlowQuery" id="FlowQuery" cols="87" class="input" onkeypress="javascript:document.frmAddEditFlow.btnVerfyFlowQuery.src='images/btnValidate.gif';document.frmAddEditFlow.btnVerfyFlowQuery.style.cursor = 'hand';;document.frmAddEditFlow.valFlowQuery.value='Y';"><%=myHTMLEncode(FlowQuery)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminDocFlowLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(8, 'FlowQuery', '<%=Request("FlowID")%>', <% If Request("FlowID") <> "" Then %>null<% Else %>document.frmAddEditFlow.FlowQueryDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFlowQuery" alt="<%=getadminDocFlowLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmAddEditFlow.valFlowQuery.value == 'Y')VerfyQuery('Flow');">
										<input type="hidden" name="valFlowQuery" id="valFlowQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr id="tabMsg" style="display: none">
								<td colspan="3" bgcolor="#F3FBFE">
								<div style="width: 100%; background-color: #E2F3FC; ">
								<font face="Verdana" size="1" color="#4783C5">
								<input <% If NoteBuilder = "Y" Then %>checked<% End If %> type="checkbox" name="NoteBuilder" value="Y" id="fp1" onclick="javascript:ChkNoteBuilder()" class="noborder"><label for="fp1"><%=getadminDocFlowLngStr("LtxtGenNoteQry")%></label></font>
								</div>
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="15" <% If NoteBuilder <> "Y" Then %>disabled<% End If %> style="width: 100%" name="NoteQuery" id="NoteQuery" cols="87" class="input" onkeypress="javascript:if(document.frmAddEditFlow.NoteBuilder.checked){document.frmAddEditFlow.btnVerfyNoteQuery.src='images/btnValidate.gif';document.frmAddEditFlow.btnVerfyNoteQuery.style.cursor = 'hand';document.frmAddEditFlow.valNoteQuery.value='Y';}"><%=myHTMLEncode(NoteQuery)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminDocFlowLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(8, 'NoteQuery', '<%=Request("FlowID")%>', <% If Request("FlowID") <> "" Then %>null<% Else %>document.frmAddEditFlow.NoteQueryDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyNoteQuery" alt="<%=getadminDocFlowLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmAddEditFlow.valNoteQuery.value == 'Y')VerfyQuery('Note');">
											<input type="hidden" name="valNoteQuery" id="valNoteQuery" value="N">
										</td>
									</tr>
								</table>
								<div style="width: 100%; background-color: #E2F3FC; ">
								<font face="Verdana" size="1">
								<strong><span class="style2"><%=getadminDocFlowLngStr("LtxtMsgText")%></span></strong></font>
								</div>
								<table border="0" cellpadding="0" width="100%" id="table28" cellspacing="0">
									<tr>
										<td><textarea rows="15" style="width: 100%" name="NoteText" cols="67" class="input"><%=myHTMLEncode(NoteText)%></textarea></td>
										<td width="120"><select <% If NoteBuilder <> "Y" Then %>disabled<% End If %> size="10" name="NoteQueryFields" class="input" style="width:120px; height:182px" ondblclick="javascript:if(!this.disabled)document.frmAddEditFlow.NoteText.value+=this.value;">
										<% If NoteBuilder = "Y" Then
										sql = "declare @LogNum int set @LogNum = -1 " & _
										"declare @LanID int set @LanID = -1 " & _
										"declare @SlpCode int set @SlpCode = -1 " & _
										"declare @dbName nvarchar(100) set @dbName = db_name() " & _
										"declare @branch int set @branch = -1 "
										
										If Left(ExecAt,1) = "D" or Left(ExecAt,1) = "R" or ExecAt = "C2" or ExecAt = "C3" Then sql = sql & "declare @CardCode nvarchar(15) set @CardCode = '' "
										If ExecAt = "O2" or ExecAt = "O3" or ExecAt = "O4" Then sql = sql & "declare @ObjectCode int "
										If Left(ExecAt, 1) = "O" and Left(ExecAt, 2) <> "OP" Then sql = sql & "declare @Entry int "
										If ExecAt = "D2" Then sql = sql & "declare @ItemCode nvarchar(20) set @ItemCode = '' declare @WhsCode nvarchar(8) set @WhsCode = N'' declare @Quantity int set @Quantity = 1 declare @Unit smallint declare @Price numeric(19,6) set @Price = 0  " 
																				
										sql = sql & NoteQuery
										sql = QueryFunctions(sql)
									set rs = conn.execute(sql)
									For each item in rs.Fields
									If item.Name <> "" Then %>
									<option value="{<%=myHTMLEncode(item.Name)%>}"><%=myHTMLEncode(item.Name)%></option>
									<% End If
										next
									End If %>
									</select></td>
										<td width="16" valign="bottom"><a href="javascript:doFldTrad('UAF', 'FlowID', '<%=Request("FlowID")%>', 'AlterNoteText', 'M', <% If Request("NewFlow") <> "Y" Then %>null<% Else %>document.frmAddEditFlow.NoteTextTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminDocFlowLngStr("DtxtTranslate")%>" border="0"></a></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr id="tabLines" style="display: none">
								<td colspan="3" bgcolor="#F3FBFE">
								<div style="width: 100%; background-color: #E2F3FC; ">
								<font face="Verdana" size="1" color="#31659C">
								<strong><%=getadminDocFlowLngStr("LtxtOptQry")%></strong></font></div>
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="15" style="width: 100%" name="LineQuery" id="LineQuery" cols="87" class="input" onkeyup="javascript:document.frmAddEditFlow.btnVerfyLineQuery.src='images/btnValidate.gif';document.frmAddEditFlow.btnVerfyLineQuery.style.cursor = 'hand';;document.frmAddEditFlow.valLineQuery.value='Y';"><%=myHTMLEncode(LineQuery)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminDocFlowLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(8, 'LineQuery', '<%=Request("FlowID")%>', <% If Request("FlowID") <> "" Then %>null<% Else %>document.frmAddEditFlow.LineQueryDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyLineQuery" alt="<%=getadminDocFlowLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmAddEditFlow.valLineQuery.value == 'Y')VerfyQuery('Line');">
											<input type="hidden" name="valLineQuery" id="valLineQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td height="30" valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminDocFlowLngStr("DtxtVariables")%></font></b></td>
						<td height="30" valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						<span id="txtVars">
						<% If ExecAt <> "" Then %>
						<% If (ExecAt <> "D1" and ExecAt <> "R1" and Left(ExecAt, 1) <> "O") or Left(ExecAt, 2) = "OP" Then %><span dir="ltr">@LogNum</span> = <%=getadminDocFlowLngStr("LtxtOLKDocKey")%><br><% End If %>
						<% If ExecAt = "O2" or ExecAt = "O3" or ExecAt = "O4" Then %><span dir="ltr">@ObjectCode</span> = <%=getadminDocFlowLngStr("LtxtObjCode")%><br><% End If %>
						<% If Left(ExecAt, 1) = "O" and Left(ExecAt, 2) <> "OP" Then %><span dir="ltr">@Entry</span> = <%=getadminDocFlowLngStr("LtxtDocEntry")%><br><% End If %>

						<span dir="ltr">@LanID</span> = <%=getadminDocFlowLngStr("DtxtLanID")%><br>
						<span dir="ltr">@SlpCode</span> = <%=getadminDocFlowLngStr("LtxtAgentCode")%><br>
						<span dir="ltr">@dbName</span> = <%=getadminDocFlowLngStr("DtxtDB")%><br>
						<span dir="ltr">@branch</span> = <%=getadminDocFlowLngStr("LtxtBranchCode")%><br>
						<% If (Left(ExecAt,1) = "D" or Left(ExecAt,1) = "R") or ExecAt = "C2" or ExecAt = "C3" Then %><span dir="ltr">@CardCode</span> = <%=getadminDocFlowLngStr("DtxtClientCode")%><br><% End If %>
						<% If (ExecAt = "D2") Then %><span dir="ltr">@ItemCode</span> = <%=getadminDocFlowLngStr("DtxtItemCode")%><br>
						<span dir="ltr">@WhsCode</span> = <%=getadminDocFlowLngStr("DtxtWhsCode")%><br>
						<span dir="ltr">@Quantity</span> = <%=getadminDocFlowLngStr("LtxtQtyInUnit")%><br>
						<span dir="ltr">@Unit</span> = <%=getadminDocFlowLngStr("DtxtUnit")%>: 1 = <%=getadminDocFlowLngStr("DtxtUnit")%>, 2 = <%=getadminDocFlowLngStr("DtxtSalUnit")%>, 3 = <%=getadminDocFlowLngStr("DtxtPackUnit")%><br>
						<span dir="ltr">@Price</span> = <%=getadminDocFlowLngStr("DtxtPrice")%>
						<% End If %>
						<% End If %>
						</span></font></td>
					</tr>
					<tr>
						<td height="30" valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<b><font size="1" color="#31659C" face="Verdana">
						<%=getadminDocFlowLngStr("DtxtFunctions")%></font></b></td>
						<td height="30" valign="top" class="style1">
						<% HideFunctionTitle = True
						functionClass="TblFlowFunction" %>
						<!--#include file="myFunctions.asp"--></td>
					</tr>
					<tr>
						<td height="30" valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<font face="Verdana" size="1" color="#31659C"><b>
						<%=getadminDocFlowLngStr("DtxtTips")%></b></font></td>
						<td height="30" valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5"><%=getadminDocFlowLngStr("LtxtTipLine1")%> <br>
						<%=getadminDocFlowLngStr("LtxtTipLine2")%><br>
						<%=getadminDocFlowLngStr("LtxtTipLine3")%></font></td>
					</tr>
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
				<input type="submit" value="<%=getadminDocFlowLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminDocFlowLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<% If Request("FlowID") < 0 Then %>
				<td width="77">
				<input type="submit" value="<%=getadminDocFlowLngStr("DtxtRestore")%>" name="btnRestore" class="OlkBtn" onclick="javascript:return confirm('<%=getadminDocFlowLngStr("LtxtValRestoreFlow")%>');"></td><% End If %>
				<td width="77">
				<input type="button" value="<%=getadminDocFlowLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminDocFlowLngStr("DtxtConfCancel")%>'))window.location.href='adminDocFlow.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminDocFlow">
	<input type="hidden" name="FlowID" value="<%=Request("FlowID")%>">
	<input type="hidden" name="LockExec" value="<%=LockExec%>">
</form>
<script language="javascript">
var txtChangeApplyToClie 	= '<%=getadminDocFlowLngStr("LtxtChangeApplyToClie")%>';
var txtFlowTypeCAlr			= '<%=getadminDocFlowLngStr("LtxtFlowTypeCAlr")%>';
var txtFlowTypeCAlr2		= '<%=getadminDocFlowLngStr("LtxtFlowTypeCAlr2")%>';
var txtFlowComExec 			= '<%=getadminDocFlowLngStr("LtxtFlowComExec")%>';
var txtExecComFlow 			= '<%=getadminDocFlowLngStr("LtxtExecComFlow")%>';
var txtValFlowNam 			= '<%=getadminDocFlowLngStr("LtxtValFlowNam")%>';
var txtValSelAgent 			= '<%=getadminDocFlowLngStr("LtxtValSelAgent")%>';
var txtValSelAgentOrClie 	= '<%=getadminDocFlowLngStr("LtxtValSelAgentOrClie")%>';
var txtValExecMom 			= '<%=getadminDocFlowLngStr("LtxtValExecMom")%>';
var txtValSelDoc 			= '<%=getadminDocFlowLngStr("LtxtValSelDoc")%>';
var txtValFlowQry 			= '<%=getadminDocFlowLngStr("LtxtValFlowQry")%>';
var txtVarFlowQryVal 		= '<%=getadminDocFlowLngStr("LtxtVarFlowQryVal")%>';
var txtValMsgQry 			= '<%=getadminDocFlowLngStr("LtxtValMsgQry")%>';
var txtValMsgQryVal 		= '<%=getadminDocFlowLngStr("LtxtValMsgQryVal")%>';
var txtValMsgText 			= '<%=getadminDocFlowLngStr("LtxtValMsgText")%>';
var txtVarAutGrpQry			= '<%=getadminDocFlowLngStr("LtxtVarAutGrpQry")%>';
var txtOLKDocKey			= '<%=getadminDocFlowLngStr("LtxtOLKDocKey")%>';
var txtDocEntry				= '<%=getadminDocFlowLngStr("LtxtDocEntry")%>';
var txtObjCode				= '<%=getadminDocFlowLngStr("LtxtObjCode")%>';
var txtAgentCode 			= '<%=getadminDocFlowLngStr("LtxtAgentCode")%>';
var txtDB 					= '<%=getadminDocFlowLngStr("DtxtDB")%>';
var txtBranchCode 			= '<%=getadminDocFlowLngStr("LtxtBranchCode")%>';
var txtClientCode 			= '<%=getadminDocFlowLngStr("DtxtClientCode")%>';
var txtItemCode 			= '<%=Replace(getadminDocFlowLngStr("DtxtItemCode"), "'", "\'")%>';
var txtWhsCode 				= '<%=getadminDocFlowLngStr("DtxtWhsCode")%>';
var txtValLineQryVal		= '<%=getadminDocFlowLngStr("LtxtValLineQryVal")%>';
var txtLanID				= '<%=getadminDocFlowLngStr("DtxtLanID")%>';
var txtQtyInUnit			= '<%=getadminDocFlowLngStr("LtxtQtyInUnit")%>';
var txtUnit					= '<%=Replace(getadminDocFlowLngStr("DtxtUnit"), "'", "\'")%>: 1 = <%=Replace(getadminDocFlowLngStr("DtxtUnit"), "'", "\'")%>, 2 = <%=Replace(getadminDocFlowLngStr("DtxtSalUnit"), "'", "\'")%>, 3 = <%=Replace(getadminDocFlowLngStr("DtxtPackUnit"), "'", "\'")%>';
var txtPrice				= '<%=getadminDocFlowLngStr("DtxtPrice")%>';
var txtOp					= '<%=getadminDocFlowLngStr("DtxtOp")%>';
</script>
<script language="javascript" src="adminDocFlow.js"></script>
	<% End If %>
	</table>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="DocFlow">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="by" value="">
	<input type="hidden" name="ExecAt" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<!--#include file="bottom.asp" -->