<!--#include file="top.asp" -->
<!--#include file="lang/adminAutGrp.asp" -->
<br>
<script type="text/javascript">
var txtConfDelGrp = '<%=getadminAutGrpLngStr("LtxtConfDelGrp")%>';
var txtOr = '<%=getadminAutGrpLngStr("DtxtOr")%>';
var txtAnd = '<%=getadminAutGrpLngStr("DtxtAnd")%>';
var txtConfDel = '<%=getadminAutGrpLngStr("LtxtConfDel")%>';
</script>
<script type="text/javascript" src="adminAutGrp.js"></script>
<% If Request("GrpID") = "" Then %>
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminAutGrpLngStr("LttlAutGrp")%></td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmGroups">
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminAutGrpLngStr("LttlAutGrpNote")%></td>
	</tr>
	<tr>
		<td >
		<table border="0" cellpadding="0"><tr class="TblRepTltSub">
				<td align="center" width="1"></td>
				<td align="center"><%=getadminAutGrpLngStr("DtxtGroup")%>&nbsp;</td>
				<td align="center"><%=getadminAutGrpLngStr("DtxtBranch")%></td>
				<td align="center" width="1"></td>
			</tr>
			<% 
			set rd = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetAutGrp" & Session("ID")
			cmd.Parameters.Refresh()
			rd.open cmd, , 3, 1
			do while not rd.eof
			GrpID = rd("GrpID") %>
			<tr class="TblRepTbl">
			  <td width="15">
				<a href="adminAutGrp.asp?GrpID=<%=GrpID%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
				<td valign="bottom"><input type="hidden" name="GrpID" value="<%=GrpID%>">
				<input type="text" id="GrpName<%=GrpID%>" name="GrpName<%=GrpID%>" size="100" value="<%=Server.HTMLEncode(rd("GrpName"))%>" onkeydown="return chkMax(event, this, 100);"></td>
				<td valign="top" align="center">
				<input type="checkbox" id="chkBranch" class="noborder" name="Branch<%=GrpID%>" value="Y" <% If rd("FilterBranch") = "Y" Then %>checked<% End If %>></td>
				<td valign="top" style="width: 15px">
				<% If rd("Verfy") = "N" Then %><a href="javascript:delGrp(<%=GrpID%>);"><img border="0" src="images/remove.gif"></a><% End If %></td>
			</tr>
			<% rd.movenext
			loop %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<% If rd.recordcount > 0 Then %><td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminAutGrpLngStr("DtxtSave")%>" name="btnSave"></td><% End If %>
				<td width="77">
				<input class="BtnRep" type="button" value="<%=getadminAutGrpLngStr("DtxtNew")%>" name="btnNew" onclick="window.location.href='adminAutGrp.asp?GrpID=New';"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="submitCmd" value="AutGrp">
	<input type="hidden" name="cmd" value="uGrp">
	</form>
</table>
<% Else
If Request("GrpID") <> "New" Then 
	GrpID = CInt(Request("GrpID"))
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetAutGrpData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@GrpID") = GrpID
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open cmd
	GroupName = rs("GrpName")
	FilterBranch = rs("FilterBranch") = "Y"
End If %>
<script language="javascript" src="js_up_down.js"></script>
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<% If Request("GrpID") = "New" Then %><%=getadminAutGrpLngStr("LttlAddAutGrp")%><% Else %><%=getadminAutGrpLngStr("LttlEditAutGrp")%><% End If %></td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmEditGroups">
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminAutGrpLngStr("LttlAutGrpEditNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminAutGrpLngStr("DtxtName")%></td>
				<td colspan="7" class="TblRepNrm"><input type="text" name="GroupName" size="88" value="<%=myHTMLEncode(GroupName)%>" onkeydown="return chkMax(event, this, 60);"></td>
			</tr>
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				</td>
				<td colspan="7" class="TblRepNrm"><input type="checkbox" class="noborder" id="chkBranch" name="chkBranch" value="Y" <% If FilterBranch Then %>checked<% End If %>><label for="chkBranch"><%=getadminAutGrpLngStr("DtxtBranch")%></label></td>
			</tr>
			<tr>
				<td style="width: 100px" class="TblRepTlt" valign="top">
				<%=getadminAutGrpLngStr("DtxtAgents")%></td>
				<td colspan="7" class="TblRepNrm">
				<table id="tblSlp" border="0">
				<tr>
					<td class="TblRepTlt" width="200"><%=getadminAutGrpLngStr("DtxtAgent")%></td>
					<td class="TblRepTlt" align="center">&nbsp;</td>
					<td class="TblRepTlt" align="center"><%=getadminAutGrpLngStr("DtxtOrder")%></td>
					<td class="TblRepTlt" width="1"></td>
				</tr>
				<% 
				If GrpID <> "" Then
					set rd = Server.CreateObject("ADODB.RecordSet")
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetAutGrpSlp" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@GrpID") = GrpID
					rd.open cmd, , 3, 1
					do while not rd.eof
					SlpCode = rd("SlpCode")
					 %>
					<tr>
						<td class="TblRepNrm"><% If rd.Bookmark = rd.recordcount Then %><script language="javascript">maxOrdr=<%=rd("Ordr")%>+1;lastSlp=<%=SlpCode%>;</script><% End If %>
					<input type="hidden" name="SlpCode" value="<%=SlpCode%>"><%=rd("SlpName")%></td>
						<td class="TblRepNrm"><select name="Op<%=SlpCode%>" id="Op<%=SlpCode%>" <% If rd.bookmark = rd.recordcount Then %>style="display: none;"<% End If %>>
						<option value="O" <% If rd("Op") = "O" Then %>selected<% End If %>><%=getadminAutGrpLngStr("DtxtOr")%></option>
						<option value="A" <% If rd("Op") = "A" Then %>selected<% End If %>><%=getadminAutGrpLngStr("DtxtAnd")%></option>
						</select>
						</td>
						<td align="center"><table cellpadding="0" border="0">
								<tr>
									<td class="TblRepNrm"><input type="text" name="Order<%=SlpCode%>" id="Order<%=SlpCode%>" size="5" style="text-align:right" class="input" value="<%=rd("Ordr")%>" onfocus="this.select()" onchange="chkThis(this, 0, 0)" onkeydown="return chkMax(event, this, 6);"></td>
									<td valign="middle">
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><img src="images/img_nud_up.gif" id="btnOrder<%=SlpCode%>Up"></td>
										</tr>
										<tr>
											<td><img src="images/spacer.gif"></td>
										</tr>
										<tr>
											<td><img src="images/img_nud_down.gif" id="btnOrder<%=SlpCode%>Down"></td>
										</tr>
									</table>
									</td>
								</tr>
							</table>
						<script language="javascript">NumUDAttach('frmEditGroups', 'Order<%=SlpCode%>', 'btnOrder<%=SlpCode%>Up', 'btnOrder<%=SlpCode%>Down');</script>
						</td>
						<td class="TblRepNrm"><img border="0" src="images/remove.gif" width="16" height="16" onclick="delSlp(this, <%=SlpCode%>);"></td>
					</tr>
					<% rd.movenext
					loop
				End If %>
					<tr>
						<td class="TblRepNrm"><%=getadminAutGrpLngStr("DtxtAdd")%></td>
						<td class="TblRepNrm" colspan="2"><select name="AddSlp" id="AddSlp" onchange="doAddSlp(this);">
						<option></option>
						<% set rd = Server.CreateObject("ADODB.RecordSet")
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetSlpFilterAutGrp" & Session("ID")
						cmd.Parameters.Refresh()
						If GrpID <> "" Then cmd("@GrpID") = GrpID Else cmd("@GrpID") = -1
						rd.open cmd, , 3, 1
						do while not rd.eof %>
						<option value="<%=rd("SlpCode")%>"><%=rd("SlpName")%></option>
						<% rd.movenext
						loop %>
						</select></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminAutGrpLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminAutGrpLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="submitCmd" value="AutGrp">
	<input type="hidden" name="cmd" value="AutGrpData">
	<input type="hidden" name="GrpID" value="<%=GrpID%>">
	<input type="hidden" name="delID" value="">
	</form>
</table>

<% End If %>
<!--#include file="bottom.asp" -->