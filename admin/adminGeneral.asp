<!--#include file="top.asp" -->
<!--#include file="lang/adminGeneral.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<% 
set rd = Server.CreateObject("ADODB.recordset")
%>
<script language="javascript" src="js_up_down.js"></script>
<script language="javascript">
function valFrm()
{
	if (document.Form1.AgentClientsFilter.value != '' && document.Form1.valAgentClientsFilter.value == 'Y')
	{
		alert('<%=getadminGeneralLngStr("LtxtValClientFilter")%>');
		document.Form1.btnVerfyFilter.focus();
		return false;
	}
	if (document.Form1.ViewDocFilter.value != '' && document.Form1.valViewDocFilter.value == 'Y')
	{
		alert('<%=getadminGeneralLngStr("LtxtValViewDocFilter")%>');
		document.Form1.btnVerfyViewDocFilter.focus();
		return false;
	}
	return true;
}
</script>

<form method="POST" action="adminsubmit.asp" name="Form1" onsubmit="javascript:return valFrm();">
	<table border="0" cellpadding="0" width="100%">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminGeneralLngStr("LttlGenOpt")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LttlGenOptNote")%></font></p>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE" height="226">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtNatLng")%></font></td>
						<td bgcolor="#F7FBFF">
						<span dir="ltr">
						<font face="Verdana" size="1" color="#4783C5">
						<% For i = 0 to UBound(myLanIndex) 
							If myApp.NatLng = myLanIndex(i)(0) Then
								Response.Write myLanIndex(i)(1)
								Exit For
							End If
						Next %></font></span></td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtNumFormat")%></font></td>
						<td bgcolor="#F7FBFF">
						<p>
						<% 
						Select Case myApp.LawsSet
							Case "CL", "CR"
								strFormat1 = "1.000.000,00"
								strFormat2 = "1,000,000.00"
							Case Else
								strFormat1 = "1,000,000.00"
								strFormat2 = "1.000.000,00"
						End Select %>
						<select size="1" name="AlterLocation" class="input">
						<option value="N"><%=strFormat1%></option>
						<option <% If myApp.AlterLocation Then %>selected<% End If %> value="Y"><%=strFormat2%></option>
						</select></p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtSlpCode")%></font></td>
						<td bgcolor="#F7FBFF">
						<p>
						<select size="1" name="SlpCode" class="input">
						<% 
						GetQuery rd, 2, "Y", "" & getadminGeneralLngStr("LoptCardAgent2") & ""
						do while not rd.eof %>
						<option <% If rd("SLPCode") = myApp.SlpCode Then %>selected<%end if%> value="<%=rd("SLPCode")%>">
						<%=myHTMLEncode(rd("SLPName"))%></option>
						<% 
						rd.movenext
						loop
				    	%></select></p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtAddDocSlpCode")%></font></td>
						<td bgcolor="#F7FBFF">
						<p>
						<select size="1" name="AddDocSlpCode" class="input">
						<option <% If myApp.AddDocSlpCode = -2 Then %>selected<% End If %> value="-2">
						-<%=getadminGeneralLngStr("LoptCardAgent")%>-</option>
						<option <% If myApp.AddDocSlpCode = -3 Then %>selected<% End If %> value="-3">
						-<%=getadminGeneralLngStr("LoptDocAgent")%>-</option>
						</select></p>
						</td>
					</tr>
					<% If 1 = 2 Then %>
					<tr>
						<td bgcolor="#F7FBFF" colspan="2">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="AsignedSLP" value="Y" <% If myApp.AsignedSLP Then %>checked<%end if %> id="AsignedSLP" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="AsignedSLP"><%=getadminGeneralLngStr("LtxtAsignedSLP")%></label></font></td>
					</tr>
					<% Else %>
					<input type="hidden" name="AsignedSLP" value="<%=myApp.AsignedSLP%>">
					<% End If %>
					<tr>
						<td bgcolor="#F7FBFF" colspan="2">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="AllowAgentAccessCDoc" value="Y" <% If myApp.AllowAgentAccessCDoc Then %>checked<%end if %> id="AllowAgentAccessCDoc" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="AllowAgentAccessCDoc"><%=getadminGeneralLngStr("LtxtAllowAgentAccessC")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF" colspan="2">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="FlowAutDirectAdd" value="Y" <% If myApp.FlowAutDirectAdd Then %>checked<%end if %> id="FlowAutDirectAdd" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="FlowAutDirectAdd"><%=getadminGeneralLngStr("LtxtFlowAutDirectAdd")%></label></font></td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF" style="height: 23px">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnRetroPoll" <% If myApp.EnRetroPoll Then %>checked<% End If %> value="Y" id="EnRetroPoll" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnRetroPoll"><%=getadminGeneralLngStr("LtxtEnRetroPoll")%></label></font></td>
						<td bgcolor="#F7FBFF" style="height: 23px">
						<p align="right">&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtEnRetroPollDays")%></font></td>
						<td bgcolor="#F7FBFF">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input type="text" name="EnRetroPollDays" id="EnRetroPollDays" size="5" value='<%=myApp.EnRetroPollDays%>' class="input" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);"></td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnEnRetroPollDaysUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnEnRetroPollDaysDown"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						<script language="javascript">NumUDAttach('Form1', 'EnRetroPollDays', 'btnEnRetroPollDaysUp', 'btnEnRetroPollDaysDown');</script>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="AddUserApr" value="Y" <% If myApp.AddUserApr Then %>checked<%end if %> id="AddUserApr" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="AddUserApr"><%=getadminGeneralLngStr("LtxtAddUserApr")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnableDelegate" value="Y" <% If myApp.EnableDelegate Then %>checked<%end if %> id="EnableDelegate" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableDelegate"><%=getadminGeneralLngStr("LtxtEnableDelegate")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnableAtt" value="Y" <% If myApp.EnableAtt Then %>checked<%end if %> id="EnableAtt" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableAtt"><%=getadminGeneralLngStr("LtxtEnableAtt")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnableBranchs" value="Y" <% If myApp.EnableBranchs Then %>checked<%end if %> id="EnableBranchs" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableBranchs"><%=getadminGeneralLngStr("LtxtEnableBranchs")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="CopyLastFCRate" value="Y" <% If myApp.CopyLastFCRate Then %>checked<%end if %> id="CopyLastFCRate" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="CopyLastFCRate"><%=getadminGeneralLngStr("LtxtCopyLastFCRate")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnBlockRClk" value="Y" <% If myApp.EnBlockRClk Then %>checked<%end if %> id="EnBlockRClk" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnBlockRClk"><%=getadminGeneralLngStr("LtxtEnBlockRClk")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF" valign="top">
						<img src="images/ganchito.gif"><font color="#4783C5" size="1" face="Verdana"> 
						<%=getadminGeneralLngStr("LtxtEnBlkRClkMsg")%> </font></td>
						<td bgcolor="#F7FBFF">
						<textarea name="EnBlkRClkMsg" class="input" cols="50" rows="2" onkeydown="return chkMax(event, this, 254);"><% If Not IsNull(myApp.EnBlkRClkMsg) Then %><%=Server.HTMLEncode(myApp.EnBlkRClkMsg)%><% End If %></textarea><a href="javascript:doFldTrad('Common', '', '', 'AlterEnBlkRClkMsg', 'M', null);"><img src="images/trad.gif" alt="<%=getadminGeneralLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
					<tr>
						<td colspan="2" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="ShowBPCountry" value="Y" <% If myApp.ShowBPCountry Then %>checked<%end if %> id="ShowBPCountry" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowBPCountry"><%=getadminGeneralLngStr("LShowBPCountry")%></label></font></td>
					</tr>
					<tr>
						<td colspan="2" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnableCSearchByLicTradNum" value="Y" <% If myApp.EnableCSearchByLicTradNum Then %>checked<%end if %> id="EnableCSearchByLicTradNum" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCSearchByLicTradNum"><%=getadminGeneralLngStr("LtxtEnableCSearchByLi")%></label></font></td>
					</tr>
					<tr>
						<td colspan="2" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font>
						<input type="checkbox" name="EnableCSearchByVatId" value="Y" <% If myApp.EnableCSearchByVatId Then %>checked<%end if %> id="EnableCSearchByVatId" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCSearchByVatId"><%=getadminGeneralLngStr("LtxtEnableCSearchByVa")%></label></font></td>						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font face="Verdana" size="1">
						<font color="#4783C5"><%=getadminGeneralLngStr("LtxtDefTab")%></font></font></td>
						<td bgcolor="#F7FBFF"><font face="Verdana" size="1">
						<select size="1" name="DefClientOPTab" class="input">
						<option <% If myApp.DefClientOPTab = 1 Then %>selected<% End If %> value="1"><%=getadminGeneralLngStr("LtxtGenData")%></option>
						<option <% If myApp.DefClientOPTab = 2 Then %>selected<% End If %> value="2"><%=getadminGeneralLngStr("DtxtActivities")%></option>
						<option <% If myApp.DefClientOPTab = 3 Then %>selected<% End If %> value="3"><%=getadminGeneralLngStr("LtxtPendOlk")%></option>
						<option <% If myApp.DefClientOPTab = 4 Then %>selected<% End If %> value="4"><%=getadminGeneralLngStr("LtxtPendSBO")%></option>
						<option <% If myApp.DefClientOPTab = 5 Then %>selected<% End If %> value="5"><%=getadminGeneralLngStr("LtxtOffers")%></option>
						</select></font></td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font face="Verdana" size="1">
						<font color="#4783C5"><%=getadminGeneralLngStr("LtxtDefActOrdr")%></font></font></td>
						<td bgcolor="#F7FBFF">
						<select size="1" name="ActOrdr1">
						<option value="1"><%=getadminGeneralLngStr("DtxtLogNum")%></option>
						<option <% If myApp.ActOrdr1 = "SlpName" Then %>selected<% End If %> value="SlpName"><%=getadminGeneralLngStr("DtxtAgent")%></option>
						<option <% If myApp.ActOrdr1 = "C1.CardCode" Then %>selected<% End If %> value="C1.CardCode"><%=getadminGeneralLngStr("DtxtCode")%></option>
						<option <% If myApp.ActOrdr1 = "CardName" Then %>selected<% End If %> value="CardName"><%=getadminGeneralLngStr("DtxtName")%></option>
						<option <% If myApp.ActOrdr1 = "GroupName" Then %>selected<% End If %> value="GroupName"><%=getadminGeneralLngStr("DtxtGroup")%></option>
						<option <% If myApp.ActOrdr1 = "Country" Then %>selected<% End If %> value="Country"><%=getadminGeneralLngStr("DtxtCountry")%></option>
						<option <% If myApp.ActOrdr1 = "CntctDateSort" Then %>selected<% End If %> value="CntctDateSort"><%=getadminGeneralLngStr("LtxtCntDate")%></option>
						<option <% If myApp.ActOrdr1 = "Action" Then %>selected<% End If %> value="Action"><%=getadminGeneralLngStr("DtxtActivity")%></option>
						</select><select size="1" name="ActOrdr2">
						<option value="asc"><%=getadminGeneralLngStr("DtxtAsc")%></option>
						<option <% If myApp.ActOrdr2 = "desc" Then %>selected<% End If %> value="desc"><%=getadminGeneralLngStr("DtxtDesc")%></option>
						</select></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminGeneralLngStr("LttlAgentClientFilter")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
			<img src="images/lentes.gif">
			<font color="#4783C5"><%=getadminGeneralLngStr("LttlAgentCLFilterNote")%></font></font></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td width="300" valign="top">
						<font face="Verdana" size="1" color="#4783C5"> 
						<%=getadminGeneralLngStr("DtxtQuery")%> - (CardCode not in)</font></td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td rowspan="2">
									<textarea rows="10" name="AgentClientsFilter" dir="ltr" cols="87" class="input" onkeydown="javascript:document.Form1.btnVerfyFilter.src='images/btnValidate.gif';document.Form1.btnVerfyFilter.style.cursor = 'hand';;document.Form1.valAgentClientsFilter.value='Y';"><% If Not IsNull(myApp.AgentClientsFilter) Then %><%=Server.HTMLEncode(myApp.AgentClientsFilter)%><% End If %></textarea>
								</td>
								<td valign="top">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminGeneralLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(1, 'AgentClientsFilter', -1, null);">
								</td>
							</tr>
							<tr>
								<td valign="bottom">
									<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminGeneralLngStr("DtxtValidate")%>" onclick="javascript:if (document.Form1.valAgentClientsFilter.value == 'Y')VerfyFilter();">
									<input type="hidden" name="valAgentClientsFilter" value="N">
								</td>
							</tr>
						</table>
						</td>
					</tr>
	
					<tr>
						<td width="300" valign="top" style="height: 14px">
						<font size="1" color="#4783C5" face="Verdana"><%=getadminGeneralLngStr("LtxtAvlVars")%></font></td>
						<td style="height: 14px">
						<font size="1" color="#4783C5" face="Verdana">
						<span dir="ltr">@SlpCode</span> = <%=getadminGeneralLngStr("LtxtAgentCode")%><br>
						<span dir="ltr">@Type</span> = <%=getadminGeneralLngStr("DtxtType")%> <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
						1: <%=getadminGeneralLngStr("LtxtBPSearch")%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 2: <%=getadminGeneralLngStr("LtxtStateAct")%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 3: 
						<%=getadminGeneralLngStr("DtxtComDocs")%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 4: <%=getadminGeneralLngStr("DtxtReceipts")%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 5: 
						<%=getadminGeneralLngStr("DtxtActivities")%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
						6: 
						<%=getadminGeneralLngStr("DtxtSO")%></font></td>
					</tr>
					</table>
			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminGeneralLngStr("LtxtViewDocFilter")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
			<img src="images/lentes.gif">
			<font color="#4783C5"><%=getadminGeneralLngStr("LtxtViewDocFilterNote")%></font></font></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td width="300" valign="top">
						<font face="Verdana" size="1" color="#4783C5"> 
						<%=getadminGeneralLngStr("DtxtQuery")%> - (DocEntry not in)</font></td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td rowspan="2">
									<textarea rows="10" name="ViewDocFilter" dir="ltr" cols="87" class="input" onkeydown="javascript:document.Form1.btnVerfyViewDocFilter.src='images/btnValidate.gif';document.Form1.btnVerfyViewDocFilter.style.cursor = 'hand';;document.Form1.valViewDocFilter.value='Y';"><% If Not IsNull(myApp.ViewDocFilter) Then %><%=Server.HTMLEncode(myApp.ViewDocFilter)%><% End If %></textarea>
								</td>
								<td valign="top">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteViewDocFilter" alt="<%=getadminGeneralLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(2, 'ViewDocFilter', -1, null);">
								</td>
							</tr>
							<tr>
								<td valign="bottom">
									<img src="images/btnValidateDis.gif" id="btnVerfyViewDocFilter" alt="<%=getadminGeneralLngStr("DtxtValidate")%>" onclick="javascript:if (document.Form1.valViewDocFilter.value == 'Y')VerfyViewDocFilter();">
									<input type="hidden" name="valViewDocFilter" value="N">
								</td>
							</tr>
						</table>
						</td>
					</tr>
	
					<tr>
						<td width="300" valign="top" style="height: 14px">
						<font size="1" color="#4783C5" face="Verdana"><%=getadminGeneralLngStr("LtxtAvlVars")%></font></td>
						<td style="height: 14px">
						<font size="1" color="#4783C5" face="Verdana">
						<span dir="ltr">@SlpCode</span> = <%=getadminGeneralLngStr("LtxtAgentCode")%><br>
						<span dir="ltr">@CardCode</span> = <%=getadminGeneralLngStr("DtxtBPCode")%><br>
						<span dir="ltr">@ObjectCode</span> = <%=getadminGeneralLngStr("DtxtObjCode")%></font></td>
					</tr>
					</table>
			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminGeneralLngStr("LttlAcctBal")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
			<img src="images/lentes.gif">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LttlAcctBalNote")%></font></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td width="325"><img src="images/ganchito.gif">
						<font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtECDays")%></font></td>
						<td>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input type="text" name="ECDays" id="ECDays" size="5" value='<%=myApp.ecdays%>' class="input" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);"></td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnECDaysUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnECDaysDown"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						<script language="javascript">NumUDAttach('Form1', 'ECDays', 'btnECDaysUp', 'btnECDaysDown');</script>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<font color="#4783C5" size="1" face="Verdana"><strong><%=getadminGeneralLngStr("DtxtAgents")%></strong></font></td>
						<td bgcolor="#F7FBFF">
						&nbsp;
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						<font color="#4783C5"><%=getadminGeneralLngStr("LtxtShowCxcOpenInvBy")%></font></font></td>
						<td bgcolor="#F7FBFF">
						<p>
						<select size="1" name="showCxcOpenInvBy" class="input" onchange="javascript:document.Form1.showCxcIncTrans.disabled=this.value == 'DocDate';document.Form1.showCxcDueDate.disabled=this.value=='DocDueDate';">
						<option <% If myApp.showCxcOpenInvBy = "DocDate" Then %>selected<% End If %>="" value="DocDate">
						<%=getadminGeneralLngStr("LoptDocDate")%></option>
						<option <% If myApp.showCxcOpenInvBy = "DocDueDate" Then %>selected<% End If %>="" value="DocDueDate">
						<%=getadminGeneralLngStr("LtxtDocDueDate")%></option>
						</select></p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="showCxcDueDate" value="Y" <% If myApp.showCxcDueDate or myApp.showCxcOpenInvBy = "DocDueDate" Then %>checked <%end if %> <% If myApp.showCxcOpenInvBy = "DocDueDate" Then %> disabled<% End If %> id="showCxcDueDate" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="showCxcDueDate"><%=getadminGeneralLngStr("LtxtShowCxcDueDate")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input <% If myApp.showCxcOpenInvBy = "DocDate" Then %>disabled<% End If %> type="checkbox" name="showCxcIncTrans" value="Y" <% If myApp.showCxcIncTrans and myApp.showCxcOpenInvBy = "DocDueDate" Then %>checked<%end if %> id="showCxcIncTrans" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="showCxcIncTrans"><%=getadminGeneralLngStr("LtxtShowCxcIncTrans")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="showCxcOpenInv" value="Y" <% If myApp.showCxcOpenInv Then %>checked<%end if %> id="showCxcOpenInv" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="showCxcOpenInv"><%=getadminGeneralLngStr("LtxtShowCxcOpenInv")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<font color="#4783C5" size="1" face="Verdana"><strong><%=getadminGeneralLngStr("DtxtClients")%></strong></font></td>
						<td bgcolor="#F7FBFF">
						&nbsp;
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						<font color="#4783C5"><%=getadminGeneralLngStr("LtxtShowCxcOpenInvBy")%></font></font></td>
						<td bgcolor="#F7FBFF">
						<p>
						<select size="1" name="showCxcOpenInvByC" class="input" onchange="javascript:document.Form1.showCxcIncTransC.disabled=this.value == 'DocDate';document.Form1.showCxcDueDateC.disabled=this.value=='DocDueDate';">
						<option <% If myApp.showCxcOpenInvByC = "DocDate" Then %>selected<% End If %>="" value="DocDate">
						<%=getadminGeneralLngStr("LoptDocDate")%></option>
						<option <% If myApp.showCxcOpenInvByC = "DocDueDate" Then %>selected<% End If %>="" value="DocDueDate">
						<%=getadminGeneralLngStr("LtxtDocDueDate")%></option>
						</select></p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="showCxcDueDateC" value="Y" <% If myApp.showCxcDueDateC or myApp.showCxcOpenInvByC = "DocDueDate" Then %>checked <%end if %> <% If myApp.showCxcOpenInvByC = "DocDueDate" Then %> disabled<% End If %> id="showCxcDueDateC" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="showCxcDueDateC"><%=getadminGeneralLngStr("LtxtShowCxcDueDate")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input <% If myApp.showCxcOpenInvByC = "DocDate" Then %>disabled<% End If %> type="checkbox" name="showCxcIncTransC" value="Y" <% If myApp.showCxcIncTransC and myApp.showCxcOpenInvByC = "DocDueDate" Then %>checked<%end if %> id="showCxcIncTransC" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="showCxcIncTransC"><%=getadminGeneralLngStr("LtxtShowCxcIncTrans")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="showCxcOpenInvC" value="Y" <% If myApp.showCxcOpenInvC Then %>checked<%end if %>="" id="showCxcOpenInvC" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="showCxcOpenInvC"><%=getadminGeneralLngStr("LtxtShowCxcOpenInv")%></label></font></td>
						<td bgcolor="#F7FBFF">
						<p>&nbsp;</p>
						</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminGeneralLngStr("DtxtLog")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
			<img src="images/lentes.gif">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminGeneralLngStr("LtxtLogDesc")%></font></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="EnableCLogLogin" value="Y" <% If myApp.EnableCLogLogin Then %>checked<%end if %> id="EnableCLogLogin" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCLogLogin"><%=getadminGeneralLngStr("LEnableCLogLogin")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="EnableVLogLogin" value="Y" <% If myApp.EnableVLogLogin Then %>checked<% end if %> id="EnableVLogLogin" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableVLogLogin"><%=getadminGeneralLngStr("LEnableVLogLogin")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="EnableCSearchFilterLog" value="Y" <% If myApp.EnableCSearchFilterLog Then %>checked<%end if %> id="EnableCSearchFilterLog" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCSearchFilterLog"><%=getadminGeneralLngStr("LEnableCSearchFilterL")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="EnableCSearchItemLog" value="Y" <% If myApp.EnableCSearchItemLog Then %>checked<%end if %> id="EnableCSearchItemLog" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCSearchItemLog"><%=getadminGeneralLngStr("LEnableCSearchItemLog")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="EnableCItemViewLog" value="Y" <% If myApp.EnableCItemViewLog Then %>checked<%end if %> id="EnableCItemViewLog" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCItemViewLog"><%=getadminGeneralLngStr("LEnableCItemViewLog")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font size="1" face="Verdana">
						</font>
						<input type="checkbox" name="EnableCItemPurLog" value="Y" <% If myApp.EnableCItemPurLog Then %>checked<%end if %> id="EnableCItemPurLog" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCItemPurLog"><%=getadminGeneralLngStr("LEnableCItemPurLog")%></label></font></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminGeneralLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	<input type="hidden" name="submitCmd" value="adminnew">
</form>
<script language="javascript">
var btnImg;
var fldVal;
function VerfyFilter()
{
	document.frmVerfyQuery.Query.value = document.Form1.AgentClientsFilter.value;
	document.frmVerfyQuery.type.value = 'AgentClientsFilter';
	btnImg = document.Form1.btnVerfyFilter;
	fldVal = document.Form1.valAgentClientsFilter;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyViewDocFilter()
{
	document.frmVerfyQuery.Query.value = document.Form1.ViewDocFilter.value;
	document.frmVerfyQuery.type.value = 'ViewDocFilter';
	btnImg = document.Form1.btnVerfyViewDocFilter;
	fldVal = document.Form1.valViewDocFilter;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyQueryVerified()
{
	btnImg.src='images/btnValidateDis.gif'
	btnImg.style.cursor = '';
	fldVal.value='N';
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="AgentClientsFilter">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<!--#include file="bottom.asp" -->