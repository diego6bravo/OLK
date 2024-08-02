<% addLngPathStr = "" %>
<!--#include file="lang/taskMonitor.asp" -->
<!--#include file="taskMonitorData.asp"-->
<%
set rd = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetInformer" & Session("ID")
cmd.Parameters.Refresh()
cmd("@UserType") = userType
cmd("@UserAccess") = Session("UserAccess")
If Session("UserAccess") = "U" Then cmd("@Groups") = myAut.AuthorizedRepGroups
rd.open cmd, , 3, 1 %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td class="TablasTituloSec" id="tdMyTtl" colspan="3">
		<img border="0" src="design/0/images/newsTitle_icon.gif">&nbsp;<%=gettaskMonitorLngStr("LtxtTaskMon")%></td>
	</tr>
	<% 
	do while not rd.eof
	Select Case rd("Type")
		Case "S"
			Select Case rd("ID")
				Case 0
					If Session("useraccess") = "P" or Session("HasActionConfAut") Then
					confCount = CInt(GetTaskMonitorInfo(1)) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS0" style="cursor: hand;<% If confCount = 0 Then %>display: none;<% End If %>" onclick="javascript:doMyLink('executeConf.asp', 'Type=A', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtMyPendActToConf")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="0"></td>
						<td align="right" id="tdTaskMonValS0"><%=confCount%></td>
					</tr>
					<% 
					End If
				Case 8
					If myAut.HasAuthorization(114) or myAut.HasAuthorization(116) or myAut.HasAuthorization(118) Then
					confCount = GetTaskMonitorInfo(9) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS8" style="cursor: hand;<% If confCount = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeConf.asp', 'Type=C', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtMyPendBPToConf")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="8"></td>
						<td align="right" id="tdTaskMonValS8"><%=confCount%></td>
					</tr>
					<% 
					End If
				Case 9
					If myAut.HasAuthorization(121) Then
					confCount = CInt(GetTaskMonitorInfo(10)) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS9" style="cursor: hand;<% If confCount = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeConf.asp', 'Type=I', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtMyPendItmToConf")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="9"></td>
						<td align="right" id="tdTaskMonValS9"><%=confCount%></td>
					</tr>
					<% 
					End If
				Case 10
					If myAut.HasAuthorization(126) Then
					confCount = CInt(GetTaskMonitorInfo(11)) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS10" style="cursor: hand;<% If confCount = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeConf.asp', 'Type=R', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtMyPendRctToConf")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="10"></td>
						<td align="right" id="tdTaskMonValS10"><%=confCount%></td>
					</tr>
					<% 
					End If
				Case 11
					If Session("useraccess") = "P" or Session("HasComDocConf") Then
					confCount = CInt(GetTaskMonitorInfo(12)) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS11" style="cursor: hand;<% If confCount = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeConf.asp', 'Type=D', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtMyPendCDToConf")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="11"></td>
						<td align="right" id="tdTaskMonValS11"><%=confCount%></td>
					</tr>
					<% 
					End If
				Case 7 
					myConf = CInt(GetTaskMonitorInfo(8)) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS7" style="<% If myConf = 0 Then %>display: none;<% End If %> ">
						<td style="width: 15px">&nbsp;</td>
						<td><%=gettaskMonitorLngStr("LtxtMyPendObjToConf")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="7"></td>
						<td align="right" id="tdTaskMonValS7"><%=myConf%></td>
					</tr>
					<% 
				Case 1
					If EnableClientActivation and myAut.HasAuthorization(89) Then %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:doMyLink('activation.asp', '', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtAnonRegAct")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="1"></td>
						<td align="right" id="tdTaskMonValS1"><%=GetTaskMonitorInfo(2)%></td>
					</tr>
					<%
					End If
				Case 2
					If myApp.EnableOCLG Then  %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:doMyLink('searchOpenedActivities.asp', 'SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtOpenActivities")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="2"></td>
						<td align="right" id="tdTaskMonValS2"><%=GetTaskMonitorInfo(3)%></td>
					</tr>
					<% 
					End If
				Case 3
					If myApp.EnableOCLG Then
					nextAct = GetTaskMonitorInfo(4)
					
					If nextAct(1) <> -1 Then
						Select Case nextAct(0)
							Case "O"
								strAction = "addActivity/goActivity.asp"
							Case "S"
								strAction = "addActivity/goEditActivity.asp"
						End Select
					End If %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS3" style="cursor: hand;<% If nextAct(1) = -1 Then %>display: none;<% End If %>" onclick="javascript:document.frmGoAct.submit();">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtNextActivity")%></td>
						<td align="right">
						<table cellpadding="0" cellspacing="0" border="0">
							<form name="frmGoAct" method="post" action="<%=strAction%>">
				              	<tr class="TablasNoticias">
				              		<td id="tdTaskMonValS3"><%=nextAct(1)%></td>
				              		<td style="width: 13px; "><img id="imgTaskMonS3" src="images/icon_activity_<%=nextAct(0)%>.gif">
					              	<input type="hidden" name="LogNum" value="<%=nextAct(1)%>"><input type="hidden" name="Card" value="<%=Server.HTMLEncode(nextAct(2))%>">
					              	<input type="hidden" name="ClgCode" value="<%=nextAct(1)%>"><input type="hidden" name="CardCode" value="<%=Server.HTMLEncode(nextAct(2))%>"></td>
				              	</tr>
				            </form>
				        </table><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="3"></td>
					</tr>
					<% 
					End If
				Case 12
					If myApp.EnableOOPR Then  %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:doMyLink('searchOpenedSO.asp', 'SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtOpenSO")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="2"></td>
						<td align="right" id="tdTaskMonValS2"><%=GetTaskMonitorInfo(13)%></td>
					</tr><% 
					End If
				Case 4
					openDocs = GetTaskMonitorInfo(5)
					If openDocs > -1 Then %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:doMyLink('searchOpenedDocs.asp', 'orden1=0&orden2=A&SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtOpenDocs")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="4"></td>
						<td align="right" id="tdTaskMonValS4"><%=openDocs%></td>
					</tr>
					<% 
					End If
				Case 5
					openPolls = GetTaskMonitorInfo(6)
					If openPolls <> "-1" Then %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:doMyLink('extPollList.asp', '', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtOpenPolls")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="5"></td>
						<td align="right"><nobr><span id="tdTaskMonValS5"><%=rs("Count")%>&nbsp;(<%=rs("Pending")%>)</span></nobr></td>
					</tr>
					<% 
					End If
				Case 6
					
					If optOfert and myAut.HasAuthorization(7) Then
					openOffers = GetTaskMonitorInfo(7) %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:doMyLink('ofertsMan.asp', 'OfertStatus=W, O', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=Replace(gettaskMonitorLngStr("LtxtOffersWaitAns"), "{0}", txtOferts)%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="6"></td>
						<td align="right"><nobr><span id="tdTaskMonValS6"><%=openOffers(0)%>&nbsp;/&nbsp;<%=openOffers(1)%></span></nobr></td>
					</tr>
					<%
					End If
				Case 13 %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS13" style="cursor: hand;<% If autCount_O = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeAut.asp', 'Type=A', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></td>
						<td><%=gettaskMonitorLngStr("LtxtAutAction")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="13"></td>
						<td align="right" id="tdTaskMonValS13"><%=autCount_O%></td>
					</tr>
					<% 
				Case 14 %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS14" style="cursor: hand;<% If autCount_C = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeAut.asp', 'Type=C', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="14"></td>
						<td><%=gettaskMonitorLngStr("LtxtAutBP")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="14"></td>
						<td align="right" id="tdTaskMonValS14"><%=autCount_C%></td>
					</tr>
					<% 
				Case 15 %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS15" style="cursor: hand;<% If autCount_I = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeAut.asp', 'Type=I', '_self');">
						<td style="width: 15px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="15"></td>
						<td><%=gettaskMonitorLngStr("LtxtAutItm")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="15"></td>
						<td align="right" id="tdTaskMonValS15"><%=autCount_I%></td>
					</tr>
					<% 
				Case 16 %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS16" style="cursor: hand;<% If autCount_R = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeAut.asp', 'Type=R', '_self');">
						<td style="width: 16px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="16" height="16"></td>
						<td><%=gettaskMonitorLngStr("LtxtAutRec")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="16"></td>
						<td align="right" id="tdTaskMonValS16"><%=autCount_R%></td>
					</tr>
					<% 
				Case 17 %>
					<tr class="TablasNoticias" onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" id="trTaskMonS17" style="cursor: hand;<% If autCount_D = 0 Then %>display: none;<% End If %> " onclick="javascript:doMyLink('executeAut.asp', 'Type=D', '_self');">
						<td style="width: 17px"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="17" height="17"></td>
						<td><%=gettaskMonitorLngStr("LtxtAutDoc")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="S"><input type="hidden" name="TaskMonID" id="TaskMonID" value="17"></td>
						<td align="right" id="tdTaskMonValS17"><%=autCount_D%></td>
					</tr>
					<% 
			End Select
		Case "U"
			strAlign = ""
			Select Case rd("Align")
				Case "C"
					strAlign = "Center"
				Case "R"
					strAlign = "Right"
				Case "L"
					strAlign = "Left"
			End Select
			
			strQuery = "declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
						"declare @LanID int set @LanID = " & Session("LanID") & " select (" & rd("Query") & ")"
			strQuery = QueryFunctions(strQuery)
			set rs = conn.execute(strQuery)
			If Not rs.Eof Then 
				strValue = rs(0) 
				If rd("HideNull") = "Y" Then
					If IsNumeric(rs(0)) Then
						Hide = (IsNull(rs(0)) or rs(0) = 0)
					Else
						Hide = IsNull(rs(0)) or rs(0) = ""
					End If
				Else
					Hide = False
				End If
			Else 
				If rd("HideNull") = "Y" Then Hide = True
				strValue = ""
			End If %>
					<tr class="TablasNoticias" id="trTaskMonU<%=rd("ID")%>" <% If Hide Then %>style="display: none; "<% End If %>  <% If Not IsNull(rd("rsIndex")) Then %>onmouseover="this.className = 'hlt';" onmouseout="this.className = 'TablasNoticias';" style="cursor: hand; " onclick="javascript:goRep(<%=rd("rsIndex")%>,<%=rd("rsVarsCount")%>);"<% End If %>>
						<td style="width: 15px"><% If Not IsNull(rd("rsIndex")) Then %><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"><% Else %>&nbsp;<% End If %></td>
						<td><%=rd("Name")%><input type="hidden" id="TaskMonType" name="TaskMonType" value="U"><input type="hidden" name="TaskMonID" id="TaskMonID" value="<%=rd("ID")%>"><input type="hidden" id="TaskMonHideNullU<%=rd("ID")%>" value="<%=rd("HideNull")%>"></td>
						<td <% If strAlign <> "" Then %> align="<%=strAlign%>"<% End If %>><nobr><span id="tdTaskMonValU<%=rd("ID")%>"><%=strValue%></span></nobr></td>
					</tr>
		<% End Select %><%	
		rd.movenext
		loop %>
					<tr class="TablasNoticias">
						<td colspan="3" align="center"><input type="button" name="btnRefreshMon" value="<%=gettaskMonitorLngStr("DtxtRefresh")%>" onclick="javascript:forceRefreshMon();"></td>
					</tr>
		</table>
<script language="javascript" src="taskMonitor.js"></script>