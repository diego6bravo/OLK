<!--#include file="lang/home.asp" -->
<!--#include file="taskMonitorData.asp"-->
<% sql = 	"select Count('A') from olkomsg " & _
			"inner join olkmsg1 on olkmsg1.olklog = olkomsg.olklog " & _
			"where OlkUser = '" & Session("vendid") & "' and OlkStatus = 'N'"
set rs = conn.execute(sql)
msgCount = rs(0)
rs.close
%>
<table border="0" cellpadding="0" cellspacing="2" width="100%">
	<tr>
		<td><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=gethomeLngStr("LtxtWelcome")%>&nbsp;<%=Session("vendnm")%>
          </font></b></td>
	</tr>
	<tr>
		<td>
		<div align="center">
			<table border="0" cellpadding="0" style="width: 100%">
				<% If (myApp.EnableORDR or myApp.EnableOQUT) Then %>
				<tr>
					<td style="width: 25px">
					&nbsp;</td>
					<td width="18">
					<p align="center"><font size="2" face="Verdana">
					<a href="operaciones.asp?cmd=pendientes">
					<img border="0" src="images/icon_pendiente.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href='operaciones.asp?cmd=pendientes&amp;SlpCodeFrom=<%=Session("vendnm")%>&amp;SlpCodeTo=<%=Session("vendnm")%>'>
					<font color="#000000"><%=gethomeLngStr("LtxtOpenDocs")%> (<%=GetTaskMonitorInfo(5)%>)</font></a></font></td>
				</tr>
				<% End If %>
				<% If myAut.HasAuthorization(1) Then %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18"><font size="2" face="Verdana">
					<a href="operaciones.asp?cmd=slist">
					<img border="0" src="images/icon_catalogo.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href="operaciones.asp?cmd=slist">
					<font color="#000000"><%=gethomeLngStr("DtxtCat")%></font></a></font></td>
				</tr>
				<% End If %>
				<% 
				crd1 = myAut.HasAuthorization(23)
				crd2 = myAut.HasAuthorization(75)
				If crd1 or crd2 or myApp.EnableOCRD Then %>
				<tr>
					<td style="width: 25px">
					&nbsp;</td>
					<td width="18">
					<p align="center"><font size="2" face="Verdana">
					<a href="operaciones.asp?cmd=searchclient">
					<img border="0" src="images/icon_busquedacliente.jpg"></a></font></td>
					<td><font size="1" face="Verdana">
					<a href="operaciones.asp?cmd=home&amp;subCmd=crd"><font color="#000000"><%=gethomeLngStr("DtxtClients")%></font></a></font></td>
				</tr>
				<% If Request("subCmd") = "crd" Then %>
				<tr>
					<td style="width: 25px; height: 34px;"></td>
					<td width="18" style="height: 34px"></td>
					<td style="height: 34px">
					<table cellpadding="0" cellspacing="0" border="0">
						<% If myApp.EnableOCRD Then %>
						<tr>
							<td>
							<a href="newclientnow.asp"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<a href="newclientnow.asp"><font color="#000000"><%=gethomeLngStr("LtxtNewClient")%></font></a></font></td>
						</tr>
						<tr>
							<td>
							<a href="operaciones.asp?cmd=pendClients"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<a href="operaciones.asp?cmd=pendClients"><font color="#000000"><%=gethomeLngStr("LtxtClientPendList")%></font></a></font></td>
						</tr>
						<% End If %>
						<% If crd1 = crd2 Then %>
						<tr>
							<td>
							<a href="operaciones.asp?cmd=searchclient"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<a href="operaciones.asp?cmd=searchclient"><font color="#000000"><%=gethomeLngStr("LtxtClientSearch")%></font></a></font></td>
						</tr>
						<% End If %>
					</table>
					</td>
				</tr>
				<% End If %>
				<% End If %>
				<% If myApp.EnableOCLG Then %>
				<tr>
					<td style="width: 25px">
					&nbsp;</td>
					<td width="18">
					<p align="center"><font size="2" face="Verdana">
					<a href="operaciones.asp?cmd=openActivities&amp;SlpCodeFrom=<%=Session("vendnm")%>&amp;SlpCodeTo=<%=Session("vendnm")%>">
					<img border="0" src="images/icon_actividades.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href='operaciones.asp?cmd=openActivities&amp;SlpCodeFrom=<%=Session("vendnm")%>&amp;SlpCodeTo=<%=Session("vendnm")%>'>
					<font color="#000000"><%=gethomeLngStr("LtxtOpenActivities")%> (<%=GetTaskMonitorInfo(3)%>)</font></a></font></td>
				</tr>
				<% End If %>
				<% inv1 = myAut.HasAuthorization(10) 'Inventory Count
				inv2 = myAut.HasInAuthorization 'Incoming Inventory
				inv3 = myAut.HasOutAuthorization 'Outgoing Inventory
				If inv1 or inv2 or inv3 Then %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18"><font size="2" face="Verdana">
					<a href='operaciones.asp?cmd=home<% If Request("subCmd") <> "inv" Then %>&amp;subCmd=inv<% End If %>'>
					<img border="0" src="images/icon_inventario.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href='operaciones.asp?cmd=home<% If Request("subCmd") <> "inv" Then %>&amp;subCmd=inv<% End If %>'>
					<font color="#000000"><%=gethomeLngStr("LtxtInv")%></font></a></font></td>
				</tr>
				<% If Request("subCmd") = "inv" Then %>
				<% If inv1 Then %>
				<tr>
					<td style="width: 25px; height: 34px;"></td>
					<td width="18" style="height: 34px"></td>
					<td style="height: 34px">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<a href="operaciones.asp?cmd=inv&amp;redir=inv"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<a href="operaciones.asp?cmd=inv&amp;redir=inv"><font color="#000000"><%=gethomeLngStr("LtxtInvRecount")%></font></a></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<% End If %>
				<% If inv2 Then %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18">&nbsp;</td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<a href="operaciones.asp?cmd=inv&amp;redir=invChkInOut&amp;Type=I"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<a href="operaciones.asp?cmd=inv&amp;redir=invChkInOut&amp;Type=I"><font color="#000000"><%=gethomeLngStr("LtxtPurchaseOrderChec")%></font></a></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<% End If %>
				<% If inv3 Then %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18">&nbsp;</td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<a href="operaciones.asp?cmd=inv&amp;redir=invChkInOut&amp;Type=O"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<a href="operaciones.asp?cmd=inv&amp;redir=invChkInOut&amp;Type=O"><font color="#000000"><%=gethomeLngStr("LtxtSalesOrderCheck")%></font></a></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<% End If %>
				<% End If %>
				<% End If
				sql = 	"select T0.SecID, IsNull(T1.AlterSecName, T0.SecName) SecName, T0.Type, " & _
						"Case T0.Type When 'L' Then T0.SecContent Else '' End Link, Case T0.Type When 'R' Then T0.SecContent Else '' End rsIndex, " & _
						"Case T0.Type When 'R' Then (select Count('') from OLKRSVars where rsIndex = Convert(int,Convert(nvarchar(100),T0.SecContent))) Else 0 End rsVarCount " & _
						"from OLKSections T0 " & _
						"left outer join OLKSectionsAlterNames T1 on T1.SecType = T0.SecType and T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
						"where T0.SecType = 'U' and T0.UserType = 'P' and T0.Status = 'A' and HideMainMenu = 'N' "

				If Session("useraccess") = "U" Then
					If myAut.AuthorizedForms <> "" Then
						sql = sql & "and T0.SecID in (" & myAut.AuthorizedForms & ") "
					Else
						sql = sql & "and 1 = 2 "
					End If
				End If

				set rs = conn.execute(sql)
				If Not rs.eof Then %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18"><font size="2" face="Verdana">
					<a href='operaciones.asp?cmd=home<% If Request("subCmd") <> "sec" Then %>&amp;subCmd=sec<% End If %>'>
					<img border="0" src="images/icon_sections.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href='operaciones.asp?cmd=home<% If Request("subCmd") <> "sec" Then %>&amp;subCmd=sec<% End If %>'>
					<font color="#000000"><%=gethomeLngStr("DtxtForms")%></font></a></font></td>
				</tr>
				<% If Request("subCmd") = "sec" Then
				do while not rs.eof
				%>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18">&nbsp;</td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<a href='operaciones.asp?cmd=sec&amp;secID=<%=rs("SecID")%>'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>
							</td>
							<td><font size="1" face="Verdana">
							<% Select Case rs("Type")
								Case "N"
									myLink = "?cmd=sec&secID=" & rs("SecID")
								Case "L"
									myLink = rs("Link")
								Case "R"
									myLink = "javascript:goRep(" & rs("rsIndex") & "," & rs("rsVarCount") & ");"
							End Select
							%>
							<a href="<%=myLink%>" <% If rs("Type") = "L" Then %>target="_blank"<% End If %>><font color="#000000"><%=rs("SecName")%></font></a></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<% rs.movenext
				loop
				End If
				End If %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18"><font size="2" face="Verdana">
					<a href="operaciones.asp?cmd=buzon">
					<img border="0" src="images/icon_buzon.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href="operaciones.asp?cmd=buzon">
					<font color="#000000"><%=gethomeLngStr("LtxtMyInbox")%><% If msgCount > 0 Then %> (<%=msgCount%>)<% End If %></font></a></font></td>
				</tr>
				<% If myAut.AuthorizedRepGroups <> "" or Session("useraccess") = "P" Then %>
				<tr>
					<td style="width: 25px">&nbsp;</td>
					<td width="18"><font size="2" face="Verdana">
					<a href="operaciones.asp?cmd=reportes">
					<img border="0" src="images/icon_reporte.jpg"></a></font></td>
					<td><font size="1" face="Verdana"><a href="operaciones.asp?cmd=reportes">
					<font color="#000000"><%=gethomeLngStr("LtxtMyReps")%></font></a></font></td>
				</tr>
				<% End If %>
			</table>
		</div>
		</td>
	</tr>
</table>

<script language="javascript">
function goRep(rsIndex, varsCount)
{
	if (varsCount == 0) document.frmReps.cmd.value='viewRep';
	else document.frmReps.cmd.value='viewRepVals';
	document.frmReps.rsIndex.value = rsIndex;
	document.frmReps.submit();
}
</script>
<form method="POST" action="operaciones.asp" name="frmReps">
<input type="hidden" name="cmd" value="viewRepVals">
<input type="hidden" name="rsIndex" value="">
</form>