<!--#include file="top.asp" -->
<!--#include file="lang/adminCustomSearch.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("OLKDB") & "]") %>

<script language="javascript" src="js_up_down.js"></script>

<table border="0" cellpadding="0" width="100%">
	<% 
	sql = "select ID, Name, Status, Ordr, Case When Convert(nvarchar(100),Query) = '' Then 'disabled' Else '' End [Disabled] from OLKCustomSearch where ObjectCode = " & Request("ObjID") & " and Status <> 'D' order by Ordr"
	rs.open sql, conn, 3, 1 %>
	<script language="javascript">
	function valFrm()
	{
		rowName = document.frmCustomSearch.rowName;
		if (rowName.length)
		{
			for (var i = 0;i<rowName.length;i++)
			{
				if (rowName[i].value == '')
				{
					alert("<%=getadminCustomSearchLngStr("LtxtValAlrNoNam")%>");
					rowName[i].focus();
					return false;
				}
			}
		}
		else
		{
			if (rowName.value == '')
			{
				alert("<%=getadminCustomSearchLngStr("LtxtValAlrNoNam")%>");
				rowName.focus();
				return false;
			}
		}
		return true;
	}
	</script>
	<form method="POST" action="adminsubmit.asp" name="frmCustomSearch" onsubmit="javascript:return valFrm();">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font size="1" face="Verdana" color="#31659C"><% Select Case CInt(Request("ObjID"))
			Case 2 %><%=getadminCustomSearchLngStr("LtxtCustomSearchCL")%>
			<% Case 4 %><%=getadminCustomSearchLngStr("LttlCustomSearch")%>
			<% End Select %>
			</font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			<font color="#4783C5"><% Select Case CInt(Request("ObjID"))
			Case 2 %><%=getadminCustomSearchLngStr("LttlCustomSearchNoteC")%>
			<% Case 4 %><%=getadminCustomSearchLngStr("LttlCustomSearchNote")%>
			<% End Select %></font></font></p>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table12">
				<tr>
					<td align="center" bgcolor="#E2F3FC" style="width: 16px">&nbsp;</td>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font size="1" face="Verdana" color="#31659C"><%=getadminCustomSearchLngStr("DtxtName")%>
					</font></b></td>
					<td align="center" bgcolor="#E2F3FC" style="width: 100px">
					<b><font size="1" face="Verdana" color="#31659C"><%=getadminCustomSearchLngStr("DtxtOrder")%>
					</font></b></td>
					<td align="center" bgcolor="#E2F3FC" style="width: 100px">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminCustomSearchLngStr("DtxtActive")%>
					</font></b></td>
					<td align="center" bgcolor="#E2F3FC" width="16">&nbsp;</td>
				</tr>
				<%
			do while not rs.eof
			ID = Replace(rs("ID"), "-", "_") %>
				<tr bgcolor="#F3FBFE">
					<td valign="top" style="width: 16px; padding-top: 4px">
					<a href="adminCustomSearchEdit.asp?ID=<%=rs("ID")%>&ObjID=<%=Request("ObjID")%>">
					<img border="0" src='images/<%=Session("rtl")%>flechaselec.gif'></a></td>
					<td valign="top">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td>
							<input style="width: 100%;" class="input" size="20" maxlength="50" value='<%=Server.HTMLEncode(RS("Name"))%>' name="rowName<%=ID%>" id="rowName" onkeydown="return chkMax(event, this, 50);">
							</td>
							<td width="16">
							<a href="javascript:doFldTrad('CustomSearch', 'ID', <%=rs("ID")%>, 'alterName', 'T', null);">
							<img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
						</tr>
					</table>
					</td>
					<td valign="top" style="width: 100px">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
							<input type="text" name="RowOrder<%=ID%>" id="RowOrder<%=ID%>" size="7" style="text-align: right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value='<%=rs("Ordr")%>'>
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td>
									<img src="images/img_nud_up.gif" id="btnRowOrder<%=ID%>Up"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td>
									<img src="images/img_nud_down.gif" id="btnRowOrder<%=ID%>Down"></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</td>
					<td style="width: 100px">
					<p align="center">
					<input <% If rs("Status") = "Y" Then %>checked<% End If %> <%=rs("Disabled")%> type="checkbox" name="rowActive<%=ID%>" value="Y" class="noborder"></p>
					</td>
					<td valign="middle" width="16">
					<% If rs("ID") >= 0 Then %>
					<a href="javascript:if(confirm('<%=getadminCustomSearchLngStr("LtxtConfRemSearch")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("Name")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&ObjID=<%=Request("ObjID")%>&ID=<%=rs("ID")%>&submitCmd=adminCustomSearch';">
					<img border="0" src="images/remove.gif" width="16" height="16"></a><% End If %></td>
				</tr>
				<input type="hidden" name="ID" value="<%=rs("ID")%>">
				<script language="javascript">NumUDAttach('frmCustomSearch', 'RowOrder<%=ID%>', 'btnRowOrder<%=ID%>Up', 'btnRowOrder<%=ID%>Down');</script>
				<% rs.movenext
				loop %>
			</table>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminCustomSearchLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td width="77">
					<input type="button" value="<%=getadminCustomSearchLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminCustomSearchEdit.asp?ObjID=<%=Request("ObjID")%>'"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="submitCmd" value="adminCustomSearch">
		<input type="hidden" name="ObjID" value="<%=Request("ObjID")%>">
		<input type="hidden" name="cmd" value="u">
	</form>
</table>
<!--#include file="bottom.asp" -->