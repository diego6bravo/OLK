<!--#include file="top.asp" -->
<!--#include file="lang/adminFooter.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<script type="text/javascript">txtValGrmNam = '<%=getadminFooterLngStr("LtxtValGrmNam")%>';</script>
<script type="text/javascript" src="adminFooter.js"></script>
<% If Request("GroupID") = "" Then %>
<form name="frmGroups" method="post" action="adminFooterSubmit.asp" onsubmit="return valFrm();">
<input type="hidden" name="cmd" value="grp">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminFooterLngStr("LttlFooter")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminFooterLngStr("LttlFooterNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0">
			<tr class="TblRepTltSub">
				<td align="center" colspan="3"><%=getadminFooterLngStr("DtxtGroup")%>&nbsp;</td>
			</tr>
			<%
			sql = "select GroupID, GroupName from OLKFooterGroups"
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			do while not rs.eof %>
			<tr class="TblRepTbl">
			  <td width="15"><input type="hidden" name="GroupID" value="<%=rs("GroupID")%>">
				<a href="adminFooter.asp?GroupID=<%=rs("GroupID")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" id="GroupName<%=rs("GroupID")%>" name="GroupName<%=rs("GroupID")%>" size="100" value="<%=Server.HTMLEncode(rs("GroupName"))%>" maxlength="100"></td>
						<td><a href="javascript:doFldTrad('FooterGroups', 'GroupID', <%=rs("GroupID")%>, 'alterGroupName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminFooterLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td style="width: 16px">
				<a href="javascript:if(confirm('<%=getadminFooterLngStr("LtxtConfDelGrp")%>'.replace('{0}', '<%=rs("GroupName")%>')))window.location.href='adminFooterSubmit.asp?cmd=remGrp&GroupID=<%=rs("GroupID")%>';">
				<img border="0" src="images/remove.gif"></a></td>
			</tr>
			<% rs.movenext
			loop
			If rs.recordcount < 4 Then %>
			<tr class="TblRepTbl">
			  <td width="15">&nbsp;</td>
				<td>
				<input type="hidden" name="GroupNameTrad">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="GroupName" size="100" maxlength="100"></td>
						<td><a href="javascript:doFldTrad('FooterGroups', 'GroupID', '', 'alterGroupName', 'T', document.frmGroups.GroupNameTrad);"><img src="images/trad.gif" alt="<%=getadminFooterLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td style="width: 16px">&nbsp;</td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminFooterLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<% Else %>
<script type="text/javascript" src="js_up_down.js"></script><%
GroupID = CInt(Request("GroupID"))
sql = "select GroupName from OLKFooterGroups"
set rs = conn.execute(sql) %>
<form name="frmGroupEdit" method="post" action="adminFooterSubmit.asp">
<input type="hidden" name="cmd" value="editGrp">
<input type="hidden" name="GroupID" value="<%=GroupID%>">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminFooterLngStr("LttlFooterGroup")%> - <%=rs("GroupName")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminFooterLngStr("LttlFooterNote")%>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" id="tblSec">
			<tr class="TblRepTltSub">
				<td align="center"><%=getadminFooterLngStr("LtxtSection")%></td>
				<td align="center"><%=getadminFooterLngStr("DtxtOrder")%></td>
				<td style="width: 16px">&nbsp;</td>
			</tr>
			<tbody id="tbSec">
			<% NewOrdr = 0
			sql = "select T0.SecID, T1.SecName, T0.OrderID " & _  
					"from OLKFooterGroupsLinks T0 " & _  
					"inner join OLKSections T1 on T1.SecType = 'U' and T1.UserType = 'C' and T1.SecID = T0.SecID " & _  
					"where T0.GroupID = " & GroupID & " " & _  
					"order by T0.OrderID "
				set rs = conn.execute(sql) 
				do while not rs.eof
				NewOrdr = CInt(rs("OrderID"))+1 %>
			<tr class="TblRepTbl" id="trSec<%=rs("SecID")%>">
				<td><%=rs("SecName")%><input type="hidden" name="SecID" value="<%=rs("SecID")%>"></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0" width="80">
					<tr>
						<td>
							<input type="text" name="OrderID<%=rs("SecID")%>" id="OrderID<%=rs("SecID")%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("OrderID")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnOrderID<%=rs("SecID")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnOrderID<%=rs("SecID")%>Down"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script type="text/javascript">NumUDAttach('frmGroupEdit', 'OrderID<%=rs("SecID")%>', 'btnOrderID<%=rs("SecID")%>Up', 'btnOrderID<%=rs("SecID")%>Down');</script>
				</td>
				<td style="width: 16px">
				<img border="0" src="images/remove.gif" onclick="javascript:if(confirm('<%=getadminFooterLngStr("LtxtConfDelSec")%>'.replace('{0}', '<%=Server.HTMLEncode(rs("SecName"))%>')))delSec(<%=rs("SecID")%>);"></td>
			</tr>
			<% rs.movenext
			loop %>
			</tbody>
			<tr class="TblRepTbl">
				<td style="width: 260px;"><select size="1" id="cmbSec" style="width: 100%;">
				<option></option>
				<% sql = "select SecID, SecName from OLKSections where UserType = 'C' and SecType = 'U' and SecID not in (select SecID from OLKFooterGroupsLinks) order by 2"
				set rs = conn.execute(sql)
				do while not rs.eof %><option value="<%=rs(0)%>"><%=rs(1)%></option><%
				rs.movenext
				loop %>
				</select></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0" width="80">
					<tr>
						<td>
							<input type="text" name="NewOrderID" id="NewOrderID" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=NewOrdr%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnNewOrderIDUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnNewOrderIDDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script type="text/javascript">NumUDAttach('frmGroupEdit', 'NewOrderID', 'btnNewOrderIDUp', 'btnNewOrderIDDown');</script>
				</td>
				<td style="width: 16px">
				<input type="button" id="btnAdd" value="+" onclick="doAdd();"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminFooterLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminFooterLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<% End If %>
<!--#include file="bottom.asp" -->