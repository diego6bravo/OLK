<!--#include file="top.asp" -->
<!--#include file="lang/adminLayout.asp" -->
<br>
<% GroupID = -99
GroupID = CInt(Request("GroupID")) %>
<script language="javascript" src="js_up_down.js"></script>
<script language="javascript" src="adminLayout.js"></script>
<form name="frmLayout" method="post" action="adminLayoutSubmit.asp">
<input type="hidden" name="cmd" value="updLay">
<input type="hidden" name="GroupID" value="<%=GroupID%>">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminLayoutLngStr("LttlLayout")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminLayoutLngStr("LttlLayoutNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0">
			<tr>
				<td bgcolor="#E2F3FC" style="width: 200px"><b>
				<font size="1" face="Verdana" color="#31659C">
				<%=getadminLayoutLngStr("DtxtGroup")%>&nbsp;</font></b></td>
				<td valign="top" bgcolor="#F3FBFE">
				<select size="1" name="cmbGroup" class="input" onchange="javascript:window.location.href='adminLayout.asp?GroupID='+this.value;">
				<option></option>
				<% sql = "select ID, Name from OLKLayout"
				set rs = conn.execute(sql)
				do while not rs.eof
				Select Case CInt(rs("ID"))
					Case -1
						strName = getadminLayoutLngStr("LtxtHomePage")
					Case -2
						strName = getadminLayoutLngStr("LtxtGeneral")
					Case Else
						strName = rs("Name")
				End Select %><option <% If GroupID = CInt(rs("ID")) Then %>selected<% End If %> value="<%=rs("ID")%>"><%=strName%></option><%
				rs.movenext
				loop %>
				</select></td>
			</tr>
		</table>
		<% 
		If GroupID <> -99 Then %>
		<table border="0" cellpadding="0">
			<% sql = "select ColID, ColType from OLKLayoutClass where ID = " & GroupID
			set rs = conn.execute(sql)
			sql = "select T0.LineID, T0.[Type], T0.TypeID, T0.ColID, T0.RowID, T0.Active, " & _  
				"Case T0.[Type] " & _  
				"	When 0 Then (select GroupName from OLKBNGroups where GroupID = T0.TypeID) " & _  
				"	When 2 Then (select Name from OLKSmallCat where ID = T0.TypeID) " & _  
				"End TypeDesc " & _  
				"from OLKLayoutLines T0 " & _  
				"where T0.ID = " & GroupID & "  " & _  
				"order by T0.ColID, T0.RowID " 
			set rd = Server.CreateObject("ADODB.RecordSet")
			rd.open sql, conn, 3, 1
			do while not rs.eof
			NewID = 0
			 %>
			<tr>
				<td style="text-align: center;color:#31659C;font-size: x-small; font-face: Verdana;background-color:#E2F3FC;font-weight: bold;" colspan="4">
				<% select case CInt(rs("ColType"))
					Case 0 %>|L:txtNone|<%
					Case 1 %><%=getadminLayoutLngStr("DtxtContent")%><%
					Case 2%><%=getadminLayoutLngStr("LtxtSideBar")%><% 
				End Select %></td>
			</tr>
			<tr style="text-align: center;color:#31659C;font-size: x-small; font-face: Verdana;background-color:#E2F3FC;">
				<td><%=getadminLayoutLngStr("DtxtType")%></td>
				<td><%=getadminLayoutLngStr("DtxtActive")%></td>
				<td><%=getadminLayoutLngStr("DtxtOrder")%></td>
				<td>&nbsp;</td>
			</tr>
			<% rd.Filter = "ColID = " & rs("ColID")
			do while not rd.eof
				LineID = Replace(rd("LineID"), "-", "_")
				NewID = CInt(rd("RowID"))+1
				%><tr style="color:#31659C;background-color: #F3FBFE;font-size:x-small; font-face: Verdana;">
				<td id="txtLineDesc<%=LineID%>"><% select case rd("Type")
					Case 0 %><%=getadminLayoutLngStr("LtxtBanner")%><%
					Case 1 %><%=getadminLayoutLngStr("LtxtNews")%><%
					Case 2 %><%=getadminLayoutLngStr("LtxtSmallCat")%><%
					Case 3 %><%=getadminLayoutLngStr("LtxtSecIndex")%><%
					Case 4 %><%=getadminLayoutLngStr("LtxtCatNav")%><%
					Case 5 %><%=getadminLayoutLngStr("LtxtCartMinRep")%><%
					Case 6 %><%=getadminLayoutLngStr("LtxtNewsLetter")%><%
					Case 7 %><%=getadminLayoutLngStr("LtxtPolls")%><%
					Case 8 %><%=getadminLayoutLngStr("LtxtHomeOferts")%><%
				End Select
				If Not IsNull(rd("TypeDesc")) Then Response.Write "&nbsp;-&nbsp;" & rd("TypeDesc")%>
				</td>
				<td style="text-align: center;"><input type="checkbox" class="noborder" name="chkActive<%=LineID%>" value="Y" <% If rd("Active") = "Y" Then %>checked<% End If %>>
				</td>
				<td style="text-align: center;"><input type="hidden" name="LineID" value="<%=rd("LineID")%>">
				<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
								<input type="text" name="RowID<%=LineID%>" id="RowID<%=LineID%>" size="4" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" value="<%=rd("RowID")%>">
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnRowID<%=LineID%>Up"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnRowID<%=LineID%>Down"></td>
								</tr>
							</table></td>
						</tr>
					</table><script language="javascript">NumUDAttach('frmLayout', 'RowID<%=LineID%>', 'btnRowID<%=LineID%>Up', 'btnRowID<%=LineID%>Down');</script>
				</td>
				<td style="width: 16px">
					<a href="javascript:if(confirm('<%=getadminLayoutLngStr("LtxtConfDelLine")%>'.replace('{0}', document.getElementById('txtLineDesc<%=LineID%>').innerText)))window.location.href='adminLayoutSubmit.asp?cmd=delLine&GroupID=<%=GroupID%>&LineID=<%=rd("LineID")%>';">
					<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
			</tr>
			<% rd.movenext
			loop
			sql = "declare @ID int set @ID = " & GroupID & " " & _  
				"select T0.[Type], Convert(int,NULL) TypeID, TypeDesc " & _  
				"from ( " & _  
				"	select 1 [Type], N'" & Replace(getadminLayoutLngStr("LtxtNews"), "'", "''") & "' TypeDesc, 1 ColType union " & _  
				"	select 3 [Type], N'" & Replace(getadminLayoutLngStr("LtxtSecIndex"), "'", "''") & "' TypeDesc, 1 ColType union " & _  
				"	select 4 [Type], N'" & Replace(getadminLayoutLngStr("LtxtCatNav"), "'", "''") & "' TypeDesc, 1 ColType union " & _  
				"	select 5 [Type], N'" & Replace(getadminLayoutLngStr("LtxtCartMinRep"), "'", "''") & "' TypeDesc, 2 ColType union " & _  
				"	select 6 [Type], N'" & Replace(getadminLayoutLngStr("LtxtNewsLetter"), "'", "''") & "' TypeDesc, 2 ColType union " & _  
				"	select 7 [Type], N'" & Replace(getadminLayoutLngStr("LtxtPolls"), "'", "''") & "' TypeDesc, 2 ColType union " & _  
				"	select 8 [Type], N'" & Replace(getadminLayoutLngStr("LtxtHomeOferts"), "'", "''") & "' TypeDesc, 1 ColType " & _  
				") T0 " & _  
				"where not exists(select '' from OLKLayoutLines where ID = @ID and [Type] = T0.[Type]) " & _  
				"and (ColType = " & rs("ColType") & " or ColType = 3) " & _  
				"union all " & _  
				"select 0, GroupID, N'" & Replace(getadminLayoutLngStr("LtxtBanner"), "'", "''") & " - ' + GroupName from OLKBNGroups where GroupID not in (select TypeID from OLKLayoutLines where ID = @ID and [Type] = 0) " & _  
				"union all " & _  
				"select 2, ID, N'" & Replace(getadminLayoutLngStr("LtxtSmallCat"), "'", "''") & " - ' + [Name] from OLKSmallCat where ID not in (select TypeID from OLKLayoutLines where ID = @ID and [Type] = 2) and CatType = 'R' " 
			set ra = Server.CreateObject("ADODB.RecordSet")
			set ra = conn.execute(sql)
			If Not ra.Eof Then %><tr style="color:#31659C;background-color: #F3FBFE;font-size:x-small; font-face: Verdana;"><td><select name="cmbAddLine<%=rs("ColID")%>" id="cmbAddLine<%=rs("ColID")%>" size="1">
			<option></option>
			<% do while not ra.eof %><option value="<%=ra("Type")%>|<%=ra("TypeID")%>"><%=ra("TypeDesc")%></option><% ra.movenext
			loop %></select>
			</td>
			<td style="text-align: center;"><input type="checkbox" name="chkNewActive<%=rs("ColID")%>" class="noborder" checked id="chkNewActive<%=rs("ColID")%>" value="Y">
			<td style="text-align: center;">
			<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
								<input type="text" name="NewRowID<%=rs("ColID")%>" id="NewRowID<%=rs("ColID")%>" size="4" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" value="<%=NewID%>">
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnNewRowID<%=rs("ColID")%>Up"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnNewRowID<%=rs("ColID")%>Down"></td>
								</tr>
							</table></td>
						</tr>
					</table><script language="javascript">NumUDAttach('frmLayout', 'NewRowID<%=rs("ColID")%>', 'btnNewRowID<%=rs("ColID")%>Up', 'btnNewRowID<%=rs("ColID")%>Down');</script>
				</td>
			<td><input type="button" name="btnAdd" id="btnAdd" value="+" onclick="javascript:doAdd(<%=GroupID%>,<%=rs("ColID")%>);"></td>
			</tr>
			<% End If 
			rs.movenext
			loop %>
		</table>
		<% End If %>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" <% If GroupID = -99 Then Response.Write "disabled" %> value="<%=getadminLayoutLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>

</table>
</form>
<!--#include file="bottom.asp" -->