<!--#include file="top.asp" -->
<!--#include file="lang/adminOps.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<script type="text/javascript" src="adminOps.js"></script>

<form name="frmOps" method="post" action="adminOpsSubmit.asp">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminOpsLngStr("LttlOpsLst")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminOpsLngStr("LttlOpsLstNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<thead>
			<tr class="TblRepTlt">
				<td align="center" width="15"><b>
				<font size="1" face="Verdana" color="#31659C"></font></b></td>
				<td align="center">
				<%=getadminOpsLngStr("DtxtName")%></td>
				<td align="center" style="width: 80px">
				<%=getadminOpsLngStr("DtxtActive")%></td>
				<td align="center" style="width: 16px"></td>
			</tr>
			</thead>
			<% 
			set rd = Server.CreateObject("ADODB.RecordSet")
			sql = "select ID, Name, (select Count('') from OLKOps where GroupID = T0.ID) [Count] from OLKOpsGrps T0 order by 2"
			rd.open sql, conn, 3, 1

			sql = "select T0.ID, T0.Name, T1.ID GrpID, T1.Name GrpName, T0.Status " & _  
					"from OLKOps T0 " & _  
					"inner join OLKOpsGrps T1 on T1.ID = T0.GroupID " & _  
					"where T0.Status <> 'D' " 
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			LastGroup = ""
			EndBody = False
			LastRep = False
			do While NOT RS.EOF
			If LastGroup <> rs("GrpName") Then
			If EndBody Then 
				Response.Write "</tbody>"
			End If
			EndBody = True %>
			<thead>
			<tr class="TblRepTltNoBold" style="cursor: hand; " onclick="javascript:doExpand(<%=rs("GrpID")%>);">
			  <td width="15" align="center">
			  	<span id="signExpand<%=rs("GrpID")%>">[+]</span>
				</td>
				<td colspan="5"><%=rs("GrpName")%></td>
				</tr>
			</thead>
			<% 
			LastGroup = rs("GrpName")
			FirstRep = True
			End If
			If FirstRep Then %>
			<tbody id="tr<%=rs("GrpID")%>" style="display: none; ">
			<% 
			FirstRep = False
			End If %>
			<tr class="TblRepTbl">
			  <td width="15">
				<a href="adminOpsEdit.asp?ID=<%=rs("ID")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
				<td><input type="hidden" name="ID" value="<%=rs("ID")%>">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="TblRep<% If Alter Then %>A<% End If %>Tbl">
						<td><input type="text" value="<%=Server.HTMLEncode(rs("Name"))%>" name="opName<%=rs("ID")%>" maxlength="100" style="width: 100%;"></td>
						<td style="width: 16px; text-align: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>;">
						<a href="javascript:doFldTrad('Ops', 'ID', '<%=rs("ID")%>', 'alterName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminOpsLngStr("DtxtTranslate")%>" border="0"></a>
						</td>
					</tr>
				</table>
				</td>
				<td style="width: 80px">
				<p align="center">
				<input class="OptionButton" type="checkbox" name="Status<%=rs("ID")%>" <% If rs("Status") = "A" Then %>checked<% End If %> value="Y"></td>
				<td style="width: 16px">
				<% If rs("ID") >= 0 Then %>
				<a href="javascript:if(confirm('<%=getadminOpsLngStr("LtxtConfDelOp")%>'.replace('{0}', '<%=rs("Name")%>')))window.location.href='adminOpsSubmit.asp?cmd=delOp&ID=<%=rs("ID")%>'">
				<img border="0" src="images/remove.gif"></a><% Else %>&nbsp;<% End If %></td>
			</tr>
	<% 	RS.MoveNext
		loop
			If EndBody Then 
				Response.Write "</tbody>"
			End If
 %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminOpsLngStr("DtxtSave")%>" name="btnSave"></td>
				<td width="77">
				<input class="BtnRep" type="button" value="<%=getadminOpsLngStr("DtxtNew")%>" name="btnNew" onclick="javascript:<% If rd.recordcount > 0 Then %>window.location.href='adminOpsEdit.asp?ID=-1'<% Else %>alert('<%=getadminOpsLngStr("LtxtValNoGrp")%>');<% End If %>"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
<input type="hidden" name="cmd" value="updOps">
</form>
<form name="frmGrp" method="post" action="adminOpsSubmit.asp">
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminOpsLngStr("DtxtGroups")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminOpsLngStr("LttlGrpNote")%></td>
	</tr>
	<tr>
		<td >
		<table border="0" cellpadding="0">
		<tr class="TblRepTltSub">
				<td align="center"><%=getadminOpsLngStr("DtxtGroup")%>&nbsp;</td>
			</tr>
			<% do while not rd.eof %>
			<tr class="TblRepTbl">
				<td valign="bottom"><input type="hidden" name="ID" value="<%=rd("ID")%>">
				<p align="center">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" id="grpName<%=rd("ID")%>" name="grpName<%=rd("ID")%>" size="100" value="<%=Server.HTMLEncode(rd("Name"))%>" max="100"></td>
						<td><a href="javascript:doFldTrad('OpsGrps', 'ID', <%=rd("ID")%>, 'alterName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminOpsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td valign="top" style="width: 15px">
				<% If rd("Count") = 0 Then %><a href="javascript:if(confirm('<%=getadminOpsLngStr("LtxtConfDelGrp")%>'.replace('{0}', '<%=Replace(rd("Name"), "'", "\'")%>')))window.location.href='adminOpsSubmit.asp?cmd=delGrp&delID=<%=rd("ID")%>'"><img border="0" src="images/remove.gif"></a><% End If %></td>
			</tr>
			<% rd.movenext
			loop %>
			<tr class="TblRep<% If Alter Then %>A<% End If %>Tbl">
				<td valign="top">
				<p align="center">
				<input type="hidden" name="GrpNameTrad">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="newGrpName" size="100" maxlength="100">
						<input type="hidden" name="newGrpNameTrad"></td>
						<td><a href="javascript:doFldTrad('OpsGrps', 'ID', '', 'alterName', 'T', document.frmGrp.newGrpNameTrad);"><img src="images/trad.gif" alt="<%=getadminOpsLngStr("DtxtTranslate")%>" border="0"></a></td>
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
				<input class="BtnRep" type="submit" value="<%=getadminOpsLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="updGrp">
</form>
<!--#include file="bottom.asp" -->