<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_arte.asp" -->

<% 
set rs = Server.CreateObject("ADODB.recordset")
CarArt = myApp.CarArt
%>
<div align="center">
	<table border="0" cellpadding="0" width="499" id="table1">
		<tr>
			<td>
			<p align="center">
			<img border="0" src="design/0/images/search_top.jpg" width="407" height="140"></td>
		</tr>
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%">
	     	<form method="POST" action="search.asp" name="frmSmallSearch">
	     		<input type="hidden" name="document" value="<%=myApp.GetDefView%>">
              	<input type="hidden" name="cmd" value="searchCatalog">
              	<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "N" Then %>ItemName<% Else %>OITM.ItemCode<% End If %>">
              	<input type="hidden" name="orden2" value="asc">
				<tr class="GeneralTlt">
					<td><%=getsearch_arteLngStr("LttlItmsSearch")%></td>
				</tr>
				<tr class="GeneralTbl">
					<td>
					<table border="0" cellpadding="0" width="100%">
						<tr>
							<td>
							<input type="text" name="string" size="80"></td>
							<td width="65">
							<input type="submit" value="<%=getsearch_arteLngStr("DbtnSearch")%>" name="B1" style="float: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"></td>
						</tr>
						<% If myApp.SearchExactA Then %>
						<tr>
							<td colspan="2">
							<p align="center">
							<font face="Verdana" size="1">
							<input type="radio" value="E" name="rdSearchAs" class="noborder" id="rdSearchAsE" <% If myApp.SearchMethodA = "E" Then %>checked<% End If %>><label for="rdSearchAsE"><%=getsearch_arteLngStr("DtxtExact")%></label>
							<input type="radio" name="rdSearchAs" class="noborder" id="rdSearchAsS" value="S" <% If myApp.SearchMethodA = "L" Then %>checked<% End If %>><label for="rdSearchAsS"><%=getsearch_arteLngStr("DtxtLike")%></label></font>
							</td>
						</tr>
						<% Else %>
						<input type="hidden" name="rdSearchAs" value="S">
						<% End If %>
					</table>
					</td>
				</tr>
				<%
				sql = "select T0.ID, IsNull(T1.AlterName, T0.Name) Name  " & _
						"from OLKCustomSearch T0 " & _
						"left outer join OLKCustomSearchAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.LanID = " & Session("LanID") & " " & _
						"where T0.ObjectCode = 4 and T0.Status = 'Y' and exists(select '' from OLKCustomSearchSession where ObjectCode = T0.ObjectCode and ID = T0.ID and SessionID = 'A') " & _
						"order by T0.Ordr "
				set rs = Server.CreateObject("ADODB.RecordSet")
				rs.open sql, conn, 3, 1
				If rs.recordcount = 1 Then %>
				<tr class="GeneralTlt">
					<td style="cursor: hand" onclick="javascript:doMyLink('adCustomSearch.asp', 'ID=<%=rs("ID")%>&adObjID=4', '');">
					<p align="right"><u><%=getsearch_arteLngStr("LtxtAdSearch")%></u></td>
				</tr>
				<% ElseIf rs.recordcount > 1 Then %>
				<tr class="GeneralTlt">
					<td style="cursor: hand" onclick="if(document.getElementById('trAdvanced').style.display==''){document.getElementById('trAdvanced').style.display='none';}else{document.getElementById('trAdvanced').style.display='';}">
					<p align="right"><u><%=getsearch_arteLngStr("LtxtAdSearch")%></u></td>
				</tr>
				<tBody id="trAdvanced" style="display: none;">
				<% do while not rs.eof %>
				<tr class="GeneralTbl">
					<td style="cursor: hand" onclick="javascript:doMyLink('adCustomSearch.asp', 'ID=<%=rs("ID")%>&adObjID=4', '');">
					<u><%=rs("Name")%></u></td>
				</tr>
				<% rs.movenext
				loop %>
				</tBody>
				<% End If %>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<input type="hidden" name="focus" value="frmSmallSearch.string">
				</form>
			</table>
			</td>
		</tr>
		</table>
</div>
<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="99" width="256" src=""></iframe>
<% set rs1 = nothing
set rs2 = nothing %>