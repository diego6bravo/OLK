<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_inte.asp" -->


  
<script language="javascript">
function getValue(myType, fld) {
if (fld.value == '') { return; } 
	updFld = fld;
	if (fld.value.indexOf('*') == -1) {
		document.frmGetValue.Type.value = myType;
		document.frmGetValue.searchStr.value = fld.value;
		document.frmGetValue.submit();
	}
	else { launchSelect(myType, fld.value); }
}
function launchSelect(myType, Value){
	var retVal = window.showModalDialog('topGetValueSelect.asp?Type=' + myType + '&Value=' + Value,'','dialogWidth:500px;dialogHeight:500px');
	if (retVal != '' && retVal != null){
		updFld.value = retVal; setTargetVal(retVal); retVal = '';
	} 
	else { 
		updFld.value = '';
	}
}
function setValue(src, value, myType){
	if (value != '') 
	{ updFld.value = value; setTargetVal(value); }
	else { if(src == 0)launchSelect(myType, updFld.value); }
}
function setTargetVal(value)
{
	if (Right(updFld.name, 4) == "From")
	{
		setFldName = Left(updFld.name, (updFld.name.length-4));
		fldTo = document.getElementById(setFldName + 'To');
		if (fldTo.value == '') { fldTo.value = value; fldTo.select(); }
	}
}
</script>
<form method="post" target="ifGetValue" name="frmGetValue" action="topGetValue.asp">
<input type="hidden" name="Type" value="">
<input type="hidden" name="searchStr" value="">
</form>
<form method="POST" action="clientsSearch.asp" name="frmSearchOCRD">
<input type="hidden" name="cmd" value="clientsSearch">
<div align="center">
	<table border="0" cellpadding="0" width="499">
		<tr>
			<td>
			<p align="center">
			<img border="0" src="design/0/images/search_top.jpg" width="407" height="140"></td>
		</tr>
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%" id="table2">
				<tr class="GeneralTlt">
					<td><%=getsearch_inteLngStr("LttlClientSearch")%></td>
				</tr>
				<tr class="GeneralTbl">
					<td>
					<table border="0" cellpadding="0" width="100%" id="table3">
						<tr>
							<td>
							<input type="text" name="string" size="47" style="width: 100%"></td>
							<td width="65">
							<input type="submit" value="<%=getsearch_inteLngStr("DbtnSearch")%>" name="search" style="float: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"></td>
						</tr>
					</table>
					</td>
				</tr>
				<%
				sql = "select T0.ID, IsNull(T1.AlterName, T0.Name) Name  " & _
						"from OLKCustomSearch T0 " & _
						"left outer join OLKCustomSearchAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.LanID = " & Session("LanID") & " " & _
						"where T0.ObjectCode = 2 and T0.Status = 'Y' and exists(select '' from OLKCustomSearchSession where ObjectCode = T0.ObjectCode and ID = T0.ID and SessionID = 'A') " & _
						"order by T0.Ordr "
				set rs = Server.CreateObject("ADODB.RecordSet")
				rs.open sql, conn, 3, 1
				If rs.recordcount = 1 Then %>
				<tr class="GeneralTlt">
					<td style="cursor: hand" onclick="javascript:doMyLink('adCustomSearch.asp', 'ID=<%=rs("ID")%>&adObjID=2', '');">
					<p align="right"><u><%=getsearch_inteLngStr("LtxtAdSearch")%></u></td>
				</tr>
				<% ElseIf rs.recordcount > 1 Then %>
				<tr class="GeneralTlt">
					<td style="cursor: hand" onclick="if(document.getElementById('trAdvanced').style.display==''){document.getElementById('trAdvanced').style.display='none';}else{document.getElementById('trAdvanced').style.display='';}">
					<p align="right"><u><%=getsearch_inteLngStr("LtxtAdSearch")%></u></td>
				</tr>
				<tBody id="trAdvanced" style="display: none;">
				<% do while not rs.eof %>
				<tr class="GeneralTbl">
					<td style="cursor: hand" onclick="javascript:doMyLink('adCustomSearch.asp', 'ID=<%=rs("ID")%>&adObjID=2', '');">
					<u><%=rs("Name")%></u></td>
				</tr>
				<% rs.movenext
				loop %>
				</tBody>
				<% End If %>
				<tr>
					<td>&nbsp;</td>
				</tr>
			</table>
			</td>
		</tr>
		</table>
	<p>
			<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="99" width="256" src=""></iframe>
</div>
</form>
<script language="javascript">
var updFld = document.frmSearchOCRD.GroupNameFrom;
</script>