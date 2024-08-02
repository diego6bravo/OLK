<!--#include file="top.asp" -->
<!--#include file="lang/adminDocBreakDown.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<br>
<% 
If Request.Form.Count > 0 Then
	If Request("Query") <> "" Then Query = "N'" & saveHTMLDecode(Request("Query"), False) & "'" Else Query = "NULL"
	sql = "update OLKCommon set DocBreDowQry = " & Query
	conn.execute(sql)

	ObjectCode = Split(Request("ObjectCode"), ", ")
	For i = 0 to UBound(ObjectCode)
		objCode = ObjectCode(i)
		If Request("chk" & objCode) = "Y" Then BreakDown = "Y" Else BreakDown = "N"
		sql = "update OLKDocConf set BreakDown = '" & BreakDown & "' where ObjectCode = " & objCode
		conn.execute(sql)
	Next
End If

sql = "select DocBreDowQry from OLKCommon"
set rs = conn.execute(sql)
Query = rs("DocBreDowQry")
rs.close

sql = "select ObjectCode, BreakDown from OLKDocConf where ObjectCode in (17)"
set rs = conn.execute(sql)

do while not rs.eof
	Select Case rs("ObjectCode")
		Case 17
			chk17 = rs("BreakDown") = "Y"
	End Select
rs.movenext
loop %>
<form method="POST" action="adminDocBreakDown.asp" name="frm" onsubmit="javascript:return valFrm();">
	<table border="0" cellpadding="0" width="100%">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminDocBreakDownLngStr("Lttl")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE" style="height: 14px">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminDocBreakDownLngStr("LttlNote")%></font></p>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td bgcolor="#E2F3FC" style="width: 120px">&nbsp;</td>
						<td valign="top" class="style6">
						<font face="Verdana" size="1" color="#4783C5"><span class="style5">
						<input type="hidden" name="ObjectCode" value="17">
						<input type="checkbox" name="chk17" class="noborder" id="chk17" <% If chk17 Then %>checked<% End If %> value="Y"></span><span class="style4"><label for="chk17"><%=getadminDocBreakDownLngStr("DtxtActive")%></label></span></font></td>				
					</tr>
					<% If 1 = 2 Then %>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
								<b><font face="Verdana" size="1" color="#31659C">
								|D:txtObjects|</font></b></td>	
								<td valign="top" class="style1">
								&nbsp;<input type="checkbox" <% If ORDR = "Y" Then %>checked<% End If %> name="ObjectCode" value="17" id="ObjectCode2" class="noborder"><label for="ObjectCode2">|D:txtSalesOrder|</label> </td>					
					</tr>
					<% End If %>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
								<b><font face="Verdana" size="1" color="#31659C">
								<%=getadminDocBreakDownLngStr("DtxtQuery")%></font></b></td>	
								<td valign="top" class="style1">
						
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea cols="78" dir="ltr" name="Query" class="input" style="width: 100%; " rows="40" onkeydown="return catchTab(this,event)" onkeypress="javascript:document.frm.btnVerfy.src='images/btnValidate.gif';document.frm.btnVerfy.style.cursor = 'hand';document.frm.valQuery.value='Y';"><%=myHTMLEncode(Query)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminDocBreakDownLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(23, 'Query', '', null);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfy" alt="<%=getadminDocBreakDownLngStr("DtxtValidate")%>" onclick="javascript:if (document.frm.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table></td>					
					</tr>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<font size="1" color="#31659C" face="Verdana"><b><%=getadminDocBreakDownLngStr("LtxtReqCols")%>:</b></font></td>
						<td valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						LineNum = <%=getadminDocBreakDownLngStr("DtxtLine")%><br>
						Quantity = <%=getadminDocBreakDownLngStr("DtxtQty")%><br>
						ShipDate = <%=getadminDocBreakDownLngStr("LtxtShipDate")%><br>
						ShipDateDiff = <%=getadminDocBreakDownLngStr("LtxtShipDateDiff")%>
						</font></td>
					</tr>
					<tr>
					<tr>
						<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<font size="1" color="#31659C" face="Verdana"><b><%=getadminDocBreakDownLngStr("DtxtFunctions")%>:</b></font></td>
						<td valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						<% HideFunctionTitle = True
						functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"-->						</font></td>
					</tr>
						<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminDocBreakDownLngStr("DtxtVariables")%></font></b></td>
						<td valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						<span dir="ltr">@LogNum</span> = 
						<%=getadminDocBreakDownLngStr("DtxtLogNum")%><br>
						<span dir="ltr">@dbName</span> = <%=getadminDocBreakDownLngStr("DtxtDB")%></font></td>		
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
					<input type="submit" value="<%=getadminDocBreakDownLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function valFrm()
{
	if (document.frm.valQuery.value == 'Y')
	{
		alert('<%=getadminDocBreakDownLngStr("LtxtValQryVal")%>');
		document.frm.btnVerfy.focus();
		return false;
	}
	
	if (document.frm.Query.value == '' && document.frm.chk17.checked)
	{
		alert('<%=getadminDocBreakDownLngStr("LtxtValQry")%>');
		document.frm.Query.focus();
		return false;
	}

}
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frm.Query.value;
	document.frmVerfyQuery.submit();
}
function VerfyQueryVerified()
{
	document.frm.btnVerfy.src='images/btnValidateDis.gif'
	document.frm.btnVerfy.cursor = '';
	document.frm.valQuery.value='N';
}

</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none">
	</iframe><input type="hidden" name="type" value="docBD">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<!--#include file="bottom.asp" -->