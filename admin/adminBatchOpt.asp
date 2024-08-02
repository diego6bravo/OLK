<!--#include file="top.asp" -->
<!--#include file="lang/adminBatchOpt.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<% conn.execute("use [" & Session("OLKDB") & "]") %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #F3FBFE;
}
.style2 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>
<script language="javascript" src="js_up_down.js"></script>
<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("NewFld") <> "Y" and Request("edit") <> "Y" Then
	sql = "select * from olkBatchRep order by RowOrder asc"
	rs.open sql, conn, 3, 1 %>
	<script language="javascript">
	function valFrm()
	{
		rowName = document.form1.rowName;
		if (rowName.length)
		{
			for (var i = 0;i<rowName.length;i++)
			{
				if (rowName[i].value == '')
				{
					alert("<%=getadminBatchOptLngStr("LtxtValAlrNoNam")%>");
					rowName[i].focus();
					return false;
				}
			}
		}
		else
		{
			if (rowName.value == '')
			{
				alert("<%=getadminBatchOptLngStr("LtxtValAlrNoNam")%>");
				rowName.focus();
				return false;
			}
		}
		return true;
	}
	</script>
	<form method="POST" action="adminsubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font size="1" face="Verdana" color="#31659C"><%=getadminBatchOptLngStr("LttlBatchOpt")%> </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		<font color="#4783C5"><%=getadminBatchOptLngStr("LttlBatchOptNote")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" bgcolor="#E2F3FC" style="width: 16px">&nbsp;</td>
				<td align="center" bgcolor="#E2F3FC" style="width: 200px"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminBatchOptLngStr("DtxtName")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminBatchOptLngStr("DtxtOrder")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtCodification")%></font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminBatchOptLngStr("DtxtField")%>&nbsp;/ 
				<%=getadminBatchOptLngStr("DtxtQuery")%></font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminBatchOptLngStr("DtxtActive")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC" width="16">&nbsp;</td>
			</tr>
			<%
			do while not rs.eof %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
			  	<a href="adminBatchOpt.asp?edit=Y&rI=<%=rs("rowIndex")%>#table20"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top" style="width: 200px">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input style="width: 100%;" class="input" size="20" value="<%=Server.HTMLEncode(RS("rowName"))%>" name="rowName<%=RS("rowIndex")%>" id="rowName" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td width="16"><a href="javascript:doFldTrad('BatchRep', 'rowIndex', <%=rs("rowIndex")%>, 'alterRowName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminBatchOptLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>			  	</td>
			  <td valign="top">
					<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder<%=rs("rowIndex")%>" id="RowOrder<%=rs("rowIndex")%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("RowOrder")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrder<%=rs("rowIndex")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrder<%=rs("rowIndex")%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td>
				<font face="Verdana" size="1">
				<select size="1" name="rowType<%=RS("rowIndex")%>" class="input">
				<option value="T"<% If rs("rowType") = "T" Then %> selected<% End If %>>
				<%=getadminBatchOptLngStr("DtxtDisabled")%></option>
				<option value="L"<% If rs("rowType") = "L" Then %> selected<% End If %>>
				<%=getadminBatchOptLngStr("DtxtLow")%></option>
				<option value="M"<% If rs("rowType") = "M" Then %> selected<% End If %>>
				<%=getadminBatchOptLngStr("DtxtMedium")%></option>
				<option value="H"<% If rs("rowType") = "H" Then %> selected<% End If %>>
				<%=getadminBatchOptLngStr("DtxtHigh")%></option>
				</select><input type="checkbox" <% If rs("rowTypeRnd") = "Y" Then %>checked<% End If %> name="rowTypeRnd<%=rs("rowIndex")%>" id="rowTypeRnd<%=rs("rowIndex")%>" value="ON" class="noborder"><font color="#31659C"><label for="rowTypeRnd<%=rs("rowIndex")%>"><%=getadminBatchOptLngStr("DtxtRndLtr")%></label></font></font></td>
				<td align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("rowField"))%>"></td>
				<td>
				<p align="center">
				<input <% If rs("rowActive") = "Y" Then %>checked<% End If %> type="checkbox" name="rowActive<%=RS("rowIndex")%>" value="ON" class="noborder"></td>
				<td valign="middle" width="16">
						<a href="javascript:if(confirm('<%=getadminBatchOptLngStr("LtxtConfRemFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("rowName")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&rI=<%=rs("rowIndex")%>&submitCmd=adminBatchOpt';">
						<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="rowIndex" value="<%=rs("rowIndex")%>">
				<script language="javascript">NumUDAttach('form1', 'RowOrder<%=rs("rowIndex")%>', 'btnRowOrder<%=rs("rowIndex")%>Up', 'btnRowOrder<%=rs("rowIndex")%>Down');</script>
				<% rs.movenext
				loop %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table22">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminBatchOptLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminBatchOptLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminBatchOpt.asp?NewFld=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminBatchOpt">
<input type="hidden" name="cmd" value="u">
  </form>
  <% Else %>
<script language="javascript">
function valFrm2()
{
	if (document.form2.customSql.value != '' && document.form2.valQuery.value == 'Y')
	{
		alert('<% If Request("Edit") = "Y" Then %><%=getadminBatchOptLngStr("LtxtValVrfyQryUpd")%><% Else %><%=getadminBatchOptLngStr("LtxtValVrfyQryAdd")%><% End If %>');
		document.form2.btnVerfyFilter.focus();
		return false;
	}
	else if (document.form2.rowName.value == '')
	{
		alert('<%=getadminBatchOptLngStr("LtxtValFldNam")%>');
		document.form2.rowName.focus();
		return false;
	}
	else if (document.form2.Custom.checked && document.form2.customSql.value == '')
	{
		alert('<%=getadminBatchOptLngStr("LtxtValQry")%>');
		document.form2.customSql.focus();
		return false;
	}
	return true;
}
</script>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("rI") = "" Then %><%=getadminBatchOptLngStr("LttlAddItemDet")%><% Else %><%=getadminBatchOptLngStr("LtxtEditItemDet")%><% End If %>  </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		<font color="#4783C5"><%=getadminBatchOptLngStr("LttlAddItemDetNote")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("edit") = "Y" Then
			sql = "select * from olkBatchRep where rowIndex = " & Request("rI") 
			set rs = conn.execute(sql) 
			rowName = rs("rowName")
			rowField = rs("rowField")
			rowType = rs("rowType")
			rowTypeRnd = rs("rowTypeRnd")
			rowActive = rs("rowActive")
			rowOrder = rs("rowOrder")
			rowTypeDec = rs("rowTypeDec")
		Else
			sql = "select IsNull(Max(rowOrder)+1, 0) from olkBatchRep"
			set rs = conn.execute(sql)
			rowOrder = rs(0)
			rowTypeDec = "P"
			rowName = "" %>
		<input type="hidden" name="rowNameTrad">
		<input type="hidden" name="customSqlDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%" id="table20">
			<tr>
				<td>
				<table border="0" cellpadding="0" width="100%" id="table20">
					<tr>
						<td align="center" bgcolor="#E2F3FC" style="width: 200px"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtOrder")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtCodification")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtDecimal")%><br>(<%=getadminBatchOptLngStr("DtxtCodification")%>)</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtField")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminBatchOptLngStr("DtxtActive")%>&nbsp;</font></b></td>
					</tr>
					<tr>
						<td valign="top" style="width: 200px" class="style1">
						<p align="center"><font face="Verdana" size="1">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input name="rowName" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(rowName)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('BatchRep', 'rowIndex', '<%=Request("rI")%>', 'alterRowName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.rowNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminBatchOptLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</font></td>
						<td valign="top" class="style1">
						<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder" id="RowOrder" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rowOrder%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrderDown"></td>
							</tr>
						</table></td>
					</tr>
				</table></td>
						<td valign="top" class="style1">
						<font face="Verdana" size="1">
						<select size="1" name="rowType" class="input" style="width: 120; height: 16">
						<option value="T"<% If rowType = "T" Then %> selected<% End If %>>
						<%=getadminBatchOptLngStr("DtxtDisabled")%></option>
						<option value="L"<% If rowType = "L" Then %> selected<% End If %>>
						<%=getadminBatchOptLngStr("DtxtLow")%></option>
						<option value="M"<% If rowType = "M" Then %> selected<% End If %>>
						<%=getadminBatchOptLngStr("DtxtMedium")%></option>
						<option value="H"<% If rowType = "H" Then %> selected<% End If %>>
						<%=getadminBatchOptLngStr("DtxtHigh")%></option>
						</select><input type="checkbox" name="rowTypeRnd" <% If rowTypeRnd = "Y" Then %>checked<% End If %> value="ON" id="rowTypeRnd" class="noborder"><font color="#31659C"><label for="rowTypeRnd"><%=getadminBatchOptLngStr("DtxtRndLtr")%></label></font></td>
						<td valign="top" class="style1">
						<p align="center">
						<select size="1" name="rowTypeDec" class="input">
						<option <% If rowTypeDec = "S" Then %>selected<% End If %> value="S"><%=getadminBatchOptLngStr("DtxtDecSum")%></option>
						<option <% If rowTypeDec = "P" Then %>selected<% End If %> value="P"><%=getadminBatchOptLngStr("DtxtDecPrice")%></option>
						<option <% If rowTypeDec = "R" Then %>selected<% End If %> value="R"><%=getadminBatchOptLngStr("DtxtDecRate")%></option>
						<option <% If rowTypeDec = "Q" Then %>selected<% End If %> value="Q"><%=getadminBatchOptLngStr("DtxtDecQty")%></option>
						<option <% If rowTypeDec = "%" Then %>selected<% End If %> value="%"><%=getadminBatchOptLngStr("DtxtDecPercent")%></option>
						<option <% If rowTypeDec = "M" Then %>selected<% End If %> value="M"><%=getadminBatchOptLngStr("DtxtDecMeasure")%></option>
						</select></td>
						<td valign="top" class="style1">
						<select <% If Request("Edit") = "Y" Then %>disabled<% End If %> size="1" name="rowField" class="input" onchange="javascript:document.form2.customSql.value=this.value;">
						<% If Request("Edit") <> "Y" Then %>
						<option></option>
						<% sql = "select name from syscolumns where id = object_id('OIBT') and name not in ('Consig', 'DataSource', 'Direction', 'Instance', 'ItemCode', 'ItemName', 'Status', 'Transfered', 'U_3dxDesc', 'UserSign', 'WhsCode', 'BatchNum') "
						   set rs = conn.execute(sql)
						   while not rs.eof %>
						<option value="OIBT.<%=RS("Name")%>">OIBT.<%=RS("Name")%></option>
						<% rs.movenext
						wend
						Else %>
						<option>---------</option>
						<% End If %>
						</select></td>
						<td valign="top" class="style1">
						<p align="center">
						<input type="checkbox" <% If rowActive = "Y" Then %>checked<% End If %> name="rowActive" value="ON" class="noborder"></td>
					</tr>
				</table>
				<table border="0" cellpadding="0" width="100%" id="table20">
					<tr>
						<td valign="top">
						<table border="0" width="100%" id="table23" cellpadding="0">
							<tr>
								<td bgcolor="#E2F3FC">
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="10" style="width: 100%" name="customSql" cols="87" class="input" onkeypress="javascript:document.form2.btnVerfyFilter.src='images/btnValidate.gif';document.form2.btnVerfyFilter.style.cursor = 'hand';document.form2.valQuery.value='Y';"><%=myHTMLEncode(rowField)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminBatchOptLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(12, 'customSql', '<%=Request("rI")%>', <% If Request("rI") <> "" Then %>null<% Else %>document.form2.customSqlDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminBatchOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top">
								<table cellpadding="0" style="width: 100%">
									<tr>
										<td valign="top" style="width: 100px" class="style2">
										<font size="1" face="Verdana">
											<strong><%=getadminBatchOptLngStr("DtxtVariables")%></strong></font></td>
										<td class="style1"><font face="Verdana" size="1" color="#4783C5"><span dir="ltr">@ItemCode</span> = <%=getadminBatchOptLngStr("DtxtItemCode")%><br>
								<span dir="ltr">@BatchNum</span> = <%=getadminBatchOptLngStr("DtxtBatchNum")%><br>
								<span dir="ltr">@WhsCode</span> = <%=getadminBatchOptLngStr("DtxtWhsCode")%><br>
								<span dir="ltr">@LanID</span> = <%=getadminBatchOptLngStr("DtxtLanID")%></font></td>
									</tr>
									<tr>
										<td valign="top" style="width: 100px" class="style2">
								<font size="1" face="Verdana">
								<strong><%=getadminBatchOptLngStr("DtxtFunctions")%></strong></font></td>
										<td class="style1"><% HideFunctionTitle = True
										functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
									</tr>
									</table>
								</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminBatchOptLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminBatchOptLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminBatchOptLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminBatchOptLngStr("LtxtConfOptCancel")%>'))window.location.href='adminBatchOpt.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="rI" value="<%=Request("rI")%>">
	<input type="hidden" name="submitCmd" value="adminBatchOpt">
	<input type="hidden" name="cmd" value="<% If Request("Edit") = "Y" Then %>e<% Else %>a<% End If %>">
	</form>
	<script language="javascript">
	NumUDAttach('form2', 'RowOrder', 'btnRowOrderUp', 'btnRowOrderDown');
	function VerfyQuery()
	{
		document.frmVerfyQuery.Query.value = document.form2.customSql.value;
		document.frmVerfyQuery.submit();
	}
	
	function VerfyQueryVerified()
	{
		//document.form2.btnVerfy.disabled = true;
		document.form2.btnVerfyFilter.src='images/btnValidateDis.gif'
		document.form2.btnVerfyFilter.style.cursor = '';
		document.form2.valQuery.value='N';
	}
	</script>
	<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
		<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
		<input type="hidden" name="type" value="batchOpt">
		<input type="hidden" name="Query" value="">
		<input type="hidden" name="parent" value="Y">
	</form>
	<% End If %>
</table>
<!--#include file="bottom.asp" -->