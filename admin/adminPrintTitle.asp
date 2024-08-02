<!--#include file="top.asp" -->
<!--#include file="lang/adminPrintTitle.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<% conn.execute("use [" & Session("OLKDB") & "]") %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #F5FBFE;
}
.style2 {
	color: #31659C;
}
.style3 {
	background-color: #F5FBFE;
}
</style>
</head>
<script language="javascript" src="js_up_down.js"></script>
<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("NewFld") <> "Y" and Request("edit") <> "Y" Then
	sql = "select * from OLKDocAddHdr order by Row, Col"
	rs.open sql, conn, 3, 1 %>
	<script language="javascript">
	function valFrm()
	{
		lineName = document.form1.lineName;
		if (lineName.length)
		{
			for (var i = 0;i<lineName.length;i++)
			{
				if (lineName[i].value == '')
				{
					alert("<%=getadminPrintTitleLngStr("LtxtValAlrNoNam")%>");
					lineName[i].focus();
					return false;
				}
			}
		}
		else
		{
			if (lineName.value == '')
			{
				alert("<%=getadminPrintTitleLngStr("LtxtValAlrNoNam")%>");
				lineName.focus();
				return false;
			}
		}
		return true;
	}
	</script>
	<form method="POST" action="adminsubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font size="1" face="Verdana" color="#31659C"><%=getadminPrintTitleLngStr("LttlPrintTitle")%> </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		<font color="#4783C5"><%=getadminPrintTitleLngStr("LttlPrintTitleNote")%> </font></font></td>
	</tr>
	<tr>
		<td bgcolor="#E2F3FC">
		<p align="justify" class="style2"> 
		<font face="Verdana" size="1"><strong><%=getadminPrintTitleLngStr("DtxtPreview")%> 
		</strong> </font></td>
	</tr>
	<tr>
		<td bgcolor="#E2F3FC">
		<iframe width="100%" name="iPreview" src="prevTitle.asp"></iframe></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" bgcolor="#E2F3FC" style="width: 16px">&nbsp;</td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminPrintTitleLngStr("DtxtName")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminPrintTitleLngStr("DtxtRow")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminPrintTitleLngStr("DtxtCol")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminPrintTitleLngStr("DtxtField")%>&nbsp;/ 
				<%=getadminPrintTitleLngStr("DtxtQuery")%></font></b></td>
				<td align="center" bgcolor="#E2F3FC"><b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminPrintTitleLngStr("DtxtAccess")%>&nbsp;</font></b></td>
				<td align="center" bgcolor="#E2F3FC" width="16">&nbsp;</td>
			</tr>
			<%
			do while not rs.eof %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
			  	<a href='adminPrintTitle.asp?edit=Y&amp;rI=<%=rs("LineIndex")%>'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input class="input" size="20" value="<%=Server.HTMLEncode(RS("Name"))%>" name="lineName<%=RS("LineIndex")%>" id="lineName" style="width: 100%; " onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('DocAddHdr', 'LineIndex', <%=rs("LineIndex")%>, 'alterName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminPrintTitleLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
			  	</td>
			  <td valign="top">
					<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="Row<%=rs("LineIndex")%>" id="Row<%=rs("LineIndex")%>" size="7" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6" value="<%=rs("Row")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRow<%=rs("LineIndex")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRow<%=rs("LineIndex")%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
			  <td valign="top">
					<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="Col<%=rs("LineIndex")%>" id="Col<%=rs("LineIndex")%>" size="7" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6" value="<%=rs("Col")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnCol<%=rs("LineIndex")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnCol<%=rs("LineIndex")%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("Query"))%>"></td>
				<td>
				<p align="center">
				<select size="1" class="input" name="Access<%=rs("LineIndex")%>">
				<option <% if rs("Access") = "T" then %>selected<%end if %> value="T">
				<%=getadminPrintTitleLngStr("DtxtAll")%></option>
				<option <% if rs("Access") = "V" then %>selected<%end if %> value="V">
				<%=getadminPrintTitleLngStr("DtxtAgent")%></option>
				<option <% if rs("Access") = "C" then %>selected<%end if %> value="C">
				<%=getadminPrintTitleLngStr("DtxtClient")%></option>
				<option <% if rs("Access") = "D" then %>selected<%end if %> value="D">
				<%=getadminPrintTitleLngStr("DtxtDisabled")%></option>
				</select></td>
				<td valign="middle" width="16">
						<a href="javascript:if(confirm('<%=getadminPrintTitleLngStr("LtxtConfRemFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("Name")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&rI=<%=rs("LineIndex")%>&submitCmd=adminPrintTitle';">
						<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="LineIndex" value="<%=rs("LineIndex")%>">
				<script language="javascript">NumUDAttach('form1', 'Row<%=rs("LineIndex")%>', 'btnRow<%=rs("LineIndex")%>Up', 'btnRow<%=rs("LineIndex")%>Down');
				NumUDAttach('form1', 'Col<%=rs("LineIndex")%>', 'btnCol<%=rs("LineIndex")%>Up', 'btnCol<%=rs("LineIndex")%>Down');</script>
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
				<input type="submit" value="<%=getadminPrintTitleLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminPrintTitleLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminPrintTitle.asp?NewFld=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminPrintTitle">
<input type="hidden" name="cmd" value="u">
  </form>
  <% Else %>
<script language="javascript">
function valFrm2()
{
	if (document.form2.Query.value != '' && document.form2.valQuery.value == 'Y')
	{
		alert('<% If Request("Edit") = "Y" Then %><%=getadminPrintTitleLngStr("LtxtValVrfyQryUpd")%><% Else %><%=getadminPrintTitleLngStr("LtxtValVrfyQryAdd")%><% End If %>');
		document.form2.btnVerfyFilter.focus();
		return false;
	}
	else if (document.form2.lineName.value == '')
	{
		alert('<%=getadminPrintTitleLngStr("LtxtValFldNam")%>');
		document.form2.lineName.focus();
		return false;
	}
	else if (document.form2.Query.value == '')
	{
		alert('<%=getadminPrintTitleLngStr("LtxtValQry")%>');
		document.form2.Query.focus();
		return false;
	}
	return true;
}
</script>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("rI") = "" Then %><%=getadminPrintTitleLngStr("LttlAddFld")%><% Else %><%=getadminPrintTitleLngStr("LtxtEditFld")%><% End If %>  </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		<font color="#4783C5"><%=getadminPrintTitleLngStr("LttlAddFldNote")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("edit") = "Y" Then
			sql = "select * from OLKDocAddHdr where LineIndex = " & Request("rI") 
			set rs = conn.execute(sql) 
			lineName = rs("Name")
			Query = rs("Query")
			Access = rs("Access")
			Row = rs("Row")
			Col = rs("Col")
			ShowName = rs("ShowName") = "Y"
		Else
			sql = "select IsNull(Max(Row)+1, 0) from OLKDocAddHdr"
			set rs = conn.execute(sql)
			Row = rs(0)
			Col = 1
			lineName = "" %>
		<input type="hidden" name="lineNameTrad">
		<input type="hidden" name="QueryDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td align="center" bgcolor="#E2F3FC" width="160"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminPrintTitleLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC" style="width: 100px"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminPrintTitleLngStr("DtxtRow")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC" style="width: 100px"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminPrintTitleLngStr("DtxtCol")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC" width="50"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminPrintTitleLngStr("DtxtField")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC" width="60"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminPrintTitleLngStr("DtxtAccess")%>&nbsp;</font></b></td>
						<td align="center" bgcolor="#E2F3FC" style="width: 160px">
						&nbsp;</td>
						<td align="center" bgcolor="#E2F3FC">&nbsp;</td>
					</tr>
					<tr>
						<td valign="top" width="160" class="style3">
						<p align="center"><font face="Verdana" size="1">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input name="lineName" class="input" value="<%=Server.HTMLEncode(lineName)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('DocAddHdr', 'LineIndex', '<%=Request("rI")%>', 'alterName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.lineNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminPrintTitleLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</font></td>
						<td valign="top" style="width: 100px" class="style3">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td>
									<input type="text" name="Row" id="Row" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6" value="<%=Row%>">
								</td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnRowUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnRowDown"></td>
									</tr>
								</table></td>
							</tr>
						</table></td>
						<td valign="top" style="width: 100px" class="style3">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td>
									<input type="text" name="Col" id="Col" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6" value="<%=Col%>">
								</td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnColUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnColDown"></td>
									</tr>
								</table></td>
							</tr>
						</table></td>
						<td valign="top" width="50" class="style3">
						<select <% If Request("Edit") = "Y" Then %>disabled<% End If %> size="1" name="Field" class="input" onchange="javascript:document.form2.Query.value=this.value;">
						<% If Request("Edit") <> "Y" Then %>
						<option></option>
						<% sql = "select name from syscolumns where id = object_id('OADM') "
						   set rs = conn.execute(sql)
						   while not rs.eof %>
						<option value="T0.<%=RS("Name")%>"><%=RS("Name")%></option>
						<% rs.movenext
						wend
						Else %>
						<option>---------</option>
						<% End If %>
						</select></td>
						<td valign="top" width="50" class="style3">
						<select size="1" class="input" name='Access'>
						<option <% if Access = "T" then %>selected<%end if %> value="T">
						<%=getadminPrintTitleLngStr("DtxtAll")%></option>
						<option <% if Access = "V" then %>selected<%end if %> value="V">
						<%=getadminPrintTitleLngStr("DtxtAgent")%></option>
						<option <% if Access = "C" then %>selected<%end if %> value="C">
						<%=getadminPrintTitleLngStr("DtxtClient")%></option>
						<option <% if Access = "D" then %>selected<%end if %> value="D">
						<%=getadminPrintTitleLngStr("DtxtDisabled")%></option>
						</select></td>
						<td valign="top" class="style1" style="width: 160px">
						<font face="Verdana" size="1" color="#31659C">
						<input type="checkbox" name="chkShowName" id="chkShowName" class="noborder" value="Y" <% If ShowName Then %>checked<% End If %>><label for="chkShowName"><%=getadminPrintTitleLngStr("LtxtShowTitle")%></label></font></td>
						<td valign="top" class="style3">
						&nbsp;</td>
					</tr>
				</table>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td valign="top">
						<table border="0" width="100%" cellpadding="0">
							<tr bgcolor="#E2F3FC">
								<td colspan="2">
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="10" style="width: 100%" name="Query" cols="87" class="input"onkeypress="javascript:document.form2.btnVerfyFilter.src='images/btnValidate.gif';document.form2.btnVerfyFilter.style.cursor = 'hand';;document.form2.valQuery.value='Y';"><%=myHTMLEncode(Query)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminPrintTitleLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(9, 'Query', '<%=Request("rI")%>', <% If Request("rI") <> "" Then %>null<% Else %>document.form2.QueryDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminPrintTitleLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr bgcolor="#E2F3FC">
								<td width="100" valign="top" style="padding-top: 1px; " class="style2">
								<font size="1" face="Verdana">
								<strong><%=getadminPrintTitleLngStr("DtxtVariables")%></strong></font></td>
								<td class="style3">
								<font face="Verdana" size="1" color="#4783C5">
								<span dir="ltr">@LanID</span> = <%=getadminPrintTitleLngStr("DtxtLanID")%><br>
								<span dir="ltr">@branch</span> = <%=getadminPrintTitleLngStr("DtxtBranch")%></font>
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
				<input type="submit" value="<%=getadminPrintTitleLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminPrintTitleLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminPrintTitleLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminPrintTitleLngStr("DtxtConfCancel")%>'))window.location.href='adminPrintTitle.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="rI" value="<%=Request("rI")%>">
	<input type="hidden" name="submitCmd" value="adminPrintTitle">
	<input type="hidden" name="cmd" value="<% If Request("Edit") = "Y" Then %>e<% Else %>a<% End If %>">
	</form>
	<script language="javascript">
	NumUDAttach('form2', 'Row', 'btnRowUp', 'btnRowDown');
	NumUDAttach('form2', 'Col', 'btnColUp', 'btnColDown');
	function VerfyQuery()
	{
		document.frmVerfyQuery.Query.value = document.form2.Query.value;
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
		<input type="hidden" name="type" value="printTitle">
		<input type="hidden" name="Query" value="">
		<input type="hidden" name="parent" value="Y">
	</form>
	<% End If %>
</table>
<!--#include file="bottom.asp" -->