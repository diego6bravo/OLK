<!--#include file="top.asp" -->
<!--#include file="lang/adminCatOpt.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<head>
<style type="text/css">
.style1 {
	background-color: #E1F3FD;
}
.style2 {
	font-weight: bold;
	background-color: #E1F3FD;
}
.style3 {
	background-color: #E2F3FC;
}
.style4 {
	background-color: #F3FBFE;
}
.style6 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style7 {
	color: #4783C5;
}
.style9 {
	color: #31659C;
}
.style10 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style11 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
	text-align: center;
}
</style>
</head>

<% 
varx = 0
conn.execute("use [" & Session("OLKDB") & "]")
If Request("OLKCType") <> "" Then OLKCType = Request("OLKCType") Else OLKCType = "T"
sql = "select Case UserType When 'C' Then N'" & getadminCatOptLngStr("DtxtClients") & "' When 'V' Then N'" & getadminCatOptLngStr("DtxtAgents") & "' End UserType, UserType As 'SubmitType', ImgMaxSize, catRows, catCols, pdfCols from OLKCatOpt where CatType = '" & OLKCType & "'"
set rs = conn.execute(sql)
%>
<script language="javascript" src="js_up_down.js"></script>
<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("edit") <> "Y" and Request("NewFld") <> "Y" Then %>
	<form method="POST" action="adminsubmit.asp" name="form1">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("LttlItmSearchCust")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			<font color="#4783C5"><%=getadminCatOptLngStr("LttlItmSearchCustNote")%></font></font></p>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" id="table7">
				<tr>
					<td width="183" class="style3">
					<font face="Verdana" size="1" color="#4783C5"><strong><span class="style9"><%=getadminCatOptLngStr("LtxtCatOf")%></span></strong></font>&nbsp;
					</td>
					<td width="183" class="style4">
					<select size="1" name="OLKCType" class="input" style="height: 16" onchange="javascript:window.location.href='adminCatOpt.asp?OLKCType='+this.value">
					<option value="T">
					<%=getadminCatOptLngStr("DtxtStore")%></option>
					<option value="C" <% if Request("OLKCType") = "C" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtCat")%></option>
					<option value="L" <% If Request("OLKCType") = "L" Then %>selected<% End If %>><%=getadminCatOptLngStr("DtxtList")%></option>
					</select></td>
				</tr>
				<tr>
					<% do while not rs.eof %>
					<td colspan="2">
					<br>
					<table border="0" cellpadding="0" id="table8">
						<tr>
							<td class="style11" colspan="2"><font face="Verdana" size="1"><strong><%=rs("UserType")%></strong></font>&nbsp;
							</td>
						</tr>
						<tr>
							<td width="144" class="style6">
							<font face="Verdana" size="1"><strong><%=getadminCatOptLngStr("LtxtImgSize")%></strong></font></td>
							<td class="style4">
							<input type="text" name="ImgMaxSize<%=rs("SubmitType")%>" size="20" class="input" style="width: 73px; height:16px; text-align: right;" onchange="chkNum(this, <%=rs("ImgMaxSize")%>, 2, document.form1.oldImgMaxSize<%=rs("SubmitType")%>)" value="<%=rs("ImgMaxSize")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);">
							<input type="hidden" name="oldImgMaxSize<%=rs("SubmitType")%>" value="<%=rs("ImgMaxSize")%>"></td>
						</tr>
						<tr>
							<td width="144" class="style6">
							<font face="Verdana" size="1"><strong><%=getadminCatOptLngStr("LtxtAmountRows")%></strong></font></td>
							<td class="style4">
							<input type="text" name="catRows<%=rs("SubmitType")%>" size="20" class="input" style="width: 73px; height:16px; text-align: right;" onchange="chkNum(this, <%=rs("catRows")%>, 1, document.form1.oldcatRows<%=rs("SubmitType")%>)" value="<%=rs("catRows")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);">
							<input type="hidden" name="oldcatRows<%=rs("SubmitType")%>" value="<%=rs("catRows")%>"></td>
						</tr>
						<% If OLKCType = "C" Then %>
						<tr>
							<td width="144" class="style6">
							<font face="Verdana" size="1"><strong><%=getadminCatOptLngStr("LtxtAmountCols")%></strong></font></td>
							<td class="style4">
							<input type="text" name="catCols<%=rs("SubmitType")%>" size="20" class="input" style="width: 73px; height:16px; text-align: right;" onchange="chkNum(this, <%=rs("catCols")%>, 1, document.form1.oldcatCols<%=rs("SubmitType")%>)" value="<%=rs("catCols")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);">
							<input type="hidden" name="oldcatCols<%=rs("SubmitType")%>" value="<%=rs("catCols")%>"></td>
						</tr>
						<tr>
							<td width="144" class="style6">
							<font face="Verdana" size="1"><strong><%=getadminCatOptLngStr("LtxtAmountCols")%> (PDF)</strong></font></td>
							<td class="style4">
							<input type="text" name="pdfCols<%=rs("SubmitType")%>" size="20" class="input" style="width: 73px; height:16px; text-align: right;" onchange="chkNum(this, <%=rs("pdfCols")%>, 1, document.form1.oldpdfCols<%=rs("SubmitType")%>)" value="<%=rs("pdfCols")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);">
							<input type="hidden" name="oldpdfCols<%=rs("SubmitType")%>" value="<%=rs("pdfCols")%>"></td>
						</tr>
						<% End If %>
					</table>
					</td>
					<% rs.movenext
				loop %>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td><hr color="#0D85C6" size="1"></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table12">
				<tr>
					<td align="center" class="style2" style="width: 16px; height: 21px;">
					</td>
					<td align="center" class="style2" style="height: 21px; width: 200px;">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtName")%></font></td>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtOrder")%></font></td>
					<% If OLKCType = "T" Then %>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtPosition2")%></font></td>
					<% End If %>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtCodification")%></font></td>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtField")%> / 
					<%=getadminCatOptLngStr("DtxtQuery")%></font></td>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtAccess")%></font></td>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("LtxtSesion")%></font></td>
					<td align="center" class="style2" style="height: 21px">
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtAlignment")%></font></td>
					<td align="center" class="style1" style="width: 16px; height: 21px;"></td>
				</tr>
				<% sql = "select * from OLK" & OLKCType & "Cart order by colordr asc"
	rs.close
	rs.open sql, conn, 3, 1
	do while not rs.eof
		FieldAdd = rs("LineIndex")
		LinkAdd = rs("LineIndex") & "&OLKCType=" & OLKCType
	   varx = varx + 1 %>
				<tr bgcolor="#F3FBFE">
					<td valign="top" style="width: 16px; padding-top: 4px">
					<a href="adminCatOpt.asp?edit=Y&LineIndex=<%=LinkAdd%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
					<td valign="top" style="width: 200px">
					<table border="0" cellpadding="0" id="table13" width="100%">
						<tr>
							<td>
							<input class="input" size="20" style="width: 100%;" value="<%=Server.HTMLEncode(RS("ColName"))%>" name="ColName<%=FieldAdd%>" onkeydown="return chkMax(event, this, 50);">
							</td>
							<td width="16"><a href="javascript:doFldTrad('<%=OLKCType%>Cart', 'LineIndex', <%=rs("LineIndex")%>, 'alterColName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminCatOptLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>
					<% colQuery = Left(rs("colQuery"), 100) %>
					<td valign="top">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
								<input type="text" name="ColOrdr<%=FieldAdd%>" id="ColOrdr<%=FieldAdd%>" size="4" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" value="<%=rs("colOrdr")%>">
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnColOrdr<%=FieldAdd%>Up"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnColOrdr<%=FieldAdd%>Down"></td>
								</tr>
							</table></td>
						</tr>
					</table>
					</td>
					<% If OLKCType = "T" Then %>
					<td valign="top">
					<select size="1" name="ColIndex<%=FieldAdd%>" class="input">
					<option value="L" <% if rs("colindex") = "L" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtLeft")%></option>
					<option value="R" <% if rs("colindex") = "R" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtRight")%></option>
					</select></td>
					<% End If %>
					<td valign="top"><nobr>
					<select size="1" name="ColType<%=FieldAdd%>" class="input">
					<option value="T" <% if rs("coltype") = "T" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtDisabled")%></option>
					<option value="L" <% if rs("coltype") = "L" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtLow")%></option>
					<option value="M" <% if rs("coltype") = "M" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtMedium")%></option>
					<option value="H" <% if rs("coltype") = "H" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtHigh")%></option>
					</select><font face="Verdana" size="1" color="#31659C"><input type="checkbox" <% if rs("coltypernd") = "Y" then %>checked<% end if %> name="colTypeRnd<%=FieldAdd%>" id="colTypeRnd<%=FieldAdd%>" value="Y" class="noborder"><label for="colTypeRnd<%=FieldAdd%>"><%=getadminCatOptLngStr("DtxtRndLtr")%></label></font></nobr></td>
					<td valign="top" align="center">
					<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(colQuery)%>"></td>
					<td valign="top">
					<select size="1" class="input" name="colAccess<%=FieldAdd%>" onchange="javascript:document.form1.ReqLogin<%=FieldAdd%>.disabled=(this.value!='T'&&this.value!='C');">
					<option <% if rs("colaccess") = "T" then %>selected<%end if %> value="T">
					<%=getadminCatOptLngStr("DtxtAll")%></option>
					<option <% if rs("colaccess") = "V" then %>selected<%end if %> value="V">
					<%=getadminCatOptLngStr("DtxtAgent")%></option>
					<option <% if rs("colaccess") = "C" then %>selected<%end if %> value="C">
					<%=getadminCatOptLngStr("DtxtClient")%></option>
					<option <% if rs("colaccess") = "D" then %>selected<%end if %> value="D">
					<%=getadminCatOptLngStr("DtxtDisabled")%></option>
					</select></td>
					<td valign="top">
					<p align="center">
					<input type="checkbox" <% If rs("ReqLogin") = "Y" Then %>checked<% End If %> <% If rs("colaccess") <> "T" and rs("colaccess") <> "C" Then %>disabled<% End If %> class="noborder" name="ReqLogin<%=FieldAdd%>" value="Y"></td>
					<td valign="top">
					<select size="1" class="input" name="colAlign<%=FieldAdd%>" style="width: 100">
					<option <% if rs("colalign") = "center" then %>selected<%end if %> value="center">
					<%=getadminCatOptLngStr("DtxtCenter")%></option>
					<option <% if rs("colalign") = "right" then %>selected<%end if %> value="right">
					<%=getadminCatOptLngStr("DtxtRight")%></option>
					<option <% if rs("colalign") = "left" then %>selected<%end if %> value="left">
					<%=getadminCatOptLngStr("DtxtLeft")%></option>
					<option <% if rs("colalign") = "justify" then %>selected<%end if %> value="justify">
					<%=getadminCatOptLngStr("DtxtJustify")%></option>
					</select></td>
					<td valign="top" style="width: 16px">
					<a href="javascript:if(confirm('<%=getadminCatOptLngStr("LtxtValDelFld")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("ColName")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&LineIndex=<%=LinkAdd%>&submitCmd=adminCatOpt';">
					<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<script language="javascript">NumUDAttach('form1', 'ColOrdr<%=FieldAdd%>', 'btnColOrdr<%=FieldAdd%>Up', 'btnColOrdr<%=FieldAdd%>Down');</script>
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
					<input type="submit" value="<%=getadminCatOptLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
					<td width="77">
					<input type="button" value="<%=getadminCatOptLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="window.location.href='adminCatOpt.asp?NewFld=Y&amp;OLKCType=<%=OLKCType%>'"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="submitCmd" value="adminCatOpt">
		<input type="hidden" name="cmd" value="u">
	</form>
	<% End If %> <% If Request("NewFld") = "Y" or Request("edit") = "Y" Then %>
	<script language="javascript">
  function valFrm2()
  {
  	if (document.form2.colName.value == ''){
  		alert('<%=getadminCatOptLngStr("LtxtValFldNam")%>');
  		document.form2.colName.focus();
  		return false;
  	} else if (document.form2.ColQuery.value == '') {
  		alert('<%=getadminCatOptLngStr("LtxtValQry")%>');
  		document.form2.ColQuery.focus();
  		return false;
  	} else if (document.form2.valColQuery.value == 'Y') {
  		alert('<%=getadminCatOptLngStr("LtxtValQryVal")%>');
  		document.form2.btnVerfyFilter.focus();
  		return false;
  	}
  	return true;
  }
  </script>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="return valFrm2()">
		<tr>
			<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("Edit") <> "Y" Then %><%=getadminCatOptLngStr("LttlAddCustFld")%><% Else %><%=getadminCatOptLngStr("LttlEditCustFld")%><% End If %></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminCatOptLngStr("LttlCustFldNote")%></font></p>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table20">
				<% If Request("edit") = "Y" Then
		sql = "select *, (select Count('A') from OLK" & OLKCType & "Cart) ColCount from OLK" & OLKCType & "Cart where LineIndex = " & Request("LineIndex")
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open sql, conn, 3, 1
			colName = rs("colName")
			colAccess = rs("colAccess")
			colQuery = rs("colQuery")
			colType = rs("colType")
			colTypeRnd = rs("colTypeRnd")
			colAlign = rs("ColAlign")
			colOrdr = rs("ColOrdr")
			If OLKCType = "T" Then colIndex = rs("ColIndex")
			colCount = rs("ColCount")
			colTypeDec = rs("colTypeDec")
			ReqLogin = rs("ReqLogin")
		Else
			sql = "select Count('A')+1 ColCount, IsNull(Max(ColOrdr)+1, 0) ColOrdr from OLK" & OLKCType & "Cart"
			set rs = conn.execute(sql)
			colAccess = "T"
			colCount = rs("ColCount")
			colOrdr = rs("ColOrdr")
			colQuery = ""
			colName = ""
			colTypeDec = "P" %>
		<input type="hidden" name="colNameTrad">
		<input type="hidden" name="ColQueryDef">
		<% End If %>
				<tr>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtName")%></font></b></td>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtOrder")%></font></b></td>
					<% If OLKCType = "T" Then %>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtPosition2")%></font></b></td>
					<% End If %>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtCodification")%></font></b></td>
					<td align="center" bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminCatOptLngStr("DtxtDecimal")%><br>(<%=getadminCatOptLngStr("DtxtCodification")%>)</font></b></td>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtField")%></font></b></td>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtAccess")%></font></b></td>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("LtxtSesion")%></font></b></td>
					<td align="center" bgcolor="#E2F3FC"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminCatOptLngStr("DtxtAlignment")%></font></b></td>
				</tr>
				<tr>
					<td valign="top" class="style4">
					<p align="left"><font face="Verdana" size="1">
					<table cellpadding="0" cellspacing="0" border="0" width="200">
						<tr>
							<td><input name="colName" style="width: 100%;" class="input" value="<%=Server.HTMLEncode(colName)%>" size="20" onkeydown="return chkMax(event, this, 50);">
							</td>
							<td width="16"><a href="javascript:doFldTrad('<%=OLKCType%>Cart', 'LineIndex', '<%=Request("LineIndex")%>', 'alterColName', 'T', <% If Request("NewFld") <> "Y" Then %>null<% Else %>document.form2.colNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCatOptLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</font></p>
					</td>
					<td valign="top" class="style4">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
								<input type="text" name="ColOrdr" id="ColOrdr" size="4" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" value="<%=colordr%>">
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnColOrdrUp"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnColOrdrDown"></td>
								</tr>
							</table></td>
						</tr>
					</table></td>
					<% If OLKCType = "T" Then %>
					<td valign="top" class="style4">
					<select size="1" name="ColIndex" class="input">
					<option value="L" <% if colindex = "L" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtLeft")%></option>
					<option value="R" <% if colindex = "R" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtRight")%></option>
					</select></td>
					<% End If %>
					<td valign="top" class="style4">
					<nobr>
					<select size="1" name="colType" class="input">
					<option value="T" <% if coltype = "T" then %> selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtDisabled")%></option>
					<option value="L" <% if coltype = "L" then %> selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtLow")%></option>
					<option value="M" <% if coltype = "M" then %> selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtMedium")%></option>
					<option value="H" <% if coltype = "H" then %> selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtHigh")%></option>
					</select><font face="Verdana" size="1" color="#31659C"><input type="checkbox" id="colTypeRnd" name="colTypeRnd" <% if coltypernd = "Y" then %>checked<% end if %> value="Y" class="noborder"><label for="colTypeRnd"><span class="style7"><%=getadminCatOptLngStr("DtxtRndLtr")%></span></label></font></nobr></td>
					<td valign="top" class="style4">
						<select size="1" name="colTypeDec" class="input">
						<option <% If colTypeDec = "S" Then %>selected<% End If %> value="S"><%=getadminCatOptLngStr("DtxtDecSum")%></option>
						<option <% If colTypeDec = "P" Then %>selected<% End If %> value="P"><%=getadminCatOptLngStr("DtxtDecPrice")%></option>
						<option <% If colTypeDec = "R" Then %>selected<% End If %> value="R"><%=getadminCatOptLngStr("DtxtDecRate")%></option>
						<option <% If colTypeDec = "Q" Then %>selected<% End If %> value="Q"><%=getadminCatOptLngStr("DtxtDecQty")%></option>
						<option <% If colTypeDec = "%" Then %>selected<% End If %> value="%"><%=getadminCatOptLngStr("DtxtDecPercent")%></option>
						<option <% If colTypeDec = "M" Then %>selected<% End If %> value="M"><%=getadminCatOptLngStr("DtxtDecMeasure")%></option>
						</select></td>
					<td valign="top" class="style4">
					<select <% if request("edit") = "Y" then %>disabled<% end if %> size="1" name="colField" class="input" onchange="document.form2.ColQuery.value=this.value;">
					<option></option>
					<% 
					if request("edit") <> "Y" then
					sql = "select 'OITM.' + name name " & _
								"from syscolumns T0 " & _
								"where id = object_id('OITM') " & _
								"union " & _
								"select 'OITW.' + name name  " & _
								"from syscolumns T0 " & _
								"where id = object_id('OITW') "
				   set rs = conn.execute(sql)
				   do while not rs.eof %>
					<option value="<%=RS("Name")%>"><%=RS("Name")%></option>
					<% rs.movenext
					loop
					Else %>
					<option>----------</option>
					<% End If %></select></td>
					<td valign="top" class="style4">
					<select size="1" name="colAccess" class="input" onchange="javascript:document.form2.ReqLogin.disabled=(this.value!='T'&&this.value!='C');">
					<option value="T" <% if colaccess = "T" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtAll")%></option>
					<option value="V" <% if colaccess = "V" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtAgent")%></option>
					<option value="C" <% if colaccess = "C" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtClient")%></option>
					<option value="D" <% if colaccess = "D" then %>selected<% end if %>>
					<%=getadminCatOptLngStr("DtxtDisabled")%></option>
					</select></td>
					<td valign="top" class="style4">
					<p align="center">
					<input type="checkbox" <% If ReqLogin = "Y" Then %>checked<% End If %> <% If colAccess <> "T" and colAccess <> "C" Then %>disabled<% End If %> class="noborder" name="ReqLogin" value="Y"></td>
					<td valign="top" class="style4">
					<select size="1" class="input" name="colAlign">
					<option <% if colalign = "center" then %>selected<%end if %> value="center">
					<%=getadminCatOptLngStr("DtxtCenter")%></option>
					<option <% if colalign = "right" then %>selected<%end if %> value="right">
					<%=getadminCatOptLngStr("DtxtRight")%></option>
					<option <% if colalign = "left" then %>selected<%end if %> value="left">
					<%=getadminCatOptLngStr("DtxtLeft")%></option>
					<option <% if colalign = "justify" then %>selected<%end if %> value="justify">
					<%=getadminCatOptLngStr("DtxtJustify")%></option>
					</select></td>
				</tr>
				<tr>
					<% Select Case Request("OLKCType")
						Case "T"
							ColSpan = 9
							DefID = 1
						Case "C", "L"
							ColSpan = 8
							DefID = 2
					End Select %>
					<td valign="top" colspan="<%=ColSpan%>">
					<table border="0" width="100%" id="table23" cellpadding="0">
						<tr>
							<td colspan="3" class="style4">
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td rowspan="2">
										<textarea cols="78" dir="ltr" style="width: 100%;" name="ColQuery" class="input" rows="6" onkeypress="javascript:document.form2.btnVerfyFilter.src='images/btnValidate.gif';document.form2.btnVerfyFilter.style.cursor = 'hand';;document.form2.valColQuery.value='Y';"><%=myHTMLEncode(ColQuery)%></textarea>
									</td>
									<td valign="top" width="1">
										<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCatOptLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(14, 'ColQuery', '<%=DefID%><%=Request("LineIndex")%>', <% If Request("LineIndex") <> "" Then %>null<% Else %>document.form2.ColQueryDef<% End If %>);">
									</td>
								</tr>
								<tr>
									<td valign="bottom" width="1">
										<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminCatOptLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valColQuery.value == 'Y')VerfyQuery();">
										<input type="hidden" name="valColQuery" value="N">
									</td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td valign="top" bgcolor="#E2F3FC" style="width: 120px" class="style10">
							<font size="1" face="Verdana">
							<strong><%=getadminCatOptLngStr("LtxtAvlTables")%></strong></font></td>
							<td class="style4"><label for="fx1">
							<font face="Verdana" size="1" color="#4783C5">OITM = 
							<%=getadminCatOptLngStr("LtxtItemsMaster")%><br>
							OITW = <%=getadminCatOptLngStr("LtxtItmMasWhs")%><br>
							@Table = <%=getadminCatOptLngStr("LtxtDocTable")%></font></label></td>
						</tr>
						<tr>
							<td valign="top" bgcolor="#E2F3FC" style="width: 120px" class="style10">
							<font size="1" face="Verdana">
							<strong><%=getadminCatOptLngStr("DtxtVariables")%></strong></font></td>
							<td class="style4"><label for="fx1">
							<font size="1" color="#4783C5" face="Verdana">
							<span dir="ltr">@LanID</span> = <%=getadminCatOptLngStr("DtxtLanID")%><br>
							<span dir="ltr">@SlpCode</span> = <%=getadminCatOptLngStr("DtxtAgentCode")%><br>
							<span dir="ltr">@CardCode</span> = <%=getadminCatOptLngStr("LtxtCCode")%><br>
							<span dir="ltr">@PriceList</span> = <%=getadminCatOptLngStr("LtxtPList")%><br>
							<span dir="ltr">@ItemCode</span> = <%=getadminCatOptLngStr("DtxtItemCode")%><br>
							<span dir="ltr">@DocNum</span> = <%=getadminCatOptLngStr("LtxtDocNum")%> (<%=getadminCatOptLngStr("LtxtCatOnly")%>)</font></label></td>
						</tr>
						<tr>
							<td valign="top" bgcolor="#E2F3FC" style="width: 120px" class="style10">
							<font size="1" face="Verdana">
							<strong><%=getadminCatOptLngStr("DtxtFunctions")%></strong></font></td>
							<td class="style4"><% HideFunctionTitle = True
							functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
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
					<input type="submit" value="<%=getadminCatOptLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
					<td width="77">
					<input type="submit" value="<%=getadminCatOptLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
					<td width="77">
					<input type="button" value="<%=getadminCatOptLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminCatOptLngStr("DtxtConfCancel")%>'))window.location.href='adminCatOpt.asp'"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td style="height: 24px"></td>
		</tr>
		<input type="hidden" name="LineIndex" value="<%=Request("LineIndex")%>">
		<input type="hidden" name="OLKCType" value="<%=OLKCType%>">
		<input type="hidden" name="submitCmd" value="adminCatOpt">
		<input type="hidden" name="cmd" value="<% If Request("Edit") = "Y" Then %>e<% Else %>a<% End If %>">
	</form>
	<script language="javascript">NumUDAttach('form2', 'ColOrdr', 'btnColOrdrUp', 'btnColOrdrDown');</script>
	<% End If %>
</table>
<% If Request("NewFld") = "Y" or Request("edit") = "Y" Then %>
<script language="javascript">
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.form2.ColQuery.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	document.form2.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.form2.btnVerfyFilter.cursor = '';
	document.form2.valColQuery.value='N';

	//document.form2.btnVerfy.disabled = true;
}
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src>
	</iframe><input type="hidden" name="type" value="catOpt">
	<input type="hidden" name="Query" value>
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %>
<script language="javascript">
function chkNum(fld, max, min, oldVal)
{
	if (!IsNumeric(fld.value))
	{
		alert("<%=getadminCatOptLngStr("DtxtValNumVal")%>");
		fld.value = oldVal.value;
	}
	else if (parseFloat(fld.value) < parseFloat(min))
	{
		alert("<%=getadminCatOptLngStr("DtxtValNumMinVal")%>".replace("{0}", min));
		fld.value = min;
	}
	else if (parseFloat(fld.value) > 32767)
	{
		alert("<%=getadminCatOptLngStr("DtxtValNumMaxVal")%>".replace("{0}", "32767"));
		fld.value = 32767;
	}
	else
	{
		fld.value = parseInt(fld.value);
	}
	oldVal.value = fld.value;
}

</script>
<!--#include file="bottom.asp" -->