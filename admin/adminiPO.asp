<!--#include file="top.asp" -->
<!--#include file="lang/adminiPO.asp" -->
<!--#include file="adminTradSubmit.asp" -->
<script language="javascript" src="js_up_down.js"></script>
<br>
<form name="form1" action="adminSubmit.asp"  method="post">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminiPOLngStr("LttliPOPList")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminiPOLngStr("LttliPOPListNote")%> </font></td>
	</tr><%
		sql = "select T0.ListNum, T0.ListName, IsNull(T1.Ordr, ROW_NUMBER() OVER(order by T1.Ordr)+(select MAX(Ordr) from OLKIPOPLIST)) Ordr, IsNull(T1.Active, 'N') Active " & _  
		"from OPLN T0 " & _  
		"left outer join OLKIPOPLIST T1 on T1.ListNum = T0.ListNum " & _  
		"order by [Ordr] " 
		set rs = conn.execute(sql) %>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0">
				<tr>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminiPOLngStr("DtxtPriceList")%></font></td>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminiPOLngStr("DtxtOrder")%></font></td>
					<td bgcolor="#F7FBFF">&nbsp;</td>
				</tr>
				<% do while not rs.eof
				ID = rs("ListNum") %>
				<tr>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5">
					<%=rs("ListName")%></font></td>
					<td bgcolor="#F7FBFF" align="center">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td>
								<input type="text" name="Order<%=ID%>" id="Order<%=ID%>" size="7" style="text-align:right" class="input"onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Ordr")%>">
							</td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnOrder<%=ID%>Up"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnOrder<%=ID%>Down"></td>
								</tr>
							</table></td>
						</tr>
					</table>
					</td>
					<td bgcolor="#F7FBFF" align="center">
					<font face="Verdana" size="1" color="#4783C5">
					<input type="checkbox" name="chkListActive<%=rs("ListNum")%>" <% If rs("Active") = "Y" Then %>checked<% End If %> id="chkListActive<%=rs("ListNum")%>" value="Y" class="noborder"><label for="chkListActive<%=rs("ListNum")%>"><%=getadminiPOLngStr("DtxtActive")%></label></font>
				<input type="hidden" name="ListNumID" value="<%=ID%>">
				<script language="javascript">NumUDAttach('form1', 'Order<%=ID%>', 'btnOrder<%=ID%>Up', 'btnOrder<%=ID%>Down');</script></td>
				</tr>
				<% rs.movenext
				loop %>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminiPOLngStr("LttliPOPItemDet")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminiPOLngStr("LttliPOPItemDetNote")%> </font></td>
	</tr><%
sql = "select T0.rowIndex, T0.rowName, ISNULL(T1.Active, 'N') Active " & _  
"from OLKItemRep T0 " & _  
"left outer join OLKIPOITMFLD T1 on T1.rowIndex = T0.rowIndex " & _  
"order by T0.[rowOrder] " 
set rs = conn.execute(sql) %>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0">
				<tr>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminiPOLngStr("DtxtField")%></font></td>
					<td bgcolor="#F7FBFF">
					&nbsp;</td>
				</tr>
				<% do while not rs.eof
				ID = Replace(rs("rowIndex"), "-", "_") %>
				<tr>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5">
					<%=rs("rowName")%></font></td>
					<td bgcolor="#F7FBFF" align="center">
					<font face="Verdana" size="1" color="#4783C5">
					<input type="checkbox" name="chkRowActive<%=ID%>" <% If rs("Active") = "Y" Then %>checked<% End If %> id="chkRowActive<%=ID%>" value="Y" class="noborder"><label for="chkRowActive<%=ID%>"><%=getadminiPOLngStr("DtxtActive")%></label>
					<input type="hidden" name="rowID" value="<%=rs("rowIndex")%>"></font></td>
				</tr>
				<% rs.movenext
				loop %>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminiPOLngStr("LttliPOPFld")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminiPOLngStr("LttliPOPFldNote")%> </font></td>
	</tr>
<% sql = "select * from OLKIPOITM"
	set rs = conn.execute(sql) %>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td valign="top" class="style5" style="width: 100px">
					<font size="1" face="Verdana" color="#31659C"><strong><%=getadminiPOLngStr("DtxtVariables")%></strong></font></td>
					<td class="style4">
					<font face="Verdana" size="1" color="#4783C5">
					<span dir="ltr">@ItemCode</span> = <%=getadminiPOLngStr("DtxtItemCode")%><br>
					<span dir="ltr">@SlpCode</span> = <%=getadminiPOLngStr("DtxtAgentCode")%><br>
					<span dir="ltr">@dbName</span> = <%=getadminiPOLngStr("DtxtDB")%><br>
					<span dir="ltr">@WhsCode</span> = <%=getadminiPOLngStr("DtxtWhsCode")%><br>
					<span dir="ltr">@LanID</span> = <%=getadminiPOLngStr("DtxtLanID")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminiPOLngStr("DtxtField")%></font></td>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5"><%=getadminiPOLngStr("DtxtQuery")%></font></td>
				</tr>
				<% For each itm in rs.Fields %>
				<tr>
					<td bgcolor="#F7FBFF" valign="top" width="200">
					<font face="Verdana" size="1" color="#4783C5">
					<%=itm.Name%></font></td>
					<td bgcolor="#F7FBFF" align="center">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td rowspan="2">
								<textarea dir="ltr" rows="10" style="width: 100%" name="qry<%=itm.Name%>" id="qry<%=itm.Name%>" cols="100" class="input" onkeypress="javascript:document.form1.btnVerfyFilter<%=itm.Name%>.src='images/btnValidate.gif';document.form1.btnVerfyFilter<%=itm.Name%>.style.cursor = 'hand';;document.form1.valQuery<%=itm.Name%>.value='Y';"><%=myHTMLEncode(itm)%></textarea>
							</td>
							<td valign="top" width="1">
								<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter<%=itm.Name%>" alt="<%=getadminiPOLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(24, 'qry<%=itm.Name%>', '', null);">
							</td>
						</tr>
						<tr>
							<td valign="bottom" width="1">
								<img src="images/btnValidateDis.gif" id="btnVerfyFilter<%=itm.Name%>" alt="<%=getadminiPOLngStr("DtxtValidate")%>" onclick="javascript:if (document.form1.valQuery<%=itm.Name%>.value == 'Y')VerfyQuery('<%=itm.Name%>');">
								<input type="hidden" name="valQuery<%=itm.Name%>" value="N">
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<% Next %>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminiPOLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminiPO">
</table>
</form>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="iPO">
	<input type="hidden" name="Field" value="">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<script type="text/javascript">
var myBtnVerfy;
var myHdVerfy;
function VerfyQuery(field)
{
	document.frmVerfyQuery.Field.value = field;
	document.frmVerfyQuery.Query.value = document.getElementById('qry' + field).value;
	myBtnVerfy = document.getElementById('btnVerfyFilter' + field);
	myHdVerfy = document.getElementById('valQuery' + field)
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	//myBtnVerfy.disabled = true;
	myBtnVerfy.src='images/btnValidateDis.gif'
	myBtnVerfy.style.cursor = '';
	myHdVerfy.value='N';
}

</script>
<!--#include file="bottom.asp" -->