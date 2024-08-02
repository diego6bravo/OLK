<!--#include file="top.asp" -->
<!--#include file="lang/adminPaySis.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	text-align: center;
}
.style2 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
	text-align: center;
}
.style4 {
	color: #4783C5;
}
</style>
</head>
<% conn.execute("use [" & Session("OLKDB") & "]") %>
<br>
<% If Request("PayTypeID") = "" Then %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminPaySisLngStr("LttlPaySys")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		</font><font face="Verdana" size="1" color="#4783C5"><%=getadminPaySisLngStr("LttlPaySysNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0">
			<tr>
				<td style="width: 16px" bgcolor="#E1F3FD"></td>
				<td style="width: 400px" bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminPaySisLngStr("DtxtCompany")%></strong></font></td>
				<td bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminPaySisLngStr("DtxtActive")%></strong></font></td>
			</tr>
			<% sql = "select T0.PayTypeID, T0.Company, IsNull(T2.Active, 'N') Active " & _
				"from OLKCommon..OLKPaymentServices T0 " & _
				"inner join OLKCommon..OLKPaymentCountries T1 on T1.PayTypeID = T0.PayTypeID " & _
				"left outer join OLKPayment T2 on T2.PayTypeID = T0.PayTypeID " & _
				"where T0.Active = 'Y' and T1.Country collate database_default in ((select top 1 Country from OADM), '*') " & _
				"order by 2"
			set rs = conn.execute(sql)
			do while not rs.eof %>
			<tr>
				<td style="width: 16px" bgcolor="#F3FBFE">
				<font size="1">
				<a href="adminPaySis.asp?PayTypeID=<%=rs("PayTypeID")%>">
				<font color="#4783C5" face="Verdana">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></font></a></font></td>
				<td style="width: 400px" bgcolor="#F3FBFE">
				<font size="1" color="#4783C5" face="Verdana"><%=rs("Company")%></font>
				</td>
				<td style="width: 100px" bgcolor="#F3FBFE" class="style1">
				<font size="1" color="#4783C5" face="Verdana">
				<% Select Case rs("Active")
				Case "N" %><%=getadminPaySisLngStr("DtxtNo")%>
				<% Case "Y" %><%=getadminPaySisLngStr("DtxtYes")%>
				<% End Select %>
				</font>
				</td>
			</tr>
			<%  rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		&nbsp;</td>
	</tr>
</table>
<% Else
sql = "select T0.PayTypeID, T0.Company, IsNull(T2.Active, 'N') Active " & _
		"from OLKCommon..OLKPaymentServices T0 " & _
		"inner join OLKCommon..OLKPaymentCountries T1 on T1.PayTypeID = T0.PayTypeID " & _
		"left outer join OLKPayment T2 on T2.PayTypeID = T0.PayTypeID " & _
		"where T1.Country collate database_default in ((select top 1 Country from OADM), '*') " & _
		"and T0.PayTypeID = " & Request("PayTypeID")
set rs = conn.execute(sql) %>
<table border="0" cellpadding="0" width="100%">
	<form name="frmPaySis" method="post" action="adminSubmit.asp" onsubmit="return valFrm();">
	<input type="hidden" name="submitCmd" value="adminPaySis">
	<input type="hidden" name="PayTypeID" value="<%=Request("PayTypeID")%>">
	<tr>
		<td bgcolor="#E1F3FD" style="height: 21px">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminPaySisLngStr("LttlPaySys")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		</font><font face="Verdana" size="1" color="#4783C5"><%=getadminPaySisLngStr("LttlPaySysFields")%></font></td>
	</tr>
	<tr>
		<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td colspan="3">
					<table style="width: 100%">
						<tr>
							<td bgcolor="#E1F3FD"><b><font face="Verdana" size="1" color="#31659C">
							<%=rs("Company")%></font></b></td>
							<td width="1" bgcolor="#E1F3FD"><nobr><input type="checkbox" <% If rs("Active") = "Y" Then %>checked<% End If %> name="chkActive" value="Y" id="chkActive" class="noborder"><label for="chkActive"><font face="Verdana" size="1" color="#31659C"><%=getadminPaySisLngStr("DtxtActive")%></font></label></nobr></td>
						</tr>
					</table>
					</td>
				</tr>
				<%
				set rd = Server.CreateObject("ADODB.RecordSet")
				sql = "select T0.FieldID, T0.FieldName, T0.InternalDesc [Desc], T0.Type, T0.Size, T0.NotNull, IsNull(T1.Value, T0.DefValue) Value, " & _
						"Case When Exists(select '' from OLKCommon..OLKPaymentFieldsValues where PayTypeID = T0.PayTypeID and FieldID = T0.FieldID) Then 'Y' Else 'N' End HasValues " & _
						"from OLKCommon..OLKPaymentFields T0 " & _
						"left outer join OLKPaymentSettings T1 on T1.PayTypeID = T0.PayTypeID and T1.FieldID = T0.FieldID " & _
						"where T0.PayTypeID = " & Request("PayTypeID")
				rs.close
				rs.open sql, conn, 3, 1
				If Not rs.Eof Then %>
				<tr>
					<td bgcolor="#E2F3FC" class="style2">
					<strong><%=getadminPaySisLngStr("DtxtDescription")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					<strong><%=getadminPaySisLngStr("DtxtValue")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					<strong><%=getadminPaySisLngStr("DtxtType")%></strong></td>
				</tr>
				<% do while not rs.eof %>
				<input type="hidden" name="FieldID" value="<%=rs("FieldID")%>">
				<tr>
					<td bgcolor="#F3FBFE" class="style4" width="200">
					<font face="Verdana" size="1"><nobr><%=rs("Desc")%><% If rs("NotNull") = "Y" Then %><font color="red">*</font><% End If %></nobr></font></td>
					<td bgcolor="#F3FBFE" class="style4">
					<% If rs("HasValues") = "N" Then %>
					<input type="text" name="Value<%=rs("FieldID")%>" <% If Not IsNull(rs("Size")) Then %>maxlength="<%=rs("Size")%>"<% End If %> <% If rs("Type") = "S" and IsNull(rs("Size")) Then %>style="width: 98%; "<% End If %> <% If rs("Type") = "N" Then %> onchange="javascript:valFldNum(this);" <% End If %> value="<%=myHTMLEncode(rs("Value"))%>">
					<% Else %>
					<select size="1" name="Value<%=rs("FieldID")%>">
					<% sql = "select Value, Description [Desc] from OLKCommon..OLKPaymentFieldsValues where PayTypeID = " & Request("PayTypeID") & " and FieldID = " & rs("FieldID") & " order by ValID"
					set rd = conn.execute(sql)
					do while not rd.eof %>
					<option <% If CStr(rd("Value")) = CStr(rs("Value")) Then %>selected<% End If %> value="<%=Server.HTMLEncode(rd("Value"))%>"><%=rd("Desc")%></option>
					<% rd.movenext
					loop %>
					</select>
					<% End If %></td>
					<td bgcolor="#F3FBFE" class="style4" width="100">
					<font face="Verdana" size="1"><nobr><% 
					Select Case rs("Type")
						Case "N" %><%=getadminPaySisLngStr("DtxtNumeric")%>
					<% Case "S" %><%=getadminPaySisLngStr("DtxtString")%>
					<% End Select 
					%>&nbsp;</nobr></font></td>
				</tr>
				<% rs.movenext
				loop
				End If %>
			</table>
			<% 
			set rc = Server.CreateObject("ADODB.RecordSet")
			sql = "select T0.CurrID, T0.CurrCode, T0.CurrDesc, IsNull(T1.Match, '') [Match] " & _
			"from OLKCommon..OLKPaymentCurrencies T0 " & _
			"left outer join OLKPaymentCurMatch T1 on T1.PayTypeID = T0.PayTypeID and T1.CurrID = T0.CurrID " & _
			"where T0.PayTypeID = " & Request("PayTypeID")
			set rd = conn.execute(sql)
			If Not rd.Eof Then %>
			<br>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td colspan="3" bgcolor="#E1F3FD">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminPaySisLngStr("DtxtCurr")%></font></b>
					</td>
				</tr>
				<tr>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminPaySisLngStr("DtxtSystem")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminPaySisLngStr("DtxtMatch")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					&nbsp;</td>
				</tr>
				<% do while not rd.eof %>
				<input type="hidden" name="CurrID" value="<%=rd("CurrID")%>">
				<tr>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<font face="Verdana" size="1"><%=rd("CurrDesc")%></font></td>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<select size="1" name="CurMatch<%=rd("CurrID")%>">
					<option></option>
					<% 
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetCurrencies" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					set rc = cmd.execute()
					do while not rc.eof %>
					<option <% If CStr(rc(0)) = CStr(rd("Match")) Then %>selected<% End If %> value="<%=Server.HTMLEncode(rc(0))%>"><%=rc(1)%></option>
					<% rc.movenext
					loop %>
					</select></td>
					<td bgcolor="#F3FBFE" class="style4">
					&nbsp;</td>
				</tr>
				<% rd.movenext
				loop %>
			</table>
			<% End If %>
			<br>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td colspan="3" bgcolor="#E1F3FD">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminPaySisLngStr("LtxtCreditCards")%></font></b>
					</td>
				</tr>
				<tr>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminPaySisLngStr("DtxtSystem")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminPaySisLngStr("DtxtMatch")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					&nbsp;</td>
				</tr>
				<% 
				set rc = Server.CreateObject("ADODB.RecordSet")
				sql = "select T0.CardID, T2.CardName, IsNull(T1.Match, -1) [Match] " & _
					"from OLKCommon..OLKPaymentCards T0 " & _
					"left outer join OLKPaymentCardMatch T1 on T1.PayTypeID = T0.PayTypeID and T1.CardID = T0.CardID " & _
					"inner join OLKCommon..OLKCards T2 on T2.CardID = T0.CardID " & _
					"where T0.PayTypeID = " & Request("PayTypeID")
				set rd = conn.execute(sql)
				do while not rd.eof %>
				<input type="hidden" name="CardID" value="<%=rd("CardID")%>">
				<tr>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<font face="Verdana" size="1"><%=rd("CardName")%></font></td>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<select size="1" name="CardMatch<%=rd("CardID")%>">
					<option></option>
					<% sql = "select CreditCard, CardName from OCRC where Locked = 'N'"
					set rc = conn.execute(sql)
					do while not rc.eof %>
					<option <% If CStr(rc(0)) = CStr(rd("Match")) Then %>selected<% End If %> value="<%=Server.HTMLEncode(rc(0))%>"><%=rc(1)%></option>
					<% rc.movenext
					loop %>
					</select></td>
					<td bgcolor="#F3FBFE" class="style4">
					&nbsp;</td>
				</tr>
				<% rd.movenext
				loop %>
			</table>
		</td>
	</tr>
	<tr bgcolor="#F5FBFE">
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminPaySisLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>

				<td width="77">
				<input type="submit" value="<%=getadminPaySisLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr></td>
				<td width="77">
				<input type="button" value="<%=getadminPaySisLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="window.location.href='adminPaySis.asp';"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		&nbsp;</td>
	</tr>
	</form>
</table>
<script type="text/javascript">
<!--
function valFldNum(fld)
{
	if (fld.value != '')
	{
		if (!IsNumeric(fld.value))
		{
			alert('<%=getadminPaySisLngStr("DtxtValNumValWhole")%>');
			fld.value = '';
			fld.focus();
		}
		else if (fld.value.indexOf('.') != -1)
		{
			alert('<%=getadminPaySisLngStr("DtxtValNumValWhole")%>');
			fld.value = '';
			fld.focus();
		}
	}
}
function valFrm()
{
	if (document.frmPaySis.chkActive.checked)
	{
		<% 
		If rs.recordcount > 0 Then
		rs.Filter = "NotNull = 'Y'"
		do while not rs.eof %>
		if (document.frmPaySis.Value<%=rs("FieldID")%>.value == '')
		{
			alert('<%=getadminPaySisLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rs("Desc"), "'", "\'")%>'));
			document.frmPaySis.Value<%=rs("FieldID")%>.focus();
			return false;
		}
		<% rs.movenext
		loop
		End If %>
	}
	return true;
}
//-->
</script>
<% End If %><!--#include file="bottom.asp" -->