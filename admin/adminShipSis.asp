<!--#include file="top.asp" -->
<!--#include file="lang/adminShipSis.asp" -->
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
<% If Request("ShipTypeID") = "" Then %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminShipSisLngStr("LttlShipSys")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		</font><font face="Verdana" size="1" color="#4783C5"><%=getadminShipSisLngStr("LttlShipSysNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0">
			<tr>
				<td style="width: 16px" bgcolor="#E1F3FD"></td>
				<td style="width: 400px" bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminShipSisLngStr("DtxtCompany")%></strong></font></td>
				<td bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminShipSisLngStr("DtxtActive")%></strong></font></td>
			</tr>
			<% sql = "select T0.ShipTypeID, T0.Company, IsNull(T2.Active, 'N') Active " & _
				"from OLKCommon..OLKShipmentServices T0 " & _
				"inner join OLKCommon..OLKShipmentCountries T1 on T1.ShipTypeID = T0.ShipTypeID " & _
				"left outer join OLKShipment T2 on T2.ShipTypeID = T0.ShipTypeID " & _
				"where T0.Active = 'Y' and T1.Country collate database_default in ((select top 1 Country from OADM), '*') " & _
				"order by 2"
			set rs = conn.execute(sql)
			do while not rs.eof %>
			<tr>
				<td style="width: 16px" bgcolor="#F3FBFE">
				<font size="1">
				<a href='adminShipSis.asp?ShipTypeID=<%=rs("ShipTypeID")%>'>
				<font color="#4783C5" face="Verdana">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></font></a></font></td>
				<td style="width: 400px" bgcolor="#F3FBFE">
				<font size="1" color="#4783C5" face="Verdana"><%=rs("Company")%></font>
				</td>
				<td style="width: 100px" bgcolor="#F3FBFE" class="style1">
				<font size="1" color="#4783C5" face="Verdana">
				<% Select Case rs("Active")
				Case "N" %><%=getadminShipSisLngStr("DtxtNo")%>
				<% Case "Y" %><%=getadminShipSisLngStr("DtxtYes")%>
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
sql = "select T0.ShipTypeID, T0.Company, IsNull(T2.Active, 'N') Active " & _
		"from OLKCommon..OLKShipmentServices T0 " & _
		"inner join OLKCommon..OLKShipmentCountries T1 on T1.ShipTypeID = T0.ShipTypeID " & _
		"left outer join OLKShipment T2 on T2.ShipTypeID = T0.ShipTypeID " & _
		"where T1.Country collate database_default in ((select top 1 Country from OADM), '*') " & _
		"and T0.ShipTypeID = " & Request("ShipTypeID")
set rs = conn.execute(sql) %>
<table border="0" cellpadding="0" width="100%">
	<form name="frmShipSis" method="post" action="adminSubmit.asp" onsubmit="return valFrm();">
	<input type="hidden" name="submitCmd" value="adminShipSis">
	<input type="hidden" name="ShipTypeID" value="<%=Request("ShipTypeID")%>">
	<tr>
		<td bgcolor="#E1F3FD" style="height: 21px">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminShipSisLngStr("LttlShipSys")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		</font><font face="Verdana" size="1" color="#4783C5"><%=getadminShipSisLngStr("LttlShipSysFields")%></font></td>
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
							<td width="1" bgcolor="#E1F3FD"><nobr><input type="checkbox" <% If rs("Active") = "Y" Then %>checked<% End If %> name="chkActive" value="Y" id="chkActive" class="noborder"><label for="chkActive"><font face="Verdana" size="1" color="#31659C"><%=getadminShipSisLngStr("DtxtActive")%></font></label></nobr></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td bgcolor="#E2F3FC" class="style2">
					<strong><%=getadminShipSisLngStr("DtxtDescription")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					<strong><%=getadminShipSisLngStr("DtxtValue")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					<strong><%=getadminShipSisLngStr("DtxtType")%></strong></td>
				</tr>
				<%
				set rd = Server.CreateObject("ADODB.RecordSet")
				sql = "select T0.FieldID, T0.FieldName, T0.InternalDesc [Desc], T0.Type, T0.Size, T0.NotNull, IsNull(T1.Value, T0.DefValue) Value, " & _
						"Case When Exists(select '' from OLKCommon..OLKShipmentFieldsValues where ShipTypeID = T0.ShipTypeID and FieldID = T0.FieldID) Then 'Y' Else 'N' End HasValues " & _
						"from OLKCommon..OLKShipmentFields T0 " & _
						"left outer join OLKShipmentSettings T1 on T1.ShipTypeID = T0.ShipTypeID and T1.FieldID = T0.FieldID " & _
						"where T0.ShipTypeID = " & Request("ShipTypeID")
				rs.close
				rs.open sql, conn, 3, 1
				do while not rs.eof %>
				<input type="hidden" name="FieldID" value="<%=rs("FieldID")%>">
				<tr>
					<td bgcolor="#F3FBFE" class="style4" width="200">
					<font face="Verdana" size="1"><nobr><%=rs("Desc")%><% If rs("NotNull") = "Y" Then %><font color="red">*</font><% End If %></nobr></font></td>
					<td bgcolor="#F3FBFE" class="style4">
					<% If rs("HasValues") = "N" Then %>
					<input type="text" name="Value<%=rs("FieldID")%>" <% If Not IsNull(rs("Size")) Then %>maxlength="<%=rs("Size")%>"<% End If %> <% If rs("Type") = "S" and IsNull(rs("Size")) Then %>style="width: 98%; "<% End If %> <% If rs("Type") = "N" Then %> onchange="javascript:valFldNum(this);" <% End If %> value="<%=myHTMLEncode(rs("Value"))%>">
					<% Else %>
					<select size="1" name="Value<%=rs("FieldID")%>">
					<% sql = "select Value, Description [Desc] from OLKCommon..OLKShipmentFieldsValues where ShipTypeID = " & Request("ShipTypeID") & " and FieldID = " & rs("FieldID") & " order by ValID"
					set rd = conn.execute(sql)
					do while not rd.eof %>
					<option <% If Not IsNull(rs("Value")) Then %><% If CStr(rd("Value")) = CStr(rs("Value")) Then %>selected<% End If %><% End If %> value="<%=Server.HTMLEncode(rd("Value"))%>"><%=rd("Desc")%></option>
					<% rd.movenext
					loop %>
					</select>
					<% End If %></td>
					<td bgcolor="#F3FBFE" class="style4" width="100">
					<font face="Verdana" size="1"><nobr><% 
					Select Case rs("Type")
						Case "N" %><%=getadminShipSisLngStr("DtxtNumeric")%>
					<% Case "S" %><%=getadminShipSisLngStr("DtxtString")%>
					<% End Select 
					%>&nbsp;</nobr></font></td>
				</tr>
				<% rs.movenext
				loop %>
			</table>
			<br>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td colspan="3" bgcolor="#E1F3FD">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminShipSisLngStr("DtxtMeasure")%></font></b>
					</td>
				</tr>
				<tr>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminShipSisLngStr("DtxtSystem")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminShipSisLngStr("DtxtMatch")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					&nbsp;</td>
				</tr>
				<% 
				set rc = Server.CreateObject("ADODB.RecordSet")
				sql = "select T0.LenID, T0.LenCode, T0.LenDesc, IsNull(T1.Match, '') [Match] " & _
				"from OLKCommon..OLKShipmentLength T0 " & _
				"left outer join OLKShipmentLengthMatch T1 on T1.ShipTypeID = T0.ShipTypeID and T1.LenID = T0.LenID " & _
				"where T0.ShipTypeID = " & Request("ShipTypeID")
				set rd = conn.execute(sql)
				do while not rd.eof %>
				<input type="hidden" name="LenID" value="<%=rd("LenID")%>">
				<tr>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<font face="Verdana" size="1"><%=rd("LenDesc")%></font></td>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<select size="1" name="LenMatch<%=rd("LenID")%>">
					<option></option>
					<% sql = "select UnitCode, UnitName from OLGT order by 2"
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
			<br>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td colspan="3" bgcolor="#E1F3FD">
					<b><font face="Verdana" size="1" color="#31659C"><%=getadminShipSisLngStr("DtxtWeight")%></font></b>
					</td>
				</tr>
				<tr>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminShipSisLngStr("DtxtSystem")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2" style="width: 120px">
					<strong><%=getadminShipSisLngStr("DtxtMatch")%></strong></td>
					<td bgcolor="#E2F3FC" class="style2">
					&nbsp;</td>
				</tr>
				<% 
				set rc = Server.CreateObject("ADODB.RecordSet")
				sql = "select T0.WeiID, T0.WeiCode, T0.WeiDesc, IsNull(T1.Match, '') [Match] " & _
				"from OLKCommon..OLKShipmentWeight T0 " & _
				"left outer join OLKShipmentWeightMatch T1 on T1.ShipTypeID = T0.ShipTypeID and T1.WeiID = T0.WeiID " & _
				"where T0.ShipTypeID = " & Request("ShipTypeID")
				set rd = conn.execute(sql)
				do while not rd.eof %>
				<input type="hidden" name="WeiID" value="<%=rd("WeiID")%>">
				<tr>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<font face="Verdana" size="1"><%=rd("WeiDesc")%></font></td>
					<td bgcolor="#F3FBFE" class="style4" style="width: 120px">
					<select size="1" name="WeiMatch<%=rd("WeiID")%>">
					<option></option>
					<% sql = "select UnitCode, UnitName from OWGT order by 2"
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
				<input type="submit" value="<%=getadminShipSisLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>

				<td width="77">
				<input type="submit" value="<%=getadminShipSisLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr></td>
				<td width="77">
				<input type="button" value="<%=getadminShipSisLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="window.location.href='adminShipSis.asp';"></td>
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
			alert('<%=getadminShipSisLngStr("DtxtValNumValWhole")%>');
			fld.value = '';
			fld.focus();
		}
		else if (fld.value.indexOf('.') != -1)
		{
			alert('<%=getadminShipSisLngStr("DtxtValNumValWhole")%>');
			fld.value = '';
			fld.focus();
		}
	}
}
function valFrm()
{
	if (document.frmShipSis.chkActive.checked)
	{
		<% 
		If rs.recordcount > 0 Then
		rs.Filter = "NotNull = 'Y'"
		rs.movefirst
		do while not rs.eof %>
		if (document.frmShipSis.Value<%=rs("FieldID")%>.value == '')
		{
			alert('<%=getadminShipSisLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rs("Desc"), "'", "\'")%>'));
			document.frmShipSis.Value<%=rs("FieldID")%>.focus();
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