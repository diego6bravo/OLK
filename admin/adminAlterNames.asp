<!--#include file="top.asp" -->
<!--#include file="lang/adminAlterNames.asp" -->
<% conn.execute("use [" & Session("OLKDB") & "]") %>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
<!--
function valFrm()
{
	s = document.frmAlertNames.Sing;
	p = document.frmAlertNames.Plur;
	for (var i = 0;i<s.length;i++)
	{
		if (s(i).value == '' && !s(i).disabled)
		{
			alert('<%=getadminAlterNamesLngStr("LtxtValAllFld")%>');
			s(i).focus();
			return false;
		}
		else if (p(i).value == '' && !p(i).disabled)
		{	
			alert('<%=getadminAlterNamesLngStr("LtxtValAllFld")%>');
			p(i).focus();
			return false;
		}
	}
	return true;
}
//-->
</script>
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
}
.style2 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
	text-align: center;
}
.style3 {
	background-color: #E2F3FC;
	color: #31659C;
	text-align: center;
}
.style4 {
	background-color: #E2F3FC;
	color: #31659C;
}
</style>
</head>

<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<form method="POST" action="adminSubmit.asp" name="frmAlertNames" onsubmit="return valFrm();">
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminAlterNamesLngStr("LttlAlterNames")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font color="#4783C5" face="Verdana" size="1"><%=getadminAlterNamesLngStr("LttlAlterNamesNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table6">
			<tr>
				<td class="style4">
		<font face="Verdana" size="1"><strong><%=getadminAlterNamesLngStr("LtxtLanguage")%></strong></font></td>
				<td width="200" colspan="2" style="width: 400px" bgcolor="#F5FBFE">
				<font face="Verdana" size="1" color="#4783C5">
		<select size="1" name="AlterLng" class="input" onchange="javascript:window.location.href='?AlterLng='+this.value;">
		<option value=""><%=getadminAlterNamesLngStr("LoptSelLng")%></option>
		<% For i = 0 to UBound(myLanIndex) %>
		<option <% If CStr(Request("AlterLng")) = CStr(myLanIndex(i)(4)) Then %>selected<% End If %> value="<%=myLanIndex(i)(4)%>"><%=myLanIndex(i)(1)%></option>
		<% Next %>
		</select></font></td>
			</tr>
			<% If Request("AlterLng") <> "" Then %>
			<tr>
				<td class="style1">
				&nbsp;</td>
				<td width="200" class="style2">
				<font face="Verdana" size="1"><strong><%=getadminAlterNamesLngStr("LtxtSingular")%></strong></font></td>
				<td width="200" class="style3">
				<font face="Verdana" size="1"><strong><%=getadminAlterNamesLngStr("LtxtPlural")%></strong></font></td>
			</tr>
			<% sql = "select * from OLKAlterNames where LanID = " & Request("AlterLng") & " order by AlterDesc asc"
			rs.open sql, conn, 3, 1
			rs.Filter = "AlterID <> 21 and AlterID <> 22 and AlterID <> 23 and AlterID <> 24"
			Alter = True
			do while not rs.eof %>
			<tr>
				<td bgcolor="#F3FBFE">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> 
				<font color="#4783C5"><%=rs("AlterDesc")%>&nbsp;</font></font></td>
				<td width="200" bgcolor="#F3FBFE">
				<input type="text" name="Sing<%=rs("AlterID")%>" id="Sing" size="30" class="input" style="width: 100%; <% If IsNull(rs("Singular")) Then %>background-color: #CCCCCC;<% End If %>" <% If IsNull(rs("Singular")) Then %> disabled <% End If %> value="<% If Not IsNull(rs("Singular")) Then %><%=Server.HTMLEncode(rs("Singular"))%><% End If %>" onkeydown="return chkMax(event, this, 50);"></td>
				<td width="200" bgcolor="#F3FBFE">
				<input type="text" name="Plur<%=rs("AlterID")%>" id="Plur" size="30" class="input" style="width: 100%; <% If IsNull(rs("Plural")) Then %>background-color: #CCCCCC;<% End If %>" <% If IsNull(rs("Plural")) Then %> disabled <% End If %> value="<% If Not IsNull(rs("Plural")) Then %><%=Server.HTMLEncode(rs("Plural"))%><% End If %>" onkeydown="return chkMax(event, this, 50);"></td>
			</tr>
			<% Alter = Not Alter
			rs.movenext
			loop %>
			<tr>
				<td class="style1">
				&nbsp;</td>
				<td width="200" class="style2">
				<font face="Verdana" size="1"><strong><%=getadminAlterNamesLngStr("DtxtClient")%></strong></font></td>
				<td width="200" class="style3">
				<font face="Verdana" size="1"><strong><%=getadminAlterNamesLngStr("DtxtAgent")%></strong></font></td>
			</tr>
			<% rs.Filter = "AlterID = 21 or AlterID = 22 or AlterID = 23 or AlterID = 24"
			Alter = True
			do while not rs.eof %>
			<tr>
				<td bgcolor="#F3FBFE">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> 
				<font color="#4783C5"><%=rs("AlterDesc")%>&nbsp;</font></font></td>
				<td width="200" bgcolor="#F3FBFE">
				<input type="text" name="Sing<%=rs("AlterID")%>" id="Sing" size="30" class="input" style="width: 100%;<% If IsNull(rs("Singular")) Then %>background-color: #CCCCCC;<% End If %>" <% If IsNull(rs("Singular")) Then %>disabled<% End If %> value="<%=Server.HTMLEncode(rs("Singular"))%>" onkeydown="return chkMax(event, this, 50);"></td>
				<td width="200" bgcolor="#F3FBFE">
				<input type="text" name="Plur<%=rs("AlterID")%>" id="Plur" size="30" class="input" style="width: 100%;<% If IsNull(rs("Plural")) Then %>background-color: #CCCCCC;<% End If %>" <% If IsNull(rs("Plural")) Then %>disabled<% End If %> value="<%=Server.HTMLEncode(rs("Plural"))%>" onkeydown="return chkMax(event, this, 50);"></td>
			</tr>
			<% Alter = Not Alter
			rs.movenext
			loop %>
			<% End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" <% If Request("AlterLng") = "" Then %>disabled<% End If %> value="<%=getadminAlterNamesLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	</table>
<input type="hidden" name="submitCmd" value="alterNames">
</form>
<!--#include file="bottom.asp" -->