<!--#include file="top.asp" -->
<!--#include file="lang/adminSecIndex.asp" -->
<% conn.execute("use [" & Session("OLKDB") & "]")
sql = "select SecIndexByX, SecIndexByY from OLKCommon"
set rs = conn.execute(sql)
SecIndexByX = rs("SecIndexByX")
SecIndexByY = rs("SecIndexByY")
rs.close %>

<head>
<style>

.tdIndex     { font-family: Verdana; font-size: 10px; color: #4783C5 }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<table border="0" cellpadding="0" width="100%" id="table1">
	<form method="POST" name="frmSecIndex" action="adminSubmit.asp" onsubmit="return valFrm();">
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminSecIndexLngStr("LttlSecIndex")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"> <%=getadminSecIndexLngStr("LttlSecIndexNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellspacing="0" width="300" id="table23">
			<tr>
				<td><font face="Verdana" size="1" color="#4783C5"><%=getadminSecIndexLngStr("LtxtRows")%></font></td>
				<td><select size="1" name="SecIndexByY" class="input" onchange="doIndex();">
				<% For i = 1 to 10 %>
				<option <% If SecIndexByY = i Then %>selected<% End If %> value="<%=i%>"><%=i%></option>
				<% Next %>
				</select></td>
				<td><font face="Verdana" size="1" color="#4783C5"><%=getadminSecIndexLngStr("LtxtCols")%></font></td>
				<td><select size="1" name="SecIndexByX" class="input" onchange="doIndex();">
				<% For i = 1 to 3 %>
				<option <% If SecIndexByX = i Then %>selected<% End If %> value="<%=i%>"><%=i%></option>
				<% Next %>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminSecIndexLngStr("LttlDistInd")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"> <%=getadminSecIndexLngStr("LttlDistIndNote")%></font></td>
	</tr>
	<% sql = "select T0.SecID, IsNull(T2.AlterSecName, T0.SecName) SecName, T1.ID OrderID " & _
			"from OLKSections T0 " & _
			"left outer join OLKSectionsAlterNames T2 on T2.SecType = T0.SecType and T2.SecID = T0.SecID and T2.LanID = " & Session("LanID") & " " & _
			"left outer join OLKSecIndex T1 on T1.SecID = T0.SecID " & _
			"where T0.SecType = 'U' and T0.Status = 'A' and T0.UserType = 'C' "
	rs.open sql, conn, 3, 1 %>
	<script language="javascript"></script>
	<tr>
		<td bgcolor="#F5FBFE" id="tdIndex">
		<table cellpadding="0" cellspacing="0" border="0" width="500">
			<% x = 0
			For i = 1 to SecIndexByY %>
			<tr>
				<% For j = 1 to SecIndexByX
					x = x + 1 %>
				<td width="33%" align="center" class="tdIndex">
					<%=x%><br>
					<select name="SecID<%=x%>" id="cmbSecID" size="1" class="input">
					<option></option>
					<% If rs.recordcount > 0 Then rs.movefirst
					do while not rs.eof %>
					<option <% If rs("OrderID") = x Then %>selected<% End If %> value="<%=rs("SecID")%>"><%=myHTMLEncode(rs("SecName"))%></option>
					<% rs.movenext
					loop %>
					</select>
				</td>
				<% Next %>
			</tr>
			<% Next %>
		</table></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table21">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSecIndexLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminSecIndexLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
					<input type="button" value="<%=getadminSecIndexLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminSecIndexLngStr("DtxtConfCancel")%>'))window.location.href='adminSec.asp?UType=C'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminSecIndex">
	</form>
</table>
<script language="javascript">
var txtValSelIndex = '<%=getadminSecIndexLngStr("LtxtValSelIndex")%>';
var txtValEqIndex = '<%=getadminSecIndexLngStr("LtxtValEqIndex")%>';
</script>
<script language="javascript">
var mySec = '<% If rs.RecordCount > 0 Then
rs.movefirst
do while not rs.eof %><% If rs.bookmark > 1 Then Response.Write "," %><%=rs("SecID")%>|<%=Replace(rs("SecName"), "'", "\'")%><% rs.movenext
loop
End If %>'.split(',');
</script>
<script language="javascript" src="adminSecIndex.js"></script>
<!--#include file="bottom.asp" -->