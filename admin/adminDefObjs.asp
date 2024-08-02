<!--#include file="top.asp" -->
<!--#include file="lang/adminDefObjs.asp" -->
<% conn.execute("use [" & Session("olkdb") & "]") %>
<script language="javascript" src="js_up_down.js"></script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E1F3FD;
}
.style2 {
	font-weight: bold;
	background-color: #E1F3FD;
}
</style>
</head>

<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td height="15"></td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmSec">
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminDefObjsLngStr("LttlCustObjs")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminDefObjsLngStr("LttlCustObjsNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table11">
			<tr>
				<td width="20" class="style1"><font size="1">&nbsp;</font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminDefObjsLngStr("DtxtName")%></font></td>
				<td align="center" class="style2">
				<font face="Verdana" size="1" color="#31659C"><%=getadminDefObjsLngStr("DtxtActive")%></font></td>
				<% If 1 = 2 Then %><td class="style1" style="width: 16px"><font size="1">&nbsp;</font></td><% End If %>
				</tr>
			<% 	sql = "select T0.ObjType, T0.ObjID,  " & _
				"Case T0.ObjType When 'S' Then T1.ObjName collate database_default Else T0.ObjName End ObjName, T0.Status  " & _
				"from OLKObjects T0 " & _
				"left outer join OLKCommon..OLKObjectsDesc T1 on T1.ObjID = T0.ObjID and T0.ObjType = 'S' and T1.LanID = " & Session("LanID") & " " & _
				"where T0.Status <> 'D' " & _
				"order by Case When T0.ObjID = 20 Then 9 When T0.ObjID >= 9 Then T0.ObjID + 1 End "
				rs.open sql, conn, 3, 1
				do while not rs.eof %>
			<tr>
				<td width="20" bgcolor="#F3FBFE"><a href="adminDefObjEdit.asp?ObjType=<%=rs("ObjType")%>&ObjId=<%=rs("ObjId")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td bgcolor="#F3FBFE">
				<font face="Verdana" size="1" color="#4783C5"><%=rs("ObjName")%></font>&nbsp;</td>
				<td bgcolor="#F3FBFE">
				<p align="center">
				<input <% If rs("Status") = "Y" Then %>checked<% End If %> type="checkbox" name="Status<%=rs("ObjType")%><%=rs("ObjID")%>" value="Y" class="noborder"></td>
				<% If 1 = 2 Then %><td bgcolor="#F3FBFE" style="width: 16px"><% If rs("ObjType") = "U" Then %><a href="javascript:if(confirm('<%=getadminDefObjsLngStr("LtxtConfDelObj")%>'.replace('{0}', '<%=rs("ObjName")%>')))window.location.href='adminSubmit.asp?submitCmd=adminObjs&uCmd=del&ObjID=<%=rs("ObjID")%>'"><img border="0" src="images/remove.gif" width="16" height="18"></a><% End If %></td><% End If %>
				</tr>
			<% rs.movenext
			loop %>
		</table>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminDefObjsLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
		<input type="hidden" name="uCmd" value="update">
	<input type="hidden" name="submitCmd" value="adminObjs">
	</form>
</table>
<!--#include file="bottom.asp" -->