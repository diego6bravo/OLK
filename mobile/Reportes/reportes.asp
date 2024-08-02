<% addLngPathStr = "Reportes/" %>
<!--#include file="lang/reportes.asp" -->
<link rel="stylesheet" href="Reportes/style.css">
<table border="0" width="100%" id="table1" cellpadding="0">
	<tr>
		<td height="4"></td>
	</tr>
	<% sql = "select Case When T0.rgIndex < 0 Then 'A' Else 'B' End Ordr1, T0.RGIndex, IsNull(T1.alterRGName, T0.rgName) rgName, "
		If Request("rsIndex") = "" Then
			sql = sql & "'N' "
		Else
			sql = sql & "Case When Exists(select 'A' from OLKRS where rgIndex = T0.rgIndex and rsIndex = " & Request("rsIndex") & ") Then 'Y' Else 'N' End "
		End If
	   sql = sql & " Verfy from OLKRG T0 " & _
	   		"left outer join OLKRGAlterNames T1 on T1.rgIndex = T0.rgIndex and T1.LanID = " & Session("LanID") & " " & _
			"where exists(select 'A' from OLKRS where rgIndex = T0.rgIndex and Active = 'Y' and LinkOnly = 'N') and T0.UserType = 'V' "
			
		If Session("useraccess") = "U" Then
			If myAut.AuthorizedRepGroups <> "" Then
				sql = sql & "and T0.rgIndex in (" & myAut.AuthorizedRepGroups & ") "
			Else
				sql = sql & "and 1 = 2 "
			End If
		End If
		
		sql = sql & " order by 1, 3"
		
		rs.open sql, conn, 3, 1
		
		set rp = Server.CreateObject("ADODB.RecordSet")
		sql = "select T0.rsIndex, IsNull(T2.alterRSName, T0.rsName) rsName, T0.rgIndex, (select Count('A') from OLKRSVars where rsIndex = T0.rsIndex)+Case rsTop When 'Y' Then 1 Else 0 End varCount " & _
				"from OLKRS T0 " & _
				"inner join OLKRG T1 on T1.rgIndex = T0.rgIndex " & _
				"left outer join OLKRSAlterNames T2 on T2.rsIndex = T0.rsIndex and T2.LanID = " & Session("LanID") & " " & _
				"where T0.Active = 'Y' and T0.LinkOnly = 'N' "
				
		'If Session("RPAccess") = "Y" Then sql = sql & "and T1.SuperUser = 'N' and exists(select 'A' from OLKRGAccess where SlpCode = " & Session("RPUID") & " and rgIndex = T1.rgIndex) "

		If Session("useraccess") = "U" Then
			If myAut.AuthorizedRepGroups <> "" Then
				sql = sql & "and T0.rgIndex in (" & myAut.AuthorizedRepGroups & ") "
			Else
				sql = sql & "and 1 = 2 "
			End If
			' sql = sql & "and T0.SuperUser = 'N' and exists(select 'A' from OLKRGAccess where SlpCode = " & Session("vendid") & " and rgIndex = T0.rgIndex) "
		End If
		
		sql = sql & " order by 2"

		
		rp.open sql, conn, 3, 1 %>
		<form method="post" action="operaciones.asp">
		<tr class="TblTltMnu">
			<td>
			<select name="rgIndex" size="1" onchange="submit();">
			<option value=""><%=getreportesLngStr("DtxtAll")%></option>
			<% do while not rs.eof %>
			<option <% If CStr(Request("rgIndex")) = CStr(rs("rgIndex")) Then %>selected<% End If %> value="<%=rs("rgIndex")%>"><%=Server.HTMLEncode(rs("rgName"))%></option>
			<% rs.movenext
			loop
			if rs.recordcount > 0 then rs.movefirst %>
			</select></td>
		</tr>
		<input type="hidden" name="cmd" value="reportes">
		</form>
		<% If Request("rgIndex") <> "" Then rs.Filter = "rgIndex = " & Request("rgIndex")
		do while not rs.eof %>
			<% If Request("rgIndex") = "" Then %>
			<tr class="TblTltMnu">
				<td>
				<img border="0" src="images/ball.gif" width="6" height="6"><%=Server.HTMLEncode(rs("rgName"))%>
				</td>
			</tr>
			<% End If %>
			<% rp.Filter = "rgIndex = " & rs("rgIndex")
			do while not rp.eof %>
			<tr class="TblAfueraMnu">
				<td>
				<a href="javascript:goRep(<%=rp("rsIndex")%>,<%=rp("varCount")%>);">
				<img border="0" src="images/arrow_menu.gif" width="9" height="6"><%=Server.HTMLEncode(rp("rsName"))%></a>
				</td>
			</tr>
			<% rp.movenext
			loop %>
		<% rs.movenext
		loop %>
</table>
<script language="javascript">
function goRep(rsIndex, varsCount)
{
	if (varsCount == 0) document.frmReps.cmd.value='viewRep';
	else document.frmReps.cmd.value='viewRepVals';
	document.frmReps.rsIndex.value = rsIndex;
	document.frmReps.submit();
}
</script>
<form method="POST" action="operaciones.asp" name="frmReps">
<input type="hidden" name="cmd" value="viewRepVals">
<input type="hidden" name="rsIndex" value="">
</form>