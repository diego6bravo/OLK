<!--#include file="top.asp" -->
<!--#include file="lang/adminLanguages.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style2 {
	background-color: #E1F3FD;
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
	text-align: center;
	font-weight: bold;
}
.style3 {
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
	background-color: #F7FBFF;
}
.style4 {
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
	background-color: #F7FBFF;
	text-align: right;
	}
</style>
</head>
<% conn.execute("use [" & Session("olkdb") & "]")
sql = "select LanID from OLKDisLng"
rs.open sql, conn, 3, 1 %>
<br>
<table border="0" cellpadding="0" width="100%" id="table3">
<form method="post" action="adminSubmit.asp">
	<tr>
		<td bgcolor="#E7F3FF">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminLanguagesLngStr("LttlLanguages")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminLanguagesLngStr("LttlLanguagesNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table4" style="width: 100%">
			<tr>
				<td class="style2" style="width: 60">
			#</td>
				<td class="style2">
			<%=getadminLanguagesLngStr("LtxtLanguage")%></td>
				<td class="style2" style="width: 100px">
			<%=getadminLanguagesLngStr("LtxtSign")%></td>
				<td class="style2" style="width: 100px">
			<%=getadminLanguagesLngStr("DtxtActive")%></td>
			</tr>
			<% For i = 0 to UBound(myLanIndex)
			rs.Filter = "LanID = " & myLanIndex(i)(4) %>
			<% If myApp.NatLng <> myLanIndex(i)(2) Then %>
			<input type="hidden" name="LanID" value="<%=myLanIndex(i)(4)%>">
			<% Else %>
			<input type="hidden" name="NatLng" value="<%=myLanIndex(i)(4)%>">
			<% End If %>
			<tr>
				<td class="style4" style="width: 60">
				<label for="LanID<%=myLanIndex(i)(4)%>">
				<font face="Verdana" size="1" color="#4783C5"><%=myLanIndex(i)(4)%></font></label></td>
				<td class="style3">
				<label for="LanID<%=myLanIndex(i)(4)%>">
				<font face="Verdana" size="1" color="#4783C5"><%=myLanIndex(i)(1)%><% If myApp.NatLng = myLanIndex(i)(2) Then %>&nbsp;<%=getadminLanguagesLngStr("LtxtNatLng")%><% End If %></font></label></td>
				<td class="style3" style="width: 100px">
				<label for="LanID<%=myLanIndex(i)(4)%>">
				<font face="Verdana" size="1" color="#4783C5"><%=UCase(myLanIndex(i)(2))%></font></label></td>
				<td align="center" class="style3" style="width: 100px">
				<input type="checkbox" name="LanID<%=myLanIndex(i)(4)%>" id="LanID<%=myLanIndex(i)(4)%>" value="<%=myLanIndex(i)(4)%>" <% If myApp.NatLng = myLanIndex(i)(2) Then %>disabled <% End If %> <% If rs.recordcount = 0 Then %> checked<% End If %> class="noborder">&nbsp;</td>
			</tr>
			<% Next %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr>
				<td width="77"><input type="submit" value="<%=getadminLanguagesLngStr("DtxtSave")%>" name="btnSave" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:77; height:23; font-weight:bold"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminLang">
	</form>
</table>
<!--#include file="bottom.asp" -->