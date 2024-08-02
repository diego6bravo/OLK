<!--#include file="../lang.asp"-->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<!--#include file="lang/passwordReset.asp" -->
<style type="text/css">
.style1 {
	text-align: center;
	font-size: medium;
}
</style>
</head>

<body style="text-align: center" topmargin="0">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr>
		<td height="80" class="style1"><font face="Verdana"><strong><:-CmpName-></strong></font></td>
	</tr>
	<tr>
		<td>
		<p><font face="Verdana" size="2" color="#1E3E57">
		<%=getpasswordResetLngStr("LtxtNewPassword")%>:<br><:-Password-></font></p></td>
	</tr>
	<tr>
		<td>
		<p align="right"><font face="Verdana" size="1"><br>
		<:-CompnyAddr-><br>
		<%=getpasswordResetLngStr("DtxtPhone")%>: <:-Phone1->/<:-Phone2->
		<%=getpasswordResetLngStr("DtxtFax")%>: <:-Fax-><br>
		<a href="mailto:<:-E_Mail->"><:-E_Mail-></a>
		</font>
</td>
	</tr>
</table>

</body>

</html>
