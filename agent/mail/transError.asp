<!--#include file="../lang.asp"-->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<!--#include file="lang/transError.asp" -->
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
		<strong><%=gettransErrorLngStr("DtxtLogNum")%> &lt;:-LogNum-&gt;</strong><br>
		<%=gettransErrorLngStr("DtxtBP")%>: <:-CardCode-><br>
		<%=gettransErrorLngStr("DtxtName")%>: <:-CardName-><br>
		<%=gettransErrorLngStr("DtxtType")%>: <:-DocType-><br>
		<br>
		<strong><%=gettransErrorLngStr("DtxtError")%></strong><br>
		<%=gettransErrorLngStr("DtxtDate")%>: <:-EndDate-><br>
		<%=gettransErrorLngStr("DtxtHour")%>: <:-EndTime-><br>
		<%=gettransErrorLngStr("DtxtCode")%>: <:-ErrCode-><br>
		<%=gettransErrorLngStr("DtxtDescription")%>: <:-ErrMessage->
		</font></p></td>
	</tr>
</table>

</body>

</html>
