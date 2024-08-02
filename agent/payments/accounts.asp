<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/accounts.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<%
set rs = Server.CreateObject("ADODB.RecordSet")
GetQuery rs, 8, null, null %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getaccountsLngStr("LttlBankTrans")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>
<script language="javascript">
function setCuenta(Acctcode, AcctName)
{
	opener.setCuenta(Acctcode, AcctName);
	window.close();
}
</script>
<script language="javascript" src="../general.js"></script>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" style="background-color: #EDF5FE">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td><%=getaccountsLngStr("LttlSelAcct")%></td>
	</tr>
	<% while not rs.eof %>
	<tr class="GeneralTbl" onclick="javascript:setCuenta('<%=rs("AcctCode")%>','<%=myHTMLEncode(Replace(rs("AcctName"), "'", "\'"))%>')" style="cursor: hand" onmouseover="javascript:this.className='GeneralTblHigh'" onmouseout="javascript:this.className='GeneralTbl'">
		<td>
		<%=rs("AcctName")%></td>
	</tr>
	<% rs.movenext
	wend %>
</table>

</body>
<% conn.close %>
</html>