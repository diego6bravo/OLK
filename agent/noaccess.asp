<% addLngPathStr = "" %>
<!--#include file="lang/noaccess.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<script language="javascript" src="general.js"></script>
</head>

<body>
<script language="javascript">
<!--
<% If Request("ErrCode") = "" Then %>
alert('<%=getnoaccessLngStr("LtxtValNoAccess")%>');
<% ElseIf Request("ErrCode") = "0" Then %>
alert('<%=getnoaccessLngStr("LtxtValDocAccess")%>');
<% End If %>
window.close();
//-->
</script>
</body>

</html>
