<!--#include file="langIndex.inc" -->
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->

<head>

<title></title>
</head>
<body>
<script type="text/javascript">
<!--
opener.setNewFldTrad('<%=Request("PageID")%>{S}<%=Request("FieldID")%>{S}<%=Replace(myHTMLEncode(Request("txtDefinition")), "'", "\'")%>');
window.close();

//-->
</script>
</body>

</html>
