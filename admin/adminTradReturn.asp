<!--#include file="langIndex.inc" -->
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->

<head>

<title></title>
</head>
<%
retVal = ""
For i = 0 to UBound(myLanIndex)
	If i > 0 Then retVal = retVal & "{/}"
	retVal = retVal & myLanIndex(i)(4) & "{=}" & Request("txt" & myLanIndex(i)(4))
Next
%>
<body>
<script type="text/javascript">
<!--
opener.setNewFldTrad('<%=Replace(myHTMLEncode(retVal), "'", "\'")%>');
window.close();

//-->
</script>
</body>

</html>
