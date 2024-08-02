<html>
<body onbeforeunload="opener.clearWin();">
<SCRIPT LANGUAGE="JavaScript">
opener.changepic('<%=Request("filename")%>');
window.close()
</SCRIPT>
</body>
</html>