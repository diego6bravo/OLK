<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/cardsSist.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<%
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select CrTypeCode, IsNull(CrTypeName, '') CrTypeName, MinCredit, MinToPay, MaxValid, InstalMent from OCRP "
set rs = conn.execute(sql)%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getcardsSistLngStr("LttlSelCardSis")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>
<script language="javascript">
function setSistPag(CrTypeCode, CrTypeName, MinCredit, MinToPay, MaxValid, InstalMent)
{
	opener.setSistPag(CrTypeCode, CrTypeName, MinCredit, MinToPay, MaxValid, InstalMent);
	window.close();
}
</script>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onblur="opener.clearWin();">

<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><%=getcardsSistLngStr("LtxtSelPymntSys")%></td>
	</tr>
	<% while not rs.eof %>
	<tr class="GeneralTbl" onclick="javascript:setSistPag('<%=myHTMLEncode(rs("CrTypeCode"))%>','<%=myHTMLEncode(rs("CrTypeName"))%>','<%=rs("MinCredit")%>','<%=rs("MaxValid")%>','<%=rs("InstalMent")%>','<%=rs("InstalMent")%>');" style="cursor: hand" onmouseover="javascript:this.className='GeneralTblHigh'" onmouseout="javascript:this.className='GeneralTbl'">
		<td>
		<%=myHTMLEncode(rs("CrTypeName"))%></td>
	</tr>
	<% rs.movenext
	wend %>
</table>

</body>
<% conn.close %>
</html>