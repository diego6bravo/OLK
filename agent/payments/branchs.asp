<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/branchs.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
           set rs = Server.CreateObject("ADODB.RecordSet")
           sql = "SELECT IsNull(Branch, '') Branch, IsNull(Account, '') Account FROM OCRB where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and BankCode = '" & Request("BankCode") & "' and Branch like '" & Replace(Request("Value"),"*","%") & "'"
           set rs = conn.execute(sql)%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getbranchsLngStr("LttlSelBankBranch")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">

</head>


<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr>
		<td colspan="3" class="GeneralTlt"><%=getbranchsLngStr("LttlSelBranch")%></td>
	</tr>
	<tr class="GeneralTblBold2">
		<td align="center">
		<%=getbranchsLngStr("DtxtBranch")%></td>
		<td align="center">
		<%=getbranchsLngStr("DtxtAccount")%></td>
	</tr>
	<% If not rs.eof then 
	do while not rs.eof %>
	<tr class="GeneralTbl" onclick="javascript:opener.setBranch('<%=rs("Branch")%>','<%=rs("Account")%>');window.close();" style="cursor: hand" onmouseover="javascript:this.className='GeneralTblHigh'" onmouseout="javascript:this.className='GeneralTbl'">
		<td><%=rs("Branch")%></td>
		<td>
		<%=rs("Account")%>&nbsp;</td>
	</tr>
	<% rs.movenext
	loop 
	else %>
	<tr>
		<td colspan="2" class="GeneralTbl">
		<%=getbranchsLngStr("LtxtErrNoBranch")%></td>
	</tr>
	<% end if %>
	</table>

</body>
<% conn.close %>
</html>