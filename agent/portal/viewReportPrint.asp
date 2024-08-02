<%@ Language=VBScript %>
<html>
<!-- #include file="../chkLogin.asp" -->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../authorizationClass.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


Dim myAut
set myAut = New clsAuthorization

%>

<!--#include file="lang/viewReportPrint.asp" -->
<%
'Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
sql = "select SelDes, DirectRate from OLKCommon cross join oadm"
set rs = conn.execute(sql)
If Session("UserType") = "C" Then SelDes = rs("SelDes") Else SelDes = 0
imgAddPath = "../" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="../design/<%=SelDes%>/style/stylenuevo.css">
<!--#include file="../licid.inc"-->
<% If Request("Excell") <> "Y" and Request("itemSmallRep") <> "Y" Then %>
<link rel="stylesheet" href="../design/<%=SelDes%>/style/stylenuevo.css"><% End If %>

</head>
<body topmargin="0">
<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<script type="text/javascript">
<!--
function saveDoc(cmd)
{
	switch (cmd)
	{
		case 'Print':
			document.all.item('tblSave').style.display = 'none';
			window.print();
			document.all.item('tblSave').style.display = '';
			break;
		case 'PDF':
			document.frmExcell.action = 'viewRepPDF.asp';
			document.frmExcell.Excell.value = 'N';
			document.frmExcell.submit();
			break;
		case 'Excell':
			document.frmExcell.action = 'viewReportPDF.asp';
			document.frmExcell.Excell.value = 'Y';
			document.frmExcell.submit();
			break;
	}
}
//-->
</script>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tblSave">
	<tr>
		<td align="right">
		<a href="#" onclick="javascript:saveDoc('Print');"><img alt="<%=getviewReportPrintLngStr("DtxtPrint")%>" border="0" src="../images/print_OLK.gif"></a>&nbsp;
		<% If Session("UserType") = "C" or Session("UserType") = "V" and myAut.HasAuthorization(65) Then %><a href="#" onclick="javascript:saveDoc('PDF');"><img alt="<%=getviewReportPrintLngStr("DtxtExpPDF")%>" border="0" src="../images/pdf_OLK.gif"></a>&nbsp;<% End If %>
		<% If Session("UserType") = "C" or Session("UserType") = "V" and myAut.HasAuthorization(64) Then %><a href="#" onclick="javascript:saveDoc('Excell');"><img alt="<%=getviewReportPrintLngStr("LtxtExpToExcell")%>" border="0" src="../images/excell.gif"></a><% End If %>
		</td>
	</tr>
</table>
<% End If %>
<!--#include file="../lcidReturn.inc"-->
<% viewRepPDF = True %>
<!--#include file="../repVars.inc" -->
<!--#include file="viewReport.asp"-->
<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<form name="frmExcell" method="post" action="" target="_blank">
<% For each itm in Request.Form
If itm <> "itemSmallRep" Then %><input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>"><% 
End If 
Next %>
<input type="hidden" name="Excell" value="">
</form>
<% End If %>

</body>
<% set rs = nothing
conn.close %></html>