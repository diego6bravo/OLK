<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% End Select
%>
<% addLngPathStr = "" %>
<!--#include file="lang/report.asp" -->
<% printCmd = printCmd & "printShowLegend(false);printStory('printPage', '" & SelDes & "');printShowLegend(true);"		
pdfCmd = "saveRepPdf('N');" %>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td align="right">
		<a href="#" onclick="javascript:<%=printCmd%>"><img alt="<%=getreportLngStr("DtxtPrint")%>" border="0" src="<% If alterPrintOLK = "" Then %>images/print_OLK.gif<% Else %><%=alterPrintOLK%><% End If %>"></a>&nbsp;
		<a href="#" onclick="javascript:<%=pdfCmd%>"><img alt="<%=getreportLngStr("DtxtExpPDF")%>" border="0" src="<% If alterPdfOLK = "" Then %>images/pdf_OLK.gif<% Else %><%=alterPdfOLK%><% End If %>"></a>
		&nbsp;<a href="#" onclick="javascript:saveRepPdf('Y');"><img alt="<%=getreportLngStr("DtxtExpToExcell")%>" border="0" src="images/excell.gif"></a>
		</td>
	</tr>
</table>
<div id="printPage">
<!--#include file="portal/viewReport.asp"-->
</div>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>