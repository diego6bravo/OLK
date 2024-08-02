<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% 
If (Session("UserName") = "-Anon-" or not optBasket) Then Response.Redirect "default.asp"
Case "V" %><!--#include file="agentTop.asp"-->
<% 
If Not comDocsMenu Then Response.Redirect "unauthorized.asp"
End Select %>
<% addLngPathStr = "" %>
<!--#include file="lang/cartConfirm.asp" -->
<% If (Session("UserName") = "-Anon-" or not optBasket) and userType = "C" Then Response.Redirect "default.asp"
If setCustTtl Then printCmd = printCmd & "setTtlBg(true);"
printCmd = printCmd & "printStory('printPage', '" & SelDes & "');"
If setCustTtl Then printCmd= printCmd & "setTtlBg(false);"

pdfCmd = "saveInPwd('" & Request("Status") & "');"
 %>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td align="right">
		<a href="#" onclick="javascript:<%=printCmd%>"><img alt="<%=getcartConfirmLngStr("DtxtPrint")%>" border="0" src="<% If alterPrintOLK = "" Then %>images/print_OLK.gif<% Else %><%=alterPrintOLK%><% End If %>"></a>&nbsp;
		<a href="#" onclick="javascript:<%=pdfCmd%>"><img alt="<%=getcartConfirmLngStr("DtxtExpPDF")%>" border="0" src="<% If alterPdfOLK = "" Then %>images/pdf_OLK.gif<% Else %><%=alterPdfOLK%><% End If %>"></a>
		</td>
	</tr>
</table>
<div id="printPage">
<!--#include file="cartSubmitConfirm.asp" -->
</div>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
