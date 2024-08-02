<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/cuentas.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getcuentasLngStr("LtxtAccount")%></title>
<meta name="VI60_defaultClientScript" content="JavaScript">
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
<!--
function setCuenta(Acctcode, AcctName)
{
	opener.setCuenta(Acctcode, AcctName, '<%=Request("update")%>');
	window.close();
}
//-->
</script>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

           set rs = Server.CreateObject("ADODB.RecordSet")
           If Left(Request("update"),2) = "ca" or Left(Request("update"),5) = "OIRca" then
           	GetQuery rs, 8, null, null
           ElseIf Left(Request("update"),2) = "ch" or Left(Request("update"),2) = "cr" or Left(Request("update"),5) = "OIRch" or Left(Request("update"),5) = "OIRcr" then
           	GetQuery rs, 9, null, null
           End if 
           %>
</head>

<body topmargin="0" leftmargin="0" onbeforeunload="javascript:opener.clearWin();">

<table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%" class="popupTtl"><%=getcuentasLngStr("LtxtAccount")%> - <%
		Select Case Left(Request("update"),2)
		Case "ca"
		Response.write "" & getcuentasLngStr("LtxtCash") & ""
		Case "ch"
		Response.write "" & getcuentasLngStr("LtxtCheques") & ""
		Case "cr"
		Response.write "" & getcuentasLngStr("LtxtCreditCard") & ""
		End Select
		%></td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
      <tr>
        <td class="popupOptDesc"><%=getcuentasLngStr("LcolAccount")%></td>
        <td class="popupOptDesc"><%=getcuentasLngStr("DtxtDescription")%></td>
      </tr>
  <% do while not rs.eof %>
      <tr style="cursor: hand" onclick="javascript:setCuenta('<%=Server.HTMLEncode(rs("AcctCode"))%>','<%=Replace(Server.HTMLEncode(rs("AcctName")), "'", "\'")%>')">
        <td class="popupOptValue"><font size="1" face="Verdana"><%=Server.HTMLEncode(rs("AcctCode"))%></font></td>
        <td class="popupOptValue"><font size="1" face="Verdana"><%=Server.HTMLEncode(rs("AcctName"))%></font></td>
      </tr>
  	<% rs.movenext
	loop %>
    </table>
    </td>
  </tr>
</table>

</body>
<% set rs = nothing
conn.close%>
</html>