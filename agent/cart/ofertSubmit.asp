<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->

<!--#include file="../lcidReturn.inc" -->
<!--#include file="../myHTMLEncode.asp" -->
<html>
<!--#include file="../loadAlterNames.asp" -->
<%
Select Case Request("cmd")
	Case "newOfert"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKOfertsNew" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ItemCode") = Request("Item")
		cmd("@CardCode") = Session("UserName")
		cmd("@BasePrice") = CDbl(getNumericOut(Request("BasePrice")))
		cmd("@OfertDiscount") = CDbl(getNumericOut(Request("ofertDiscount")))
		cmd("@OfertPrice") = CDbl(getNumericOut(Request("ofertPrice")))
		cmd("@Quantity") = CDbl(getNumericOut(Request("ofertQuantity")))
		If Request("ofertNote") <> "" Then cmd("@OferNote") = Request("ofertNote")
		If Request("ofertLimit") <> "" Then cmd("@OfertLimit") = Request("ofertLimit")
		cmd.execute()
	Case "contraOfert"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKOfertsUpdate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@OfertIndex") = Request("ofertIndex")
		cmd("@OfertDiscount") = CDbl(getNumericOut(Request("ofertDiscount")))
		cmd("@OfertPrice") = CDbl(getNumericOut(Request("ofertPrice")))
		cmd("@Quantity") = CDbl(getNumericOut(Request("ofertQuantity")))
		If Request("ofertNote") <> "" Then cmd("@OferNote") = Request("ofertNote")
		If Request("ofertLimit") <> "" Then cmd("@OfertLimit") = Request("ofertLimit")
		cmd("@cmd") = "C"
		cmd.execute()
	Case "updateOfert"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKOfertsUpdate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@OfertIndex") = Request("ofertIndex")
		cmd("@OfertDiscount") = CDbl(getNumericOut(Request("ofertDiscount")))
		cmd("@OfertPrice") = CDbl(getNumericOut(Request("ofertPrice")))
		cmd("@Quantity") = CDbl(getNumericOut(Request("ofertQuantity")))
		If Request("ofertNote") <> "" Then cmd("@OferNote") = Request("ofertNote")
		If Request("ofertLimit") <> "" Then cmd("@OfertLimit") = Request("ofertLimit")
		cmd("@cmd") = "U"
		cmd.execute()
	Case "acceptOfert"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKOfertCmd" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@OfertIndex") = Request("ofertIndex")
		cmd("@cmd") = "A"
		cmd("@SlpCode") = Session("vendid")
		cmd.execute()
	Case "rejectOfert"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKOfertCmd" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@OfertIndex") = Request("ofertIndex")
		cmd("@cmd") = "R"
		cmd("@SlpCode") = Session("vendid")
		cmd.execute()
	Case "AgentContraOfert"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKOfertsUpdate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@OfertIndex") = Request("ofertIndex")
		cmd("@OfertDiscount") = CDbl(getNumericOut(Request("responseDiscount")))
		cmd("@OfertPrice") = CDbl(getNumericOut(Request("responsePrice")))
		cmd("@Quantity") = CDbl(getNumericOut(Request("responseQuantity")))
		If Request("responseNote") <> "" Then cmd("@OferNote") = Request("responseNote")
		If Request("responseLimit") <> "" Then cmd("@OfertLimit") = Request("responseLimit")
		cmd("@cmd") = "A"
		cmd.execute()
End Select
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=txtOfert%></title>
</head>

<% If Request("cmd") = "contraOfert" or Request("cmd") = "updateOfert" or Request("cmd") = "newOfert" Then %>
<script language="javascript">
window.location.href = '../oferts.asp';
</script>
<% ElseIf Request("cmd") = "acceptOfert" or Request("cmd") = "rejectOfert" or Request("cmd") = "AgentContraOfert" then %>
<script language="javascript">
window.location.href = '../ofertsMan.asp?cmd=<%=Request("redir")%>&page=<%=Request("page")%>';
</script>
<% end if %>
<body onLoad="//setTimeout(window.close, 0000)">
</body>
</html>