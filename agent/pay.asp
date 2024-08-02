<%@ Language=VBScript %>
<!-- #include file="chkLogin.asp" -->
<!--#include file="lang.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="lang/pay.asp" -->
<% 
If Request("document") = "" Then
	If Request.Cookies("catMethod") = "" Then Response.Cookies("catMethod") = "T"
Else
	Response.Cookies("catMethod") = Request("document")
End If

If Session("UserType") = "V" Then 
	response.redirect "agent.asp"
ElseIf Session("OLKAdmin") Then
	response.redirect "admin/admin.asp?cmd=home"
End If
%>
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<%
Session("cart") = "cart"
response.buffer = true
Dim cmd
Dim RetVal
Dim db
db = Session("olkdb")
varx = 1

Dim myAut
set myAut = New clsAuthorization

           set rs = Server.CreateObject("ADODB.recordset")
           set rg = Server.CreateObject("ADODB.recordset")
           set rx = server.createobject("ADODB.RecordSet")
           set ra = Server.Createobject("ADODB.RecordSet")

 %>
<!--#include file="loadAlterNames.asp" -->
<% 
sql = 	"select top 1 ClientMenu, " & _
		"Version, " & _
		"IsNull(dbo.OLKGetTrans(" & Session("LanID") & ", 'OADM', 'PrintHeadr', '1', PrintHeadr),IsNull(CompnyName, '')) CmpName, " & _
		"SelDes, DirectRate, " & _
		"IsNull((select AlterEnBlkRClkMsg from OLKCommonAlterNames where LanID = " & Session("LanID") & "), IsNull(EnBlkRClkMsg, '')) EnBlkRClkMsg, " & _
		"(select Count('A') from OLKOMSG X0 " & _
		"inner join OLKMSG1 X1 on X1.OlkLog = X0.OlkLog " & _
		"where X1.OlkUser = N'" & saveHTMLDecode(Session("UserName"), False) & "' and OlkUserType = 'C' and OlkStatus = 'N') newMsgCount, " & _
		"SecIndexByX, SecIndexByY, Case When Exists(select 'A' from OLKCatNavIndex) Then 'Y' Else 'N' End VerfyNavIndex, ShowCSearchTree, ShowCAdSearch "
		
If Session("UserName") <> "-Anon-" Then sql = sql & " ,dbo.olkGetTrans(" & Session("LanID") & ", 'OCRD', 'CardName', ocrd.CardCode, ocrd.CardName) CardName  "

sql = sql & " from OLKCommon cross join oadm "

If Session("UserName") <> "-Anon-" Then sql = sql & " cross join ocrd where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "

sql = sql & " order by CurrPeriod desc"

set rg = conn.execute(sql)

If Session("UserName") <> "-Anon-" Then
	CardName = rg("cardname")
Else
	CardName = txtClient & " " & LtxtAnon
End If
ShowCSearchTree = rg("ShowCSearchTree") = "Y"
ShowCAdSearch = rg("ShowCAdSearch") = "Y"
QryGroup = myApp.CarArt
Version = rg("Version")
CmpName = rg("CmpName")
SelDes = rg("SelDes")
DirectRate = rg("DirectRate")
optNav = rg("VerfyNavIndex") = "Y"
EnBlkRClkMsg = rg("EnBlkRClkMsg")
newMsgCount = rg("newMsgCount")
SecIndexByX = rg("SecIndexByX")
SecIndexByY = rg("SecIndexByY")

MainDoc = "default.asp"

If SelDes = "6" Then
	If myApp.TopLogo = "" or IsNull(myApp.TopLogo) Then 
		logoImg = "design/6/swf/logo.jpg"
	Else
		logoImg = "imagenes/" & Session("olkdb") & "/" & myApp.TopLogo
	End If
	xmlPathLogo = GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	xmlPathLogo = Replace(xmlPathLogo, "default.asp", "getXmlLogo.asp?logoImg=" & logoImg)
End If

If Request.Form("LogNum") <> "" Then Session("RetVal") = Request.Form("LogNum")
%>
<!--#include file="clearItem.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Olk - <%=Replace(getpayLngStr("LtxtClientInterface"), "{0}", Server.HTMLEncode(txtClient))%></title>
<script language="javascript" src="general.js"></script>
<script language="javascript">
<% If Request("err") = "inv" Then %>
alert("|L:txtErrItmInv|")
<% ElseIf Request("err") = "tax" Then %>
	alert('|L:txtErrNoTaxGrp|'.replace('{0}', '<%=LCase(txtTax)%>'));
<% ElseIf Request("errMInv") <> "" Then %>
	var alertMsg = '|L:txtErrMultItmInv|: \n<%=Replace(Request("errMInv"), "'", "\'")%>'
	alert(alertMsg);
<% End If %>
var EnRClickBlk = <% If myApp.EnBlockRClk Then %>true<% Else %>false<% End If %>;
var RClickBlkMsg = '<%=Replace(Replace(EnBlkRClkMsg, "'", "\'"), VbCrLf, "\n")%>';
</script>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
<% If Session("rtl") <> "" Then %><link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/hebrew.css"><% End If %>
<link rel="stylesheet" type="text/css" media="all" href="design/<%=SelDes%>/style/style_cal.css" title="winter" />
</head>
<!--#include file="langIndex.inc" -->
<script language="javascript">
var curDir = '<% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>';
var rtl = '<%=Session("rtl")%>';
var txtValFldMaxChar = 	'|D:txtValFldMaxChar|';
</script>
<script language="javascript" src="rightclick.js"></script>
<script language="javascript" src="clientes.js"></script>
<script language="vbscript" src="clientes.vbs"></script>

<body onfocus="javascript:chkWin()" onload="javascript:chkWin();" topmargin="0" leftmargin="0">
<!--#include file="licid.inc"-->
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="740" bgcolor="#FFFFFF">
	<tr>
		<td>
		<!--#include file="cartSiteStep3.asp" -->
		</td>
	</tr>
</table>
</div>
<% If setCustTtl Then %>
<script language="javascript" src="setTltBg.js.asp?custTtlBgL=<%=custTtlBgL%>&custTtlBgM=<%=custTtlBgM%>"></script>
<% End If %>
<script language="javascript">
<% If setCustTtl Then %>setTtlBg(false);<% End If %>
</script>
<!--#include file="linkForm.asp"-->
</body>
<script>
setAllSize();
</script>
</html>