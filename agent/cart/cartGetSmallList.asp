<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<title></title>
<script language="javascript" src="../general.js"></script>
</head>
<%
set rs = Server.CreateObject("ADODB.RecordSet")


Select Case Request("Group")
	Case "G"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetCartAvlExpns" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@LogNum") = Session("RetVal")

		Session("CartGroup") = Request("Group")
	Case "V"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetCartTopXItems" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@Top") = myApp.Top10Items
		cmd("@CarArt") = myApp.CarArt

		Session("CartGroup") = Request("Group")
	Case "CardGrp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetCrdGroups" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@CardType") = Request("GrpCardType")
End Select
rs.open cmd, , 3, 1
%>
<body>
<script language="javascript">
var SmallList 
SmallList = parent.getAddItemCode();
var i = 0;
<% do while not rs.eof %>
SmallList.options[i++] = new Option(clearHTMLChar('<%=Replace(Server.HTMLEncode(rs("DispItem")),"'","\'")%>'),clearHTMLChar('<%=Server.HTMLEncode(rs(0))%>'));
<% rs.movenext
loop %>
</script>
</body>

</html>
