<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

ID = CInt(Replace(Request("ExecAt"), "OP", ""))

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetOpVars" & Session("ID")
cmd.Parameters.Refresh
cmd("@ID") = ID

set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1

do while not rs.eof
	If rs.bookmark > 1 Then Response.Write "{S}"
	Response.Write rs(0) & "{C}" & rs(1)
rs.movenext
loop
%>