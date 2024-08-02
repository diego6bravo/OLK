<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRSCSS" & Session("ID")
cmd.Parameters.Refresh()
cmd("@rsIndex") = Request("rsIndex")
set rs = Server.CreateObject("ADODB.RecordSet")
set rs = cmd.execute()
do while not rs.eof
	Response.Write rs(0) & VbCrLf
rs.movenext
loop
%>

