<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")

sql = "select name from syscolumns where id = (select id from sysobjects where name = N'" & Request("TableID") & "')"
rs.open sql, conn, 3, 1
do while not rs.eof
	If rs.bookmark > 1 Then Response.Write ", "
	Response.Write rs(0)
rs.movenext
loop
%>