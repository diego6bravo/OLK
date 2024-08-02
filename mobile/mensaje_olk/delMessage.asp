<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

sql = "update olkmsg1 set olkstatus = 'D' where olklog in (" & Request("delLog") & ") and olkuser = '" & Session("vendid") & "'"
conn.execute(sql)
conn.close
response.redirect "../operaciones.asp?cmd=buzon"
%>