<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKUpdateMsgStatus" & Session("ID")
cmd.Parameters.Refresh()
cmd("@Status") = Request("status")
cmd("@LogNum") = Request("olklog")
cmd.execute()
conn.close

Select Case userType
	Case "V" 
		response.redirect "../agent.asp?onlyMsg=" & Request("onlyMsg")
	Case "C" 
		response.redirect "../messages.asp"
End Select

%>