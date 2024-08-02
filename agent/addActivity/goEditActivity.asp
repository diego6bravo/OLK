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
<%
ClgCode = CLng(Request("ClgCode"))

set oCmd = Server.CreateObject("ADODB.Command")
ocmd.ActiveConnection = connCommon
oCmd.CommandText = "DBOLKEditActivity" & Session("ID")
oCmd.CommandType = &H0004
oCmd.Parameters.Refresh()
oCmd("@ClgCode") = ClgCode
oCmd("@SlpCode") = Session("vendid")
oCmd.Execute()

RetVal = oCmd.Parameters.Item(0).value

Session("ActRetVal") = RetVal
Session("UserName") = oCmd("@CardCode").value

Response.Redirect "../agentActivity.asp"
%>