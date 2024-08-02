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
ID = CLng(Request("ID"))

set oCmd = Server.CreateObject("ADODB.Command")
ocmd.ActiveConnection = connCommon
oCmd.CommandText = "DBOLKEditSO" & Session("ID")
oCmd.CommandType = &H0004
oCmd.Parameters.Refresh()
oCmd("@ID") = ID
oCmd("@SlpCode") = Session("vendid")
oCmd.Execute()

RetVal = oCmd.Parameters.Item(0).value

Session("SORetVal") = RetVal
Session("UserName") = oCmd("@CardCode").value

Response.Redirect "../so.asp"
%>