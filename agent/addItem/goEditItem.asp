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

set oCmd = Server.CreateObject("ADODB.Command")
ocmd.ActiveConnection = connCommon
oCmd.CommandText = "DBOLKEditItem" & Session("ID")
oCmd.CommandType = &H0004
oCmd.Parameters.Refresh()
oCmd("@ItemCode") = Request("ItemCode")
oCmd("@SlpCode") = Session("vendid")
If Request("Duplicate") = "Y" Then oCmd("@Duplicate") = "Y"
oCmd.Execute()

RetVal = oCmd.Parameters.Item(0).value

Session("ItmRetVal") = RetVal

Response.Redirect "../agentItem.asp"
%>