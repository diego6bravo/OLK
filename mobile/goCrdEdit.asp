<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set oCmd = Server.CreateObject("ADODB.Command")
ocmd.ActiveConnection = connCommon
oCmd.CommandText = "DBOLKEditClient" & Session("ID")
oCmd.CommandType = &H0004
oCmd.Parameters.Refresh()
oCmd("@CardCode") = Request("CardCode")
oCmd("@SlpCode") = Session("vendid")
oCmd.Execute()

RetVal = oCmd.Parameters.Item(0).value

Session("CrdRetVal") = RetVal
Session("UserName") = Request("CardCode")

Response.Redirect "operaciones.asp?cmd=newClient"
%>