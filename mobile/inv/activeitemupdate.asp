<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
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
cmd.CommandText = "DBOLKInvCount"
cmd.Parameters.Refresh()
cmd("@dbID") = Session("ID")
cmd("@WhsCode") = Session("bodega")
cmd("@ItemCode") = Request.Form("Item")
cmd("@Counted") = Request.Form("t1")
cmd.execute()
response.redirect "../operaciones.asp?cmd=searchitem&btnSearch=Y"
%>