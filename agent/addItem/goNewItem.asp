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
cmd.CommandText = "DBOLKCreateNewItem" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SlpCode") = Session("vendid")
cmd("@branchIndex") = Session("branch")
cmd.execute
RetVal = cmd.Parameters.Item(0).Value
Session("ItmRetVal") = RetVal
Session("RetVal") = ""
Session("PayRetVal") = ""

conn.close
Session("cart") = ""
Session("PayCart") = False
Response.Redirect "../agentItem.asp"

%>