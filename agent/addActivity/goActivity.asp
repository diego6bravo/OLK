<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../chkLogin.asp" -->
<%
LogNum = CLng(Request("LogNum"))
set rs = Server.CreateObject("ADODB.recordset")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKConfLoadObj"
cmd.Parameters.Refresh()
cmd("@dbID") = Session("ID")
cmd("@ObjCode") = 33
cmd("@LogNum") = LogNum
set rs = cmd.execute()
If Not rs.Eof Then
	Session("ActRetVal") = CLng(Request("LogNum"))
	Session("UserName") = rs("CardCode")
	Session("RetVal") = ""
	Session("PayRetVal") = ""
	Session("cart") = ""
	Session("PayCart") = False
	
	Response.Redirect "../agentActivity.asp"
Else
%>Wrong Parameters<% End If %>
