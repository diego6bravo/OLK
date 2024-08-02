<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="myHTMLEncode.asp"-->
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<%
Session("UserName") = Request("CardCode")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCreateActivity" & Session("ID")
cmd.Parameters.Refresh
cmd("@CardCode") = saveHTMLDecode(Session("UserName"), True)
cmd("@SlpCode") = Session("vendid")
cmd.execute

Session("ActRetVal") = cmd("@LogNum")
Session("RetVal") = ""
Session("PayRetVal") = ""
Session("cart") = ""
Session("PayCart") = False
Session("ActReadOnly") = False

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
cmd.Parameters.Refresh
cmd("@sessiontype") = "A"
cmd("@transtype") = "N"
cmd("@object") = 33 
cmd("@LogNum") = Session("ActRetVal")
cmd("@CurrentSlpCode") = Session("vendid")
cmd("@Branch") = Session("branch")
cmd.execute()

conn.close

Response.Redirect "operaciones.asp?cmd=activity"

%>