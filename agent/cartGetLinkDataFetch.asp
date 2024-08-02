<!--#include file="myHTMLEncode.asp"-->
<% If Session("RetVal") = "" Then response.redirect "default.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
set rs = server.createobject("ADODB.RecordSet")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCartRepLinkVars" & Session("ID")
cmd.Parameters.Refresh
cmd("@LogNum") = Session("RetVal")
cmd("@LineNum") = Request("LineNum")
cmd("@ID") = Request("ID")
cmd("@PriceList") = Session("PriceList")
cmd("@CardCode") = Session("username")
cmd("@SlpCode") = Session("vendid")

set rs = cmd.execute()

For each f in rs.Fields
	Response.Write f.Name & "{C}" & f & "{R}"
Next
%>