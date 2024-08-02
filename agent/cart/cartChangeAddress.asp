<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLENcode.asp"-->
<% 
set rs = Server.CreateObject("ADODB.RecordSet")

set rs = Server.CreateObject("ADODB.recordset")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetFormatedAddress" & Session("ID")
cmd.Parameters.Refresh()
cmd("@CardCode") = Session("UserName")
cmd("@Type") = Request("AdresType")
cmd("@Address") = Request("Address")
cmd("@LanID") = Session("LanID")
cmd("@UserType") = userType
cmd("@OP") = "O"
set rs = cmd.execute()
Response.Write rs(0) & "{S}" & rs(1)

 %>
