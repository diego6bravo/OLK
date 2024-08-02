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
oCmd.CommandText = "DBOLKEditActivity" & Session("ID")
oCmd.CommandType = &H0004
oCmd.Parameters.Refresh()
oCmd("@ClgCode") = Request("ClgCode")
oCmd("@SlpCode") = Session("vendid")
oCmd.Execute()

RetVal = oCmd.Parameters.Item(0).value
Session("UserName") = Request("CardCode")

If RetVal <> -1 Then
	Session("ActRetVal") = RetVal
	Session("ActReadOnly") = False
Else
	Session("ActRetVal") = Request("ClgCode")
	Session("ActReadOnly") = True
End If
	
Response.Redirect "operaciones.asp?cmd=activity"
%>