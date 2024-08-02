<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="authorizationClass.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = new clsAuthorization

%>

<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<!--#include file="myHTMLEncode.asp"-->

<%
response.buffer = true

If Request("edit") <> "Y" Then
	cmd.CommandText = "DBOLKCreateNewCard" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@SlpCode") = Session("vendid")
	cmd("@branchIndex") = Session("branch")
	cmd("@SessionType") = "A"
	If Session("UserAccess") = "U" Then
		If myAut.HasAuthorization(45) Then
			cmd("@CardType") = "C"
		ElseIf myAut.HasAuthorization(78) Then
			cmd("@CardType") = "S"
		ElseIf myAut.HasAuthorization(77) Then
			cmd("@CardType") = "L"
		End If
	End If
	cmd.execute()
	
	RetVal = cmd.Parameters.Item(0).Value
	Session("CrdRetVal") = RetVal
	Session("RetVal") = ""
Else
	set oCmd = Server.CreateObject("ADODB.Command")
	ocmd.ActiveConnection = connCommon
	oCmd.CommandText = "DBOLKEditClient" & Session("ID")
	oCmd.CommandType = &H0004
	oCmd.Parameters.Refresh()
	oCmd("@CardCode") = Session("username")
	oCmd("@SlpCode") = Session("vendid")
	oCmd.Execute()
	
	RetVal = oCmd.Parameters.Item(0).value
	
	Session("CrdRetVal") = RetVal
	Session("RetVal") = ""
End If

conn.close
Session("cart") = ""
Response.Redirect "operaciones.asp?cmd=newClient"

%>