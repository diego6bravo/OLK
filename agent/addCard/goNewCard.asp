<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<!--#include file="../authorizationClass.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = new clsAuthorization

%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->

<%
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
Session("PayRetVal") = ""
Session("cart") = ""
Session("PayCart") = False
conn.close
Response.Redirect "../agentClient.asp"

%>