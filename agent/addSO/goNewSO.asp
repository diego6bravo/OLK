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
<!--#include file="../myHTMLEncode.asp"-->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCreateSO" & Session("ID")
cmd.Parameters.Refresh
cmd("@CardCode") = saveHTMLDecode(Session("UserName"), True)
cmd("@SlpCode") = Session("vendid")
cmd.execute

Session("SORetVal") = cmd("@LogNum")
Session("RetVal") = ""
Session("PayRetVal") = ""
Session("cart") = ""
Session("PayCart") = False

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
cmd.Parameters.Refresh
cmd("@sessiontype") = "A"
cmd("@transtype") = "N"
cmd("@object") = 97 
cmd("@LogNum") = Session("SORetVal")
cmd("@CurrentSlpCode") = Session("vendid")
cmd("@Branch") = Session("branch")
cmd.execute()

conn.close

Response.Redirect "../so.asp"

%>