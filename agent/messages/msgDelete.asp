<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKDelMsg" & Session("ID")
cmd.Parameters.Refresh()
cmd("@UserType") = userType

Select Case userType
	Case "V" 
		user = Session("vendid")
		MainDoc = "agent.asp"
	Case "C" 
		user = Session("UserName")
		MainDoc = "messages.asp"
End Select

cmd("@UserName") = user

If Request.QueryString("olklog") <> "" Then
	cmd("@LogNum") = Request.QueryString("olklog")
ElseIf Request("DelLog") <> "" Then
	cmd("@LogNum") = Request("DelLog")
End If
cmd.execute()

conn.close
response.redirect "../" & MainDoc & "?p=" & Request("p") & "&msgOrdr1=" & Request("msgOrdr1") & "&msgOrdr2=" & Request("msgOrdr2") & "&onlyMsg=" & Request("onlyMsg")
 %>