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
dim varx
dim varxx
varx = "0"
           set rd = Server.CreateObject("ADODB.recordset")
           set rs = Server.CreateObject("ADODB.recordset")
If Request("Exp") <> "Y" or Request("Exp") = "Y" and myApp.SVer = "5" Then
	sql = "delete R3_ObsCommon..doc1 where lognum = " & Session("RetVal") & " and LineNum = " & Request.QueryString("Line") & _
		  " delete olksaleslines where lognum = " & Session("RetVal") & " and LineNum = " & Request("Line")
ElseIf Request("Exp") = "Y" and myApp.SVer >= "6" Then
	sql = "delete R3_ObsCommon..doc3 where lognum = " & Session("RetVal") & " and LineNum = " & Request.QueryString("Line")
End If
conn.execute(sql)
conn.close 
If userType = "C" Then
response.redirect "../default.asp?cmd=cart&update=Y"
ElseIf userType = "V" Then
response.redirect "../cart.asp?cmd=" & Request("redir")
End If
%>