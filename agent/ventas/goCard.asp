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
           set rs = Server.CreateObject("ADODB.recordset")
           
           Session("CrdRetVal") = Request("LogNum")
           Session("ItmRetVal") = ""
           Session("cart") = ""
           Session("PayCart") = False
           Session("RetVal") = ""
           
           sql = "select Status from R3_ObsCommon..TLOG where LogNum = " & Request("LogNum")
           set rs = conn.execute(sql)
           If rs("status") = "H" Then
	           sql = "declare @LogNum int set @LogNum = " & Request("LogNum") & " " & _
	           		"update r3_obscommon..tlog set status = 'R' where lognum = @LogNum " & _
	           		"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'C1' and ObjectEntry = @LogNum and Status in ('O', 'E') "
	           conn.execute(sql)
           end if
           Response.Redirect "../agentClient.asp"
           conn.close
%>
