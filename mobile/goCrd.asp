<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Session("CrdRetVal") = Request("LogNum")
Session("RetVal") = ""
Session("cart") = ""

If Request("status") = "H" Then
	sql = "update r3_obscommon..tlog set status = 'R' where lognum = " & Session("CrdRetVal") & " " & _
        	"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'C1' and ObjectEntry = " & Session("CrdRetVal") & " and Status in ('O', 'E') "
	conn.execute(sql)
end if

Response.Redirect "operaciones.asp?cmd=newClient"

%>