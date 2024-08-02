<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/go.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<% 
Session("RetVal") = ""
Session("PayRetVal") = Request("doc") 
Session("UserName") = saveHTMLDecode(Request("cl"), True)
Session("PayCart") = False
If Request("status") = "H" Then
		sql = 	"declare @LogNum int set @LogNum = " & Request("doc") & " " & _
				"update r3_obscommon..tlog set status = 'R' where lognum = @LogNum " & _
				"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'R2' and ObjectEntry = @LogNum and Status in ('O', 'E') "
           conn.execute(sql)
           conn.close
           end if
response.redirect "../agentPayment.asp" %>
