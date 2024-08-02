<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

         
set rs = Server.CreateObject("ADODB.recordset")
sql = "select OLKCommon.dbo.DBOLKGetCardPList" & Session("ID") & "(N'" & Request("cl") & "', 'V') listnum, PriceList from R3_ObsCommon..TLOGControl where LogNum = " &  Request("doc")
set rs = conn.execute(sql)

If IsNull(rs("PriceList")) Then
	sql = "update R3_ObsCommon..TLOGControl set PriceList = " & rs("ListNum") & " where LogNum = " &  Request("doc") & " and PriceList is null"
	conn.execute(sql)
End If

Session("UserName") = Request("cl")
Session("RetVal") = Request("doc")
If IsNull(rs("PriceList")) Then
	Session("PList") = RS("listnum")
Else
	Session("PList") = rs("PriceList")
End If

If Request("status") = "H" Then
	sql = "update r3_obscommon..tlog set status = 'R' where lognum = " & Session("RetVal") & " " & _
        	"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'D3' and ObjectEntry = " & Session("RetVal") & " and Status in ('O', 'E') "

	conn.execute(sql)
end if

Response.Redirect "operaciones.asp?cmd=cart"
%>