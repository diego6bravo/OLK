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

set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetExtPollSubmitList" & Session("ID")
cmd.Parameters.Refresh()
cmd("@AdPollID") = Request("AdPollID")
set rs = cmd.execute()

do while not rs.eof
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKSaveExtSubmit" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@AdPollID") = Request("AdPollID")
	cmd("@CardCode") = Request("CardCode")
	cmd("@LineID") = rs("LineID")
	cmd("@Answer") = Request("qc" & rs("LineID"))
	If Request("N" & rs("LineID")) <> "" Then cmd("@Notes") = Request("N" & rs("LineID"))
	If Request("CntctCode") <> "" Then cmd("@CntctCode") = Request("CntctCode")
	cmd.execute()
rs.movenext
loop

conn.close
response.redirect "../extPollOpen.asp?AdPollID=" & Request("AdPollID")
 %>
