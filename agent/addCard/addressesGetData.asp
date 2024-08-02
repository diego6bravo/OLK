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
set rs = Server.CreateObject("ADODB.RecordSet")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

Select Case Request.Form("Type")
	Case "S"
		cmd.CommandText = "DBOLKGetCountryStates" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@Code") = Request.Form("Country")	
End Select

rs.open cmd, , 3, 1

retVal = ""
do while not rs.eof
	If retVal <> "" Then retVal = retVal & "{S}"
	retVal = retVal & rs(0) & "{C}" & rs(1)
rs.movenext
loop
Response.Write retVal
%>