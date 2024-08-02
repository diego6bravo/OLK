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
<!--#include file="../lcidReturn.inc" -->
<%
set rs = Server.CreateObject("ADODB.RecordSet")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

myType = Request("Type")
Value = Request("Value")

select case myType 
	Case "chkCode"
		cmd.CommandText = "DBOLKValidateNewCard" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@CardCode") = Value
		set rs = cmd.execute()
		Response.Write rs(0)
	Case "crdGroups"
		cmd.CommandText = "DBOLKGetCrdGroups" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@CardType") = Value
		set rs = cmd.execute()
		retVal = ""
		do while not rs.eof
			If retVal <> "" Then retVal = retVal & "{S}"
			retVal = retVal & rs(0) & "{V}" & rs(1)
		rs.movenext
		loop
		Response.Write retVal
End Select


%>