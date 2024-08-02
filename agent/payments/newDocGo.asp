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
<!--#include file="../authorizationClass.asp"-->
<%
response.buffer = true

Dim myAut
set myAut = New clsAuthorization


If myApp.CopyLastFCRate Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCopyLastFCRate" & Session("ID")
	cmd.Parameters.Refresh()
	cmd.execute()
End If

set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SysID") = "ORCT"
cmd("@ObsID") = "TPMT"
set rs = cmd.execute()
If rs(0) = "Y" Then Response.Redirect "../configErr.asp?errCmd=Pay"

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckAgentObjectCreation" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjID") = 24
cmd("@CardCode") = Session("UserName")
cmd("@UserAccess") = Session("UserAccess")
cmd("@SlpCode") = Session("vendid")
rs.close
rs.open cmd, , 3, 1
        
If rs("AsignedSLP") = "Y" Then Response.Redirect "../configErr.asp?errCmd=AsignedSLP"
For each itm in rs.Fields
	if itm = "Y" Then Response.Redirect "../configErr.asp?errCmd=Pay"
next

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCreateNewPayment" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SlpCode") = Session("vendid")
cmd("@branchIndex") = Session("branch")
cmd("@SessionType") = "A"
cmd("@CardCode") = Session("UserName")
cmd.execute()

RetVal = cmd.Parameters.Item(0).Value
Session("PayRetVal") = RetVal
Session("PayCart") = False
Session("RetVal") = ""

conn.close
Response.Redirect "../agentPayment.asp"
%>