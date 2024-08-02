<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


ExecAt = Request.Form("ExecAt")
arrVars = Split(Request.Form("Variables"), "{S}")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
cmd.Parameters.Refresh()
cmd("@UserType") = userType
cmd("@ExecAt") = ExecAt
cmd("@AgentID") = Session("vendid")
cmd("@LanID") = Session("LanID")
cmd("@branch") = Session("branch")

Select Case ExecAt
	Case "O0"
		cmd("@ObjectEntry") = arrVars(0)
		cmd("@SetLogNumConf") = "N"
	Case "O1", "O7"
		cmd("@ObjectEntry") = arrVars(0)
		cmd("@Series") = arrVars(1)
		cmd("@SetLogNumConf") = "N"
	Case "O2", "O3", "O4" 
		cmd("@ObjectCode") = arrVars(0)
		cmd("@ObjectEntry") = arrVars(1)
		cmd("@SetLogNumConf") = "N"
	Case "D3" 
		cmd("@ObjectEntry") = Session("RetVal")
		cmd("@SetLogNumConf") = "Y"
	Case "R2" 
		cmd("@ObjectEntry") = Session("PayRetVal")
		cmd("@SetLogNumConf") = "Y"
	Case "A1" 
		cmd("@ObjectEntry") = Session("ItmRetVal")
		cmd("@SetLogNumConf") = "Y"
	Case "C1" 
		cmd("@ObjectEntry") = Session("CrdRetVal")
		cmd("@SetLogNumConf") = "Y"
	Case "C2" 
		cmd("@ObjectEntry") = Session("ActRetVal")
		cmd("@SetLogNumConf") = "Y"
End Select

cmd.execute()
id = cmd("@ID")

chkNote = False
arrFlow = Split(Request.Form("FlowID"), ", ")
If Request.Form("FlowNotes") <> "" Then 
	arrFlowNotes = Split(Request.Form("FlowNotes"), "{S}")
	chkNote = True
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCreateUAFControlDetail" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ID") = id
For i = 0 to UBound(arrFlow)	
	If chkNote Then If arrFlowNotes(i) <> "" Then cmd("@Note") = arrFlowNotes(i)
	cmd("@FlowID") = arrFlow(i)
	cmd.execute()
Next

Response.Write "ok"
%>