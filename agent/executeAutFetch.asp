<!--#include file="myHTMLEncode.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

myType = Request.Form("Type")
ID = Request.Form("ID")
LineID = Request.Form("LineID")
FlowID = Request.Form("FlowID")
DirectRate = myApp.DirectRate

Select Case myType
	Case "S" 'Submit
		Note = Request.Form("Note")
		Status = Request.Form("Status")
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKExecuteAuthorization" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@ID") = ID
		cmd("@LineID") = LineID
		cmd("@FlowID") = FlowID
		cmd("@Status") = Status
		cmd("@UserSign") = Session("vendid")
		If Note <> "" Then cmd("@Note") = Note
		
		cmd.Execute()
		
		Response.Write "ok{S}" & cmd("@Allready") 
	Case "N"
		sql = "select Note from OLKUAFControl where ID = " & ID
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		Response.Write rs(0)
End Select
        
%>