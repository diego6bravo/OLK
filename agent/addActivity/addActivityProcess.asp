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

LogNum = Session("ActRetVal")
Field = Request.Form("Field")
FieldType = Request.Form("FieldType")
Value = Request.Form("Value")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKProcessAJAX"
cmd.Parameters.Refresh()
cmd("@dbID") = Session("ID")
cmd("@LogNum") = LogNum
cmd("@TableID") = "TCLG"
cmd("@FieldID") = Field
cmd("@FieldType") = FieldType
If Value <> "" Then
	Select Case FieldType
		Case "S"
			cmd("@ValueText") = Value
		Case "N"
			cmd("@ValueNumeric") = CDbl(Value)
		Case "I"
			cmd("@ValueInt") = CLng(Value)
		Case "D"
			cmd("@ValueDate") = SaveCmdDate(Value)
		Case "T"
			cmd("@ValueDate") = SaveCmdTime(Value)
			cmd("@FieldType") = "D"
	End Select
End If
cmd.execute()

Response.Write "ok"

%>