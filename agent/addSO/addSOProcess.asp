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

LogNum = Session("SORetVal")
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
cmd("@TableID") = "TOPR"
cmd("@FieldID") = Field
cmd("@FieldType") = FieldType
If Value <> "" Then
	Select Case FieldType
		Case "S"
			cmd("@ValueText") = Value
		Case "N"
			cmd("@ValueNumeric") = CDbl(getNumericOut(Value))
		Case "I"
			cmd("@ValueInt") = CLng(getNumeric(Value))
		Case "D"
			cmd("@ValueDate") = SaveCmdDate(Value)
		Case "T"
			cmd("@ValueDate") = SaveCmdTime(Value)
			cmd("@FieldType") = "D"
	End Select
End If
cmd.execute()

Select Case Field
	Case "ChnCrdCode"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSOSetChnCrdCode" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = cmd.execute()
		
		Response.Write rs(0)
		Response.Write "{S}"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@CardCode") = Value
		set rs = cmd.execute()
		do while not rs.eof
			Response.Write rs("CntctCode") & "{C}" & rs("Name") & "{V}"
		rs.movenext
		loop
		
	Case Else
		Response.Write "ok"
End Select

%>