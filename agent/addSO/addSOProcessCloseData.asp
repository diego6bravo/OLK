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
set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSOSetCoseData"
cmd.Parameters.Refresh()
cmd("@LogNum") = LogNum
cmd("@FieldID") = CInt(Request.Form("FieldID"))
If Request.Form("OpenDate") <> "" Then cmd("@OpenDate") = SaveCmdDate(Request.Form("OpenDate"))
If Request.Form("DifNum") <> "" Then cmd("@DifNum") = Request.Form("DifNum")
If Request.Form("DifType") <> "" Then cmd("@DifType") = Request.Form("DifType")
If Request.Form("PredDate") <> "" Then cmd("@PredDate") = SaveCmdDate(Request.Form("PredDate"))
set rs = cmd.execute()
Response.Write rs("PredDateQty") & "{S}" & rs("DifType") & "{S}" & rs("PredDate")

%>