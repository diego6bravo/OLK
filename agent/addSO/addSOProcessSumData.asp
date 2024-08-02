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
cmd.CommandText = "DBOLKSOSetSumData"
cmd.Parameters.Refresh()
cmd("@LogNum") = LogNum
cmd("@FieldID") = CInt(Request.Form("FieldID"))
If Request.Form("MaxSumLoc") <> "" Then cmd("@MaxSumLoc") = CDbl(getNumeric(Request.Form("MaxSumLoc")))
If Request.Form("WtSumLoc") <> "" Then cmd("@WtSumLoc") = CDbl(getNumeric(Request.Form("WtSumLoc")))
If Request.Form("SumProfL") <> "" Then cmd("@SumProfL") = CDbl(getNumeric(Request.Form("SumProfL")))
If Request.Form("PrcntProf") <> "" Then cmd("@PrcntProf") = CDbl(getNumeric(Request.Form("PrcntProf")))
set rs = cmd.execute()
If Not IsNull(rs("MaxSumLoc")) Then Response.Write FormatNumber(CDbl(rs("MaxSumLoc")), myApp.SumDec)
Response.Write "{S}"
If Not IsNull(rs("WtSumLoc")) Then Response.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)
Response.Write "{S}"
If Not IsNull(rs("SumProfL")) Then Response.Write FormatNumber(CDbl(rs("SumProfL")), myApp.SumDec)
Response.Write "{S}" 
If Not IsNull(rs("PrcntProf")) Then Response.Write FormatNumber(CDbl(rs("PrcntProf")), myApp.SumDec)

%>