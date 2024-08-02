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
Line = CInt(Request("Line"))
set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSOSetAmtGrs" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = LogNum
cmd("@Line") = Line
set rs = cmd.execute()
If Not IsNull(rs("MaxSumLoc")) Then Response.Write FormatNumber(CDbl(rs("MaxSumLoc")), myApp.SumDec)
Response.Write "{S}"
If Not IsNull(rs("WtSumLoc")) Then Response.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)
Response.Write "{S}" 
If Not IsNull(rs("PrcntProf")) Then Response.Write FormatNumber(CDbl(rs("PrcntProf")), myApp.SumDec)
Response.Write "{S}"
If Not IsNull(rs("SumProfL")) Then Response.Write FormatNumber(CDbl(rs("SumProfL")), myApp.SumDec)

%>