<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->

<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
          
set rs = Server.CreateObject("ADODB.RecordSet")
myType = Request.Form("Type")

Select Case myType
	Case "cash"
		GetQuery rs, 8, null, null
	Case "check", "credit"
		GetQuery rs, 9, null, null
End Select

retVal = ""
do while not rs.eof
	If retVal <> "" Then retVal = retVal & "{S}"
	retVal = retVal & rs("AcctCode") & "{C}" & rs("AcctName") & "{C}" & rs("AcctDisp")
rs.movenext
loop

Response.Write retVal
%>