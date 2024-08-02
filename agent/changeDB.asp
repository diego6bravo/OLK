<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="lang.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<%
Session.Abandon

dbID = CInt(Request("dbID"))

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

isUpdated = myApp.IsDBUpdated(dbID)

retVal = GetYN(isUpdated) & "{S}"

If isUpdated Then
	myApp.LoadDBConfigData(dbID)
	retVal = retVal & GetYN(myApp.EnableBranchs) & "{S}"
	
	If myApp.EnableBranchs Then
		myApp.ConnectDatabase Session("olkdb")
		cmd.ActiveConnection = connCommon
		cmd.CommandText = "DBOLKGetBranchList" & dbID
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then retVal = retVal & "{O}"
			retVal = retVal & rs(0) & "{C}" & rs(1)
		rs.movenext
		loop
	End If
	
	Response.Write retVal
End If
%>