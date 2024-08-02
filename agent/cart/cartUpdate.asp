<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="../chkLogin.asp"-->
<!--#include file="../lcidReturn.inc"-->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../authorizationClass.asp"-->
<%
dim varxx
Dim varpricex

Dim myAut
set myAut = New clsAuthorization

varErr = ""

Session("CartGroup") = Request("CartGroup")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKUpdateDocSpecial"
cmd.Parameters.Refresh()

If Request("Draft") = "Y" Then
	cmd("@LogNum") = Session("RetVal")
	cmd("@cmd") = "D"
	cmd.execute()
End If

If Request("Authorize") = "Y" and Request("R1") = "17" Then
	cmd("@LogNum") = Session("RetVal")
	cmd("@cmd") = "C"
	cmd.execute()
End If

If Session("PayCart") Then
	cmd("@LogNum") = Session("PayRetVal")
	If Request("saldofuera") = "Y" Then cmd("@cmd") = "SF" Else cmd("@cmd") = "SFC"
	cmd.execute()
End If

If Request("DelLine") <> "" and Request("btnDelLines") <> "" Then 
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKDocDelLines" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	
	arrDel = Split(Request("DelLine"), ", ")
	
	For i = 0 to UBound(arrDel)
		cmd("@LineNum") = arrDel(i)
		cmd.Execute()
	Next
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
cmd.Parameters.Refresh
Select Case userType
	Case "C"
		cmd("@sessiontype") = "C"
	Case "V"
		cmd("@sessiontype") = "A"
End Select
cmd("@object") = Request("R1")
cmd("@LogNum") = Session("RetVal")
cmd("@transtype") = "A"
cmd("@CurrentSlpCode") = Session("vendid")
cmd("@Branch") = Session("branch")
cmd.execute()

conn.close 

If (Request("cartSubmit") = "I2" or Request("cartSubmit") = "I3") and varErr = "" Then
	Session("NotifyAdd") = True
	response.redirect "../cartSubmit.asp?submit=y&I=" & Request("cartSubmit") & "&DocConf=" & Request("DocConf") & "&Confirm=" & Request("Confirm")
Else
	If userType = "V" Then
		If Request("document") = "B" Then AddRedir = "&document=B&String=" & Request("String")
		response.redirect "../cart.asp?cmd=" & Request("redir") & "&update=Y" & varErr & "&ViewMode=" & Request("ViewMode") & AddRedir
	ElseIf userType = "C" then
		response.redirect "../cart.asp?update=Y" & varErr & "&ViewMode=" & Request("ViewMode")
	End If
End If

%>
