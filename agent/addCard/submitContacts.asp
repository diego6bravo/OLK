<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="../lcidReturn.inc" -->
<%
set rs = Server.CreateObject("ADODB.RecordSet")

If Request("btnDel") = "" Then
	set rv = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFCmdSave" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = "OCPR"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rv.open cmd, , 3, 1
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCrdSaveCntData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("CrdRetVal")
	If Request("Op") <> "add" Then cmd("@LineNum") = Request("Op")
	cmd("@NewName") = Request("NewName")
	If Request("Position") <> "" Then cmd("@Position") = Request("Position")
	If Request("Address") <> "" Then cmd("@Address") = Request("Address")
	If Request("Title") <> "" Then cmd("@Title") = Request("Title")
	If Request("Tel1") <> "" Then cmd("@Tel1") = Request("Tel1")
	If Request("Tel2") <> "" Then cmd("@Tel2") = Request("Tel2")
	If Request("Cellolar") <> "" Then cmd("@Cellolar") = Request("Cellolar")
	If Request("Fax") <> "" Then cmd("@Fax") = Request("Fax")
	If Request("EMail") <> "" Then cmd("@E_MailL") = Request("EMail")
	If Request("Pager") <> "" Then cmd("@Pager") = Request("Pager")
	If Request("Notes1") <> "" Then cmd("@Notes1") = Request("Notes1")
	If Request("Notes2") <> "" Then cmd("@Notes2") = Request("Notes2")
	If Request("Password") <> "" Then cmd("@Password") = Request("Password")
	If Request("BirthPlace") <> "" Then cmd("@BirthPlace") = Request("BirthPlace")
	If Request("BirthDate") <> "" Then cmd("@BirthDate") = SaveCmdDate(Request("BirthDate"))
	If Request("Gender") <> "" Then cmd("@Gender") = Request("Gender")
	If Request("Profession") <> "" Then cmd("@Profession") = Request("Profession")
	If Request("SetDef") = "Y" Then cmd("@SetDef") = "Y" Else cmd("@SetDef") = "N"
	
	do while not rv.eof
		strVal = Request(rv("InsertID"))
		If strVal <> "" Then
			Select Case rv("TypeID") 
				Case "B" 
					cmd("@" & rv("InsertID")) = CDbl(getNumericOut(strVal))
				Case "D" 
					cmd("@" & rv("InsertID")) = SaveCmdDate(strVal)
				Case Else
					cmd("@" & rv("InsertID")) = strVal
			End Select
		End If
	rv.movenext
	loop
	
	cmd.execute()
	
	op = cmd.Parameters.Item(0).value
Else
	cmd.CommandText = "DBOLKCrdRemCnt"
	cmd.Paramets.Refresh()
	cmd("@LogNum") = Session("CrdRetVal")
	cmd("@LineNum") = CInt(Request("Op"))
	cmd.execute()
End If

If Request("btnApply") <> "" Then
	Response.Redirect "contacts.asp?Op=" & op
Else
	Response.Redirect "contacts.asp"
End If
%>