<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp"-->
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="../lcidReturn.inc" -->
<%
set rs = Server.CreateObject("ADODB.RecordSet")

If Request("btnDel") = "" Then
	set rv = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFCmdSave" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = "CRD1"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rv.open cmd, , 3, 1
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCrdSaveAddData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("CrdRetVal")
	If Request("Op") <> "add" Then cmd("@LineNum") = Request("Op")
	cmd("@NewAddress") = Request("NewAddress")
	cmd("@AdresType") = Request("AdresType")
	If Request("Street") <> "" Then cmd("@Street") = Request("Street")
	If Request("Block") <> "" Then cmd("@Block") = Request("Block")
	If Request("City") <> "" Then cmd("@City") = Request("City")
	If Request("ZipCode") <> "" Then cmd("@ZipCode") = Request("ZipCode")
	If Request("County") <> "" Then cmd("@County") = Request("County")
	If Request("Country") <> "" Then cmd("@Country") = Request("Country")
	If Request("State") <> "" Then cmd("@State") = Request("State")
	If Request("TaxCode") <> "" Then cmd("@TaxCode") = Request("TaxCode")
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
	
	op = cmd.Parameters.Item(0).Value

Else
	cmd.CommandText = "DBOLKCrdRemAdd"
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("CrdRetVal")
	cmd("@LineNum") = Request("Op")
	cmd("@AdresType") = Request("AdresType")
	cmd.execute()
End If

If Request("btnApply") <> "" Then
	Response.Redirect "addresses.asp?AdresType=" & Request("AdresType") & "&Op=" & op
Else
	Response.Redirect "addresses.asp?AdresType=" & Request("AdresType")
End If
%>