<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="myHTMLEncode.asp" -->
<!--#include file="adminTradSave.asp"-->
<!--#include file="repVars.inc" -->
<%
set rs = server.createobject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")

GroupID = CInt(Request("GroupID"))

Select Case Request("cmd")
	Case "addLine"
		Active = "N"
		If Request("Active") = "Y" Then Active = "Y"
		ID = "-1"
		If Request("ID") <> "" Then ID = CInt(Request("ID"))
		sql = "declare @ID int set @ID = " & GroupID & " " & _
				"declare @LineID int set @LineID = IsNull((select Max(LineID)+1 from OLKLayoutLines where ID = @ID), 0) " & _
				"insert OLKLayoutLines(ID, LineID, [Type], [TypeID], ColID, RowID, Active) " & _
				"values(@ID, @LineID, " & CInt(Request("Type")) & ", " & ID & ", " & CInt(Request("ColID")) & ", " & CInt(Request("RowID")) & ", '" & Active & "')"
		conn.execute(sql)
	case "delLine"
		sql = "delete OLKLayoutLines where ID = " & GroupID & " and LineID = " & Request("LineID")
		conn.execute(sql)
	case "updLay"
		If Request("LineID") <> "" Then
			arrLineID = Split(Request("LineID"), ", ")
			For i = 0 to UBound(arrLineID)
				LineID = CInt(arrLineID(i))
				rLineID = Replace(LineID, "-", "_")
				Active = "N"
				If Request("chkActive" & rLineID) = "Y" Then Active = "Y"
				sql = "update OLKLayoutLines set RowID = " & CInt(Request("RowID" & rLineID)) & ", Active = '" & Active & "' where ID = " & GroupID & " and LineID = " & LineID
				conn.execute(sql)
			Next
		End If
End select

Response.Redirect "adminLayout.asp?GroupID=" & GroupID

%>
