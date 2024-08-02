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

CatType = Request("CatType")

Select Case Request("cmd")
	Case "del"
		sql = "declare @ID int set @ID = " & Request("ID") & " " & _
		"delete OLKSmallCat where ID = @ID " & _
		"delete OLKSmallCatAlterNames where ID = @ID " & _
		"delete OLKLayoutLines where [Type] = 2 and [TypeID] = @ID "
		conn.execute(sql)
	case "upd"
		If Request("ID") <> "" Then
			arrID = Split(Request("ID"), ", ")
			For i = 0 to UBound(arrID)
				ID = arrID(i)
				If Request("chkStatus" & ID) = "Y" Then strStatus = "A" Else strStatus = "N"
				sql = "update OLKSmallCat set [Status] = '" & strStatus & "' where ID = " & ID
				conn.execute(sql)
			Next
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGenSmallCat" & Session("ID")
			cmd.Parameters.Refresh
			cmd.execute()
		End If
	Case "edit"
		If Request("strStatus") = "Y" Then strStatus = "A" Else strStatus = "N"
		If Request("Subtitle") <> "" Then Subtitle = "N'" & Request("Subtitle") & "'" Else Subtitle = "NULL"
		If Request("editID") <> "" Then
			ID = Request("editID")
			sql = "update OLKSmallCat set Name = N'" & saveHTMLDecode(Request("strName"), False) & "', Subtitle = " & Subtitle & ", " & _
				"Direction = '" & Request("Direction") & "', [Top] = " & Request("Top") & ", Query = N'" & saveHTMLDecode(Request("Query"), False) & "', Status = '" & strStatus & "' " & _
				"where ID = " & ID
			conn.execute(sql)
		Else 
			sql = "declare @ID int set @ID = IsNull((select Max(ID)+1 from OLKSmallCat), 0) select @ID ID " & _
				"insert OLKSmallCat(ID, Name, Subtitle, Direction, [Top], Query, Status, CatType) " & _
				"values(@ID, N'" & saveHTMLDecode(Request("strName"), False) & "', " & Subtitle & ", '" & Request("Direction") & "', " & Request("Top") & ", N'" & saveHTMLDecode(Request("Query"), False) & "', '" & strStatus & "', '" & CatType & "') "
			set rs = conn.execute(sql)
			ID = rs(0)
			If Request("nameTrad") <> "" Then
				SaveNewTrad Request("nameTrad"), "SmallCat", "ID", "alterName", ID
			End If

			If Request("subtitleTrad") <> "" Then
				SaveNewTrad Request("subtitleTrad"), "SmallCat", "ID", "alterSubtitle", ID
			End If

		End If
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGenSmallCat" & Session("ID")
		cmd.Parameters.Refresh
		cmd.execute()
		cmd("@ID") = ID
		cmd.execute()
		
		If Request("btnApply") <> "" Then Response.Redirect "adminSmallCat.asp?CatType=" & CatType & "&editID=" & ID
End select

Response.Redirect "adminSmallCat.asp?CatType=" & CatType

%>
