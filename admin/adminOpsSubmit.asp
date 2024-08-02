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
<%
set rs = server.createobject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")

Select Case Request("cmd")
	Case "delOp"
		sql = "declare @ID int set @ID = " & Request("ID") & " " & _
			"delete OLKOps where ID = @ID " & _
			"delete OLKOpsAlterNames where ID = @ID " & _
			"delete OLKOpsLines where ID = @ID " & _
			"delete OLKOpsLinesAlterNames where ID = @ID "
		conn.execute(sql)
		Response.Redirect "adminOps.asp"
	Case "delGrp"
		sql = "declare @ID int set @ID = " & Request("delID") & " " & _
		"delete OLKOpsGrps where ID = @ID " & _
		"delete OLKOpsGrpsAlterNames where ID = @ID"
		conn.execute(sql)
		Response.Redirect "adminOps.asp"
	Case "updGrp"
		If Request("ID") <> "" Then
			arrID = Split(Request("ID"), ", ")
			For i = 0 to UBound(arrID)
				sql = "update OLKOpsGrps set Name = N'" & saveHTMLDecode(Request("grpName" & arrID(i)), False) & "' where ID = " & arrID(i)
				conn.execute(sql)
			Next
			If Request("newGrpName") <> "" Then
				sql = "declare @ID int set @ID = IsNull((select Max(ID)+1 from OLKOpsGrps), 0) " & _
					"select @ID ID insert OLKOpsGrps(ID, Name) values(@ID, N'" & saveHTMLDecode(Request("newGrpName"), False) & "')"
				set rs = conn.execute(sql)
				
				If Request("newGrpNameTrad") <> "" Then
					SaveNewTrad Request("newGrpNameTrad"), "OpsGrps", "ID", "alterName", rs(0)
				End If

			End If
		End If
		Response.Redirect "adminOps.asp"
	Case "updOps"
		If Request("ID") <> "" Then
			arrID = Split(Request("ID"), ", ")
			For i = 0 to UBound(arrID)
				If Request("Status" & arrID(i)) = "Y" Then Status = "A" Else Status = "N"
				sql = "update OLKOps set Name = N'" & saveHTMLDecode(Request("opName" & arrID(i)), False) & "', Status = '" & Status & "' where ID = " & arrID(i)
				conn.execute(sql)
			Next
		End If
		Response.Redirect "adminOps.asp"
	Case "editOP"
		ID = CInt(Request("ID"))
		If Request("TargetObjectID") <> "" Then TargetObjectID = Request("TargetObjectID") Else TargetObjectID = "NULL"
		If Request("Status") = "Y" Then opStatus = "A" Else opStatus = "N"
		If Request("Filter") <> "" Then opFilter = "N'" & saveHTMLDecode(Request("Filter"), False) & "'" Else opFilter = "NULL"
		If ID = -1 Then
			If Request("chkGenNewDoc") = "Y" Then GenNewDoc = "Y" Else GenNewDoc = "N"
			
			sql = "declare @ID int set @ID = IsNull((select Max(ID)+1 from OLKOps), 0) select @ID ID " & _
					"insert OLKOps(ID, Name, GroupID, ObjectID, Operation, TrgtObjID, Status, GenNewDoc) " & _
					"values(@ID, N'" & saveHTMLDecode(Request("opName"), False) & "', " & Request("GroupID") & ", " & Request("ObjectID") & ", " & Request("Operation") & ", " & TargetObjectID & ", '" & opStatus & "', '" & GenNewDoc & "')"
			set rs = conn.execute(sql)
			ID = rs(0)
		Else
			sql = "update OLKOps set Name = N'" & saveHTMLDecode(Request("opName"), False) & "', GroupID = " & Request("GroupID") & ", Status = '" & opStatus & "' where ID = " & Request("ID")
			conn.execute(sql)
		End if
		If Request("btnApply") <> "" Then
			Response.Redirect "adminOpsEdit.asp?ID=" & ID
		Else
			Response.Redirect "adminOps.asp"
		End If
	Case "editOPDet"
		ID = CInt(Request("ID"))
		TypeID = CInt(Request("TypeID"))
		lineTypeID = TypeID
		LineID = Split(Request("LineID"), ", ")
		For i = 0 to UBound(LineID)
			arrData = Split(LineID(i), "_")
			line = CInt(arrData(0))
			If arrData(1) <> "" Then lineTypeID = arrData(1)
			
			If Request("alterType" & LineID(i)) <> "" Then TypeID = Request("alterType" & LineID(i))
			
			sql = "update OLKOpsLines set StyleID = " & Request("styleID" & LineID(i)) & ", " & _
					"ColID = " & Request("colID" & LineID(i)) & ", " & _
					"Ordr = " & Request("orderID" & LineID(i))
			If CInt(Request("styleID" & LineID(i))) = 4 Then
				sql = sql & ", AliasDesc = N'" & saveHTMLDecode(Request("aliasDesc" & LineID(i)), False) & "' "
			End If
			sql = sql & " where ID = " & ID & " and TypeID = " & lineTypeID & " and LineID = " & line
			conn.execute(sql)
		Next
		
		Response.Redirect "adminOpsEdit.asp?ID=" & ID & "&Type=" & TypeID
	Case "editOPFldDet"
		ID = CInt(Request("ID"))
		TypeID = CInt(Request("TypeID"))
		
		ColID = CInt(Request("colID"))
		orderID = CInt(Request("orderID"))
		AliasDesc = saveHTMLDecode(Request("AliasDesc"), False)
		
		If Request("LineID") <> "" Then
			AliasID = saveHTMLDecode(Request("AliasID"), False)
			StyleID = 4
			LineID = CInt(Request("LineID"))
			
			sql = "update OLKOpsLines set ColID = " & ColID & ", Ordr = " & orderID & ", " & _
					"StyleID = " & StyleID & ", AliasDesc = N'" & AliasDesc & "', " & _
					"AliasID = N'" & AliasID & "' where ID = " & ID & " and TypeID = " & TypeID & " and LineID = " & LineID
			conn.execute(sql)
		Else
			If Request("fldID") = "Custom" Then
				AliasID = saveHTMLDecode(Request("AliasID"), False)
				StyleID = 4
			Else
				AliasID = Split(Request("fldID"), "{S}")(3)
				alterType = Split(Request("fldID"), "{S}")(4)
				If alterType <> "" Then TypeID = CInt(alterType)
				StyleID = Request("StyleID")
			End If
			sql = "declare @ID int set @ID = " & ID & " " & _
					"declare @TypeID int set @TypeID = " & TypeID & " " & _
					"declare @LineID int set @LineID = IsNull((select Max(LineID) + 1 from OLKOpsLines where ID = @ID and TypeID = @TypeID), 0) " & _
					"select @LineID LineID " & _
					"insert OLKOpsLines(ID, TypeID, LineID, AliasID, AliasDesc, StyleID, ColID, Ordr) " & _
					"values(@ID, @TypeID, @LineID, N'" & AliasID & "', " & _
					"N'" & AliasDesc & "', " & StyleID & ", " & ColID & ", " & orderID & ") "
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			LineID = rs("LineID")	
		End If
		
		If Request("btnApply") <> "" Then
			Response.Redirect "adminOpsEdit.asp?ID=" & ID & "&Type=" & TypeID & "&LineID=" & LineID
		Else
			Response.Redirect "adminOpsEdit.asp?ID=" & ID & "&Type=" & TypeID
		End If
	Case "delLine"
		ID = CInt(Request("ID"))
		TypeID = CInt(Request("Type"))
		LineID = CInt(Request("LineID"))
		sql = "declare @ID int set @ID = " & ID & " declare @TypeID int set @TypeID = " & TypeID & " declare @LineID int set @LineID = " & LineID & " " & _
		"delete OLKOpsLines where ID = @ID and TypeID = @TypeID and LineID = @LineID " & _
		"delete OLKOpsLinesAlterNames where ID = @ID and TypeID = @TypeID and LineID = @LineID "
		conn.execute(sql)
		Response.Redirect "adminOpsEdit.asp?ID=" & ID & "&Type=" & TypeID
End Select

%>