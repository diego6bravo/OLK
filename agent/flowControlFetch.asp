<!--#include file="myHTMLEncode.asp"--><!--#include file="lcidReturn.inc"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
          
set rFlow = Server.CreateObject("ADODB.RecordSet")
set rChk = Server.CreateObject("ADODB.RecordSet")
set rGen = Server.CreateObject("ADODB.RecordSet")


ExecAt = Request("ExecAt")
arrVars = Split(Request("Variables"), "{S}")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetDFList" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ExecAt") = ExecAt
If ExecAt = "D1" Then cmd("@ObjectCode") = arrVars(0)
cmd("@UserType") = userType
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid")
If Session("RetVal") <> "" Then cmd("@LogNum") = Session("RetVal")
set rFlow = cmd.execute()

retVal = ""
Sub LoadCmdParams
	cmd.Parameters.Refresh()

	cmd("@LanID") = Session("LanID")
	cmd("@SlpCode") = Session("vendid")
	cmd("@dbName") = Session("olkdb")
	cmd("@branch") = Session("branch")
	cmd("@UserType") = userType
	
	Select Case ExecAt
		Case "O0", "O1", "O7" ' Aprove Sales Order, Convert Quotation to Sales Order, Convert Sales Order to Invoice
			cmd("@Entry") = arrVars(0)
		Case "O2", "O3", "O6" ' Close  Object, Cancel Object, Remove Object
			cmd("@ObjectCode") = arrVars(0) 
			cmd("@Entry") = arrVars(1)
		Case "D2" ' Add Item
			cmd("@LogNum") = Session("RetVal")
			cmd("@CardCode") = Session("UserName")
			cmd("@ItemCode") = arrVars(0)
			If arrVars(1) <> "null" and arrVars(1) <> "" Then cmd("@Quantity") = CDbl(getNumericOut(arrVars(1)))
			If arrVars(2) <> "null" and arrVars(2) <> "" Then cmd("@Unit") = arrVars(2)
			If arrVars(3) <> "null" and arrVars(3) <> "" Then cmd("@Price") = CDbl(getNumericOut(arrVars(3)))
			If arrVars(4) <> "null" and arrVars(4) <> "" Then cmd("@WhsCode") = arrVars(4)
			If UBound(arrVars) > 4 Then 
				If arrVars(5) = "Y" Then 
					cmd("@All") = "Y"
				End If
			End If
		Case "D3" ' LtxtDocConf
			cmd("@LogNum") = Session("RetVal")
		Case "R1" ' LtxtCreation	******clean*******
		Case "R2" ' LtxtRcpConf
			cmd("@LogNum") = Session("PayRetVal")
		Case "A1" ' LtxtItmConf
			cmd("@LogNum") = Session("ItmRetVal")
		Case "C1" ' LtxtClientConf
			cmd("@LogNum") = Session("CrdRetVal")
		Case "C2" ' LtxtActivityConf
			cmd("@LogNum") = Session("ActRetVal")
		Case "C3" ' LtxtActivityConf
			cmd("@LogNum") = Session("SORetVal")
	End Select	
	
	Select Case ExecAt
		Case "C2", "C3", "R1", "R2", "D3", "D1"
			cmd("@CardCode") = Session("UserName")
	End Select
End Sub
do while not rFlow.eof
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(rFlow("FlowID"), "-", "_")
	LoadCmdParams

	set rChk = cmd.execute()
	If not rChk.eof then
		If Not IsNull(rChk(0)) Then
			If lcase(rChk(0)) = lcase("True") Then
				If retVal <> "" Then retVal = retVal & "{F}"
				
				Note = rFlow("NoteText")
				Note = Replace(Note, VbCrLf, "<br>")
				If rFlow("NoteBuilder") = "Y" Then
					cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(rFlow("FlowID"), "-", "_") & "_msg"
					LoadCmdParams
					set rGen = cmd.execute()
					If Not rGen.Eof Then
						For each fld in rGen.Fields
							If Not IsNull(fld) Then Note = Replace(Note, "{" & fld.Name & "}", fld) Else Note = Replace(Note, "{" & fld.Name & "}", "")
						Next
					End If
				End If
				
				strTable = ""
				If rFlow("LineQry") = "Y" Then
					cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(rFlow("FlowID"), "-", "_") & "_line"
					LoadCmdParams
					set rGen = cmd.execute()
					
					If Not rGen.Eof Then
						strLine = ""
						For each fld in rGen.Fields
							If strLine <> "" Then strLine = strLine & "{C}"
							strLine = strLine & fld.Name
						Next
						strTable = strLine
						do while not rGen.Eof
							strLine = ""
							For each fld in rGen.Fields
								If strLine <> "" Then strLine = strLine & "{C}"
								strLine = strLine & fld
							Next
							strTable = strTable & "{R}" & strLine
						rGen.movenext
						loop
					End If
				End If
				
				retVal = retVal & rFlow("FlowID") & "{S}" & rFlow("Name") & "{S}" & rFlow("Type") & "{S}" & strTable & "{S}" & Note & "{S}" & rFlow("Draft") & "{S}" & rFlow("Authorize")
				If rFlow("Type") = 0 Then Exit do
			End If
		End If
	End If
rFlow.movenext
loop

Response.Write retVal
%>