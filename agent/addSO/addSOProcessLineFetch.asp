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
set rs = Server.CreateObject("ADODB.RecordSet")

LogNum = Session("SORetVal")
Line = CInt(Request("Line"))
DataType = CInt(Request("DataType"))
Field = Request("Field")
FieldType = Request("FieldType")
Value = Request("Value")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSOSetLineData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = LogNum
cmd("@Line") = Line
cmd("@DataType") = DataType
cmd("@FieldID") = Field
If Value <> "" Then
	Select Case FieldType
		Case "S"
			cmd("@ValueText") = Value
		Case "N"
			cmd("@ValueNumeric") = CDbl(getNumericOut(Value))
		Case "I"
			cmd("@ValueInt") = CLng(getNumeric(Value))
		Case "D"
			cmd("@ValueDate") = SaveCmdDate(Value)
		Case "T"
			cmd("@ValueDate") = SaveCmdTime(Value)
			cmd("@FieldType") = "D"
	End Select
End If
set rs = cmd.execute()

Select Case DataType
	Case 1
		Select Case Field
			Case "Step_Id"
				With Response
					.Write FormatNumber(CDbl(rs("ClosePrcnt")), myApp.PercentDec)
					.Write "{S}"
					.Write FormatNumber(CDbl(rs("MaxSumLoc")), myApp.SumDec)
					.Write "{S}"
					.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)
				End With
			Case "ClosePrcnt"
				With Response
					.Write FormatNumber(CDbl(rs("ClosePrcnt")), myApp.PercentDec)
					.Write "{S}"
					.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)
				End With
			Case "MaxSumLoc", "WtSumLoc"
				With Response
					.Write FormatNumber(CDbl(rs("MaxSumLoc")), myApp.SumDec)
					.Write "{S}"
					.Write FormatNumber(CDbl(rs("WtSumLoc")), myApp.SumDec)
					.Write "{S}"
					.Write FormatNumber(CDbl(rs("SumProfL")), myApp.SumDec)
				End With
		End Select
	Case 2
		Select Case Field
			Case "ParterId"
				If Value <> "" Then
					With Response
						.Write rs("OrlCode")
						.Write "{S}"
						.Write rs("RelatCard")
						.Write "{S}"
						.Write rs("Memo")
					End With
				Else
					Response.Write "Rem"
				End If
		End Select
	Case 3
		Select Case Field
			Case "CompetId"
				If Value <> "" Then
					With Response
						.Write rs("Memo")
						.Write "{S}"
						.Write rs("ThreatLevl")
					End With
				Else
					Response.Write "Rem"
				End If
		End Select
	Case 4
		Select Case Field
			Case "NewInt", "NewReason"
				Response.Write rs(0)
			Case "NewComp"
				With Response
					.Write rs(0)
					.Write "{S}"
					.Write rs(1)
					.Write "{S}"
					.Write rs(2)
				End With
			Case "NewBP"
				With Response
					.Write rs(0)
					.Write "{S}"
					.Write rs(1)
					.Write "{S}"
					.Write rs(2)
					.Write "{S}"
					.Write rs(3)
				End With
			Case "NewStage"
				With Response
					.Write rs(0)
					.Write "{S}"
					If Not IsNull(rs(1)) Then .Write FormatNumber(CDbl(rs(1)), myApp.PercentDec)
					.Write "{S}"
					If Not IsNull(rs(2)) Then .Write FormatNumber(CDbl(rs(2)), myApp.SumDec)
					.Write "{S}"
					If Not IsNull(rs(3)) Then .Write FormatNumber(CDbl(rs(3)), myApp.SumDec)
					.Write "{S}"
					.Write rs(4)
					.Write "{S}"
					.Write rs(5)
					.Write "{S}"
					.Write rs(6)
					.Write "{S}"
					.Write rs(7)
				End With
		End Select
End Select
%>