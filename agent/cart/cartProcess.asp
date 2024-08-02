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
<!--#include file="../authorizationClass.asp"-->
<%

LogNum = Session("RetVal")
Field = Request("Field")
FieldType = Request("FieldType")
Value = Request("Value")
LineID = Request("LineID")
LineType = Request("LineType")
SumDec = myApp.SumDec

Select Case Field
	Case "GroupNum"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSetGroupNum" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@GroupNum") = Value
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		strResp = FormatDate(rs(0), False)
	Case "CheckSumInv"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartCheckSumInv" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@Lines") = Request("Lines")
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		strResp = rs(0)
	Case "CheckLinesInv"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartCheckLinesInv" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@Lines") = Request("Lines")
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		strResp = ""
		do while not rs.eof
			If strResp <> "" Then strResp = strResp & "{S}"
			strResp = strResp & rs(0) & "{C}" & rs(1)
		rs.movenext
		loop
	Case "IsLocked"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetCardIsLocked" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		strResp = rs(0)
	Case "ObjectCode"
		Dim myAut
		set myAut = New clsAuthorization
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveObjType" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@Object") = Value
		
		strSeries = myAut.GetObjectProperty(Value, "S")
		If strSeries <> "NULL" Then cmd("@Series") = strSeries
		
		cmd.execute()
		
		strResp = FormatDate(cmd("@DocDueDate"), False) & "{S}" & cmd("@ChkInvOP")
	Case "GroupNumAppyList"
		Session("PriceList") = Value
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKApplyPListLines" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("RetVal")
		cmd("@Pricelist") = Value
		cmd("@CardCode") = Session("UserName")
		cmd("@UserType") = userType
		cmd.execute()
	Case "DocDiscount"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveDiscount" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@DiscPrcnt") = CDbl(getNumericOut(Value))
		cmd("@PercentDec") = myApp.PercentDec
		cmd("@SlpCode") = Session("vendid")
		cmd("@UserAccess") = Session("UserAccess")
		cmd.execute()
		Response.Write FormatNumber(CDbl(cmd("@DiscPrcnt")), myApp.PercentDec)
		Response.Write "{S}"
		Response.Write cmd("@ErrMaxDisc")
	Case "DpmPrcnt"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveDPM"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@DPMPrcnt") = CDbl(getNumericOut(Value))
		cmd("@PercentDec") = myApp.PercentDec
		cmd("@SlpCode") = Session("vendid")
		cmd("@UserAccess") = Session("UserAccess")
		cmd.execute()
		Response.Write FormatNumber(CDbl(cmd("@DPMPrcnt")), myApp.PercentDec)
		Response.Write "{S}"
		Response.Write cmd("@ErrMaxDPM")
	Case "DocDate"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveDocDate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@DocDate") = SaveCmdDate(Value) 
		cmd.execute()
		Response.Write FormatDate(cmd("@DocDueDate"), False)
		Response.Write "{S}"
		Response.Write cmd("@ErrDocDate")
	Case "DocCur"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveCurrency" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@Currency") = Value
		cmd("@MainCur") = myApp.MainCur
		cmd("@DirectRate") = GetYN(myApp.DirectRate)
		cmd.execute()
	Case Else
		myTable = "TDOC" 
		cmdText = "DBOLKProcessAJAX"
		
		Select Case LineType
			Case "I"
				myTable = "DOC1"
				cmdText = "DBOLKProcessAJAXLine"
			Case "E"
				myTable = "DOC3"
				cmdText = "DBOLKProcessAJAXLine"
		End Select
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = cmdText
		cmd.Parameters.Refresh()
		cmd("@dbID") = Session("ID")
		cmd("@LogNum") = Session("RetVal")
		If LineType = "I" or LineType = "E" Then cmd("@Line") = LineID
		cmd("@TableID") = myTable
		cmd("@FieldID") = Field
		cmd("@FieldType") = FieldType
		
		If Value <> "" Then
			Select Case FieldType
				Case "S"
					cmd("@ValueText") = Value
				Case "N"
					cmd("@ValueNumeric") = CDbl(getNumericOut(Value))
				Case "D"
					cmd("@ValueDate") = SaveCmdDate(Value)
			End Select
		End If
		
		cmd.execute()
End Select
			
Select Case Field
	Case "Quantity"
		If LineType = "I" Then
			set rs = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKCheckCartLineLock" & Session("ID")
			cmd.Parameters.Refresh()
			set rs = cmd.execute()
			Response.Write "ok|" & rs(0) & "|" & rs(1)
		Else
			Response.Write "ok"	
		End If
	Case "GroupNumAppyList" 
		set rs = Server.CreateObject("ADODB.RecordSet")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCardGetLinesData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@MainCur") = myApp.MainCur
		cmd("@SumDec") = myApp.SumDec
		cmd("@DirectRate") = GetYN(myApp.DirectRate)
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		If LineID <> "" Then cmd("@Lines") = LineID
		set rs = cmd.execute()
		
		strResp = ""
		do while not rs.eof
			If strResp <> "" Then strResp = strResp & "{S}"
			strResp = strResp & rs(0) & "{C}" & FormatNumber(CDbl(rs(1)), myApp.PriceDec) & "{C}" & FormatNumber(CDbl(rs(2)), myApp.PriceDec) & "{C}" & rs(3) & "{C}" & FormatNumber(CDbl(rs(4)), myApp.PriceDec) & "{C}" & FormatNumber(CDbl(rs(5)), SumDec)
		rs.movenext
		loop
		
		strResp = strResp & "|"
		
		If LineID <> "" Then
		
			set rs = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKCardGetLinesSumData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LogNum") = LogNum
			cmd("@MainCur") = myApp.MainCur
			cmd("@DirectRate") = GetYN(myApp.DirectRate)
			cmd("@SumDec") = SumDec
			cmd("@Lines") = LineID
			set rs = cmd.execute()
			
			strResp = strResp & FormatNumber(CDbl(rs(0)), SumDec)
		End If
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		If Session("PayCart") Then cmd("@MC") = "Y"
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		
		strResp = strResp & "|" & FormatNumber(CDbl(rs(0)), SumDec) & "{S}" & FormatNumber(CDbl(rs(1)), SumDec) & "{S}" & FormatNumber(CDbl(rs(2)), SumDec) & "{S}" & FormatNumber(CDbl(rs(3)), SumDec) & "{S}" & FormatNumber(CDbl(rs(4)), SumDec)
		If Session("PayCart") Then Response.Write "{S}" & FormatNumber(CDbl(rs(5)), SumDec)

		Response.Write "ok|" & strResp
	Case "ShipToCode", "PayToCode"
		set rs = Server.CreateObject("ADODB.RecordSet")
		myType = "S"
		If Field = "PayToCode" Then myType = "B"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetFormatedAddress" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@CardCode") = Session("UserName")
		cmd("@Type") = myType
		cmd("@Address") = Value
		cmd("@LanID") = Session("LanID")
		cmd("@UserType") = userType
		cmd("@OP") = "O"
		set rs = cmd.execute()
		Response.Write "ok|" & rs(0)
	Case "DocDueDate" 
		set rs = Server.CreateObject("ADODB.RecordSet") 
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKDBGetCardDueDate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		set rs = cmd.execute()
		Response.Write "ok|" & rs(0) & "|" & rs(1)
	Case "DocDiscount"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		If Session("PayCart") Then cmd("@MC") = "Y"
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		
		Response.Write "{S}" & FormatNumber(CDbl(rs("Discount")), SumDec) & "{S}" & FormatNumber(CDbl(rs("Tax")), SumDec) & "{S}" & FormatNumber(CDbl(rs("DocTotal")), SumDec)
		If Session("PayCart") Then Response.Write "{S}" & FormatNumber(CDbl(rs(5)), SumDec)
	Case "DpmPrcnt"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		If Session("PayCart") Then cmd("@MC") = "Y"
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		
		Response.Write "{S}" & FormatNumber(CDbl(rs("DPM")), SumDec) & "{S}" & FormatNumber(CDbl(rs("Tax")), SumDec) & "{S}" & FormatNumber(CDbl(rs("DocTotal")), SumDec)
		If Session("PayCart") Then Response.Write "{S}" & FormatNumber(CDbl(rs(5)), SumDec)
	Case "DocDate", "GroupNum"
		WriteTaxAndTotal
	Case "DocCur"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartGetLinesTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@MainCur") = myApp.MainCur
		cmd("@SumDec") = myApp.SumDec
		cmd("@DirectRate") = GetYN(myApp.DirectRate)
		If LineID <> "" Then cmd("@Lines") = LineID
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{L}"
			Response.Write rs(0) & "{C}" & FormatNumber(CDbl(rs(1)), SumDec)
		rs.movenext
		loop
		
		Response.Write "{S}"
		
		If LineID <> "" Then
		
			set rs = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKCardGetLinesSumData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LogNum") = LogNum
			cmd("@MainCur") = myApp.MainCur
			cmd("@DirectRate") = GetYN(myApp.DirectRate)
			cmd("@SumDec") = SumDec
			cmd("@Lines") = LineID
			set rs = cmd.execute()
			
			Response.Write FormatNumber(CDbl(rs(0)), SumDec)
		End If
		
		Response.Write "{S}"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartGetExpensesTotalData"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{L}"
			Response.Write rs(0) & "{C}" & FormatNumber(CDbl(rs(1)), SumDec)
		rs.movenext
		loop
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		If Session("PayCart") Then cmd("@MC") = "Y"
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		
		Response.Write "{S}" & FormatNumber(CDbl(rs("SubTotal")), SumDec) & "{S}" & FormatNumber(CDbl(rs("Discount")), SumDec) & "{S}" & FormatNumber(CDbl(rs("Tax")), SumDec) & "{S}" & FormatNumber(CDbl(rs("DocTotal")), SumDec)
		Response.Write "{S}"
		If Session("PayCart") Then Response.Write FormatNumber(CDbl(rs(5)), SumDec)
		Response.Write "{S}"
		If myApp.MainCur <> Value Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetCartCurrRate" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = Session("RetVal")
			set rs = cmd.execute()
			Response.Write FormatNumber(CDbl(rs("Rate")), myApp.RateDec)
		Else
			Response.Write FormatNumber(0, myApp.RateDec)
		End If

	Case "LineTotal"
		Response.Write FormatNumber(CDbl(getNumericOut(Value)), myApp.PercentDec)
		If LineType = "E" Then
			WriteTaxAndTotal
		Else
			Response.Write "ok"	
		End If
	Case "ObjectCode", "CheckSumInv", "CheckLinesInv", "IsLocked", "GroupNum"
		Response.Write strResp
	Case Else
		Response.Write "ok"	
End Select

Sub WriteTaxAndTotal
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		If Session("PayCart") Then cmd("@MC") = "Y"
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		
		Response.Write "{S}" & FormatNumber(CDbl(rs("Tax")), SumDec) & "{S}" & FormatNumber(CDbl(rs("DocTotal")), SumDec)
		If Session("PayCart") Then Response.Write "{S}" & FormatNumber(CDbl(rs(5)), SumDec)
End Sub

Function GetRateFunc(ByVal i)
	Select Case DirectRate
		Case "Y"
			Select Case i
				Case 1
					GetRateFunc = "*"
				Case 2
					GetRateFunc = "/"
			End Select
		Case "N"
			Select Case i
				Case 1
					GetRateFunc = "/"
				Case 2
					GetRateFunc = "*"
			End Select
	End Select
End Function %>