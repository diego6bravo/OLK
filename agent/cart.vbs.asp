<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="chkLogin.asp" -->
<!--#include file="lang/cart.vbs.asp" -->
<!--#include file="myHTMLEncode.asp"-->

function fNumber(Value, Dec)
	fNumber = FormatNumber(Value,Dec)
end function 


Sub chkListNum(groupNum, Index)
	Dim value
	value = ListNum(Index)
	
	doProc "GroupNum", "N", groupNum

	If value <> PList Then
		document.frmCart.NewPList.value = value
		If Confirm("<%=getcartvbsLngStr("LtxtApplyPListByPTerm")%>") Then
			doApplyPriceList value
			PList = value
		End If
	End If
End Sub

sub setDocDueDate(i)
	If cartObject = 13 Then
		If i > 0 Then j = i else j = 0
		varDocDueDate = DateAdd("d", ExtraDays(j), DocDateValue)
		SetDueDateValue varDocDueDate 
	End If
end sub

Function DocDateValue
	strDate = document.frmCart.DocDate.value
	DocDateValue = Mid(strDate,<%=InStr(myApp.DateFormat, "MM")%>, 2)+"/"+Mid(strDate,<%=InStr(myApp.DateFormat, "dd")%>,2)+"/"+Mid(strDate, <%=InStr(myApp.DateFormat, "yyyy")%>,4)
End Function

Sub SetDueDateValue(value)
	strDueDate = "<%=myApp.DateFormat%>"
	strDueDate = Replace(strDueDate, "yyyy", Year(value))
	strDueDate = Replace(strDueDate, "MM", Right("0" & Month(value), 2))
	strDueDate = Replace(strDueDate, "dd", Right("0" & Day(value), 2))
	document.frmCart.DocDueDate.value = strDueDate
End Sub

sub setOQUTDueDate()
	varDocDueDate = DateAdd("m", 1, DocDateValue)
	SetDueDateValue varDocDueDate 
end sub

sub chkThis(Field, FType, EditType, FSize)
Select Case FType
	Case "A"
		If Len(Field.value) > FSize Then
			Alert(Replace("<%=getcartvbsLngStr("DtxtValFldMaxChar")%>", "{0}", FSize))
			Field.value = Left(Field.value,FSize)
		End If
	Case "N"
		Select Case Trim(EditType)
			Case ""
				If Not IsNumeric(getNumericVB(Field.value)) and Field.value <> "" Then
					Field.value = ""
					Alert("<%=getcartvbsLngStr("DtxtValNumVal")%>")
				ElseIf CDbl(getNumericVB(Field.Value))-CInt(getNumericVB(Field.Value)) <> 0 Then
					Field.value = ""
					Alert("<%=getcartvbsLngStr("DtxtValNumValWhole")%>")
				End If
			Case "T"
		End Select
	Case "B"
		If Not IsNumeric(getNumericVB(Field.value)) and Field.value <> "" Then
			Field.value = ""
			Alert("<%=getcartvbsLngStr("DtxtValNumVal")%>")
		Else
			Select Case Trim(EditType)
				Case "R"
					Field.Value = FormatNumber(getNumericVB(Field.Value), RateDec)
				Case "S"
					Field.Value = FormatNumber(getNumericVB(Field.Value), SumDec)
				Case "P"
					Field.Value = FormatNumber(getNumericVB(Field.Value), PriceDec)
				Case "Q"
					Field.Value = FormatNumber(getNumericVB(Field.Value), myApp.QtyDec)
				Case "%"
					Field.Value = FormatNumber(getNumericVB(Field.Value), PercentDec)
				Case "M"
					Field.Value = FormatNumber(getNumericVB(Field.Value), MeasureDec)
			End Select
		End If
End Select
end sub
