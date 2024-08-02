<%
Function GetFormatSep()
	GetFormatSep = Mid(FormatNumber(1000, 2),2,1)
End Function

Function GetFormatDec()
	GetFormatDec = Mid(FormatNumber(1000, 2),6,1)
End Function

Function ClearFormatNumber(ByVal Value, ByVal Symbol)
	strClearFormatNumber = Value
	strClearFormatNumber = Replace(strClearFormatNumber, Symbol, "")
	strClearFormatNumber = Right(strClearFormatNumber, Len(strClearFormatNumber)-1)
	strClearFormatNumber = getNumeric(strClearFormatNumber)
	ClearFormatNumber = strClearFormatNumber
End Function

Function getNumeric(value)
	strGetNumeric = CStr(value)
	strGetNumeric = Replace(strGetNumeric, GetFormatSep, "")
	strGetNumeric = Replace(strGetNumeric, GetFormatDec, ".")
	getNumeric = strGetNumeric
End Function

Function getNumericOut(value)
	strGetNumericOut = CStr(value)
	strGetNumericOut = Replace(strGetNumericOut, GetFormatSep, "")
	getNumericOut = strGetNumericOut
End Function

%>