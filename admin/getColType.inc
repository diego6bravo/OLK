<%
Function getColType(ByVal ColType)
Select Case ColType
Case 0
	getColType = "adEmpty"
Case 16
	getColType = "adTinyInt"
Case 2
	getColType = "adSmallInt"
Case 3
	getColType = "adInteger"
Case 20
	getColType = "adBigInt"
Case 17
	getColType = "adUnsignedTinyInt"
Case 18
	getColType = "adUnsignedSmallInt"
Case 19
	getColType = "adUnsignedInt"
Case 21
	getColType = "adUnsignedBigInt"
Case 4
	getColType = "adSingle"
Case 5
	getColType = "adDouble"
Case 6
	getColType = "adCurrency"
Case 14
	getColType = "adDecimal"
Case 131
	getColType = "adNumeric"
Case 11
	getColType = "adBoolean"
Case 10
	getColType = "adError"
Case 132
	getColType = "adUserDefined"
Case 12
	getColType = "adVariant"
Case 9
	getColType = "adIDispatch"
Case 13
	getColType = "adIUnknown"
Case 72
	getColType = "adGUID"
Case 7
	getColType = "adDate"
Case 133
	getColType = "adDBDate"
Case 134
	getColType = "adDBTime"
Case 135
	getColType = "adDBTimeStamp"
Case 8
	getColType = "adBSTR"
Case 129
	getColType = "adChar"
Case 200
	getColType = "adVarChar"
Case 201
	getColType = "adLongVarChar"
Case 130
	getColType = "adWChar"
Case 202
	getColType = "adVarWChar"
Case 203
	getColType = "adLongVarWChar"
Case 128
	getColType = "adBinary"
Case 204
	getColType = "adVarBinary"
Case 205
	getColType = "adLongVarBinary"
Case 136
	getColType = "adChapter"
Case 64
	getColType = "adFileTime"
Case 138
	getColType = "adPropVariant"
Case 139
	getColType = "adVarNumeric"
End Select
End Function

Function getColTypeVal(ByVal ColType)
If ColType = 129 or ColType = 200 or ColType = 201 or ColType = 130 or ColType = 202 or ColType = 203 Then
	getColTypeVal = "T"
ElseIf ColType = 16 or ColType = 2 or ColType = 3 or ColType = 20 or ColType = 17 or ColType = 18 or _
	ColType = 19 or ColType = 21 or ColType = 4 or ColType = 5 or ColType = 14 or ColType = 131 or ColType = 139 Then
	getColTypeVal = "N"
ElseIf ColType = 7 or ColType = 133 or ColType = 134 or ColType = 135 Then
	getColTypeVal = "D"
Else
	getColTypeVal = "U"
End If
End Function %>