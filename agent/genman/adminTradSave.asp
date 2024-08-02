<%
Sub SaveNewTrad(TradValues, Table, ColumnID, ColumnName, ID)
	sql = "declare @LanID int "
	arrTrad = Split(TradValues, "{/}")
	
	arrCol = Split(ColumnID, ",")
	arrVal = Split(ID, ",")

	For j = 0 to UBound(arrTrad)
		LanID = Split(arrTrad(j), "{=}")(0)
		Value = Split(arrTrad(j), "{=}")(1)
		If Value <> "" Then
			sql = sql & "set @LanID = " & LanID & " "
			'sql = sql & "insert OLK" & Table & "AlterNames(LanID, " & ColumnID & ", " & ColumnName & ") " & _
			'			"values(" & LanID & ", " & ID & ", N'" & saveHTMLDecode(Value, False) & "') "
			
			sql = sql & "if not exists(select 'A' from OLK" & Table & "AlterNames where LanID = @LanID "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
				Else
					sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
				End If
			Next

			sql = sql & ") begin " & _
			"	insert OLK" & Table & "AlterNames(LanID, "
			
			For i = 0 to UBound(arrCol)
				sql = sql & arrCol(i) & ", "
			Next
			
			sql = sql & ColumnName & ") " & _
			"	values(@LanID, "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & arrVal(i) & ", "
				Else
					sql = sql & "N'" & arrVal(i) & "', "
				End If
			Next
			
			sql = sql & "N'" & saveHTMLDecode(Value, False) & "') " & _
			"end else begin " & _
			"	update OLK" & Table & "AlterNames set " & ColumnName & " = N'" & saveHTMLDecode(Value, False) & "' " & _
			"	where LanID = @LanID "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
				Else
					sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
				End If
			Next

			sql = sql & "end "

		End If
	Next
	If sql <> "" Then conn.execute(sql)
End Sub

%>
