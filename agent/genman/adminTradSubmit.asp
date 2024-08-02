<html>
<body>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<%

Table = Request("Table")
arrCol = Split(Request("ColumnID"), ",")
arrVal = Split(Request("ID"), ",")

ColumnName = Request("ColumnName")

sql = "declare @LanID int "
For j = 0 to UBound(myLanIndex)
	LanID = myLanIndex(j)(4)
	sql = sql & "set @LanID = " & LanID & " "
	If Request("txt" & LanID) <> "" Then
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
		
		sql = sql & "N'" & saveHTMLDecode(Request("txt" & LanID), False) & "') " & _
		"end else begin " & _
		"	update OLK" & Table & "AlterNames set " & ColumnName & " = N'" & saveHTMLDecode(Request("txt" & LanID), False) & "' " & _
		"	where LanID = @LanID "
		
		For i = 0 to UBound(arrCol)
			If IsNumeric(arrVal(i)) Then
				sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
			Else
				sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
			End If
		Next

		sql = sql & "end "
	Else
		sql = sql & "	update OLK" & Table & "AlterNames set " & ColumnName & " = NULL " & _
		"	where LanID = @LanID "
		
		For i = 0 to UBound(arrCol)
			If IsNumeric(arrVal(i)) Then
				sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
			Else
				sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
			End If
		Next
	End If
Next

conn.execute(sql)
%>
<script type="text/javascript">window.close();</script>
</body>
</html>
