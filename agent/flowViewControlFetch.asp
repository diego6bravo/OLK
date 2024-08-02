<!--#include file="myHTMLEncode.asp"--><!--#include file="lcidReturn.inc"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

ID = CInt(Request.Form("ID"))
myType = Request.Form("Type")
Select Case myType
	Case "L"
		sql = 	"select T1.ID, T1.RequestDate, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(1, 'OSLP', 'SlpName', T2.SlpCode, T2.SlpName) RequestUserSign, T1.ConfirmDate, " & _
				"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(1, 'OSLP', 'SlpName', T3.SlpCode, T3.SlpName) ConfirmUserSign, T1.Note, T1.Status " & _  
				"from OLKUAFControl T0 " & _  
				"inner join OLKUAFControl T1 on T1.ExecAt = T0.ExecAt and IsNull(T1.ObjectCode, -1) = IsNull(T1.ObjectCode, -1) and T1.ObjectEntry = T0.ObjectEntry " & _  
				"inner join OSLP T2 on T2.SlpCode = T1.RequestUserSign " & _
				"left outer join OSLP T3 on T3.SlpCode = T1.ConfirmUserSign " & _
				"where T0.ID = " & ID & " " & _  
				"order by T1.RequestDate desc " 
		set rs = conn.execute(sql)
		strRetVal = ""
		do while not rs.eof
			If strRetVal <> "" Then strRetVal = strRetVal & "{L}"
			
			strDate = FormatDate(rs(1), True) & "&nbsp;"
			strTime = FormatTime(rs(1))
			strDate = strDate & strTime 'Left(strTime, Len(strTime)-8) & Right(strTime, 5)
			
			If Not IsNull(rs(3)) Then
				strConfDate = FormatDate(rs(3), True) & "&nbsp;"
				strConfTime = FormatTime(rs(1))
				strConfDate = strConfDate & strConfTime 'Left(strConfTime, Len(strConfTime)-8) & Right(strConfTime, 5)
			End If
			strRetVal = strRetVal & rs(0) & "{S}" & strDate & "{S}" & rs(2) & "{S}" & strConfDate & "{S}" & rs(4) & "{S}" & rs(5) & "{S}" & rs(6)
		rs.movenext
		loop
		Response.Write strRetVal
	Case "D"
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rd = Server.CreateObject("ADODB.RecordSet")
		sql = "select ExecAt, ObjectCode, ObjectEntry from OLKUAFControl where ID = " & ID
		set rd = conn.execute(sql)
		execAt = rd("ExecAt")
		objCode = rd("ObjectCode")
		objEntry = rd("ObjectEntry")
		
		sql = 	"select T1.FlowID, IsNull(T2.AlterName, T1.Name) Name, NoteBuilder, NoteQuery, NoteText, LineQuery, T0.Note " & _  
				"from OLKUAFControl1 T0 " & _  
				"inner join OLKUAF T1 on T1.FlowID = T0.FlowID " & _  
				"left outer join OLKUAFAlterNames T2 on T2.FlowID = T1.FlowID and T2.LanID = " & Session("LanID") & " " & _  
				"where T0.ID = " & ID & " " & _  
				"order by [Order] " 
		set rs = conn.execute(sql)
		strRetVal = ""
		
		sqlBase = 	"declare @LanID int set @LanID = " & Session("LanID") & " " & _
					"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @branch int set @branch = " & Session("branch") & " "
							
		Select Case execAt
			Case "O0", "O1" ' Aprove Sales Order, Convert Quotation to Sales Order
				sqlBase = sqlBase & "declare @Entry int set @Entry = " & objEntry & " "
			Case "O2", "O3", "O4" ' Close  Object, Cancel Object, Remove Object
				sqlBase = sqlBase & "declare @ObjectCode int set @ObjectCode = " & objCode & " " & _
									"declare @Entry int set @Entry = " & objEntry & " "
			Case "C1", "A1", "R2", "D3"
				LogNum = objEntry
				sqlBase = sqlBase & "declare @LogNum int set @LogNum = " & LogNum & " "
		End Select
		
		do while not rs.eof
			If strRetVal <> "" Then strRetVal = strRetVal & "{F}"
			Note = rs("NoteText")
			If rs("NoteBuilder") = "Y" Then
				NoteQry = sqlBase & rs("NoteQuery")
				NoteQry = QueryFunctions(NoteQry)
				set rGen = conn.execute(NoteQry)
				If Not rGen.Eof Then
					For each fld in rGen.Fields
						Note = Replace(Note, "{" & fld.Name & "}", fld)
					Next
				End If
			End If
			Note = Replace(Note, VbCrLf, "<br>")
			
			strTable = ""
			If Not IsNull(rs("LineQuery")) Then
				LineQry = sqlBase & rs("LineQuery")
				LineQry = QueryFunctions(LineQry)
				set rGen = conn.execute(LineQry)
				
				If Not rGen.Eof Then
					strLine = ""
					For each fld in rGen.Fields
						If strLine <> "" Then strLine = strLine & "{C}"
						strLine = strLine & fld.Name
					Next
					strLine = strLine & "{H}"
					strTable = strLine
					do while not rGen.Eof
						strLine = ""
						For each fld in rGen.Fields
							If strLine <> "" Then strLine = strLine & "{C}"
							strLine = strLine & fld
						Next
						strLine = strLine & "{R}"
					rGen.movenext
					loop
				End If
			End If
			
			strRetVal = strRetVal & rs("Name") & "{S}" & strTable & "{S}" & Note & "{S}" & rs("Note")
		rs.movenext
		loop
		
		Response.Write strRetVal
End Select
%>