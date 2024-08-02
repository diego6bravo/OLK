<!--#include file="chkLogin.asp"-->
<!--#include file="lang/verfyQuery.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="getColType.inc"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
set rt = Server.CreateObject("ADODB.RecordSet")

on error resume next

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKUpdateCheckDB" & Session("ID")
cmd.Parameters.Refresh()
cmd("@Type") = Request("Type")
cmd("@TypeID") = Request("TypeID")
cmd("@ID") = Request("ID")
cmd("@IsValid") = "Y"

Select Case CInt(Request("Type"))
	Case 0
		sqlTest = "declare @SlpCode int declare @LanID int select ("
		
		sql = "select Query from OLKInformer where Type = 'U' and ID = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
		
		sqlTest = sqlTest & " )"
		
		set rs = conn.execute(sqlTest)
		
		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If
	Case 1
		sqlTest = "declare @LanID int set @LanID = -1 declare @branch int set @branch = -1 select ("
		
		sql = "select Query from OLKDocAddHdr where LineIndex = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
		
		sqlTest = sqlTest & " ) from OADM T0"
		
		set rs = conn.execute(sqlTest)
		
		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If
	Case 2
		sqlTest = "declare @CardCode nvarchar(15) set @CardCode = '' " & _
					"declare @SlpCode int set @SlpCode = -1 " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @LanID int set @LanID = -1 " & _
					"select ("

		
		sql = "select rowField from olkCardRep where RowIndex = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
		
		sqlTest = sqlTest & " ) from OCRD where CardCode = 'X'"
		
		set rs = conn.execute(sqlTest)
		
		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If
	Case 3
		sqlTest = "declare @PriceList int set @PriceList = 1 " & _
			"declare @CardCode nvarchar(15) set @CardCode = '' " & _
			"declare @SlpCode int set @SlpCode = -1 " & _
			"declare @dbName nvarchar(100) set @dbName = db_name() " & _
			"declare @WhsCode nvarchar(8) set @WhsCode = '1' " & _
			"declare @ItemCode nvarchar(20) set @ItemCode = '' " & _
			"declare @LanID int set @LanID = -1 " & _
			"declare @Quantity numeric(19,6) set @Quantity = 1 " & _
			"declare @Unit int set @Unit = 2 " & _
			"declare @Price numeric(19,6) set @Price = 1 " & _
			"select ("

		
		sql = "select rowField from olkItemRep where RowIndex = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
		
		sqlTest = sqlTest & " ) from OITM where ItemCode = 'X'"
		
		set rs = conn.execute(sqlTest)
		
		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If
	Case 4
		sqlTest = "declare @LanID int set @LanID = -1 select ("

		
		sql = "select rowField from olkBatchRep where RowIndex = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
		
		sqlTest = sqlTest & ") from OIBT " & _
			"left outer join R3_ObsCommon..DOC2 T1 on T1.LogNum = -1 and T1.LineNum = -1 and T1.BatchNum = OIBT.BatchNum collate database_default " & _
			"where OIBT.ItemCode = 'X' and OIBT.BatchNum = 'X'"
			
		sqlTest = Replace(sqlTest, "@ItemCode", "OIBT.ItemCode")
		sqlTest = Replace(sqlTest, "@BatchNum", "OIBT.BatchNum")
		sqlTest = Replace(sqlTest, "@WhsCode", "OIBT.WhsCode")
		
		set rs = conn.execute(sqlTest)
		
		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If
	Case 5

		sqlTest = "declare @LanID int set @LanID = -1 declare @LogNum int set @LogNum = -1 " & _
				"declare @dbName nvarchar(100) set @dbName = '' declare @branch int set @branch = -1 declare @SlpCode int set @SlpCode = -1 "
		
		If Request("TypeID") <> "OITM" Then
			sqlTest = sqlTest & "declare @CardCode nvarchar(15) set @CardCode = '' "
		End If
		
		If Request("TypeID") = "OINV" or Request("TypeID") = "INV1" Then
			sqlTest = sqlTest & "declare @PriceList int set @PriceList = -1 "
		End If
		
		If Request("TypeID") = "INV1" Then
			sqlTest = sqlTest & "declare @ItemCode nvarchar(15) set @ItemCode = '' declare @WhsCode nvarchar(8) set @WhsCode = '' "
		End If
		
		sql = "select SqlQuery from OLKCUFD where TableID = '" & Request("TypeID") & "' and FieldID = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
		
		set rs = conn.execute(sqlTest)

		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If

	Case 6, 7, 8
		sqlTest = "declare @PriceList int set @PriceList = 1 " & _
			"declare @CardCode nvarchar(15) set @CardCode = '' " & _
			"declare @LanID int set @LanID = -1 " & _
			"declare @SlpCode int set @SlpCode = -1 " & _
			"declare @DocNum int select ("
		
		Select Case CInt(Request("Type"))
			Case 6
				tblType = "T"
			Case 7
				tblType = "C"
			Case 8
				tblType = "L"
		End Select
			
		sql = "select ColQuery from OLK" & tblType & "Cart where LineIndex = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))
	
		sqlTest = sqlTest & ") from OITM inner join OITW on OITW.ItemCode = OITM.ItemCode and WhsCode = '01' where OITM.ItemCode = 'X' "
		sqlTest = Replace(sqlTest, "@ItemCode", "OITM.ItemCode")
		sqlTest = Replace(sqlTest, "@Table", "INV")

		set rs = conn.execute(sqlTest)

		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If


	Case 9
		sqlTestU = "declare @LogNum int set @LogNum = 1 declare @CardCode nvarchar(15) set @CardCode = '' declare @LanID int set @LanID = -1 select ("
		sqlTestS = "declare @DocEntry int set @DocEntry = 1 declare @CardCode nvarchar(15) set @CardCode = '' declare @LanID int set @LanID = -1 select ("
		
		sql = "select RowQuery, SystemQuery from OLKCMREP where RowType = '" & Request("TypeID") & "' and LineIndex = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTestU = sqlTestU & rs(0) & ")"
		sqlTestS = sqlTestS & rs(1) & ")"
		
		If Not IsNull(rs(0)) Then
			sqlTestU = QueryFunctions(sqlTestU)
			
			set rt = conn.execute(sqlTestU)
			
			If err.Number = 0 Then
				Response.Write "ok{Q}"
			Else
				Response.Write "err{S}" & err.Number & "{S}" & err.description & "{Q}"
				cmd("@IsValid") = "N"
				cmd("@ErrCode") = err.Number
				cmd("@ErrMessage") = err.description
			End If
		Else
			Response.Write "ok{Q}"
		End If
		
		If Not IsNull(rs(1)) Then
			sqlTestS = QueryFunctions(sqlTestS)
			sqlTestS = Replace(sqlTestS, "{Table}", "INV")
			set rt = conn.execute(sqlTestS)
			
			If err.Number = 0 Then
				Response.Write "ok"
			Else
				Response.Write "err{S}" & err.Number & "{S}" & err.description
				cmd("@IsValid") = "N"
				cmd("@ErrCode") = err.Number
				cmd("@ErrMessage") = err.description
			End If
		Else
			Response.Write "ok"
		End If
	Case 10
		sqlTest = "select ("

		sql = "select Query from OLKObjConfCols where TypeID = '" & Request("TypeID") & "' and [ID] = " & Request("ID")
		set rs = conn.execute(sql)
		
		sqlTest = sqlTest & QueryFunctions(rs(0))

		sqlTest = sqlTest & ") [CUSTCOL] "
			
		fld = "LogNum"
		If Request("TypeID") = "A" Then fld = "ID"
		sqlTest = Replace(sqlTest, "@" & fld, 1)
		
		Select Case Request("TypeID")
			Case "C"
				sqlTest = sqlTest & " from R3_ObsCommon..TCRD where LogNum = -1"
			Case "I"
				sqlTest = sqlTest & " from R3_ObsCommon..TITM where LogNum = -1"
			Case "D"
				sqlTest = sqlTest & " from R3_ObsCommon..TDOC where LogNum = -1"
			Case "R"
				sqlTest = sqlTest & " from R3_ObsCommon..TPMT where LogNum = -1"
		End Select
		
		set rs = conn.execute(sqlTest)

		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If
	Case 11
			sql = "select varIndex, varVar, varType, varDataType, varMaxChar, DefValBy, DefValDate, DefValValue, OLKCommon.dbo.DBOLKGetRSVarBaseIndex" & Session("ID") & "(rsIndex, varIndex) BaseIndex from OLKRSVars where rsIndex = " & Request("ID")
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			
			sqlTest = ""
			rs.Filter = "varType <> 'CL'"
			do while not rs.eof
				myVar = "@" & rs("varVar")
				If rs("varDataType") = "nvarchar" Then
					sqlTest = sqlTest & "declare " & myVar & " nvarchar(" & rs("varMaxChar") & ") "
				Else
					sqlTest = sqlTest & "declare " & myVar & " " & rs("varDataType") & " "
				End If

				sqlTest = sqlTest & "set " & myVar & " = "
				Select Case rs("DefValBy")
					Case "V"
						Select Case rs("varDataType")
							Case "int", "numeric"
								sqlTest = sqlTest & rs("DefValValue") & " "
							Case "datetime"
								sqlTest = sqlTest & "Convert(datetime,'" & SaveSqlDate(FormatDate(rs("DefValDate"), False)) & "',120) "
							Case "nvarchar"
								sqlTest = sqlTest & "N'" & rs("DefValValue") & "' "
						End Select
					Case "Q"
						set rsVal = Server.CreateObject("ADODB.RecordSet")
						sqlVal = getRSVariables(rs("BaseIndex")) & " " & rs("DefValValue")
						set rsVal = conn.execute(sqlVal)
						Select Case rs("varDataType")
							Case "int", "numeric"
								sqlTest = sqlTest & rsVal(0) & " "
							Case "datetime"
								sqlTest = sqlTest & "Convert(datetime,'" & SaveSqlDate(FormatDate(rsVal(0), False)) & "',120) "
							Case "nvarchar"
								sqlTest = sqlTest & "N'" & rsVal(0) & "' "
						End Select
					Case Else
						Select Case rs("varDataType")
							Case "nvarchar"
								sqlTest = sqlTest & "'' "
							Case "datetime"
								sqlTest = sqlTest & "'01/01/01' "
							Case "numeric", "int"
								sqlTest = sqlTest & "0 "
						End Select
				End Select
			rs.movenext
			loop
			If Request("TypeID") = "C" Then
				sqlTest = sqlTest & " declare @CardCode nvarchar(15) set @CardCode = '' "
			ElseIf Request("TypeID") = "V" Then
				sqlTest = sqlTest & " declare @SlpCode int set @SlpCode = -1 "
			End If
			sqlTest = sqlTest & " declare @LanID int set @LanID = -1 "
			
			set rd = Server.CreateObject("ADODB.RecordSet")
			sql = "select rsQuery, RSTop from OLKRS where rsIndex = " & Request("ID")
			set rd = conn.execute(sql)
			RSTop = rd("RSTop") = "Y"
			
			sqlTest = sqlTest & rd(0)
			
			If RSTop Then
				sqlTest = Replace(sqlTest, "@top", 1)
			End If
			rs.Filter = "varType = 'CL'"
			do while not rs.eof
				sqlTest = Replace(sqlTest, "@" & rs("varVar"), "'1', '2'")
			rs.movenext
			loop

			set rs = conn.execute(QueryFunctions(sqlTest))
	
			If err.Number = 0 Then
				Response.Write "ok"
			Else
				Response.Write "err{S}" & err.Number & "{S}" & err.description
				cmd("@IsValid") = "N"
				cmd("@ErrCode") = err.Number
				cmd("@ErrMessage") = err.description
			End If
	Case 12
		cmd("@ID2") = Request("ID2")
		set rd = Server.CreateObject("ADODB.RecordSet")
		sql = "select varQuery, OLKCommon.dbo.DBOLKGetRSVarBaseIndex" & Session("ID") & "(rsIndex, varIndex) BaseIndex from OLKRSVars where rsIndex = " & Request("ID") & " and varIndex = " & Request("ID2")
		set rd = conn.execute(sql)
		
		sqlTest = getRSVariables(rd("BaseIndex"))
		sqlTest = sqlTest & " declare @LanID int set @LanID = -1 "
	
		
		sqlTest = sqlTest & rd(0)
		
		set rs = conn.execute(QueryFunctions(sqlTest))

		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If

	Case 13, 14, 15

		ExecAt = Request("TypeID")
		sqlTest = "declare @SlpCode int set @SlpCode = -1 " & _
		"declare @dbName nvarchar(100) set @dbName = db_name() " & _
		"declare @branch int set @branch = -1 " & _
		"declare @LanID int set @LanID = -1 "
		
		If ExecAt <> "D1" and ExecAt <> "R1" Then sqlTest = sqlTest & "declare @LogNum int set @LogNum = -1 "
		If ExecAt = "O2" or ExecAt = "O3" or ExecAt = "O4" Then sqlTest = sqlTest & "declare @ObjectCode int "
		If Left(ExecAt, 1) = "O" Then sqlTest = sqlTest & "declare @Entry int "
		
		If Left(ExecAt,1) = "D" or Left(ExecAt,1) = "R" or ExecAt = "C2" or ExecAt = "C3" Then sqlTest = sqlTest & "declare @CardCode nvarchar(15) set @CardCode = '' "
		If ExecAt = "D2" Then 
			sqlTest = sqlTest & "declare @ItemCode nvarchar(15) set @ItemCode = '' " & _
						"declare @WhsCode nvarchar(8) set @WhsCode = '' " & _
						"declare @Quantity numeric(19,6) set @Quantity = 0 " & _
						"declare @Unit smallint declare @Price numeric(19,5) set @Price = 0 "
		End If	
		
		Select Case CInt(Request("Type"))
			Case 13
				fldQuery = "Query"
			Case 14
				fldQuery = "NoteQuery"
			Case 15
				fldQuery = "LineQuery"
		End Select
		set rd = Server.CreateObject("ADODB.RecordSet")
		sql = "select " & fldQuery & " from OLKUAF where FlowID = " & Request("ID")
		set rd = conn.execute(sql)
		sqlTest = sqlTest & rd(0)
		
		sqlTest = QueryFunctions(sqlTest)
		set rs = conn.execute(sqlTest)

		If err.Number = 0 Then
			Response.Write "ok"
		Else
			Response.Write "err{S}" & err.Number & "{S}" & err.description
			cmd("@IsValid") = "N"
			cmd("@ErrCode") = err.Number
			cmd("@ErrMessage") = err.description
		End If


		
		
	Case Else
		Response.Write "err{S}-1{S}Not Checked"
End Select

cmd.execute()

Function getRSVariables(ByVal baseIndex)
	strRSVariables = ""
	Select Case Request("TypeID") 
		Case "C"
			strRSVariables = "declare @CardCode nvarchar(15) set @CardCode = '' "
		Case "V"
			strRSVariables = "declare @SlpCode int set @SlpCode = -1 "
	End Select
	If baseIndex <> "-1" Then %>
	<!--#include file="repVars.inc"-->
<%	
		sql2 = "select '@' + varVar varVar, varDataType, varMaxChar, DefValBy, DefValDate, DefValValue, OLKCommon.dbo.DBOLKGetRSVarBaseIndex" & Session("ID") & "(rsIndex, varIndex) BaseIndex from OLKRSVars where rsIndex = " & Request("ID") & " and varIndex in (" & baseIndex & ")"
		set rs = conn.execute(sql2)
		do while not rs.eof
			If rs("varDataType") = "nvarchar" Then 
				MaxVar = "(" & rs("varMaxChar") & ")"
			ElseIf rs("varDataType") = "numeric" Then
				MaxVar = "(19,6)"
			Else
				MaxVar = ""
			End If
			strRSVariables = strRSVariables & "declare " & rs("varVar") & " " & rs("varDataType") & " " & MaxChar & " "
			
			strRSVariables = strRSVariables & "set " & rs("varVar") & " = "
			Select Case rs("DefValBy")
				Case "V"
					Select Case rs("varDataType")
						Case "int", "numeric"
							strRSVariables = strRSVariables & rs("DefValValue") & " "
						Case "datetime"
							strRSVariables = strRSVariables & "Convert(datetime,'" & SaveSqlDate(FormatDate(rs("DefValDate"), False)) & "',120) "
						Case "nvarchar"
							strRSVariables = strRSVariables & "N'" & rs("DefValValue") & "' "
					End Select
				Case "Q"
					set rsVal = Server.CreateObject("ADODB.RecordSet")
					sqlVal = getRSVariables(rs("BaseIndex")) & " " & rs("DefValValue")
					set rsVal = conn.execute(sqlVal)
					Select Case rs("varDataType")
						Case "int", "numeric"
							strRSVariables = strRSVariables & rsVal(0) & " "
						Case "datetime"
							strRSVariables = strRSVariables & "Convert(datetime,'" & SaveSqlDate(FormatDate(rsVal(0), False)) & "',120) "
						Case "nvarchar"
							strRSVariables = strRSVariables & "N'" & rsVal(0) & "' "
					End Select
				Case Else
					Select Case rs("varDataType")
						Case "nvarchar"
							strRSVariables = strRSVariables & "'' "
						Case "datetime"
							strRSVariables = strRSVariables & "'01/01/01' "
						Case "numeric", "int"
							strRSVariables = strRSVariables & "0 "
					End Select
			End Select
		rs.movenext
		loop
	End If
	getRSVariables = strRSVariables
End Function

%>
