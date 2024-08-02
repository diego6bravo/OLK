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

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<script language="javascript" src="general.js"></script>
</head>
<%
If Request("Query") <> "" Then
	For i = 0 to UBound(myLanIndex)
		myItm = myLanIndex(i)
		If myItm(4) = CStr(Session("LanID")) Then
			sql = "set language " & myItm(5)
			conn.execute(sql)
			Exit For
		End If
	Next
	sql = ""

	Select Case Request("Type")
		Case "AutoGenCode"
			Select Case CInt(Request("obj"))
				Case 2
					fld = "Card"
				Case 4
					fld = "Item"
			End Select
			
			sql = "declare @LogNum int set @LogNum = -1 declare @" & fld & "Code nvarchar(20) set @" & fld & "Code = ("
		Case "OpFilter"
			Operation = CInt(Request("Operation"))
			ObjectID = CLng(Request("ObjectID"))
			
			Select Case Operation 
				Case 6
					sql = "select T1, T2 from OLKDocConf where ObjectCode = " & ObjectID
					set rs = Server.CreateObject("ADODB.RecordSet")
					set rs = conn.execute(sql)
					sql = "select 1 from " & rs(0) & " inner join " & rs(1) & " on " & rs(1) & ".DocEntry = " & rs(0) & ".DocEntry " & _
							"where " & rs(0) & ".DocEntry = 1 and "
				Case Else
					sql = "declare @DocEntry int set @DocEntry = 1 select 1 where 1 in ("
			End Select
		Case "OpFld"
			sql = "select ("
		Case "objConfCols"
			sql = "select ("
		Case "docBD"
			sql = "declare @LogNum int declare @dbName nvarchar(100) declare @BreakTable table(LineNum int, Quantity numeric(19,6), ShipDate datetime, ShipDateDiff int) "
		Case "ItemRec"
			sql = "declare @LogNum int declare @CardCode nvarchar(20) declare @SlpCode int declare @branch int declare @WhsCode nvarchar(8) declare @LanID int " & _
					"select top 1 [Query].* " & _
					"from OITM X_____0 " & _
					"inner join ("
		Case "MenuGroupFormula"
			sql = "select top 1 '' from OITM "
			If Request("subType") = "D" Then
				sql = sql & "inner join [" & Request("TableID") & "] X0 on 1 = 1 "
			End If
			sql = sql & " where "
		Case "TaskMon"
			sql = "declare @SlpCode int declare @LanID int select ("
		Case "secSubRS"
			sql = "declare @CardCode nvarchar(15) set @CardCode = '' declare @LanID int set @LanID = -1 "
		Case "GetSeries"
			sql = "EXEC OLKCommon..DBOLKGetQuery" & Session("ID") & " 4, "
		Case "printTitle"
			sql = "declare @LanID int set @LanID = -1 declare @branch int set @branch = -1 select ("
		Case "ChecksFilter"
			sql = "declare @SlpCode int set @SlpCode = -1 declare @branch int set @branch = -1 select '' from OACT T0 where AcctCode not in (" 
		Case "CustomBarCode"
			Select Case Request("CodeBarsQryMethod")
				Case "R"
					sql = "select ("
				Case "I"
					sql = "select CodeBars from ("
			End Select
		Case "newQueryVar"
			sql = "declare @LanID int "
			Select Case Request("UserType")
				Case "C"
					sql = sql & "declare @CardCode nvarchar(15) "
				Case "V"
					sql = sql & "declare @SlpCode int "
			End Select
			sql = sql & "declare @"
		Case "CustSearchVar"
			sql = "declare @SlpCode int declare @branch int declare @"
		Case "editQuery"
			sql = "select varVar, varType, varDataType, varMaxChar, DefValBy, DefValDate, DefValValue from OLKRSVars where rsIndex = " & Request("rsIndex")
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			sql = ""
			rs.Filter = "varType <> 'CL'"
			do while not rs.eof
				myVar = "@" & rs("varVar")
				If rs("varDataType") = "nvarchar" Then
					sql = sql & "declare " & myVar & " nvarchar(" & rs("varMaxChar") & ") "
				Else
					sql = sql & "declare " & myVar & " " & rs("varDataType") & " "
				End If

				sql = sql & "set " & myVar & " = "
				Select Case rs("DefValBy")
					Case "V"
						Select Case rs("varDataType")
							Case "int", "numeric"
								sql = sql & rs("DefValValue") & " "
							Case "datetime"
								sql = sql & "Convert(datetime,'" & SaveSqlDate(FormatDate(rs("DefValDate"), False)) & "',120) "
							Case "nvarchar"
								sql = sql & "N'" & rs("DefValValue") & "' "
						End Select
					Case "Q"
						set rsVal = Server.CreateObject("ADODB.RecordSet")
						sqlVal = getRSVariables(Request("baseIndex")) & " " & rs("DefValValue")
						set rsVal = conn.execute(sqlVal)
						Select Case rs("varDataType")
							Case "int", "numeric"
								sql = sql & rsVal(0) & " "
							Case "datetime"
								sql = sql & "Convert(datetime,'" & SaveSqlDate(FormatDate(rsVal(0), False)) & "',120) "
							Case "nvarchar"
								sql = sql & "N'" & rsVal(0) & "' "
						End Select
					Case Else
						Select Case rs("varDataType")
							Case "nvarchar"
								sql = sql & "'' "
							Case "datetime"
								sql = sql & "'01/01/01' "
							Case "numeric", "int"
								sql = sql & "0 "
						End Select
				End Select
			rs.movenext
			loop
			If Request("UserType") = "C" Then
				sql = sql & " declare @CardCode nvarchar(15) set @CardCode = '' "
			ElseIf Request("UserType") = "V" Then
				sql = sql & " declare @SlpCode int set @SlpCode = -1 "
			End If
			sql = sql & " declare @LanID int set @LanID = -1 "
		Case "editCustomSearch"
			sql = ""
				
			Select Case CInt(Request("ObjID"))
				Case 2
					sql = sql & "select top 1 T0.CardCode " & _
							"from OCRD T0 " & _
							"left outer join OCRY T1 on T1.Code = T0.Country " & _
							"left outer join OCRG T2 on T2.GroupCode = T0.GroupCode where "
				Case 4
					sql = sql & "select top 1 OITM.ItemCode " & _
							"from OITM " & _
							"inner join OITW on OITW.ItemCode = OITM.ItemCode " & _
							"inner join OITB on OITB.ItmsGrpCod = OITM.ItmsGrpCod " & _
							"inner join OMRC on OMRC.FirmCode = OITM.FirmCode where "
			End Select
		Case "CustSearchQry"
			sql = "select Variable, [Type], DataType, MaxChar from OLKCustomSearchVars where ID = " & Request("ID")
			rs.open sql, conn, 3, 1
			sql = ""
			rs.Filter = "Type <> 'CL'"
			do while not rs.eof
				If rs("DataType") = "nvarchar" Then
					sql = sql & "declare @" & rs("Variable") & " nvarchar(" & rs("MaxChar") & ") "
				Else
					sql = sql & "declare @" & rs("Variable") & " " & rs("DataType") & " "
				End If
			rs.movenext
			loop
			sql = sql & " declare @LanID int set @LanID = -1 "
		Case "invOpt"
			sql = "declare @PriceList int set @PriceList = 1 " & _
			"declare @CardCode nvarchar(15) set @CardCode = '' " & _
			"declare @SlpCode int set @SlpCode = -1 " & _
			"declare @dbName nvarchar(100) set @dbName = db_name() " & _
			"declare @WhsCode nvarchar(8) set @WhsCode = '1' " & _
			"declare @ItemCode nvarchar(20) set @ItemCode = '' " & _
			"declare @LanID int set @LanID = -1 " & _
			"declare @Quantity numeric(19,6) set @Quantity = 1 " & _
			"declare @Unit int set @Unit = 2 " & _
			"declare @Price numeric(19,6) set @Price = 1 " & _
			"declare @branchIndex int set @branchIndex = -1 " & _
			"select ("
		Case "iPO"
			sql = "declare @SlpCode int set @SlpCode = -1 " & _
			"declare @dbName nvarchar(100) set @dbName = db_name() " & _
			"declare @WhsCode nvarchar(8) set @WhsCode = '1' " & _
			"declare @ItemCode nvarchar(20) set @ItemCode = '' " & _
			"declare @LanID int set @LanID = -1 " & _
			"select ("
		Case "cartOpt"
			sql = "declare @LogNum int set @LogNum = -1 " & _
			"declare @CardCode nvarchar(15) set @CardCode = '' " & _
			"declare @SlpCode int set @SlpCode = -1 " & _
			"declare @dbName nvarchar(100) set @dbName = db_name() " & _
			"declare @LanID int set @LanID = -1 " & _
			"select ("
		Case "crdOpt"
			sql = "" & _
					"declare @SlpCode int set @SlpCode = -1 " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @LanID int set @LanID = -1 " & _
					"select ("
		Case "batchOpt"
			sql = "declare @LanID int set @LanID = -1 select ("
		Case "catOpt"
			sql = "declare @PriceList int set @PriceList = 1 " & _
			"declare @CardCode nvarchar(15) set @CardCode = '' " & _
			"declare @LanID int set @LanID = -1 " & _
			"declare @SlpCode int set @SlpCode = -1 " & _
			"declare @DocNum int select ("
		Case "DocFlow"
			ExecAt = Request("ExecAt")
			sql = "declare @SlpCode int set @SlpCode = -1 " & _
			"declare @dbName nvarchar(100) set @dbName = db_name() " & _
			"declare @branch int set @branch = -1 " & _
			"declare @LanID int set @LanID = -1 "
			
			If ExecAt <> "D1" and ExecAt <> "R1" Then sql = sql & "declare @LogNum int set @LogNum = -1 "
			If ExecAt = "O2" or ExecAt = "O3" or ExecAt = "O4" Then sql = sql & "declare @ObjectCode int "
			If Left(ExecAt, 1) = "O" and Left(ExecAt, 2) <> "OP" Then sql = sql & "declare @Entry int "
			
			If Left(ExecAt,1) = "D" or Left(ExecAt,1) = "R" or ExecAt = "C2" or ExecAt = "C3" Then sql = sql & "declare @CardCode nvarchar(15) set @CardCode = '' "
			If ExecAt = "D2" Then 
				sql = sql & "declare @ItemCode nvarchar(15) set @ItemCode = '' " & _
							"declare @WhsCode nvarchar(8) set @WhsCode = '' " & _
							"declare @Quantity numeric(19,6) set @Quantity = 0 " & _
							"declare @Unit smallint declare @Price numeric(19,5) set @Price = 0 "
			End If	
			
			If Left(Request("by"), 8) = "GrpValue" Then
				sql = sql & "select SlpCode from OSLP where SlpCode in ("
			End If
		Case "MailMsg"
			sql = "declare @CardCode nvarchar(15) set @CardCode = '' " & _
					"declare @LanID int set @LanID = -1 "
		Case "minrep"
			sql = "declare @LogNum int set @LogNum = 1 declare @CardCode nvarchar(15) set @CardCode = '' declare @LanID int set @LanID = -1 select ("
		Case "minrepSys"
			sql = "declare @DocEntry int set @DocEntry = 1 declare @CardCode nvarchar(15) set @CardCode = '' declare @LanID int set @LanID = -1 select ("
		Case "GenFilter"
			sql = 	" " & _
					"select 'A' from OITM T____0  where T____0.ItemCode not in ("
		Case "ClientCatFilter", "NavQry"
			sql = 	"declare @LanID int set @LanID = -1 " & _
					"select 'A' from OITM T____0  where T____0.ItemCode not in ("
		Case "SearchTreeFilter"
			sql = 	"select 'A' from OITM T____0 where T____0.ItemCode not in ("
		Case "AnonCatFilter"
			sql = "select 'A' from OITM where ItemCode not in ("
		Case "AgentClientsFilter"
			sql = "declare @SlpCode int set @SlpCode = -1 declare @Type int set @Type = 1 select 'A' from OCRD where CardCode not in ("
		Case "ViewDocFilter"
			sql = "select top 1 'A' from OINV where DocEntry not in ("
		Case "newQuery"
			If Request("UserType") = "C" Then 
				sql = "declare @CardCode nvarchar(15) set @CardCode = ''  "
			ElseIf Request("UserType") = "V" Then
				sql = "declare @SlpCode int set @SlpCode = -1 "
			End IF
			sql = sql & "declare @LanID int set @LanID = -1 "
			
		Case "RSVar", "RSVarDef"
			sql = getRSVariables(Request("baseIndex"))
			sql = sql & " declare @LanID int set @LanID = -1 "
		Case "CustSearchVarQry", "CustVarDefVal"
			If Request("BaseID") <> "" Then 
				sql2 = "select '@' + Variable Variable, DataType, MaxChar from OLKCustomSearchVars where ID = " & Request("ID") & " and VarID in (" & Request("BaseID") & ")"
				set rs = conn.execute(sql2)
				do while not rs.eof
					If rs("DataType") = "nvarchar" Then 
						MaxVar = "(" & rs("MaxChar") & ")"
					ElseIf rs("DataType") = "numeric" Then
						MaxVar = "(19,6)"
					Else
						MaxVar = ""
					End If
					sql = sql & "declare " & rs("Variable") & " " & rs("DataType") & " " & MaxChar & " "
				rs.movenext
				loop
			End If
			sql = sql & " declare @LanID int set @LanID = -1  declare @branch int declare @SlpCode int set @SlpCode = -1 declare @CardCode nvarchar(20) "
		Case "Banner"
			sql = "declare @CardCode nvarchar(15) set @CardCode = N'' " & _
			"declare @LanID int set @LanID = -1 "
		Case "NavImgQry"
			sql = "declare @CardCode nvarchar(15) set @CardCode = N'' declare @SlpCode int set @SlpCode = -1 "
		Case "CUFD"
			TableID = Request("TableID")
			FieldID = CInt(Request("FieldID"))
			
			sql = "declare @LanID int set @LanID = -1 declare @LogNum int set @LogNum = -1 " & _
					"declare @dbName nvarchar(100) set @dbName = '' declare @branch int set @branch = -1 declare @SlpCode int set @SlpCode = -1 "
			
			If TableID <> "OITM" Then
				sql = sql & "declare @CardCode nvarchar(15) set @CardCode = '' "
			End If
			
			If TableID = "OINV" or TableID = "INV1" Then
				sql = sql & "declare @PriceList int set @PriceList = -1 "
			End If
			
			If TableID = "INV1" Then
				sql = sql & "declare @LineNum int declare @ItemCode nvarchar(15) set @ItemCode = '' declare @WhsCode nvarchar(8) set @WhsCode = '' "
			End If
			
			
			If TableID = "CRD1" and FieldID = -2 and Request("Query") <> "" Then
				sql = sql & " select top 1 Code, Name from OCRY where "
			ElseIf TableID = "INV1" and (FieldID = -20 or FieldID >= -7 and FieldID <= -3) and Request("Query") <> "" Then
				sql = sql & " select top 1 OcrCode, OcrName from OOCR where "
			End If
	End Select
	
	sql = sql & Request("Query")

	Select Case Request("Type")
		Case "OpFld"
			TypeID = CInt(Request("TypeID"))
			set rs = Server.CreateObject("ADODB.RecordSet")
			Operation = CInt(Request("Operation"))
			ObjectID = CLng(Request("ObjectID"))
			
			strQry = "select T1, T2 from OLKDocConf where ObjectCode = " & ObjectID
			set rs = conn.execute(strQry)
			T1 = rs("T1")
			T2 = rs("T2")
			
			Select Case Operation
				Case 6
					sql = sql & ") from [" & T1 & "] inner join [" & T2 & "] on [" & T2 & "].DocEntry = [" & T1 & "].DocEntry  where [" & T1 & "].DocEntry = 1 and LineNum = 1"
				Case Else
					Select Case ObjectID
						Case 2
						Case 4
						Case Else
							Select Case TypeID
								Case 0, 2
									sql = sql & ") from [" & T1 & "] where DocEntry = 1"
								Case 1
									sql = sql & ") from [" & T1 & "] inner join [" & T2 & "] on [" & T2 & "].DocEntry = [" & T1 & "].DocEntry  where [" & T1 & "].DocEntry = 1 and LineNum = 1"
							End Select
					End Select
			End Select
		Case "OpFilter"
			Select Case Operation 
				Case 6
					'sql = sql & ""
				Case Else
					sql = sql & ")"

			End Select
		Case "AutoGenCode"
			sql = sql & ")"
		Case "objConfCols"
			sql = sql & ") [CUSTCOL] "
			
			fld = "LogNum"
			If Request("TypeID") = "A" Then fld = "ID"
			sql = Replace(sql, "@" & fld, 1)
			
			Select Case Request("TypeID")
				Case "C"
					sql = sql & " from R3_ObsCommon..TCRD where LogNum = -1"
				Case "I"
					sql = sql & " from R3_ObsCommon..TITM where LogNum = -1"
				Case "D"
					sql = sql & " from R3_ObsCommon..TDOC where LogNum = -1"
				Case "R"
					sql = sql & " from R3_ObsCommon..TPMT where LogNum = -1"
				Case Else
					If Request("OpObj") <> "" Then
						Select Case CLng(Request("OpObj"))
							Case 2
								sql = sql & " from R3_ObsCommon..TCRD where LogNum = -1"
							Case 4
								sql = sql & " from R3_ObsCommon..TITM where LogNum = -1"
							Case 24
								sql = sql & " from R3_ObsCommon..TPMT where LogNum = -1"
							Case Else
								sql = sql & " from R3_ObsCommon..TDOC where LogNum = -1"
						End Select
					End If
			End Select
		Case "TaskMon"
			sql = sql & ") [ALIAS] "
		Case "ItemRec"
			sql = sql & ") [Query] on [Query].ItemCode = X_____0.ItemCode "
			sql = Replace(sql, "@ItemCode", "N'test'")
		Case "MenuGroupFormula"
			sql = sql & " = '1' "
		Case "printTitle"
			sql = sql & ") from OADM T0 "
		Case "ChecksFilter"
			sql = sql & ")"
		Case "CustomBarCode"
			Select Case Request("CodeBarsQryMethod")
				Case "R"
					sql = sql & ")"
				Case "I"
					sql = sql & ") X0 "
			End Select
				sql = Replace(sql, "@CodeBars", "N'0000000'")
		Case "newQueryVar", "CustSearchVar"
			sql = sql & " nvarchar(1) set @" & Request("Query") & " = N'test'"
		Case "invOpt"
			sql = sql & ") from OITM where ItemCode = 'X'"
		Case "iPO"
			sql = sql & ") from OITM where ItemCode = 'X'"
		Case "cartOpt"
			sql = sql & ") from R3_ObsCommon..DOC1 " & _
			"inner join OITM on OITM.ItemCode = DOC1.ItemCode collate database_default " & _
			"inner join OlkSalesLines T2 on T2.LogNum = DOC1.Lognum and T2.LineNum = DOC1.LineNum " & _
			"inner join R3_ObsCommon..TDOC on TDOC.LogNum = DOC1.LogNum " & _
			"where DOC1.LogNum = -1 "
			
			sql = Replace(sql, "@WhsCode", "DOC1.WhsCode")
			sql = Replace(sql, "@Quantity", "DOC1.Quantity*Case DOC1.UseBaseUn When 'N' Then OITM.NumInSale Else 1 End")
			sql = Replace(sql, "@Unit", "T2.SaleType")
			sql = Replace(sql, "@PriceList", 1)
			sql = Replace(sql, "@Price", "DOC1.Price")
			sql = Replace(sql, "@ItemCode", "OITM.ItemCode")
			
			
		Case "crdOpt"
			sql = sql & ") from OCRD where CardCode = 'X'"
			sql = Replace(sql, "@CardCode", "N''")
		Case "batchOpt"
			sql = sql & ") from OIBT " & _
			"left outer join R3_ObsCommon..DOC2 T1 on T1.LogNum = -1 and T1.LineNum = -1 and T1.BatchNum = OIBT.BatchNum collate database_default " & _
			"where OIBT.ItemCode = 'X' and OIBT.BatchNum = 'X'"
			sql = Replace(sql, "@ItemCode", "OIBT.ItemCode")
			sql = Replace(sql, "@BatchNum", "OIBT.BatchNum")
			sql = Replace(sql, "@WhsCode", "OIBT.WhsCode")
		Case "catOpt"
			sql = sql & ") from OITM inner join OITW on OITW.ItemCode = OITM.ItemCode and WhsCode = '01' where OITM.ItemCode = 'X' "
			sql = Replace(sql, "@ItemCode", "OITM.ItemCode")
			sql = Replace(sql, "@Table", "INV")
		Case "minrep"
			sql = sql & ")"
		Case "minrepSys"
			sql = sql & ")"
			sql = Replace(sql, "{Table}", "INV")
		Case "GenFilter", "ClientCatFilter", "NavQry", "SearchTreeFilter"
			sql = sql & ")"
			sql = Replace(sql, "@CardCode", "''")
			sql = Replace(sql, "@SlpCode", "-1")
			sql = Replace(sql, "@UserType", "''")
		Case "AnonCatFilter"
			sql = sql & ")"
		Case "AgentClientsFilter"
			sql = sql & ")"
		Case "ViewDocFilter"
			sql = sql & ")"
			sql = Replace(sql, "@CardCode", "''")
			sql = Replace(sql, "@SlpCode", "-1")
			sql = Replace(sql, "@ObjectCode", "17")
		Case "editQuery"
			If Request("rsTop") = "Y" Then
				sql = Replace(sql, "@top", 1)
			End If
			rs.Filter = "varType = 'CL'"
			do while not rs.eof
				sql = Replace(sql, "@" & rs("varVar"), "'1', '2'")
			rs.movenext
			loop
		Case "editCustomSearch"
			If Request("ID") <> "" Then
				rs.close
				sqlVar = "select Variable, [Type], DataType, MaxChar from OLKCustomSearchVars where ID = " & Request("ID")
				rs.open sqlVar, conn, 3, 1
				rs.Filter = "Type <> 'S'"
				
				do while not rs.eof
					If rs("Type") <> "CL" Then
						sql = Replace(sql, "@" & rs("Variable"), "'1'")
					Else
						sql = Replace(sql, "@" & rs("Variable"), "'1', '2'")
					End If
				rs.movenext
				loop
				rs.Filter = "Type = 'S'"
				If rs.recordcount > 0 Then sql = Replace(sql, "@SystemFilters", "1 = 2")
				Select Case CInt(Request("ObjID"))
					Case 2
						sql = Replace(sql, "OCRD.", "T0.")
						sql = Replace(sql, "OCRY.", "T1.")
						sql = Replace(sql, "OCRG.", "T2.")
					Case 4
				End Select
				sql = Replace(sql, "@SlpCode", 1)
				sql = Replace(sql, "@branch", 1)
				sql = Replace(sql, "@LanID", 1)
				sql = Replace(sql, "@CardCode", "''")
			End If
		Case "CustSearchQry"
			rs.Filter = "Type = 'CL'"
			do while not rs.eof
				sql = Replace(sql, "@" & rs("Variable"), "'1', '2'")
			rs.movenext
			loop
		Case "RSVar", "CustSearchVarQry", "docBD"
			sql = QueryFunctions(sql)
	End Select
	'Functions
	Select Case  Request("Type")
		Case "batchOpt", "crdOpt", "invOpt", "objConfCols", "cartOpt", _
				"minrep", "minrepSys", "catOpt", "DocFlow", "TaskMon", "editQuery", "secSubRS"
					
				If Left(Request("by"), 8) = "GrpValue" Then
					sql = sql & ")"
				End If

				sql = QueryFunctions(sql)
	End Select

End If
%>
<body <% If Request("Query") = "" Then %>onload="parent.VerfyQueryVerified();"<% End If %>>
<% 

errMsg = ""
dupColName = ""
noColName = False
noParName = False
err2Cols = False
errType = False

If Request("Query") <> "" Then
	On Error Resume Next
	set rs = conn.execute(sql)
	If Err.Number <> 0 Then
		If Request("Type") <> "newQueryVar" and Request("Type") <> "CustSearchVar" Then
			errMsg = Replace(Replace(Err.Description,"\","\\"),"'","\'")
		Else
			errMsg = Replace(getverfyQueryLngStr("LtxtValidVar"), "{0}", Request("Query"))
		End If %>
	<script language="javascript">alert('<%=errMsg%>')</script>
	<% Else
	
	Select Case Request("Type")
		Case "newQuery", "editQuery"
			For i = 0 to rs.Fields.Count -1
				For j = 0 to rs.Fields.Count - 1
					If j <> i Then
						If rs.Fields(i).Name = rs.Fields(j).Name Then
							dupColName = rs.Fields(i).Name
							Exit For
						End If
					End If
					If rs.Fields(i).Name = "" Then 
						noColName = True
						Exit For
					End If
					If InStr(rs.Fields(i).Name, "(") or InStr(rs.Fields(i).Name, ")") Then
						noParName = True
						Exit For
					End If
				Next
				If dupColName <> "" Then Exit For
			Next
		Case "RSVar", "CustSearchVarQry"
			If rs.Fields.Count < 2 Then
				err2Cols = True
			End If
		Case "editCustomSearch"
			If InStr(LCase(sql), "order by") <> 0 Then
				endQuery = ""
			    Do While (InStr(LCase(sql), "order by") <> 0)
			        If InStr(Right(sql, Len(sql) - InStr(LCase(sql), "order by")), ")") = 0 Then
			            endQuery = endQuery & Left(sql, InStr(LCase(sql), "order by") - 1)
			            sql = ""
						isOrderWrong = True
						errMsg = "" & getverfyQueryLngStr("LtxtErrOrderBy") & ""
			        Else
			            endQuery = endQuery & Left(sql, InStr(LCase(sql), "order by") + 7)
			            sql = Right(sql, Len(sql) - InStr(LCase(sql), "order by") - 7)
			        End If
			    Loop
			    If sql <> "" Then endQuery = endQuery & sql
			End If
		Case "ItemRec"
			
			For i = 0 to rs.Fields.Count -1
				Select Case rs.Fields(i).Name
					Case "ItemCode" 
						colItemCode = True
					Case "Quantity"
						colQuantity = True
					Case "Locked"
						colLocked = True
					Case "Checked"
						colChecked = True
					Case "WhsCode"
						colWhsCode = True
					Case "Comment"
						colComment = True
				End Select
				
				For j = 0 to rs.Fields.Count - 1
					If j <> i Then
						If rs.Fields(i).Name = rs.Fields(j).Name Then
							dupColName = rs.Fields(i).Name
							Exit For
						End If
					End If
					If rs.Fields(i).Name = "" Then 
						noColName = True
						Exit For
					End If
					If InStr(rs.Fields(i).Name, "(") or InStr(rs.Fields(i).Name, ")") Then
						noParName = True
						Exit For
					End If
				Next
				If dupColName <> "" Then Exit For
			Next
			If Not colItemCode or not colQuantity or not colLocked or not colChecked or not colWhsCode or not colComment Then
				isColMissing = True
				
				errMsg = getverfyQueryLngStr("LtxtMisCols") & ": \n"
				If Not colItemCode Then errMsg = errMsg & "ItemCode\n"
				If Not colQuantity Then errMsg = errMsg & "Quantity\n"
				If Not colLocked Then errMsg = errMsg & "Locked\n"
				If Not colChecked Then errMsg = errMsg & "Checked\n"
				If Not colWhsCode Then errMsg = errMsg & "WhsCode\n"
				If Not colComment Then errMsg = errMsg & "Comment"
			End If
		Case "docBD"
			If InStr(sql, "insert @BreakTable(LineNum, Quantity, ShipDate, ShipDateDiff)") = 0 Then
				errMsg = "insert @BreakTable(LineNum, Quantity, ShipDate, ShipDateDiff)"
			End If
	End Select
	
	If (Request("Type") = "RSVar" or Request("Type") = "CustSearchVarQry") and not err2Cols or (Request("Type") = "RSVarDef" or Request("Type") = "CustVarDefVal") Then
		If Request("varQueryField") = "" Then
			colType = getColTypeVal(rs.Fields(0).Type)
		Else
			colType = getColTypeVal(rs.Fields(CStr(Request("varQueryField"))).Type)
		End If
		If (Request("varDataType") = "numeric" or Request("varDataType") = "int") and colType <> "N" Then
			errType = True
		ElseIf Request("varDataType") = "datetime" and colType <> "D" Then
			errType = True
		End If
	End If
End If %>
<script language="javascript">
<% If dupColName = "" and not noColName and not err2Cols and not noParName and not errType and errMsg = "" and Request("Type") <> "GetSeries" Then %>
	parent.VerfyQueryVerified();
	<% If Request("Type") = "DocFlow" and Request("By") = "Note" Then %>
		var QueryFields = parent.getQueryFields();
		for (var i = QueryFields.length-1;i>=0;i--)
		{
			QueryFields.remove(i);
		}
		var o = 0;
		<% For each item in rs.Fields
		If item.Name <> "" Then %>
		QueryFields.options[o++] = new Option('<%=myHTMLEncode(item.Name)%>', '{<%=myHTMLEncode(item.Name)%>}');
		<% End If
		Next %>
	<% ElseIf Request("Type") = "MailMsg" Then %>
		var HeaderFields = parent.getHeaderQueryFields();
		var MsgFields = parent.getMsgQueryFields();
		for (var i = HeaderFields.length-1;i>=0;i--)
		{
			HeaderFields.remove(i);
			MsgFields.remove(i);
		}
		var o = 0;
		<% For each itm in rs.Fields
		If itm.Name <> "" Then %>
		HeaderFields.options[o] = new Option('<%=myHTMLEncode(itm.Name)%>', '{<%=myHTMLEncode(itm.Name)%>}');
		MsgFields.options[o++] = new Option('<%=myHTMLEncode(itm.Name)%>', '{<%=myHTMLEncode(itm.Name)%>}');
		<% End If
		Next %>
	<% ElseIf Request("Type") = "CUFD" or Request("Type") = "RSVar" or Request("Type") = "CustSearchVarQry" Then %>
		<% If Request("Type") = "CUFD" and CInt(Request("FieldID")) >= 0 Then %>
		var cmbSqlQueryField = parent.getSqlQueryField();
		var selVal = cmbSqlQueryField.value;
		for (var i = cmbSqlQueryField.length-1;i>=0;i--)
		{
			cmbSqlQueryField.remove(i);
		}
		var o = 0;
		<% For each itm in rs.Fields
		If itm.Name <> "" Then %>
		cmbSqlQueryField.options[o++] = new Option('<%=Replace(itm.Name, "'", "\'")%>', '<%=Replace(itm.Name, "'", "\'")%>');
		<% End If
		Next %>
		if (selVal != '') cmbSqlQueryField.value = selVal;
		<% End If %>
	<% End If %>
<% ElseIf Request("Type") = "GetSeries" Then %>
	var series = parent.getObjSeries();
	var o = 0;
	<% do while not rs.eof %>
	series.options[o++] = new Option('<%=rs(1)%>', '<%=rs(0)%>');
	<% rs.movenext
	loop %>
	series.disabled = false;
<% ElseIf dupColName <> "" Then %>
alert('<%=getverfyQueryLngStr("LtxtErrCols")%>'.replace('{0}', '<%=Replace(dupColName, "'", "\'")%>'));
<% ElseIf noColName  Then %>
alert('<%=getverfyQueryLngStr("LtxtErrColsNames")%>'.replace('{0}', '<%=Replace(dupColName, "'", "\'")%>'));
<% ElseIf noParName Then %>
alert('<%=getverfyQueryLngStr("LtxtErrColsPar")%>');
<% ElseIf err2Cols Then %>
alert('<%=getverfyQueryLngStr("Ltxt2Cols")%>');
<% ElseIf isColMissing or isOrderWrong Then %>
alert('<%=errMsg%>');
<% ElseIf errType Then %>
<% Select Case Request("varDataType")
	Case "datetime"
		varDataTypeDesc = getverfyQueryLngStr("DtxtDate")
	Case "int"
		varDataTypeDesc = getverfyQueryLngStr("LtxtNumWhole")
	Case "numeric"
		varDataTypeDesc = getverfyQueryLngStr("DtxtNumeric")
End Select
varColDesc = "1"
If Request("varQueryField") <> "" Then varColDesc = Request("varQueryField")
 %>
alert('<%=getverfyQueryLngStr("LtxtColType")%>'.replace('{0}', '<%=varColDesc%>').replace('{1}', '<%=varDataTypeDesc%>'));
<% End If %>
</script>
<% End If %>
</body>
<% If Request("Query") <> "" Then
conn.close
set rs = nothing
End If

Function getRSVariables(ByVal baseIndex)
	strRSVariables = ""
	Select Case Request("UserType") 
		Case "C"
			strRSVariables = "declare @CardCode nvarchar(15) set @CardCode = '' "
		Case "V"
			strRSVariables = "declare @SlpCode int set @SlpCode = -1 "
	End Select
	If Request("baseIndex") <> "" Then %>
	<!--#include file="repVars.inc"-->
<%	
		sql2 = "select '@' + varVar varVar, varDataType, varMaxChar, DefValBy, DefValDate, DefValValue, OLKCommon.dbo.DBOLKGetRSVarBaseIndex" & Session("ID") & "(rsIndex, varIndex) BaseIndex from " & repTbl & "RSVars where rsIndex = " & Request("rsIndex") & " and varIndex in (" & baseIndex & ")"
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
					sqlVal = ""
					If rs("BaseIndex") <> "-1" Then sqlVal = getRSVariables(rs("BaseIndex"))
					sqlVal = sqlVal & " " & rs("DefValValue")
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
End Function %>
</html>
