<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="myHTMLEncode.asp" -->
<!--#include file="adminTradSave.asp"-->
<!--#include file="repVars.inc" -->
<%
set rs = server.createobject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")

Select Case Request("Cmd")
	Case "repRestore"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKRestoreRS" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@rsIndex") = Request("rsIndex")
		cmd.execute()

		Response.Redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex")
	Case "uActive"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetRSAdmList" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@UserType") = Request("UserType")
		If Request("cmbRg") <> "" Then cmd("@rgIndex") = Request("cmbRg")
		set rs = cmd.execute()
		
		LoadCmd "OLKAdminRS"
		do while not rs.eof
			If Request("Active" & rs(0)) = "Y" Then Active = "Y" Else Active = "N"
			cmd("@Active") = Active
			cmd("@rsIndex") = rs(0)
			cmd.execute()
		rs.movenext
		loop
		response.redirect "adminReps.asp?uType=" & Request("UserType") & "&rgIndex=" & Request("cmbRg")
	Case "uGrp"
		sql = "select rgIndex from OLKRG where UserType = '" & Request("UserType") & "'"
		If Request("UserType") = "V" Then sql = sql & " and rgIndex >= 0"
		set rs = conn.execute(sql)
		sql = ""
		do while not rs.eof
			sql = sql & " update OLKRG set rgName = N'" & saveHTMLDecode(Request("rgName" & rs("rgIndex")), False) & "', SuperUser = N'" & Request("SuperUser" & rs("rgIndex")) & "' where rgIndex = " & rs("rgIndex")
		rs.movenext
		loop
		If sql <> "" Then conn.execute(sql)
		
		If Request("rgName") <> "" Then
			sql =	" declare @rgIndex int set @rgIndex = IsNULL((select max(rgIndex)+1 from OLKRG),0) " & _
					" select @rgIndex rgIndex " & _
					"insert OLKRG(rgIndex, rgName, SuperUser, UserType) values(@rgIndex, N'" & saveHTMLDecode(Request("rgName"), False) & "', N'" & Request("SuperUser") & "', N'" & Request("UserType") & "')"
			set rs = conn.execute(sql)
			
			If Request("rgNameTrad") <> "" Then
				SaveNewTrad Request("rgNameTrad"), "RG", "rgIndex", "alterRGName", rs(0)
			End If
		End If
	Case "remRG"
		sql = "delete OLKRG where rgIndex = " & Request("delIndex") & _
				" delete OLKRGAlterNames where rgIndex = " & Request("delIndex")
		If Request("UserType") = "C" Then
			sql = sql & " delete OLKClientsRGAccess where rgIndex = " & Request("delIndex")
		End If
		conn.execute(sql)
	Case "remRS"
		LoadCmd "OLKAdminRS"
		cmd("@rsIndex") = Request("rsIndex")
		cmd("@Action") = "D"
		cmd.execute()
	Case "uRep"
		If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
		If Request("LinkOnly") = "Y" Then LinkOnly = "Y" Else LinkOnly = "N"
		If Request("rsTop") = "Y" Then rsTop = "Y" Else rsTop = "N"
		
		LoadCmd "OLKAdminRS"
		response.write Request("rsIndex")
		cmd("@rsIndex") = Request("rsIndex")
		cmd("@rsName") = saveHTMLDecode(Request("rsName"), True)
		cmd("@rsDesc") = saveHTMLDecode(Request("rsDesc"), True)
		cmd("@rsQuery") = saveHTMLDecode(Request("rsQuery"), True)
		cmd("@rsTop") = rsTop
		cmd("@rsTopDef") = Request("rsTopDef")
		cmd("@rgIndex") = Request("rgIndex")
		cmd("@Refresh") = Request("cmbRefresh")
		cmd("@Active") = Active
		cmd("@LinkOnly") = LinkOnly
		cmd("@Action") = "U"
		
		If Request("saveAs") = "Y" Then
			cmd("@SaveAs") = "Y"
			cmd("@SaveAsRSName") = saveHTMLDecode(Request("saveAsName"), True)
			cmd("@SaveAsRGIndex") = Request("saveAsRG")
		End If
		cmd.execute()
		
		If Request("saveAs") <> "Y" Then
			rsIndex = Request("rsIndex")
			
			set rs = Server.CreateObject("ADODB.RecordSet")
			sql = "select varVar, varDataType from OLKRSvars where rsIndex = " & Request("rsIndex")
			set rs = conn.execute(sql)
			sql = "declare @LanID int "
			do while not rs.eof
				sql = sql & "declare @" & rs("varVar") & " " & rs("varDataType") & " set @" & rs("VarVar") & " = "
				Select Case rs("varDataType")
					Case "nvarchar"
						sql = sql & "'' "
					Case "datetime"
						sql = sql & "'01/01/01' "
					Case "float"
						sql = sql & "0 "
					Case "numeric"
						sql = sql & "0 "
					Case "int"
						sql = sql & "0 "
				End Select
			rs.movenext
			loop
			If Request("UserType") = "C" Then
				sql = sql & " declare @CardCode nvarchar(15) set @CardCode = '' "
			ElseIf Request("UserType") = "V" Then
				sql = sql & " declare @SlpCode int set @SlpCode = -1 "
			End If
			sqlQuery = "select rsQuery, rsTop from OLKRS where rsIndex = " & Request("rsIndex")
			set rs = conn.execute(sqlQuery)
			sqlQuery = rs("rsQuery")
			sql = sql & sqlQuery
			If rs("rsTop") = "Y" Then sql = Replace(sql, "@top", 1)
			sql = QueryFunctions(sql)
			set rs = conn.execute(sql)
			
			colFound = ""
			For each fld in rs.Fields
				If colFound <> "" Then colFound = colFound & ", "
				colFound = colFound & "N'" & Replace(fld.Name, "'", "''") & "'"
			Next
			
			If colFound <> "" Then
				sql = "delete OLKRSTotals where rsIndex = " & Request("rsIndex") & " and colName not in (" & saveHTMLDecode(colFound, True) & ")"
				conn.execute(sql)
			End If
		
		Else
			rsIndex = cmd("@rsIndex")
		End If
		
		conn.close
		
		If Request("btnApply") <> "" or Request("saveAs") = "Y" Then
			response.redirect "adminRepEdit.asp?rsIndex=" & rsIndex & "&repCmd=" & Request("repCmd")
		Else
			response.redirect "adminReps.asp?uType=" & Request("UserType")
		End If
	Case "AdmColors"
		LoadCmd "OLKAdminRSAdmColors"
		cmd("@rsIndex") = Request("rsIndex")
		
		If Request("ColID") <> "" Then
			ColID = Split(Request("ColID"), ", ")
			For i = 0 to UBound(ColID)
				ColorID = Replace(Split(ColID(i), "_")(0), "Lst", "")
				LineID = Split(ColID(i), "_")(1)
				If Request("ColActive" & ColID(i)) = "Y" Then Active = "Y" Else Active = "N"
				
				cmd("@ColorID") = ColorID
				cmd("@LineID") = LineID
				cmd("@Alias") = saveHTMLDecode(Request("ColAlias" & ColID(i)), True)
				cmd("@Ordr") = Request("ColOrdr" & ColorID)
				cmd("@Ordr2") = Request("ColOrdr2" & ColID(i))
				cmd("@Active") = Active
				
				cmd.execute()
			Next
			
		End If
		
		Response.Redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=repColor&#tblRepColors"
	Case "repColor"
		LoadCmd "OLKAdminRSColors"
		cmd("@rsIndex") = Request("rsIndex")
		
		If Request("ColorID") = "" or Request("AlterOf") <> "" Then
			If Request("ColorID") = "" Then
				cmd("@LineID") = 0
			Else
				cmd("@ColorID") = Request("ColorID")
				If Request("AlterOf") = "Y" Then cmd("@AlterOf") = "Y"
			End If
		Else
			cmd("@ColorID") = Request("ColorID")
			cmd("@LineID") = Request("LineID")
		End If
		
		cmd("@Alias") 		= saveHTMLDecode(Request("Alias"), True)
		cmd("@colName")		= saveHTMLDecode(Split(Request("colName"),"{|}")(0), True)
		cmd("@colType")		= Split(Request("colName"),"{|}")(1)
		cmd("@colOp")		= Request("colOp")
		
		If Request("colOp") <> "N" and Request("colOp") <> "NN" Then
			colOpBy = Request("colOpBy")
			If Request("colOpBy") = "F" Then
				If Request("colValCol") <> "" Then cmd("@colValue") = saveHTMLDecode(Split(Request("colValCol"),"{|}")(0), True)
			Else
				If Split(Request("colName"),"{|}")(1) <> "D" Then
					If Request("colValVal") <> "" Then cmd("@colValue") = saveHTMLDecode(Request("colValVal"), True)
				Else
					If Request("colValDat") <> "" Then cmd("@colValDate") = SaveSqlDate(Request("colValDat"))
				End If
			End if
		Else
			colOpBy = "F"
		End If
		
		cmd("@colOpBy") = colOpBy 
		
		If Request("FontFace") <> "" and Request("FontFace") <> " " Then cmd("@FontFace") = Request("FontFace")
		If Request("FontSize") <> "" and Request("FontSize") <> " " Then cmd("@FontSize") = Request("FontSize")
		If Request("ForeColor") <> "" Then cmd("@ForeColor") = Request("ForeColor")
		If Request("BackColor") <> "" Then cmd("@BackColor") = Request("BackColor")
		If Request("FontAlign") <> "" Then cmd("@FontAlign") = Request("FontAlign")
		
		If Request("FontBold") 	= "Y" Then FontBold = "Y" Else FontBold = "N"
		If Request("FontItalic") = "Y" Then FontItalic = "Y" Else FontItalic = "N"
		If Request("FontUnderline") = "Y" Then FontUnderline = "Y" Else FontUnderline = "N"
		If Request("FontStrike") = "Y" Then FontStrike = "Y" Else FontStrike = "N"
		If Request("FontBlink") = "Y" Then FontBlink = "Y" Else FontBlink = "N"
		If Request("ApplyTo") = "R" Then ApplyToRow = "Y" Else ApplyToRow = "N"
		If Request("ApplyTo") = "A" Then cmd("@ApplyToCol") = saveHTMLDecode(Request("ApplyToCol"), True)
		If Request("Active") = "Y" Then Active = "Y" Else Active = "N"
		
		cmd("@FontBold")		= FontBold
		cmd("@FontItalic") 		= FontItalic
		cmd("@FontUnderline") 	= FontUnderline
		cmd("@FontStrike") 		= FontStrike
		cmd("@FontBlink") 		= FontBlink
		cmd("@ApplyToRow")		= ApplyToRow
		cmd("@Active")			= Active
		cmd("@Ordr")			= Request("Ordr")
		cmd("@Ordr2")			= Request("Ordr2")
		
		cmd.execute()
		
		ColorID = cmd("@ColorID")
		LineID = cmd("@LineID")
		
		If Request("AliasTrad") <> "" Then
			SaveNewTrad Request("AliasTrad"), "RSColors", "rsIndex,ColorID,LineID", "AlterAlias", Request("rsIndex") & "," & ColorID & "," & LineID
		End If

		If Request("btnSave") <> "" Then
			response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=repColor&#editColor"
		Else
			response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=repColor&ColorID=" & ColorID & "&LineID=" & LineID & "&LineNum=" & Request("LineNum") & "&AlterNum=" & Request("AlterNum") & "&#editColor"
		End If
	Case "addVar"
		If Request("varNotNull") = "Y" Then varNotNull = "Y" Else varNotNull = "N"
		If Request("varShowRep") = "Y" Then varShowRep = "Y" Else varShowRep = "N"
		
		LoadCmd "OLKAdminRSVars"
		cmd("@rsIndex") 		= Request("rsIndex")
		cmd("@varName") 		= saveHTMLDecode(Request("varName"), True)
		cmd("@varVar") 			= saveHTMLDecode(Request("varVar"), True)
		cmd("@varType") 		= Request("varType")
		cmd("@varDataType") 	= Request("varDataType")
		If Request("varQuery") <> "" and Request("varQueryBy") = "Q" Then cmd("@varQuery") = saveHTMLDecode(Request("varQuery"), True)
		If Request("varQueryField") <> "" Then cmd("@varQueryField") = saveHTMLDecode(Request("varQueryField"), True)
		If Request("varMaxChar") <> "" Then cmd("@varMaxChar") = Request("varMaxChar")
		cmd("@varNotNull") 		= varNotNull
		cmd("@varDefVars") 		= Request("varQueryBy")
		cmd("@varShowRep") 		= varShowRep
		cmd("@DefValBy") = Request("varDefBy")
		If Request("varDefBy") = "V" and Request("varDataType") <> "datetime" Then
			cmd("@DefValValue") = Request("varDefValValue")
		ElseIf Request("varDefBy") = "V" and Request("varDataType") = "datetime" Then
			cmd("@DefValDate") = SaveSqlDate(Request("varDefValDate"))
		ElseIf Request("varDefBy") = "Q" Then
			cmd("@DefValValue") = Request("varDefValQuery")
		End If
		cmd("@Ordr") = Request("Ordr")
		
		cmd.execute()
		varIndex = cmd("@varIndex")
		
		If Request("varNameTrad") <> "" Then
			SaveNewTrad Request("varNameTrad"), "RSVars", "rsIndex,varIndex", "alterVarName", Request("rsIndex") & "," & varIndex
		End If
		
		If Request("varQueryDef") <> "" Then
			SaveNewDef Request("varQueryDef"), CStr(Request("rsIndex")) & CStr(varIndex)
		End If
		
		If Request("varDefValueDef") <> "" Then
			SaveNewDef Request("varDefValueDef"), CStr(Request("rsIndex")) & CStr(varIndex)
		End If

		If Request("varQueryBy") = "F" Then
			LoadCmd "OLKAdminRSVarsVals"
			cmd("@rsIndex") = Request("rsIndex")
			cmd("@varIndex") = varIndex

			ArrVal = Split(Request("varQuery"),VbCrLf)
			for i = 0 to UBound(ArrVal)
				ArrVal2 = Split(ArrVal(i),",")
				cmd("@valValue") = saveHTMLDecode(ArrVal2(0), True)
				cmd("@valText") = saveHTMLDecode(ArrVal2(1), True)
				cmd.execute()
			next
		End If
		
		LoadCmd "OLKAdminRSVarsBase"
		cmd("@rsIndex") = Request("rsIndex")
		cmd("@varIndex") = varIndex

		If Request("baseVar") <> "" Then
			ArrVal = Split(Request("baseVar"), ", ")
			For i = 0 to UBound(ArrVal)
				cmd("@baseIndex") = ArrVal(i)
				cmd.execute()
			Next
		End If
		
		If Request("btnSave") <> "" Then
			response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=variables"
		Else
			response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&editIndex=" & varIndex & "&repCmd=variables&#editVar"
		End If 
	Case "editVar"
		If Request("varNotNull") = "Y" Then varNotNull = "Y" Else varNotNull = "N"
		If Request("varShowRep") = "Y" Then varShowRep = "Y" Else varShowRep = "N"
		
		LoadCmd "OLKAdminRSVars"
		cmd("@rsIndex") 		= Request("rsIndex")
		cmd("@varIndex")		= Request("varIndex")
		cmd("@varName") 		= saveHTMLDecode(Request("varName"), True)
		cmd("@varVar") 			= saveHTMLDecode(Request("varVar"), True)
		cmd("@varType") 		= Request("varType")
		cmd("@varDataType") 	= Request("varDataType")
		If Request("varQuery") <> "" and Request("varQueryBy") = "Q" Then cmd("@varQuery") = saveHTMLDecode(Request("varQuery"), True)
		If Request("varQueryField") <> "" Then cmd("@varQueryField") = saveHTMLDecode(Request("varQueryField"), True)
		If Request("varMaxChar") <> "" Then cmd("@varMaxChar") = Request("varMaxChar")
		cmd("@varNotNull") 		= varNotNull
		cmd("@varDefVars") 		= Request("varQueryBy")
		cmd("@varShowRep") 		= varShowRep
		cmd("@DefValBy") = Request("varDefBy")
		If Request("varDefBy") = "V" and Request("varDataType") <> "datetime" Then
			cmd("@DefValValue") = Request("varDefValValue")
		ElseIf Request("varDefBy") = "V" and Request("varDataType") = "datetime" Then
			cmd("@DefValDate") = SaveSqlDate(Request("varDefValDate"))
		ElseIf Request("varDefBy") = "Q" Then
			cmd("@DefValValue") = Request("varDefValQuery")
		End If
		cmd("@Ordr") = Request("Ordr")
		cmd("@Action") = "U"
		
		cmd.execute()
		
		If Request("varQueryBy") = "F" Then
			ClearTableData "RSVarsVals", Request("rsIndex"), Request("varIndex")
			LoadCmd "OLKAdminRSVarsVals"
			cmd("@rsIndex") = Request("rsIndex")
			cmd("@varIndex") = Request("varIndex")

			ArrVal = Split(Request("varQuery"),VbCrLf)
			for i = 0 to UBound(ArrVal)
				ArrVal2 = Split(ArrVal(i),",")
				cmd("@valValue") = saveHTMLDecode(ArrVal2(0), True)
				cmd("@valText") = saveHTMLDecode(ArrVal2(1), True)
				cmd.execute()
			next
		End If
		
		ClearTableData "RSVarsBase", Request("rsIndex"), Request("varIndex")
		LoadCmd "OLKAdminRSVarsBase"
		cmd("@rsIndex") = Request("rsIndex")
		cmd("@varIndex") = Request("varIndex")

		If Request("baseVar") <> "" Then
			ArrVal = Split(Request("baseVar"), ", ")
			For i = 0 to UBound(ArrVal)
				cmd("@baseIndex") = ArrVal(i)
				cmd.execute()
			Next
		End If
		
		If Request("btnSave") <> "" Then
			response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=variables"
		Else
			response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&editIndex=" & Request("varIndex") & "&repCmd=variables&#editVar"
		End If 
	Case "admVars"
		LoadCmd "OLKAdminRSVars"
		cmd("@rsIndex") = Request("rsIndex")
		
		arrIndex = Split(Request("varIndex"), ", ")
		For i = 0 to UBound(arrIndex)
			varIndex = arrIndex(i)
			cmd("@Action") = "ADM"
			cmd("@varIndex") = varIndex
			cmd("@varName") = saveHTMLDecode(Request("varName" & varIndex), True)
			cmd("@Ordr") = Request("Ordr" & varIndex)
			cmd.execute()
		Next
		response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=variables"

	Case "remRSVar"
		LoadCmd "OLKAdminRSVars"
		cmd("@rsIndex") = Request("rsIndex")
		cmd("@varIndex") = Request("varIndex")
		cmd("@Action") = "D"
		cmd.execute()
		response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=variables"
	Case "remRSCol"
		LoadCmd "OLKAdminRSColors"
		cmd("@rsIndex") = Request("rsIndex")
		cmd("@ColorID") = Request("ColorID")
		If Request("LineID") <> "" Then cmd("@LineID") = Request("LineID")
		cmd("@Action") = "D"
		cmd.execute()
		response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&repCmd=repColor"
	Case "newRep"
		LoadCmd "OLKAdminRS"
		cmd("@rsName") = saveHTMLDecode(Request("rsName"), True)
		cmd("@rsDesc") = saveHTMLDecode(Request("rsDesc"), True)
		cmd("@rsQuery") = saveHTMLDecode(Request("rsQuery"), True)
		cmd("@rgIndex") = Request("rgIndex")
		cmd("@Refresh") = Request("cmbRefresh")
		cmd.execute()
		rsIndex = cmd("@rsIndex")
		
		If Request("rsNameTrad") <> "" Then
			SaveNewTrad Request("rsNameTrad"), "RS", "rsIndex", "alterRSName", rsIndex
		End If
		
		If Request("rsDescTrad") <> "" Then
			SaveNewTrad Request("rsDescTrad"), "RS", "rsIndex", "alterRSDesc", rsIndex
		End If
		
		If Request("rsQueryDef") <> "" Then
			SaveNewDef Request("rsQueryDef"), rsIndex
		End If
		
		conn.close
		response.redirect "adminRepEdit.asp?rsIndex=" & rsIndex
	Case "varUp"
		sql = "update OLKRSVars set varIndex = -1 where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("varIndex") & _
			" update OLKRSVars set varIndex = varIndex + 1 Where rsIndex = " & Request("rsIndex") & " and varIndex = -1 + " & Request("varIndex") & _
			" update OLKRSVars set varIndex = -1 + " & Request("varIndex") & " where rsIndex = " & Request("rsIndex") & " and varIndex = -1 "  & _
			" update OLKRSVarsVals set varIndex = -1 where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("varIndex") & _
			" update OLKRSVarsVals set varIndex = varIndex + 1 Where rsIndex = " & Request("rsIndex") & " and varIndex = -1 + " & Request("varIndex") & _
			" update OLKRSVarsVals set varIndex = -1 + " & Request("varIndex") & " where rsIndex = " & Request("rsIndex") & " and varIndex = -1"
		conn.execute(sql)
		LastUpdate()
		response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex")
	Case "varDown"
		sql = "update OLKRSVars set varIndex = -1 where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("varIndex") & _
			" update OLKRSVars set varIndex = varIndex - 1 Where rsIndex = " & Request("rsIndex") & " and varIndex = 1 + " & Request("varIndex") & _
			" update OLKRSVars set varIndex = 1 + " & Request("varIndex") & " where rsIndex = " & Request("rsIndex") & " and varIndex = -1 "  & _
			" update OLKRSVarsVals set varIndex = -1 where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("varIndex") & _
			" update OLKRSVarsVals set varIndex = varIndex - 1 Where rsIndex = " & Request("rsIndex") & " and varIndex = 1 + " & Request("varIndex") & _
			" update OLKRSVarsVals set varIndex = 1 + " & Request("varIndex") & " where rsIndex = " & Request("rsIndex") & " and varIndex = -1 "
		conn.execute(sql)
		LastUpdate()
		response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex")
	Case "repTotals"
		sql = "delete OLKRSTotals where rsIndex = " & Request("rsIndex") & " "
		If Request("arrColName") <> "" Then sql = sql & " and colName not in (" & Request("arrColName") & ")"

		conn.execute(sql)
		sql = "select varVar, varDataType from OLKRSvars where rsIndex = " & Request("rsIndex")
		set rs = conn.execute(sql)
		sql = "declare @LanID int "
		do while not rs.eof
			sql = sql & "declare @" & rs("varVar") & " " & rs("varDataType") & " set @" & rs("VarVar") & " = "
			Select Case rs("varDataType")
				Case "nvarchar"
					sql = sql & "'' "
				Case "datetime"
					sql = sql & "'01/01/01' "
				Case "float"
					sql = sql & "0 "
				Case "numeric"
					sql = sql & "0 "
				Case "int"
					sql = sql & "0 "
			End Select
		rs.movenext
		loop
		If Request("UserType") = "C" Then
			sql = sql & " declare @CardCode nvarchar(15) set @CardCode = '' "
		ElseIf Request("UserType") = "V" Then
			sql = sql & " declare @SlpCode int set @SlpCode = -1 "
		End If
		sqlQuery = "select rsQuery, rsTop from OLKRS where rsIndex = " & Request("rsIndex")
		set rs = conn.execute(sqlQuery)
		sqlQuery = rs("rsQuery")
		sql = sql & sqlQuery
		If rs("rsTop") = "Y" Then sql = Replace(sql, "@top", 1)
		sql = QueryFunctions(sql)
		set rs = conn.execute(sql)
		
		LoadCmd "OLKAdminRSTotals"
		cmd("@rsIndex") = Request("rsIndex")
		
		For i = 1 to rs.Fields.count
			colName = i
			If Request("colSum" & colName) = "Y" Then colSum = "Y" Else colSum = "N"
			If Request("colNB" & colName) = "Y" Then colNB = "Y" Else colNB = "N"
			
			cmd("@colName") 	= rs.Fields(i-1).Name
			cmd("@colTotal") 	= Request("Action" & colName)
			cmd("@colAlign") 	= Request("Align" & colName)
			cmd("@colFormat") 	= Request("Format" & colName)
			cmd("@colSum") 		= colSum
			cmd("@colShow")		= Request("Show" & colName)
			cmd("@colNB")		= colNB
			cmd.execute()
		next
		conn.execute(sql)
		response.redirect "adminRepEdit.asp?rsIndex=" & Request("rsIndex") & "&#tblRepTotals"
End Select

Sub LastUpdate()
	sql = "update OLKRS set LastUpdate = getdate() where rsIndex = " & Request("rsIndex")
	conn.execute(sql)
End Sub

set rs = nothing
conn.close
response.redirect "adminReps.asp?uType=" & Request("UserType") & "&rgIndex=" & Request("rgIndex")
%> 
