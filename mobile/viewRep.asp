<!--#include file="lang/viewRep.asp" -->

<link rel="stylesheet" href="Reportes/style.css">
<% HaveVals = False

set rd = server.createobject("ADODB.RecordSet")
set rcCopy = Server.CreateObject("ADODB.RecordSet")
sql = "select IsNull(alterRSName, rsName) rsName, IsNull(alterRSDesc, rsDesc) rsDesc, rsQuery, rsTop, Refresh, " & _
"Case When Exists(select 'A' from OLKRSTotals where rsIndex = T0.rsIndex and colTotal <> 'D') Then 'Y' Else 'N' End Total, " & _
"Case When Exists(select 'A' from OLKRSTotals where rsIndex = T0.rsIndex and colTotal <> 'D' and colShow in ('A','T')) Then 'Y' Else 'N' End TotalTop, " & _
"Case When Exists(select 'A' from OLKRSTotals where rsIndex = T0.rsIndex and colTotal <> 'D' and colShow in ('A','B')) Then 'Y' Else 'N' End TotalBottom " & _
"from OLKRS T0 " & _
"left outer join OLKRSAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.LanID = " & Session("LanID") & " " & _
"where T0.rsIndex = " & Request("rsIndex")
set rs = conn.execute(sql)
rsName = rs("rsName")
rsDesc = rs("rsDesc")
sqlQuery = QueryFunctions(rs("rsQuery"))
Refresh = rs("Refresh")
rsTop = rs("rsTop") = "Y"
If rs("Total") = "Y" Then setTotal = True else setTotal = False
If rs("TotalTop") = "Y" Then setTotalTop = True Else setTotalTop = False
If rs("TotalBottom") = "Y" Then setTotalBottom = True Else setTotalBottom = False
sqlVars = " "

sql = "select T0.varIndex, varVar, IsNull(alterVarName, varName) varName, varType, varDataType, varMaxChar, varShowRep " & _
	"from OLKRSVars T0 " & _
	"left outer join OLKRSVarsAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = T0.varIndex and T1.LanID = " & Session("LanID") & " " & _
	"where T0.rsIndex = " & Request("rsIndex")
set rVal = Server.CreateObject("ADODB.RecordSet")
rVal.open sql, conn, 3, 1

If rVal.recordcount > 0 Then
	do while not rVal.eof
		If rVal("varType") <> "CL" Then
			If rVal("varDataType") = "nvarchar" Then 
				MaxVar = "(" & rVal("varMaxChar") & ")"
			ElseIf rVal("varDataType") = "numeric" Then
				MaxVar = "(19,6)"
			Else
				MaxVar = ""
			End If
			sqlVars = sqlVars & "declare @" & rVal("varVar") & " " & rVal("varDataType") & MaxVar & " set @" & rVal("varVar") & " = "

			If Request("var" & rVal("varIndex")) <> "" Then 
				strVal = saveHTMLDecode(Request("var" & rVal("varIndex")), False)
				If rVal("varType") = "DP" Then sqlVars = sqlVars & "Convert(datetime,'" & SaveSqlDate(strVal) & "',120)" Else sqlVars = sqlVars & "N'" & strVal & "' " 
			Else 
				sqlVars = sqlVars & "NULL "
			End If
		End If
		If varShowRep <> "" then varShowRep = varShowRep  & "{|} "
		If rVal("varShowRep") = "Y" Then 
			varShowRep = varShowRep & "<B>" & rVal("varName") & ":</b> "
			If rVal("varType") <> "CL" Then
				varShowRep = varShowRep & Request("var" & rVal("varIndex"))
			Else
				varShowRep = varShowRep & Request("var" & rVal("varIndex") & "Desc")
			End If
		End If
		HaveVals = True
	rVal.movenext
	loop
Else
	sql = " "
End If
sqlVars = sqlVars& "declare @SlpCode int set @SlpCode = " & Session("vendid") & " declare @LanID int set @LanID = " & Session("LanID") & " "

sql = sqlVars & sqlQuery

If rsTop Then sql = Replace(sql, "@top", Request("varTop"))

rVal.Filter = "varType = 'CL'"
do while not rVal.eof
	sql = Replace(sql, "@" & rVal("varVar"), Request("var" & rVal("varIndex")))
rVal.movenext
loop

set rs = conn.execute(sql) 

If Session("RetVal") = "" Then
	linkType = "Case When T0.linkType = 'A' and T0.linkObject in (4, 5) Then 'N' Else T0.linkType End"
Else
	linkType = "T0.linkType"
End If

sql = 	"select T0.colName, T4.alterColName,  " & _
		"Case T0.colAlign When 'L' Then 'left' When 'C' Then 'center' When 'R' Then 'right' End colAlign,  " & _
		"T0.colFormat, T0.colTotal, T0.colSum, T0.colShow, " & linkType & " linkType, Case T0.linkType When 'F' Then T0.linkObjectPocket Else T0.linkObject End linkObject, " & _
		"Case T0.linkType When 'F'Then T0.linkLinkPocket Else T0.linkLink End linkLink, T0.linkPopup,  " & _
		"Replace(T1.rsName,'""""','""""""""') linkObjectTtl, "
	
Select Case Session("useraccess")
	Case "P"
		sql = sql & "'Y' "
	Case "U"
		If myAut.AuthorizedRepGroups <> "" Then
			sql = sql & "Case When T1.rgIndex in (" & myAut.AuthorizedRepGroups & ") Then 'Y' Else 'N' End "
		Else
			sql = sql & "'Y' "
		End If
End Select
		
sql = sql & " linkObjectAccess from OLKRSTotals T0 "

If userType = "V" Then
	sql = sql & "inner join OLKAgentsAccess S0 on S0.SlpCode = " & Session("vendid") & " "
	fldUid = "SlpCode"
End If
		
sql = sql & "left outer join OLKRS T1 on T1.rsIndex = T0.linkObject and T0.linkType = 'R' " & _
		"left outer join OLKRG T2 on T2.rgIndex = T1.rgIndex "

sql = sql & "left outer join OLKRSTotalsAlterNames T4 on T4.rsIndex = T0.rsIndex and T4.colName = T0.colName and T4.LanID = " & Session("LanID") & " " & _
			"where T0.rsIndex = " & Request("rsIndex") & " "
rd.open sql, conn, 3, 1

set rl = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.colName, T0.varId, T0.valBy, T2.varType, IsNull(T0.valValue,OLKCommon.dbo.DBOLKDateFormat" & Session("ID") & "(T0.valValDat)) valValue, T2.varIndex " & _
"from OLKRSLinksVars T0 " & _
"left outer join OLKRSTotals T1 on T1.rsIndex = T0.rsIndex and T1.colName = T0.colName " & _
"left outer join OLKRSVars T2 on T2.rsIndex = T1.linkObject and T2.varVar = T0.varId " & _
"where T0.rsIndex = " & Request("rsIndex") & " " & _
"and (T0.valValue is not null or T0.valValDat is not null) "
rl.open sql, conn, 3, 1

If setTotal Then 
	set rt = server.createobject("ADODB.RecordSet")
	
	endQuery = ""
	If (InStr(LCase(sqlQuery), "order by") <> 0) Then
	    Do While (InStr(LCase(sqlQuery), "order by") <> 0)
	        If InStr(Right(sqlQuery, Len(sqlQuery) - InStr(LCase(sqlQuery), "order by")), ")") = 0 Then
	            endQuery = endQuery & Left(sqlQuery, InStr(LCase(sqlQuery), "order by") - 1)
	            sqlQuery = ""
	        Else
	            endQuery = endQuery & Left(sqlQuery, InStr(LCase(sqlQuery), "order by") + 7)
	            sqlQuery = Right(sqlQuery, Len(sqlQuery) - InStr(LCase(sqlQuery), "order by") - 7)
	        End If
	    Loop
	    If sqlQuery <> "" Then endQuery = endQuery & sqlQuery
	Else
	    endQuery = sqlQuery
	End If
	sqlQuery = endQuery
	endQuery = ""
	rd.Filter = ""
	rd.movefirst
	hCols = 0
	sql = sqlVars & " select "
	do while not rd.eof
		If rd("colFormat") <> "H" Then
			If sql <> sqlVars & " select " Then sql = sql & ", "
			If rd("colTotal") = "D" Then sql = sql & "'' As '" & rd("colName") & "'" Else sql = sql & rd("colTotal") & "([" & rd("colName") & "]) As '" & rd("colName") & "'"
		Else
			hCols = hCols + 1
		End If
	rd.movenext
	loop
	sql = sql & " from(" & sqlQuery & ") As Table1"
	'on error resume next
	If rsTop Then sql = Replace(sql, "@top", Request("varTop"))
	rVal.Filter = "varType = 'CL'"
	do while not rVal.eof
		sql = Replace(sql, "@" & rVal("varVar"), Request("var" & rVal("varIndex")))
	rVal.movenext
	loop
	set rt = conn.execute(sql)
End If

FldFilter = ""
For each Field in rs.Fields
	If FldFilter <> "" Then FldFilter = FldFilter & ", "
	FldFilter = FldFilter & "N'" & Field.Name & "'"
Next

set rc = Server.CreateObject("ADODB.RecordSet")
sql = "select ColorID, LineID, Alias, colName, colType, colOp, colOpBy, colValue, colValDate, FontFace, FontSize, ForeColor, FontBold, FontItalic, FontUnderLine, " & _
	"FontStrike, FontBlink, BackColor, ApplyToRow, ApplyToCol, Active " & _
	"from OLKRSColors " & _
	"where rsIndex = " & Request("rsIndex") & " and Active = 'Y' " & _
	"and colName in (" & FldFilter & ") " & _
	"and (colOpBy <> 'F' or colOpBy = 'F' and colValue in (" & FldFilter & ") or colOp in ('N', 'NN')) " & _
	"order by Ordr, Ordr2 "
rc.open sql, conn, 3, 1

If Err.Number = 0 Then %>
<div align="center">
	<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
		<tr class="TblTlt">
			<td>&nbsp;<%=getviewRepLngStr("LtxtReps")%></td>
		</tr>
		<tr class="TablasNoticiasTitle">
			<td>
			<%=Server.HTMLEncode(rsName)%>&nbsp;</td>
		</tr>
		<% If rsDesc <> "" and Not IsNull(rsDesc) Then %>
		<tr class="TablasNoticias">
			<td>
			<%=Server.HTMLEncode(rsDesc)%>&nbsp;</td>
		</tr>
		<% End If %>
		<tr class="TablasNoticiasTitle">
			<td><b><%=getviewRepLngStr("DtxtExecTime")%>:</b> <%=FormatDate(Now(), True)%>&nbsp;<%=FormatTime(Now())%></td>
		</tr>
		<tr>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" width="100%" id="table2" style="font-family: Verdana; font-size: 10px">
				<% If varShowRep <> "" Then
					ArrVal = Split(varShowRep, "{|} ")
					For i = 0 to UBound(ArrVal) %>
						<tr>
							<td colspan="<%=rs.Fields.Count%>"><%=ArrVal(i)%>&nbsp;</td>
						</tr>
					<% next
					End If %>
					<tr class="TblTltMnu">
				<% For each Field in rs.Fields
				   rd.Filter = "colName = '" & Field.Name & "'"
				   If Not rd.Eof Then
				   	If rd("colFormat") = "H" Then ShowCol = False Else ShowCol = True
				   	If IsNull(rd("alterColName")) Then
				   		fldName = Field.Name
				   	Else
				   		fldName = rd("alterColName")
				   	End If
				   Else
				   	ShowCol = True
				   	fldName = Field.Name
				   End If
				   If ShowCol Then
				   		myColCount = myColCount + 1 %>
				   		<% If Not rd.Eof Then
				   			If rd("linkType") <> "N" or rd("colFormat") = "EML" Then
				   			hCols = hCols - 1 %><td>&nbsp;</td>
				   			<% End If
				   			End If %>
						<td align="center"><%=fldName%>&nbsp;</td>
				<% End If
				   next %>
					</tr>
					<% If setTotal and setTotalTop Then
					rd.Filter = "colShow = 'A' or colShow = 'T'"
					If rd.RecordCount > 0 Then %>
					<% ShowRepTotal "T" %>
					<tr class="TablasNoticias">
						<td colspan="<%=rs.Fields.Count-hCols%>" class="RepTotalCel" height="4">&nbsp;</td>
					</tr>
					<% End If
					rd.Filter = ""
					End If %>
				<%  rd.Filter = "colSum = 'Y'"
					Redim repColSum(rd.RecordCount)
					curLine = 0
					do while not rs.eof
						curColSum = 0
						curLine = curLine + 1 %>
						<tr class="TablasNoticias">
					<% 
						For each Field in rs.Fields
						strVal = ""
						rd.Filter = "colName = '" & Field.Name & "'"
						ignoreRowFormat = False
						DoColorRowBlink = False
						If Not rd.Eof Then
							If rd("colFormat") = "H" Then ShowCol = False Else ShowCol = True
							If rd("colSum") = "Y" Then
								If IsNull(Field) Then colSumVal = 0 Else colSumVal = CDbl(Field)
								If curLine > 1 Then
									repColSum(curColSum) = repColSum(curColSum) + colSumVal
								Else
									repColSum(curColSum) = colSumVal
								End If
								strVal = repColSum(curColSum)
								curColSum = curColSum + 1
								ignoreRowFormat = True
							End If
						End If
						If ShowCol Then
							If Not rd.Eof Then 
								If rd("colFormat") = "EML" Then
									doEMail = True
								Else 
									doEMail = False
								End If
							Else
								doEmail = False
							End If %>
							<% If doEmail Then %>
							<td width="16"><% If Not IsNull(Field) Then %><a href="mailto:<%=Field%>"><img border="0" src="images/mail.gif"></a><% End If %></td>
							<% End If %>
							<%=doLink%>
							<td <% If Not rd.Eof Then %>align="<%=rd("colAlign")%>"<% End If %> style="<%=doFormat(ignoreRowFormat)%>">
							<% If DoColorRowBlink Then %><blink><% End If %>
							<% 	If rd.Eof Then 
									strVal = Field
								Else
									If rd("colSum") = "Y" Then
										doFormatStr = strVal
									ElseIf Not IsNull(Field) Then
										doFormatStr = Field
									Else 
										doFormatStr = ""
									End If
									If doFormatStr <> "" Then
								 		Select Case rd("colFormat")
											Case "R"
												strVal = FormatNumber(CDbl(doFormatStr),myApp.RateDec)
											Case "S"
												strVal = FormatNumber(CDbl(doFormatStr),myApp.SumDec)
											Case "P"
												strVal = FormatNumber(CDbl(doFormatStr),myApp.PriceDec)
											Case "Q"
												strVal = FormatNumber(CDbl(doFormatStr),myApp.QtyDec)
											Case "%"
												strVal = FormatNumber(CDbl(doFormatStr),myApp.PercentDec)
											Case "M"
												strVal = FormatNumber(CDbl(doFormatStr),myApp.MeasureDec)
											Case "IS"
												strVal = NumInItem(CDbl(doFormatStr), CDbl(rs("NumInSale")))
											Case "IB"
												strVal = NumInItem(CDbl(doFormatStr), CDbl(rs("NumInBuy")))
											Case "IM"
												strVal = "<img src=""pic.aspx?filename=" & doFormatStr & "&MaxSize=80&dbName=" & Session("olkdb") & """>"
											Case Else
												If Left(rd("colFormat"),2) = "DT" and IsNumeric(Right(rd("colFormat"),1)) Then
													Select Case CInt(Right(rd("colFormat"),1))
														Case 1
															strVal = FormatDateTime(doFormatStr, 1)
														Case 2
															strVal = FormatDate(doFormatStr, true)
														Case 3
															strVal = FormatTime(doFormatStr)
														Case 4
															strVal = FormatDateTime(doFormatStr, 4)
													End Select
												ElseIf Left(rd("colFormat"),1) = "D" and IsNumeric(Right(rd("colFormat"),1)) Then
													strVal = FormatNumber(CDbl(doFormatStr),CInt(Right(rd("colFormat"),1)))
												Else
													strVal = doFormatStr
												End If
										End Select
									End If
								End If %>
								<%=strVal%><% If DoColorRowBlink Then %></blink><% End If %></td>
					<% End If
					   next %>
					</tr>
				<% rs.movenext
				loop %>
				<% If setTotal and setTotalBottom Then
					rd.Filter = "colShow = 'A' or colShow = 'B'"
					If rd.RecordCount > 0 Then  %>
					<tr class="TablasNoticias">
						<td colspan="<%=rs.Fields.Count-hCols%>" class="RepTotalCel" height="4">&nbsp;</td>
					</tr>
					<% ShowRepTotal "B"
					End If
					rd.Filter = "" %>
				<% End If %>
				</table>
			</div>
			</td>
		</tr>
		</table>
</div>
<form action="operaciones.asp" method="post" name="frmReload">
<% For each item in Request.Form
If item <> "btnReload" Then %>
<input type="hidden" name="<%=item%>" value="<%=Request(item)%>">
<% End If
Next
For each item in Request.QueryString %>
<input type="hidden" name="<%=item%>" value="<%=Request(item)%>">
<% Next %>
<div align="left">
	<table border="0" cellpadding="0" width="315" id="table3">
		<tr>
			<td width="99">
			<p align="center">
			<input type="submit" value="<%=getviewRepLngStr("DtxtRefresh")%>" name="btnReload"></td>
			<% If HaveVals Then %><td>
			<input type="submit" value="<%=getviewRepLngStr("DtxtNew")%>" name="btnNewQuery" onclick="javascript:document.frmReload.cmd.value='viewRepVals'"></td><% End If %>
		</tr>
	</table>
</div>
</form>
<% Else %>
<div align="center"><font color="red" face="Verdana" size="2"><p><%=getviewRepLngStr("LtxtQueryErr")%>: </p><p><font color="black"><%=Err.Description%></font></p></font></div>
<% End If %>
<%
Function NumInItem(ByVal Qty, ByVal QtyInItem)
	P1 = Fix(Qty/QtyInItem)
	P2 = CInt((Qty/QtyInItem-Fix(Qty/QtyInItem))*QtyInItem)
	RetVal = P1
	If QtyInItem <> 1 Then RetVal = RetVal & "(" & P2 & ")"
	NumInItem = RetVal
End Function
%>
<% Sub ShowRepTotal(ByVal Pos) %>
	<tr class="TablasNoticias">
	<% For each Field in rs.Fields
	rd.Filter = "colName = '" & Field.Name & "'"
	If not rd.Eof Then
		If rd("colFormat") = "H" Then ShowCol = False Else ShowCol = True
		If ShowCol Then If rd("colShow") <> "A" and rd("colShow") <> Pos Then BlankCol = True Else BlankCol = False
		'BlankCol = False
	Else
		ShowCol = True
		BlankCol = True
	End if
	If ShowCol Then
	If Not BlankCol Then %>
		<td <% If Not rd.Eof Then %>align="<%=rd("colAlign")%>" <% If rd("linkType") <> "N" or rd("colFormat") = "EML" Then %>colspan="2"<% End If %><% End If %>>
		<% 	Field = rt(Field.name)
			If rd.Eof Then 
				strVal = Field
			Else
				If Not IsNull(Field) and Field <> "" Then
			 		Select Case rd("colFormat")
						Case "R"
							strVal = FormatNumber(CDbl(Field),myApp.RateDec)
						Case "S"
							strVal = FormatNumber(CDbl(Field),myApp.SumDec)
						Case "P"
							strVal = FormatNumber(CDbl(Field),myApp.PriceDec)
						Case "Q"
							strVal = FormatNumber(CDbl(Field),myApp.QtyDec)
						Case "%"
							strVal = FormatNumber(CDbl(Field),myApp.PercentDec)
						Case "M"
							strVal = FormatNumber(CDbl(Field),myApp.MeasureDec)
						Case Else
							If Left(rd("colFormat"),2) = "DT" and IsNumeric(Right(rd("colFormat"),1)) Then
								Select Case CInt(Right(rd("colFormat"),1))
									Case 1
										strVal = FormatDateTime(Field, 1)
									Case 2
										strVal = FormatDate(Field, true)
									Case 3
										strVal = FormatTime(Field)
									Case 4
										strVal = FormatDateTime(Field, 4)
								End Select
							ElseIf Left(rd("colFormat"),1) = "D" and IsNumeric(Right(rd("colFormat"),1)) Then
								strVal = FormatNumber(CDbl(Field),CInt(Right(rd("colFormat"),1)))
							Else
								strVal = Field
							End If
					End Select
				Else
					strVal = ""
				End If
			End If %>
			<%=strVal%>&nbsp;</td>
<% Else %>
	<td <% If Not rd.Eof Then %><% If rd("linkType") <> "N" or rd("colFormat") = "EML" Then %>colspan="2"<% End If %><% End If %>>&nbsp;</td>
<% End If
   End If
   next %></tr><% End Sub
   
Function doFormat(ByVal ColSum)
	LastColorID = -1
	colFormat = ""
	If Not ColSum Then
		rc.Filter = "(ColName = '" & Field.Name & "' and ApplyToCol = '' or ApplyToCol = '" & Field.Name & "') or ApplyToRow = 'Y'"
	Else
		rc.Filter = "(ColName = '" & Field.Name & "' and ApplyToCol = '') and ApplyToRow = 'N'"
	End If
	If Not rc.Eof Then
		rc.movefirst
		do while not rc.eof
			If CInt(LastColorID) <> rc("ColorID") Then
			colFormat = AnalizeFormat(rc("ColorID"), CStr(rc("ColName")), colFormat, rc("colOp"), rc("colOpBy"), _
						rc("colValue"), rc("colType"), rc("colValDate"), rc("FontFace"), rc("FontSize"), _
						rc("ForeColor"), rc("FontBold"), rc("FontItalic"), rc("FontUnderline"), rc("FontStrike"), rc("FontBlink"), rc("BackColor"))
			End If
		rc.movenext
		loop
	End If
	doFormat = colFormat 
End Function

Function AnalizeFormat(ByVal ColorID, ByVal ColName, ByVal curFormat, ByVal colOp, ByVal colOpBy, ByVal colValue, ByVal colType, ByVal colValueDate, ByVal FontFace, ByVal FontSize, _
						ByVal ForeColor, ByVal FontBold, ByVal FontItalic, ByVal FontUnderline, ByVal FontStrike, ByVal FontBlink, ByVal BackColor)
	colFormat = curFormat
	If colOp <> "N" and colOp <> "NN" Then 
		If colOpBy = "F" Then 
			If colType = "N" Then
				compVal = CDbl(rs(CStr(colValue)))
			Else
				compVal = CStr(rs(CStr(colValue)))
			End If
		ElseIf colType = "D" Then
			compVal = colValueDate
		ElseIf colType = "N" Then
			compVal = CDbl(colValue)
		Else
			compVal = CStr(colValue)
		End If
	End If
	If rd.Eof Then 
		If IsNull(rs(ColName)) Then
			mainVal = rs(ColName)
		ElseIf colType = "N" Then 
			mainVal = CDbl(rs(ColName))
		ElseIf colType = "D" Then
			mainVal = rs(ColName)
		Else
			mainVal = CStr(rs(ColName))
		End If
	Else
		If rd("ColSum") = "Y" Then
			If IsNull(strVal) Then
				mainVal = strVal
			End If
		Else
			If IsNull(rs(ColName)) Then 
				mainVal = rs(ColName)
			End If
		End If
		If colType = "N" Then
			If rd("colSum") = "Y" Then
				If Not IsNull(strVal) Then mainVal = CDbl(strVal) 
			Else 
				If Not IsNull(rs(ColName)) Then mainVal = CDbl(rs(ColName))
			End If
		ElseIf colType = "D" Then
			If rd("ColSum") = "Y" Then 
				If Not IsNull(strVal) Then mainVal = strVal
			Else 
				If Not IsNull(rs(ColName)) Then mainVal = rs(ColName)
			End If
		Else
			If rd("colSum") = "Y" Then 
				If Not IsNull(strVal) Then mainVal = CStr(strVal) 
			Else 
				If Not IsNull(rs(ColName)) Then mainVal = CStr(rs(ColName))
			End If
		End If
	End If
	apFormat = False
	Select Case colOp
		Case "="
			If mainVal = compVal Then apFormat = True
		Case "<>"
			If mainVal <> compVal Then apFormat = True
		Case ">"
			If mainVal > compVal Then apFormat = True
		Case "<"
			If mainVal < compVal Then apFormat = True
		Case ">="
			If mainVal >= compVal Then apFormat = True
		Case "<="
			If mainVal <= compVal Then apFormat = True
		Case "N"
			If IsNull(mainVal) Then apFormat = True
		Case "NN"
			If Not IsNull(mainVal) Then apFormat = True
	End Select
	If apFormat Then 
		If FontBlink = "Y" Then DoColorRowBlink = True
		colFormat = getRCFormat(colFormat, FontFace, FontSize, ForeColor, FontBold, FontItalic, FontUnderline, FontStrike, BackColor)
		LastColorID = ColorID
	End If
	AnalizeFormat = colFormat
End Function 

Function getRCFormat(ByVal curFormat, ByVal FontFace, ByVal FontSize, ByVal ForeColor, ByVal FontBold, ByVal FontItalic, ByVal FontUnderline, ByVal FontStrike, ByVal BackColor)
	retVal = curFormat
	If Not IsNull(FontFace) Then retVal = retVal & "font-family:" & FontFace & "; "
	If Not IsNull(FontSize) Then 
		retVal = retVal & "font-size: "
		Select Case FontSize
			Case 1
				retVal = retVal & "8"
			Case 2
				retVal = retVal & "10"
			Case 3
				retVal = retVal & "12"
			Case 4
				retVal = retVal & "14"
			Case 5
				retVal = retVal & "18"
			Case 6
				retVal = retVal & "24"
			Case 7
				retVal = retVal & "36"
		End Select
		retVal = retVal & "pt; "
	End If 
	If Not IsNull(ForeColor) Then retVal = retVal & "color:" & ForeColor & "; "
	If FontBold = "Y" Then retVal = retVal & "font-weight:bold; "
	If FontItalic = "Y" Then retVal = retVal & "font-style:italic; "
	If FontUnderline = "Y" and FontStrike = "Y" Then 
		retVal = retVal & "text-decoration:underline line-through; "
	ElseIf FontUnderline = "Y" and FontStrike = "N" Then
		retVal = retVal & "text-decoration:underline; "
	ElseIf FontUnderline = "N" and FontStrike = "Y" Then
		retVal = retVal & "text-decoration:line-through; "
	End If
	If Not IsNull(BackColor) Then retVal = retVal & "background-color:" & BackColor & "; "
	getRCFormat = retVal 
End Function

Private Function doLink
	retVal = ""
	If Not rd.Eof Then
		If rd("linkType") <> "N" Then retVal = "<td width=""15"" style=""" & doFormat(ignoreRowFormat) & """><a href=""javascript:"
		Select Case rd("linkType")
			Case "L", "F"
				myLink = rd("linkLink")
				For each myLnkItm in rs.Fields
					If Not IsNull(myLnkItm) Then
						myLink = Replace(myLink, "{" & myLnkItm.Name & "}", Replace(myLnkItm, "'", "\'"))
					ElseIf InStr(myLink, "{" & myLnkItm.Name & "}") <> 0 Then
						'myLink = Replace(myLink, "{" & myLnkItm.Name & "}", "")
						'Hay que agergarla opcion en el enlace checkbox si permite variables nulas"
						doLink = "<td width=""15"">&nbsp;</td>"
						Exit Function
					End If
				Next
				Select Case rd("linkType") 
					Case "F" 
						myLink = "operaciones.asp?cmd=sec&SecID=" & rd("linkObject") & "&" & myLink
						retVal = retVal & "window.location.href='" & myLink & "';"
					Case "L"
						retVal = retVal & "window.open('" & myLink & "');"
				End Select
			Case "O"
				rl.Filter = "colName = '" & Field.Name & "'"
				myLink = "doDetail(" & rd("linkObject") & ", '"
				If rl("valBy") = "F" Then
					myLink = myLink & rs(CStr(rl("valValue")))
				ElseIf rl("valBy") = "V" Then
					myLink = myLink & rl("valValue")
				End If
				myLink = myLink & "');"
				retVal = retVal & myLink
			Case "R"
				If rd("linkObjectAccess") = "Y" Then
					myLink = "goRep" & Replace(Replace(rd("colName"), " ", ""),"#","Sharp") & "("
					rl.Filter = "colName = '" & rd("colName") & "' and valBy = 'F'"
					myLinkVars = ""
					do while not rl.eof
						If myLinkVars <> "" Then myLinkVars = myLinkVars & ", "
						If Not IsNull(rs(CStr(rl("valValue")))) Then
							If rl("varType") <> "CL" Then
								myLinkVars = myLinkVars & "'" & Replace(rs(CStr(rl("valValue"))),"'","\'") & "'"
							Else
								myLinkVars = myLinkVars & "'\'" & Replace(rs(CStr(rl("valValue"))),"'","\'") & "\''"
							End If
						Else
							doLink = "<td width=""15"" style=""" & doFormat(ignoreRowFormat) & """>&nbsp;</td>"
							Exit Function
						End If
					rl.movenext
					loop
					myLink = myLink & myLinkVars & ")"
				Else
					myLink = myLink & "alert('" & getviewRepLngStr("LtxtNoRepAccess") & "');"
				End If
				myLink = myLink & """ alt=""Enlace a reporte: " & rd("linkObjectTtl")
				retVal = retVal & myLink
		End Select
		If rd("linkType") <> "N" Then retVal = retVal & """><img border=""0"" src=""images/" &Session("rtl") & "flechaselec.gif"" style=""cursor: hand""></a></td>"
	End If
	doLink = retVal
End Function %>
<% rd.Filter = "linkType = 'O'"
If rd.recordcount > 0 Then %>
<form target="_blank" method="post" name="frmViewDetail" action="">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="">
<input type="hidden" name="ViewOnly" value="Y">
<input type="hidden" name="CardCode" value="">
</form>
<form name="frmOp" method="post" action="operaciones.asp">
<input type="hidden" name="cmd" value="">
<input type="hidden" name="card" value="">
<input type="hidden" name="c1" value="">
<input type="hidden" name="item" value="">
<input type="hidden" name="view" value="Y">
<input type="hidden" name="price" value="no">
</form>
<script language="javascript">
function doDetail(ObjectCode, Entry)
{
	document.frmViewDetail.AddPath.value = '';
	if (ObjectCode == 13 || ObjectCode == 17 || ObjectCode == 23 || ObjectCode == 15 || ObjectCode == 16 || ObjectCode == 14 ||
			ObjectCode == 22 || ObjectCode == 20 || ObjectCode == 21 || ObjectCode == 18 || ObjectCode == 19)
	{
		document.frmViewDetail.action = "cxcdocdetail.asp";
		document.frmViewDetail.DocEntry.value = Entry;
		document.frmViewDetail.DocType.value = ObjectCode;
		document.frmViewDetail.submit();
	}
	else if (ObjectCode == 2)
	{
		document.frmOp.cmd.value = 'datos';
		document.frmOp.card.value = Entry;
		document.frmOp.c1.value = Entry;
		document.frmOp.submit();
	}
	else if (ObjectCode == 4)
	{
		document.frmOp.cmd.value = '<% If Session("RetVal") = "" Then %>itemdetails<% Else %>addcart<% End If %>';
		document.frmOp.item.value = Entry;
		document.frmOp.submit();
	}
}

function Start(page) {
<% If userType = "C" Then %>
OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width=482,height=310");
<% ElseIf userType = "V" Then %>
OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,resizable=no,scrollbars=yes,width=598,height=410,status=yes");
<% End If %>
OpenWin.focus()
}
</script>
<% End If
rd.Filter = "linkType = 'R'"
do while not rd.eof
rlVars = "" %>
<form method="post" action="operaciones.asp" name="frmRep<%=Replace(Replace(rd("colName"), " ", ""),"#","Sharp")%>">
<input type="hidden" name="rsIndex" value="<%=rd("linkObject")%>">
<% rl.Filter = "colName = '" & rd("colName") & "'"
do while not rl.eof
If rl("valBy") = "F" Then
	If rlVars <> "" Then rlVars = rlVars & ", "
	rlVars = rlVars & "v" & rl("varId")
End If %>
<input type="hidden" name="var<%=rl("varIndex")%>" value="<% If rl("valBy") = "V" Then Response.Write rl("valValue") %>">
<% rl.movenext
loop %>
<input type="hidden" name="cmd" value="viewRep">
</form>
<% rl.Filter = "colName = '" & rd("colName") & "' and valBy = 'F'" %>
<script language="javascript">
function goRep<%=Replace(Replace(rd("colName"), " ", ""),"#","Sharp")%>(<%=rlVars%>)
{
	<% do while not rl.eof %>document.frmRep<%=Replace(Replace(rd("colName"), " ", ""),"#","Sharp")%>.var<%=rl("varIndex")%>.value = v<%=rl("varId")%>;
	<% rl.movenext
	loop %>
	document.frmRep<%=Replace(Replace(rd("colName"), " ", ""),"#","Sharp")%>.submit();
}
</script>
<% 
rd.movenext
loop %>
<% If Refresh > 0 Then %>
<script language="javascript">
setTimeout("reloadRep()", <%=Refresh*1000*60%>);
function reloadRep()
{
	document.frmReload.btnReload.click();
}
</script>
<% End If %>