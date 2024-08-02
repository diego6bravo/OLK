<%
Function QueryFunctions(ByVal value)
	strQryFunction = value
	strQryFunction = Replace(strQryFunction, "OLKCode(", "OLKCommon.dbo.DBOLKCode" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKDocTotal(", "OLKCommon.dbo.DBOLKDocTotal" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKDocTotalBox(", "OLKCommon.dbo.DBOLKDocTotalBox" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKDocTotalVol(", "OLKCommon.dbo.DBOLKDocTotalVol" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKDocTotalWeight(", "OLKCommon.dbo.DBOLKDocTotalWeight" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKEncodeBreakLines(", "OLKCommon.dbo.OLKEncodeBreakLines(")
	strQryFunction = Replace(strQryFunction, "OLKDateFormat(", "OLKCommon.dbo.DBOLKDateFormat" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKFormatNumber(", "OLKCommon.dbo.DBOLKFormatNumber" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKGetCufv(", "OLKCommon.dbo.DBOLKGetCufv" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKGetTrans(", "OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKInv(", "OLKCommon.dbo.DBOLKInv" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKItemInvVal(", "OLKCommon.dbo.DBOLKItemInvVal" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKObjectName(", "OLKCommon.dbo.DBOLKObjectName" & Session("ID") & "(")
	strQryFunction = Replace(strQryFunction, "OLKNumIn(", "OLKCommon.dbo.OLKNumIn(")
	strQryFunction = Replace(strQryFunction, "OLKSplit(", "OLKCommon.dbo.OLKSplit(")
	QueryFunctions = strQryFunction 
End Function

Function JBool(ByVal value)
	If value Then JBool = "true" Else JBool = "false"
End Function

Function GetYN(ByVal value)
	If value Then GetYN = "Y" Else GetYN = "N"
End Function

Function JScriptHRefEncode(ByVal value)
	JScriptHRefEncoderetVal = value
	JScriptHRefEncoderetVal = Replace(JScriptHRefEncoderetVal, "'", "\'")
	JScriptHRefEncoderetVal = Replace(JScriptHRefEncoderetVal, """", """""")
	JScriptHRefEncode = JScriptHRefEncoderetVal
End Function

Function StringFormat(sVal, aArgs)
	For stringI=0 To UBound(aArgs)
		sVal = Replace(sVal,"{" & CStr(stringI) & "}",aArgs(stringI))
	Next
	StringFormat = sVal
End Function

Function GetNumberStep(dec)
	strGetNumberStep = "0"
	
	If dec > 0 Then
		strGetNumberStep = "0."
		For i = 1 to dec
			If i <> dec Then
				strGetNumberStep = strGetNumberStep & "0"
			Else
				strGetNumberStep = strGetNumberStep & "1"
			End If
		Next
	End If
	
	GetNumberStep = strGetNumberStep
End Function

Function GetCalendarFormatString
	GetCalendarFormatString = Replace(Replace(Replace(myApp.DateFormat, "dd", "%d"), "MM", "%m"), "yyyy", "%Y")
End Function

Function GetJQueryCalendarFormat
	GetJQueryCalendarFormat = Replace(myApp.DateFormat, "MM", "mm")
End Function

Function SaveSqlDate(ByVal value)
	If IsNull(value) or value = "" Then Exit Function
	
	value = Replace(value, "'", "")

	dim strDFormat
	
	strDFormat = myApp.DateFormat
	
	strDYear = Mid(value, InStr(strDFormat, "yyyy"), 4)
	strDMonth = Mid(value, InStr(strDFormat, "MM"), 2)
	strDDay = Mid(value, InStr(strDFormat, "dd"), 2)
	
	SaveSqlDate = StringFormat("{0}-{1}-{2}", Array(strDYear, strDMonth, strDDay))
End Function

Function SaveCmdDate(ByVal value)
	If IsNull(value) or value = "" Then Exit Function
	
	value = Replace(value, "'", "")

	dim strDFormat
	
	strDFormat = myApp.DateFormat
	
	strDYear = Mid(value, InStr(strDFormat, "yyyy"), 4)
	strDMonth = Mid(value, InStr(strDFormat, "MM"), 2)
	strDDay = Mid(value, InStr(strDFormat, "dd"), 2)
	
	SaveCmdDate = DateSerial(strDYear,strDMonth,strDDay)
End Function


Function SaveCmdTime(ByVal value)
	If IsNull(value) or value = "" Then Exit Function
	
	arrTime = Split(value, ",")
	myTimeH = arrTime(0)
	myTimeM = arrTime(1)
	myTimeS = arrTime(2)
	If myTimeS = "AM" and CInt(myTimeH) = 12 Then
		myTimeH = 0
	End If

	myRetTime = Now()
	myRetTime = DateAdd("h", CInt(myTimeH), myDate)
	myRetTime = DateAdd("n", CInt(myTimeM), myDate)

	SaveCmdTime = myRetTime 
End Function


Function FormatTime(ByVal strDate)
	If IsNull(strDate) or Not IsDate(strDate) Then Exit Function

	strHour = DatePart("h", strDate)
	strMinute = DatePart("n", strDate)
	strSign = "am"
	
	If strHour = 0 Then
		strHour = 12
	ElseIf strHour > 11 Then
		If strHour > 12 Then strHour = strHour - 12
		strSign = "pm"
	End If
	
	If Len(strMinute) = 1 Then strMinute = "0" & strMinute 
	
	strTimeRetVal = strHour & ":" & strMinute & " " & strSign
	
	
	FormatTime = strTimeRetVal	

End Function

Function FormatDate(ByVal strDate, ByVal noBr)
	If IsNull(strDate) or Not IsDate(strDate) Then Exit Function
  
	dim strDFormat
	
	strDFormat = myApp.DateFormat
	
	dim strDateRetVal
	
	strDateRetVal  = Replace(strDFormat, "dd", right("0" & DatePart("d",strDate),2))
	strDateRetVal = Replace(strDateRetVal, "MM",  right("0" & DatePart("m",strDate),2))
	strDateRetVal = Replace(strDateRetVal, "yyyy", DatePart("yyyy",strDate))
	
	If noBr Then strDateRetVal = "<nobr>" & strDateRetVal & "</nobr>"

	FormatDate = strDateRetVal
End Function

iisVer = Request.ServerVariables("SERVER_SOFTWARE")
iisVer = Right(iisVer, Len(iisVer)-InStr(iisVer, "/"))
If iisVer >= "6.0" Then Response.CodePage = 65001
Response.CharSet = "UTF-8"

Function getMyLng
	Select Case Session("LanID")
		Case 2
			getMyLng = "ES-LA"
		Case 3
			getMyLng = "HE"
		Case 6
			getMyLng = "PT-BR"
		Case 8
			getMyLng = "FR"
		Case 10
			getMyLng = "DE"
		Case 13
			getMyLng = "RU"
		Case Else
			getMyLng = "EN-US"
	End Select
End Function

Function myHTMLDecode(strVal)
	retValHTMLDecode = strVal
	retValHTMLDecode = Replace(retValHTMLDecode, "&lt;", "<")
	retValHTMLDecode = Replace(retValHTMLDecode, "&gt;", ">")
	retValHTMLDecode = Replace(retValHTMLDecode, "&amp;", "&")
	retValHTMLDecode = Replace(retValHTMLDecode, "&quot;", """")

	myHTMLDecode = retValHTMLDecode
End Function

Function myHTMLEncode(strVal)
	retValHTMLEncode = strVal
	If Not IsNull(retValHTMLEncode) Then
		retValHTMLEncode = Server.HTMLEncode(retValHTMLEncode)
		retValHTMLEncode = Replace(retValHTMLEncode, "&lt;", "<")
		retValHTMLEncode = Replace(retValHTMLEncode, "&gt;", ">")
		retValHTMLEncode = Replace(retValHTMLEncode, "&amp;", "&")
		retValHTMLEncode = Replace(retValHTMLEncode, "&quot;", """")
		myHTMLEncode = retValHTMLEncode
	Else
		myHTMLEncode = ""
	End If
End Function

Function myQSEncode(strVal)
	retValQSEncode = myHTMLEncode(strVal)
	retValQSEncode = Replace(retValQSEncode, "&", "{AND}")
	retValQSEncode = Replace(retValQSEncode, "#", "{SHARP}")
	myQSEncode = retValQSEncode
End Function

Function myQSDecode(strVal)
	retValQSDecode = Replace(strVal, "{AND}", "&")
	retValQSDecode = Replace(retValQSDecode, "{SHARP}", "#")
	retValQSDecode = saveHTMLDecode(retValQSDecode, False)
	myQSDecode = retValQSDecode
End Function

Function mySearchString(strVal)
	mySearchString = Replace(strVal, "_", "[_]")
	mySearchString = Replace(strVal, "'", "''")
End Function

Function saveHTMLDecode(strVal, IsCommand)
	strRetVal = strVal
	
	If Not IsNull(strVal) Then
		strRetVal = Replace(strRetVal, "&#225;", "á")
		strRetVal = Replace(strRetVal, "&#233;", "é")
		strRetVal = Replace(strRetVal, "&eacute;", "é")
		strRetVal = Replace(strRetVal, "&#237;", "í")
		strRetVal = Replace(strRetVal, "&iacute;", "í")
		strRetVal = Replace(strRetVal, "&#243;", "ó")
		strRetVal = Replace(strRetVal, "&#250;", "ú")
		strRetVal = Replace(strRetVal, "&#241;", "ñ")
		strRetVal = Replace(strRetVal, "&#191;", "¿")
		strRetVal = Replace(strRetVal, "&#193;", "Á")
		strRetVal = Replace(strRetVal, "&#201;", "É")
		strRetVal = Replace(strRetVal, "&#205;", "Í")
		strRetVal = Replace(strRetVal, "&#211;", "Ó")
		strRetVal = Replace(strRetVal, "&#218;", "Ú")
		strRetVal = Replace(strRetVal, "&#209;", "Ñ")
		strRetVal = Replace(strRetVal, "&#220;", "Ü")
		strRetVal = Replace(strRetVal, "&#225;", "á")
		strRetVal = Replace(strRetVal, "&#233;", "é")
		strRetVal = Replace(strRetVal, "&#237;", "í")
		strRetVal = Replace(strRetVal, "&#243;", "ó")
		strRetVal = Replace(strRetVal, "&#250;", "ú")
		strRetVal = Replace(strRetVal, "&#241;", "ñ")
		strRetVal = Replace(strRetVal, "&#252;", "ü")
		strRetVal = Replace(strRetVal, "&#192;", "À")
		strRetVal = Replace(strRetVal, "&#194;", "Â")
		strRetVal = Replace(strRetVal, "&#195;", "Ã")
		strRetVal = Replace(strRetVal, "&#202;", "Ê")
		strRetVal = Replace(strRetVal, "&#212;", "Ô")
		strRetVal = Replace(strRetVal, "&#213;", "Õ")
		strRetVal = Replace(strRetVal, "&#224;", "à")
		strRetVal = Replace(strRetVal, "&#226;", "â")
		strRetVal = Replace(strRetVal, "&#227;", "ã")
		strRetVal = Replace(strRetVal, "&#234;", "ê")
		strRetVal = Replace(strRetVal, "&#232;", "è")
		strRetVal = Replace(strRetVal, "&#244;", "ô")
		strRetVal = Replace(strRetVal, "&#245;", "õ")
		strRetVal = Replace(strRetVal, "&#252;", "ü")
		strRetVal = Replace(strRetVal, "&#251;", "û")
		strRetVal = Replace(strRetVal, "&#199;", "Ç")
		strRetVal = Replace(strRetVal, "&#231;", "ç")
		strRetVal = Replace(strRetVal, "&quot;", "\")
		strRetVal = Replace(strRetVal, "&amp;", "&")
		strRetVal = Replace(strRetVal, "&#161;", "¡")
		strRetVal = Replace(strRetVal, "&#1488;", "א")
		strRetVal = Replace(strRetVal, "&#1489;", "ב")
		strRetVal = Replace(strRetVal, "&#1490;", "ג")
		strRetVal = Replace(strRetVal, "&#1491;", "ד")
		strRetVal = Replace(strRetVal, "&#1492;", "ה")
		strRetVal = Replace(strRetVal, "&#1493;", "ו")
		strRetVal = Replace(strRetVal, "&#1494;", "ז")
		strRetVal = Replace(strRetVal, "&#1495;", "ח")
		strRetVal = Replace(strRetVal, "&#1496;", "ט")
		strRetVal = Replace(strRetVal, "&#1497;", "י")
		strRetVal = Replace(strRetVal, "&#1498;", "ך")
		strRetVal = Replace(strRetVal, "&#1499;", "כ")
		strRetVal = Replace(strRetVal, "&#1500;", "ל")
		strRetVal = Replace(strRetVal, "&#1501;", "ם")
		strRetVal = Replace(strRetVal, "&#1502;", "מ")
		strRetVal = Replace(strRetVal, "&#1503;", "ן")
		strRetVal = Replace(strRetVal, "&#1504;", "נ")
		strRetVal = Replace(strRetVal, "&#1505;", "ס")
		strRetVal = Replace(strRetVal, "&#1506;", "ע")
		strRetVal = Replace(strRetVal, "&#1507;", "ף")
		strRetVal = Replace(strRetVal, "&#1508;", "פ")
		strRetVal = Replace(strRetVal, "&#1509;", "ץ")
		strRetVal = Replace(strRetVal, "&#1510;", "צ")
		strRetVal = Replace(strRetVal, "&#1511;", "ק")
		strRetVal = Replace(strRetVal, "&#1512;", "ר")
		strRetVal = Replace(strRetVal, "&#1513;", "ש")
		strRetVal = Replace(strRetVal, "&#1514;", "ת")
		strRetVal = Replace(strRetVal, "&#1523;", "׳")
		strRetVal = Replace(strRetVal, "&#1524;", "״")
		strRetVal = Replace(strRetVal, "&#8362;", "₪")
		If Not IsCommand Then strRetVal = Replace(strRetVal, "'", "''")
	End If
	
	saveHTMLDecode = strRetVal
End Function

Function myJavascriptEncode(strVal)
	retValJavascriptEncode = strVal

	retValJavascriptEncode = Replace(retValJavascriptEncode, VbCrLf, "\n")
	retValJavascriptEncode = Replace(retValJavascriptEncode, "'", "\'")

	myJavascriptEncode = retValJavascriptEncode
End Function

Function getMid(midStr, startStr, endStr)
	addCount = Len(startStr)+7
	getMid = Mid(midStr, InStr(midStr, "<!--" & startStr & "-->")+addCount, InStr(midStr, "<!--" & endStr & "-->")-InStr(midStr, "<!--" & startStr & "-->")-addCount)
End Function

Function getFullMid(midStr, startStr, endStr)
	addCount = (Len(endStr)+7)
	getFullMid = Mid(midStr, InStr(midStr, "<!--" & startStr & "-->"), InStr(midStr, "<!--" & endStr & "-->")-InStr(midStr, "<!--" & startStr & "-->")+addCount)
End Function 

Function GetSelDes

	strSelDes = "0"
	
	If userType = "C" Then
		set oCmd = Server.CreateObject("ADODB.Command")
		oCmd.ActiveConnection = connCommon
		oCmd.CommandType = &H0004
		oCmd.CommandText = "DBOLKGetSelDes" & Session("ID")
		oCmd.Parameters.Refresh()
		set rSelDes = Server.CreateObject("ADODB.RecordSet")
		set rSelDes = oCmd.execute()
		strSelDes = rSelDes(0)
		set rSelDes = Nothing
		set oCmd = Nothing
	End If
	
	GetSelDes = strSelDes
End Function

Sub GetQuery(RecordSet, oType, Var1, Var2)
	If Not Session("OLKAdmin") or Session("OLKAdmin") and Request("dbID") = "" Then dbID = Session("ID") Else dbID = CInt(Request("dbID"))
	set oCmd = Server.CreateObject("ADODB.Command")
	ocmd.ActiveConnection = connCommon
	oCmd.CommandText = "DBOLKGetQuery" & dbID
	oCmd.CommandType = &H0004
	oCmd.Parameters.Refresh()
	oCmd("@Type").Value = oType
	oCmd("@Var1").Value = Var1
	oCmd("@Var2").Value = Var2
	oCmd("@LanID").Value = Session("LanID")
	set RecordSet = oCmd.execute()
End Sub

Sub GetAdminQuery(RecordSet, oType, Var1, Var2)
	set oCmd = Server.CreateObject("ADODB.Command")
	ocmd.ActiveConnection = connCommon
	oCmd.CommandText = "DBOLKGetAdminQuery" & Session("ID")
	oCmd.CommandType = &H0004
	oCmd.Parameters.Refresh()
	oCmd("@Type").Value = oType
	oCmd("@Var1").Value = Var1
	oCmd("@Var2").Value = Var2
	oCmd("@LanID").Value = Session("LanID")
	set RecordSet = oCmd.execute()
End Sub

Sub ClearTableData(TableID, IndexID, Index2ID)
	set oCmd = Server.CreateObject("ADODB.Command")
	ocmd.ActiveConnection = connCommon
	oCmd.CommandText = "DBOLKClearData" & Session("ID")
	oCmd.CommandType = &H0004
	oCmd.Parameters.Refresh()
	oCmd("@TableID") = TableID
	oCmd("@Index") = IndexID
	If Index2ID <> "" Then oCmd("@Index2") = Index2ID
	oCmd.execute()
End Sub

Sub LoadCmd(CommandText)
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DB" & CommandText & Session("ID")
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Refresh()
End Sub

Function CheckAgentClientFilter(CardCode, myType)
	If myApp.AgentClientsFilter <> "" Then
		set oCmdChk = Server.CreateObject("ADODB.Command")
		oCmdChk.ActiveConnection = connCommon
		oCmdChk.CommandText = "DBOLKCheckAgentClientFilter" & Session("ID")
		oCmdChk.CommandType = &H0004
		oCmdChk.Parameters.Refresh()
		oCmdChk("@CardCode") = CardCode
		oCmdChk("@SlpCode") = Session("vendid")
		oCmdChk("@Type") = myType
		oCmdChk.execute()
		CheckAgentClientFilter = CInt(oCmdChk.Parameters.Item(0).value) = 1
	Else
		CheckAgentClientFilter = True
	End If
End Function

Function IsIn(Values, Value)
	myArray = Split(Values, ", ")
	myRetVal = False
	For chkVar = 0 to UBound(myArray)
		If myArray(chkVar) = Value Then
			myRetVal = True
			Exit For
		End If
	Next
	IsIn = myRetVal
End Function

Function GetHTTPStr
	If Request.ServerVariables("HTTPS") = "off" Then GetHTTPStr = "http://" Else GetHTTPStr = "https://"
End Function

Function GetBoolStr(ByVal Val)
	If Val Then GetBoolStr = "True" Else GetBoolStr = "False"
End Function 

Function GetCardType(ByVal Val)
	tmpVal = -1
	
	If Left(Val, 1) = "4" and (Len(Val) = 13 or Len(Val) = 16) Then
		tmpVal = 1 'Visa
	End If
	
	If tmpVal = -1 and Left(Val, 1) = "3" and Len(Val) = 14 Then
		If (Left(Val, 2) = "36" or Left(Val, 2) = "38") or (Left(Val, 3) >= 300 and Left(Val, 3) <= 305) Then
			tmpVal = 3 'Dinners Club
		End If
	End If
	
	If tmpVal = -1 and Left(Val, 1) = "5" and Len(Val) = 16 Then
		If Left(Val, 2) >= 51 and Left(Val, 2) <= 55 Then
			tmpVal = 0 'Master Card
		End If
	End If
	
	If tmpVal = -1 and Left(Val, 1) = "3" and Len(Val) = 15 Then
		If Left(val, 2) = "34" or Left(Val, 2) = "37" Then
			tmpVal = 2 'American Express
		End If
	End If
	
	GetCardType = tmpVal
End Function

Function doPagingStr(str)

	If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" and iPageCount > 1 Then
		pagStr = Left(str, InStr(str, "<!--startPrev-->")-1)
		
		If CatType = "C" Then pagStr = Replace(pagStr, "{ColSpan}", (catCols+1))
		
		pagStr = pagStr & getNextBack(str, 1)
		
		pagStr = pagStr & getMid(str, "endPrevAll", "startCurPage")
		
		If iPageCount > 1 then
			For I = fromI To toI
				If I = iPageCurrent Then
					pagStr = pagStr & Replace(getMid(str, "startCurPage", "endCurPage"), "{iLink}", I)
				Else
					pagStr = pagStr & Replace(getMid(str, "startLinkPage", "endLinkPage"), "{iLink}", I)
				End If
				pagStr = pagStr & "&nbsp;"
			Next
		end if
		
		pagStr = pagStr & getMid(str, "endLinkPage", "startNext")
		
		pagStr = pagStr & getNextBack(str, 2)
		
		pagStr = pagStr & Right(str, Len(str)-InStr(str, "<!--endNextAll-->")-16)
	Else
		pagStr = ""
	End If
	
	doPagingStr = pagStr
End Function

Function getNextBack(str, part)
	retVal = ""
	
	If Session("rtl") = "" and part = 1 or Session("rtl") <> "" and part = 2 Then
		If iPageCurrent > 1 Then
			retVal = retVal & Replace(getMid(str, "startPrev", "endPrev"), "{iLink}", (iPageCurrent - 1))
		End If
		
		If iCurNext > 1 Then
			retVal = retVal & Replace(getMid(str, "startPrevAll", "endPrevAll"), "{iLink}", ((iCurNext-1)*15))
		End If
	Else
		If iPageCurrent < iPageCount Then
			retVal = retVal & Replace(getMid(str, "startNext", "endNext"), "{iLink}", (iPageCurrent + 1))
		End If
		
		If iCurNext < iCurMax Then
			retVal = retVal & Replace(getMid(str, "startNextAll", "endNextAll"), "{iLink}", ((iCurNext*15)+1))
		End If
	End If
	
	getNextBack = retVal
End Function

Function GetWhsCode(ByVal ItemTable)
	retWhs = ""
	If Session("AgentWhs") = "##" Then
		If Not IsNull(Session("BranchWhs")) Then
			retWhs = Session("BranchWhs")
		Else
			retWhs = myApp.WhsCode
		End If
	Else
		retWhs = Session("AgentWhs")
	End If
	
	If myApp.ManageItmWhs Then 
		retWhs = "IsNull(" & ItemTable & ".DfltWH, N'" & saveHTMLDecode(retWhs, False) & "')"
	Else
		retWhs = "N'" & saveHTMLDecode(retWhs, False) & "'"
	End If
	
	GetWhsCode = retWhs
End Function


%>
