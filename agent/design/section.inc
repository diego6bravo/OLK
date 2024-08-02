<% 

SecID = CInt(Request("SecID"))

If Request("SecID") = "" Then
	Response.Redirect "unauthorized.asp"
ElseIf Session("useraccess") = "U" Then
	If Not myAut.HasFormAuthorization(SecID) Then Response.Redirect "errNoAccess.asp?type=form&secID=" & SecID
End If

sql = 	"select IsNull(T1.AlterSecName, T0.SecName) SecName, IsNull(T1.AlterSecContent, T0.SecContent) SecContent, T0.ReqLogin, T0.Status, T0.UserType, " & _
		"T0.Form, IsNull(T1.AlterFormScript, T0.FormScript) FormScript, T0.FormQry, FormQryLoop, " & _
		"IsNull(T1.AlterFormConfirmContent, T0.FormConfirmContent) FormConfirmContent, FormQryRS, SecContentEnableQry, SecContentQry, ApplyCSS " & _
		"from OLKSections T0 " & _
		"left outer join OLKSectionsAlterNames T1 on T1.SecType = T0.SecType and T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
		"where T0.SecType = 'U' and T0.SecID = " & SecID
set rs = conn.execute(sql)
If rs("ReqLogin") = "Y" and Session("UserName") = "-Anon-" or rs("Status") <> "A" or rs("UserType") <> "C" and userType = "C" or rs("UserType") <> "A" and userType = "V" Then Response.Redirect MainDoc & "?cmd=home"
ApplyCSS = rs("ApplyCSS") = "Y"

set rData = Server.CreateObject("ADODB.RecordSet")

If ApplyCSS Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ObjType") = "S"
	cmd("@ObjID") = 9
	cmd("@UserType") = userType
	set ra = cmd.execute()
	strContent = ra("ObjContent")
	strContent = Replace(strContent, "{SelDes}", SelDes)
	strContent = Replace(strContent, "{rtl}", Session("rtl"))
	strContent = Replace(strContent, "{SecName}", rs("SecName"))
Else
	strContent = "{SecContent}"
End If
If Request.Form.Count = 0 or Request.Form.Count > 0 and Request("newLng") <> "" Then
	Session("SubmitForm") = True
	SecContent = rs("SecContent")
	SecContent = Replace(SecContent, "<A href", "<A class=""LinkTop"" href")
	SecContent = Replace(SecContent, "<a href", "<A class=""LinkTop"" href")
	SecContent = Replace(SecContent, "{SelDes}", SelDes)
	SecContent = Replace(SecContent, "{dbName}", Session("olkdb"))
	
	For each itm in Request.QueryString
		SecContent = Replace(SecContent, "{" & itm & "}", Request.QueryString(itm))
	Next
	
	sql = "select LineID, Query from OLKSectionsRS where SecType = 'U' and SecID = " & SecID & " and Target = 'C'"
	set ra = conn.execute(sql)
	do while not ra.eof
		LineID = ra("LineID")
		If InStr(SecContent, "startRS" & LineID) <> 0 Then
			sql = SectionBuildQuery(ra("Query"), Request.QueryString)
			set rData = conn.execute(sql)

			strData = ""
			
			do while not rData.eof
				strLine = getMid(SecContent, "startRS" & LineID, "endRS" & LineID)
				
				For f = 0 to rData.Fields.Count-1
					If Not IsNull(rData(f)) Then
						strLine = Replace(strLine, "{" & rData(f).Name & "}", rData(f))
					Else
						strLine = Replace(strLine, "{" & rData(f).Name & "}", "")
					End If
				Next
				
				strData = strData & strLine
			rData.movenext
			loop
			
			SecContent = Replace(SecContent, getFullMid(SecContent, "startRS" & LineID, "endRS" & LineID), strData)
		End If
	ra.movenext
	loop
	
	If rs("SecContentEnableQry") = "Y" Then
		sql = SectionBuildQuery(rs("SecContentQry"), Request.QueryString)
		set rData = conn.execute(sql)
		
		SectionBindData
		
	End If

	strContent = Replace(strContent, "{SecContent}", SecContent)

Else
	SecContent = rs("FormConfirmContent")
	SecContent = Replace(SecContent, "<A href", "<A class=""LinkTop"" href")
	SecContent = Replace(SecContent, "<a href", "<A class=""LinkTop"" href")
	SecContent = Replace(SecContent, "{SelDes}", SelDes)
	SecContent = Replace(SecContent, "{dbName}", Session("olkdb"))
	For each itm in Request.Form
		SecContent = Replace(SecContent, "{" & itm & "}", Request(itm))
	Next
	
	If Session("SubmitForm") or rs("FormQryRS") = "Y" Then
		If Request("LoopID") <> "" Then loopGuide = Split(Request("LoopGuide"), ", ")
		sql = SectionBuildQuery(rs("FormQry"), Request.Form)
		If rs("FormQryRS") = "N" Then
			conn.execute(sql)
		Else
			set rData = conn.execute(sql)
			SectionBindData
		End If
		If Request.Form("LoopType") <> "" Then LoopType = Request.Form("LoopType") Else LoopType = "G"
		If Session("SubmitForm") Then
			If Request("LoopID") <> "" Then
				loopID = Split(Request("LoopID"), ", ")
				If LoopType = "A" Then
					Redim loopValues(UBound(loopGuide))
					For i = 1 to UBound(loopGuide)
						If Request(loopGuide(i)) <> "" Then 
							loopValues(i) = Split(Request(loopGuide(i)), ", ") 
						Else 
							Redim cleanArr(UBound(loopID))
							loopValues(i) = cleanArr
						End If
					Next

				End If
				For i = 0 to UBound(loopID)
					sql = SectionBuildQuery(rs("FormQryLoop"), Request.Form)
					If rs("FormQryRS") <> "N" Then
						If Not rData.eof Then
							For f = 0 to rData.Fields.Count-1
								If Not IsNull(rData(f)) Then
									sql = Replace(sql, "{" & rData(f).Name & "}", rData(f))
								Else
									sql = Replace(sql, "{" & rData(f).Name & "}", "")
								End If
							Next
						End If
					End If

					For j = 0 to UBound(loopGuide)
						If j = 0 Then
							sql = Replace(sql, "{" & loopGuide(j) & "}", loopID(i))
						Else
							Select Case LoopType
								Case "G" 'Guided
									sql = Replace(sql, "{" & loopGuide(j) & "}", Request(loopGuide(j) & loopID(i)))
								Case "A"
									sql = Replace(sql, "{" & loopGuide(j) & "}", loopValues(j)(i))
							End Select
						End If
					Next
					conn.execute(sql)
				Next
			End If
		End If
		Session("SubmitForm") = False
	End If
	
	sql = "select LineID, Query from OLKSectionsRS where SecType = 'U' and SecID = " & SecID & " and Target = 'F'"
	set ra = conn.execute(sql)
	do while not ra.eof
		LineID = ra("LineID")
		If InStr(SecContent, "startRS" & LineID) <> 0 Then
			sql = SectionBuildQuery(ra("Query"), Request.Form)
			set rData = conn.execute(sql)

			strData = ""
			
			do while not rData.eof
				strLine = getMid(SecContent, "startRS" & LineID, "endRS" & LineID)
				
				For f = 0 to rData.Fields.Count-1
					If Not IsNull(rData(f)) Then
						strLine = Replace(strLine, "{" & rData(f).Name & "}", rData(f))
					Else
						strLine = Replace(strLine, "{" & rData(f).Name & "}", "")
					End If
				Next
				
				strData = strData & strLine
			rData.movenext
			loop
			
			SecContent = Replace(SecContent, getFullMid(SecContent, "startRS" & LineID, "endRS" & LineID), strData)
		End If
	ra.movenext
	loop

	strContent = Replace(strContent, "{SecContent}", SecContent)
End If
Response.Write strContent
If Not IsNull(rs("FormScript")) Then Response.Write Replace(rs("FormScript"), "{dbName}", Session("olkdb"))


Sub SectionBindData
	If Not rData.eof Then
		recCount = 0
		If InStr(SecContent, "startLoop") <> 0 Then
			strData = ""
			do while not rData.eof
				recCount = recCount + 1
				strLine = getMid(SecContent, "startLoop", "endLoop")
				For f = 0 to rData.Fields.Count-1
					If Not IsNull(rData(f)) Then
						strLine = Replace(strLine, "{" & rData(f).Name & "}", rData(f))
					Else
						strLine = Replace(strLine, "{" & rData(f).Name & "}", "")
					End If
				Next
				strData = strData & strLine
			rData.movenext
			loop
			rData.movefirst
			SecContent = Replace(SecContent, getFullMid(SecContent, "startLoop", "endLoop"), strData)
		Else
			For f = 0 to rData.Fields.Count-1
				If Not IsNull(rData(f)) Then
					SecContent = Replace(SecContent, "{" & rData(f).Name & "}", rData(f))
				Else
					SecContent = Replace(SecContent, "{" & rData(f).Name & "}", "")
				End If
			Next
		End If
		If InStr(SecContent, "startNoData") <> 0 Then
			SecContent = Replace(SecContent, getFullMid(SecContent, "startNoData", "endNoData"), "")
		End If
		If InStr(SecContent, "startRecordCount") <> 0 Then
			SecContent = Replace(SecContent, "{RecordCount}", recCount)
		End If
	Else
		If InStr(SecContent, "startLoop") <> 0 Then
			SecContent = Replace(SecContent, getFullMid(SecContent, "startLoop", "endLoop"), "")
		End If
		If InStr(SecContent, "startRecordCount") <> 0 Then
			SecContent = Replace(SecContent, getFullMid(SecContent, "startRecordCount", "endRecordCount"), "")
		End If
	End If
End Sub

Function SectionBuildQuery(Query, LoopObj)
	retVal = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
			"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
			"declare @branchIndex int set @branchIndex = " & Session("branch") & " " & _
			"declare @LanID int set @LanID = " & Session("LanID") & " " & Query
	If Request("LoopGuide") = "" Then
		For each itm in LoopObj
			retVal = Replace(retVal, "{" & itm & "}", LoopObj(itm))
		Next
	Else
		For each itm in LoopObj
			guideFound = False
			For each guide in loopGuide
				If itm = guide Then
					guideFound = True
					Exit For
				End if
			Next
			If Not guideFound Then
				retVal = Replace(retVal, "{" & itm & "}", LoopObj(itm))
			End If
		Next
	End If
	retVal = QueryFunctions(retVal)
	retVal = Replace(retVal, "{dbID}", Session("ID"))
	SectionBuildQuery = retVal
End Function

%>