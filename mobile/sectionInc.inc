<% 

If Session("useraccess") = "U" Then
	If Not myAut.HasFormAuthorization(Request("secID")) Then Response.Redirect "operaciones.asp?cmd=noaccess&type=form&secID=" & Request("secID")
End If

sql = 	"select IsNull(T1.AlterSecName, T0.SecName) SecName, IsNull(T1.AlterSecContent, T0.SecContent) SecContent, T0.ReqLogin, T0.Status, T0.UserType, " & _
		"T0.Form, IsNull(T1.AlterFormScript, T0.FormScript) FormScript, T0.FormQry, FormQryLoop, " & _
		"IsNull(T1.AlterFormConfirmContent, T0.FormConfirmContent) FormConfirmContent, FormQryRS, SecContentEnableQry, SecContentQry " & _
		"from OLKSections T0 " & _
		"left outer join OLKSectionsAlterNames T1 on T1.SecType = T0.SecType and T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
		"where T0.SecType = 'U' and T0.SecID = " & Request("SecID")
set rs = conn.execute(sql)
If rs("Status") <> "A" or rs("UserType") <> "P" Then Response.Redirect MainDoc & "operaciones.asp?cmd=home"

set rData = Server.CreateObject("ADODB.RecordSet")
SecName = rs("SecName")


	Session("SubmitForm") = True
	SecContent = rs("SecContent")
	SecContent = Replace(SecContent, "<A href", "<A class=""LinkTop"" href")
	SecContent = Replace(SecContent, "<a href", "<A class=""LinkTop"" href")
	SecContent = Replace(SecContent, "{SelDes}", SelDes)
	SecContent = Replace(SecContent, "{dbName}", Session("olkdb"))
	
	For each itm in Request.QueryString
		SecContent = Replace(SecContent, "{" & itm & "}", Request.QueryString(itm))
	Next
	
	sql = "select LineID, Query from OLKSectionsRS where SecType = 'U' and SecID = " & Request("SecID")
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
	
	Response.Write SecContent
	
If Not IsNull(rs("FormScript")) Then Response.Write Replace(rs("FormScript"), "{dbName}", Session("olkdb"))

Sub SectionBindData
	If Not rData.eof Then
	If InStr(SecContent, "startLoop") <> 0 Then
		strData = ""
		do while not rData.eof
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
ElseIf InStr(SecContent, "startLoop") <> 0 Then
	SecContent = Replace(SecContent, getFullMid(SecContent, "startLoop", "endLoop"), "")
End If
End Sub

Function SectionBuildQuery(Query, LoopObj)
	retVal = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
			"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
			"declare @LanID int set @LanID = " & Session("LanID") & " " & Query
	For each itm in LoopObj
		retVal = Replace(retVal, "{" & itm & "}", LoopObj(itm))
	Next
	retVal = QueryFunctions(retVal)
	retVal = Replace(retVal, "{dbID}", Session("ID"))
	SectionBuildQuery = retVal
End Function
 %>