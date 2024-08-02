<% addLngPathStr = "portal/" %>
<% If viewRepPDF Then addLngPathStr = "" %>
<!--#include file="lang/viewReport.asp" -->
<% 
Session("RepLogNum") = ""

If Not viewRepPDF and Request("pop") <> "Y" Then
	curUrl = GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"report.asp","")
Else
	If (Request("itemSmallRep") = "Y" or Request("pop") = "Y") and Not viewRepPDF Then
		curUrl = GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"viewreportprint.asp","")
	Else
		curUrl = GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"portal/viewreportpdf.asp","")
	End If
End If

LtxtLinkToRep = getviewReportLngStr("LtxtLinkToRep")
rsIndex = CInt(Request("rsIndex"))

HaveVals = False

set rs = server.createobject("ADODB.RecordSet")
set rd = server.createobject("ADODB.RecordSet")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRSExec" & Session("ID")
cmd.Parameters.Refresh()
cmd("@rsIndex") = rsIndex
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid")
set rs = cmd.execute()
rsName = rs("rsName")
rsDesc = rs("rsDesc")
sqlQuery = QueryFunctions(rs("rsQuery"))
rsTop = rs("rsTop") = "Y"
Refresh = rs("Refresh")
setTotal = rs("Total") = "Y"
setTotalTop = rs("TotalTop") = "Y"
setTotalBottom = rs("TotalBottom") = "Y"
HCols = rs("HCols")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRSExecVars" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@rsIndex") = rsIndex
set rVal = Server.CreateObject("ADODB.RecordSet")
rVal.open cmd, , 3, 1

varCount = rVal.recordcount
conn.CommandTimeout = 0

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKRS_" & Session("ID") & "_" & Replace(rsIndex, "-", "_")
LoadViewRepParams
rs.close
rs.open cmd, , 3, 1

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRSExecTotals" & Session("ID")
cmd.Parameters.Refresh()
cmd("@rsIndex") = rsIndex
cmd("@LanID") = Session("LanID")
If Session("RetVal") <> "" Then cmd("@LogNum") = Session("RetVal")
cmd("@UserAccess") = Session("useraccess")
cmd("@UserType") = userType
cmd("@SlpCode") = Session("vendid")
If myAut.AuthorizedRepGroups <> "" Then cmd("@Groups") = myAut.AuthorizedRepGroups
rd.open cmd, , 3, 1

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRSExecLinksVars" & Session("ID")
cmd.Parameters.Refresh()
cmd("@rsIndex") = rsIndex
set rl = Server.CreateObject("ADODB.RecordSet")
rl.open cmd, , 3, 1

If setTotal Then 
	set rt = server.createobject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKRS_" & Session("ID") & "_" & Replace(rsIndex, "-", "_") & "_total"
	LoadViewRepParams
	set rt = cmd.execute()
End If

FldFilter = ""
For each Field in rs.Fields
	If FldFilter <> "" Then FldFilter = FldFilter & ", "
	FldFilter = FldFilter & Field.Name
Next

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetExecColors" & Session("ID")
cmd.Parameters.Refresh()
cmd("@rsIndex") = rsIndex
cmd("@FldFilter") = FldFilter
set rc = Server.CreateObject("ADODB.RecordSet")
rc.open cmd, , 3, 1


If Err.Number = 0 Then

repColCount = rs.Fields.Count


Dim repCols()
Redim repCols(repColCount, 18)
For i = 0 to repColCount - 1
	colName = rs.Fields(i).Name
	rd.Filter = "colName = '" & Replace(colName, "'", "''") & "'"
	
	ignoreRowFormat = False
	DoLTR = False
	colCSS = ""
	
	colLinksChk = ""
	If Not rd.Eof Then
		repCols(i, 0) = rd("colFormat")
		repCols(i, 1) = rd("colTotal")
		repCols(i, 2) = rd("colSum")
		repCols(i, 3) = rd("colShow")
		repCols(i, 4) = rd("linkType")
		repCols(i, 5) = rd("linkObject")
		repCols(i, 6) = rd("linkLink")
		repCols(i, 7) = rd("linkPopup")
		repCols(i, 8) = rd("linkCat")
		repCols(i, 9) = rd("linkObjectTtl")
		repCols(i, 10) = rd("linkObjectAccess")
		repCols(i, 11) = rd("colTitle")
		Select Case repCols(i, 4)
			Case "F", "L"
				For each fld in rs.Fields
					If InStr(repCols(i, 6), fld.Name) <> 0 Then
						If colLinksChk <> "" Then colLinksChk = colLinksChk & ", "
						colLinksChk = colLinksChk & fld.Name
					End If
				Next
			Case Else
				rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and valBy = 'F'"
				do while not rl.eof
					If colLinksChk <> "" Then colLinksChk = colLinksChk & ", "
					colLinksChk = colLinksChk & rl("valValue")
				rl.movenext
				loop
		End Select
		repCols(i, 13) = rd.bookmark
		repCols(i, 14) = rl.RecordCount > 0 or repCols(i, 4) = "F" or repCols(i, 4) = "L"
	Else
		repCols(i, 0) = "N"
		repCols(i, 1) = "D"
		repCols(i, 2) = "N"
		repCols(i, 3) = "Y"
		repCols(i, 4) = "N"
		repCols(i, 5) = ""
		repCols(i, 6) = ""
		repCols(i, 7) = ""
		repCols(i, 8) = ""
		repCols(i, 9) = ""
		repCols(i, 10) = ""
		repCols(i, 11) = colName
		ShowCol = True
		repCols(i, 14) = False
	End If
	repCols(i, 12) = colName
	repCols(i, 15) = colLinksChk
Next
 %>
<div align="center">
	<table border="0" cellpadding="0" width="100%" style="font-family: Verdana; font-size: 10px">
		<% If tblCustTtl = "" Then %>
		<tr class="TblTlt">
			<td id="tdMyTtl">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td class="TblTlt">&nbsp;<%=getviewReportLngStr("LttlReps")%>
					</td>
					<% If Not viewRepPDF and userType = "C" Then %><td width="100" align="right"><!--#include file="../searchInc/repLegend.asp"--></td><% End If %>
				</tr>
			</table>
			</td>
		</tr>
		<% Else %>
		<%=Replace(tblCustTtl, "{txtTitle}", getviewReportLngStr("LttlReps"))%>
		<% End If %>
		<tr class="TablasNoticiasTitle">
			<td>
			<%=rsName%>&nbsp;</td>
		</tr>
		<% If rsDesc <> "" and Not IsNull(rsDesc) Then %>
		<tr class="TablasNoticias">
			<td>
			<%=rsDesc%>&nbsp;</td>
		</tr>
		<% End If %>
		<tr class="TablasNoticias">
			<td>
			<b><%=getviewReportLngStr("DtxtExecTime")%>:</b> <%=FormatDate(Now(), True)%>&nbsp;<%=FormatTime(Now())%>&nbsp;</td>
		</tr>
	</table>
	<table border="0" cellpadding="0" width="100%" style="font-family: Verdana; font-size: 10px">
		<tr>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" width="100%"style="font-family: Verdana; font-size: 10px">
				<% 
				HaveVals = rVal.recordcount > 0
				rVal.Filter = "varShowRep = 'Y'"
				If Not rVal.eof Then
					If Request("Excell") <> "Y" and Request("pdf") <> "Y" Then
						rd.Filter = "LinkType <> 'N'"
						linkCount = rd.recordcount
					Else
						linkCount = 0
					End If
					do while not rVal.eof
						%>
							<tr class="TablasNoticias">
								<td colspan="<%=repColCount+linkCount%>"><b><%=rVal("varName")%>:</b><%
								Select Case rVal("varType")
									Case "DD", "L", "CL"
										Response.Write Request("var" & rVal("varIndex") & "Desc")
									Case Else 
										Response.Write Request("var" & rVal("varIndex"))
								End Select %></td>
							</tr>
						<% rVal.movenext
						loop
					End If %>
					<tr class="<% 
					Select Case userType
						Case "V" %>FirmTlt3<% 
						Case "C" %>TblTlt<% 
					End Select%>">
				<% 	
					ShowLinks = Request("Excell") <> "Y" and Request("pdf") <> "Y"
					For i = 0 to repColCount - 1
			   			ShowCol = repCols(i, 0) <> "H"
			   			fldName = repCols(i, 12)
			   			colLink = repCols(i, 4) <> "N" or repCols(i, 0) = "EML"
			   			If colLink and ShowLinks Then hCols = hCols - 1
				   		If ShowCol Then
				   			If colLink and ShowLinks Then %><td>&nbsp;</td><% End If
						 %><td align="center"><%=repCols(i, 11)%></td>
				<% 		End If
				   		next %>
					</tr>
					<% If setTotal and setTotalTop Then
						rd.Filter = "colShow = 'A' or colShow = 'T'"
						If rd.RecordCount > 0 Then
							ShowRepTotal "T" 
							%><tr class="TablasNoticias">
								<td colspan="<%=repColCount-hCols%>" class="RepTotalCel" height="4">&nbsp;</td>
							</tr><%
						End If 
					End If 
				    rd.Filter = "colSum = 'Y'"
					Redim repColSum(rd.RecordCount)
					For i = 0 to rd.RecordCount - 1
						repColSum(i) = 0.0
					Next
					strRow = "<tr class=""TablasNoticias"">"
					For i = 0 to repColCount - 1
						colName = rs(i).Name
						rd.Filter = "colName = '" & Replace(colName, "'", "''") & "'"
						ignoreRowFormat = False
						DoLTR = False
						colCSS = ""
						
						If Not rd.Eof Then
							ShowCol = repCols(i, 0) <> "H"
							DoLTR = repCols(i, 0) = "IS" or repCols(i, 0) = "IB"
							colCSS = "rscol_" & repCols(i, 13) & " "
						End If
						If ShowCol Then
							If (repCols(i, 4) <> "N" or repCols(i, 0) = "EML") and not (Request("Excell") = "Y" or Request("pdf") = "Y") Then
	
							Select Case repCols(i, 4)
								Case "A"
									Select Case repCols(i, 5)
										Case 4
											linkImg = "design/" & SelDes & "/images/shop_icon.gif"
										Case 5
											linkImg = "design/" & SelDes & "/images/heart_icon.gif"
										Case 6
											linkImg = "design/" & SelDes & "/images/x_icon.gif"
										Case Else
											linkImg = "images/action_" & repCols(i, 5) & ".gif"
									End Select
									Select Case repCols(i, 5)
										Case 0
											strAlt = getviewReportLngStr("DtxtApprove")
											Action0 = True
										Case 3
											strAlt = getviewReportLngStr("DtxtCancel")
											Action3 = True
										Case 2
											strAlt = getviewReportLngStr("DtxtClose")
											Action2 = True
										Case 6
											strAlt = getviewReportLngStr("DtxtDelete")
											Action6 = True
										Case 1
											strAlt = getviewReportLngStr("LtxtConvQuoteSales")
											Action1 = True
										Case 5
											strAlt = getviewReportLngStr("LtxtAddToWish")
											Action5 = True
										Case 4
											strAlt = getviewReportLngStr("LtxtAddToCart")
											Action4 = True
										Case 7
											strAlt = getviewReportLngStr("LtxtConvOrderInv")
											Action7 = True
									End Select
								Case Else
									If repCols(i, 0) <> "EML" Then
										linkImg = "design/" & SelDes & "/images/" & Session("rtl") & "flecha_selec.gif"
									Else
										linkImg = "images/mail.gif"
									End If
							End Select

							If repCols(i, 0) <> "EML" Then
								strRowLink = "<td width=""16"" class=""" & colCSS & "{cvc" & colName & "}""><img src=""" & imgAddPath & linkImg & """ style=""cursor: hand"" onclick=""javascript:"

								Select Case repCols(i, 4)
									Case "L", "F"
										myLink = repCols(i, 6)
										For each myLnkItm in rs.Fields
											myLink = Replace(myLink, "{" & myLnkItm.Name & "}", "{cv" & myLnkItm.Name & "Lnk}")
										Next
										Select Case repCols(i, 4) 
											Case "F" 
												myLink = "sec.asp?SecID=" & repCols(i, 5) & "&" & myLink
												strRowLink = strRowLink & "window.location.href='" & myLink & "';"
											Case "L"
												strRowLink = strRowLink & "window.open('" & myLink & "');"
										End Select
									Case "O"
										rl.Filter = "colName = '" & Replace(colName, "'", "''") & "'"
										Select Case repCols(i, 5)
											Case -5
												rl.movenext
												myLink = myLink & "doDetail("
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case "V" 
														myLink = myLink & rl("valValue")
												End Select
												rl.movefirst
												isEntry = "N"
											Case Else
												myLink = "doDetail(" & repCols(i, 5)
												isEntry = "Y"
										End Select
										myLink = myLink & ", '"
										
										Select Case rl("valBy") 
											Case "F"
												myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
											Case "V" 
												myLink = myLink & rl("valValue")
										End Select
										
										high = ""
										If rl.recordcount > 1 and repCols(i, 5) <> 2 and repCols(i, 5) <> 4 and InStr("112,140,23,17,15,16,13,14,22,20,21,18,19,-5", repCols(i, 5)) <> 0 THen
											rl.movenext
											Select Case rl("valBy") 
												Case "F" 
													high = "{high" & rl("valValue") & "}"
												Case "V" 
													high = rl("valValue")
											End Select
										End If

										myLink = myLink & "', '" & repCols(i, 8) & "', '" & isEntry & "', '" & high & "', '" & repCols(i, 7) & "');"
										strRowLink = strRowLink & myLink
									Case "A"
										Select Case repCols(i, 5)
											Case 4
												myLink = "goAddItem('"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'ItemCode'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Quantity'"
												If Not rl.Eof Then
													Select Case rl("valBy") 
														Case "F"
															myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
														Case "V"
															myLink = myLink & rl("valValue")
													End Select
												Else
													myLink = myLink & "1"
												End If
												myLink = myLink & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Unit'"
												If Not rl.Eof Then
													Select Case rl("valBy") 
														Case "F"
															myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
														Case "V"
															myLink = myLink & rl("valValue")
													End Select
												Else
													myLink = myLink & "''"
												End If
												myLink = myLink & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Price'"
												If Not rl.Eof Then
													Select Case rl("valBy") 
														Case "F"
															myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
														Case "V"
															myLink = myLink & rl("valValue")
													End Select
												Else
													myLink = myLink & "''"
												End If
												myLink = myLink & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Locked'"
												If Not rl.Eof Then
													Select Case rl("valBy") 
														Case "F"
															myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
														Case "V"
															myLink = myLink & rl("valValue")
													End Select
												Else
													myLink = myLink & "', '"
												End If
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'WhsCode'"
												If Not rl.Eof Then
													Select Case rl("valBy") 
														Case "F"
															myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
														Case "V"
															myLink = myLink & Replace(rl("valValue"), "'", "\'")
													End Select
												Else
													myLink = myLink & ""
												End If
												
												myLink = myLink & "');"
											Case 5
												myLink = "goAddWish('"
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'ItemCode'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else 
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "');"
											Case 0
												myLink = "goApproveOrder('"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Entry'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else 
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "');"
											Case 1, 7
												Select Case repCols(i, 5)
													Case 1
														myLink = "goConvQuoteOrder('"
													Case 7
														myLink = "goConvOrderInvoice('"
												End Select
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Entry'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else 
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Series'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else 
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "');"
											Case 2, 3, 6
												myLink = "goObjAct('" & repCols(i, 5) & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'ObjectCode'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else 
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "', '"
												
												rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and varId = 'Entry'"
												Select Case rl("valBy") 
													Case "F" 
														myLink = myLink & "{cv" & rl("valValue") & "Lnk}"
													Case Else 
														myLink = myLink & rl("valValue")
												End Select
												myLink = myLink & "');"
										End Select
										strRowLink = strRowLink & myLink & """  alt=""" & strAlt
									Case "R"
										If repCols(i, 10) = "Y" Then
											myLink = "goRep" & repCols(i, 13) & "("
											rl.Filter = "colName = '" & Replace(colName, "'", "''") & "' and valBy = 'F'"
											myLinkVars = ""
											do while not rl.eof
												If myLinkVars <> "" Then myLinkVars = myLinkVars & ", "
												myLinkVars = myLinkVars & "'{cv" & rl("valValue") & "Lnk}'"
											rl.movenext
											loop
											myLink = myLink & myLinkVars & ");"
										Else
											myLink = myLink & "alert('" & getviewReportLngStr("LtxtRepNoAccess") & "');"
										End If
										myLink = myLink & """ alt=""" & LtxtLinkToRep & ": " & repCols(i, 9)
										strRowLink = strRowLink & myLink
								End Select
								strRowLink = strRowLink & """>"
							Else
								strRowLink = "<td width=""16"" class=""" & colCSS & "{cvc" & colName & "}""><a href=""mailto:{cv" & colName & "}""><img border=""0"" src=""" & imgAddPath & linkImg & """</a></td>"
							End If

							strRow = strRow & "{ColLink" & colName & "}"
							repCols(i, 16) = strRowLink
							repCols(i, 17) = "<td width=""16"" class=""" & colCSS & "{cvc" & colName & "}""></td>"
							
							End If
							
							strRow = strRow & "<td "
							If DoLTR Then strRow = strRow & "dir=""ltr"" "
							strRow = strRow & "class=""" & colCSS & "{cvc" & colName & "}"">{cv" & colName & "}</td>"
						End If
						
					next
					strRow = strRow & "</tr>"
					
					do while not rs.eof
						strLine = strRow
						rowCSS = doGetRSColCSS(ignoreRowFormat, true) & " "
						curColSum = 0
						For i = 0 to repColCount - 1
							fldVal = rs(i)
							colName = rs(i).Name
							ignoreRowFormat = False
							DoColorRowBlink = False
							rd.Filter = "colName = '" & Replace(colName, "'", "''") & "'"
							rsRowCSS = doGetRSColCSS(ignoreRowFormat, false)
							
							If repCols(i, 2) = "Y" Then
								If IsNull(fldVal) Then colSumVal = 0 Else colSumVal = CDbl(fldVal)
								repColSum(curColSum) = repColSum(curColSum) + colSumVal
								fldVal = repColSum(curColSum)
								curColSum = curColSum + 1
								ignoreRowFormat = True
							End If
							
							If Not IsNull(fldVal) and Request("Excell") <> "Y" Then 
								fldVal = GetRepFormatVal(fldVal, repCols(i, 0), False)
							ElseIf IsNull(fldVal) Then
								fldVal = ""
							End If
							
							If DoColorRowBlink Then fldVal = "<blink>" & fldVal & "</blink>"
							
							strLine = Replace(strLine, "{cv" & colName & "}", fldVal)

							If repCols(i, 17) <> "" Then
								cancelLink = False
								strLink = repCols(i, 16)
								If repCols(i, 14) Then
									arrValCol = Split(repCols(i, 15), ", ")
									For c = 0 to UBound(arrValCol)
										If IsNull(rs(arrValCol(c))) Then 
											cancelLink = True
											Exit For
										End If
										Select Case rs(arrValCol(c)).Type
											Case 135
												strLink = Replace(strLink, "{cv" & arrValCol(c) & "Lnk}", FormatDateTime(rs(arrValCol(c)), 2))
											Case Else
												strLink = Replace(strLink, "{cv" & arrValCol(c) & "Lnk}", Replace(Server.HTMLEncode(rs(arrValCol(c))), "'", "\'"))
										End Select 
									Next
								ElseIf repCols(i, 0) = "EML" Then
									If Not IsNull(rs(colName)) and rs(colName) <> "" Then
										strLink = Replace(strLink, "{cv" & colName & "Lnk}", Replace(Server.HTMLEncode(rs(colName)), "'", "\'"))
									Else
										cancelLink = True
									End If
								End If
								If Not cancelLink Then
									strLine = Replace(strLine, "{ColLink" & colName & "}", strLink)
								Else
									strLine = Replace(strLine, "{ColLink" & colName & "}", repCols(i, 17))
								End If
							End If
							
							strLine = Replace(strLine, "{cvc" & colName & "}", rowCSS & rsRowCSS) 

						Next
						Response.Write strLine & VbCrLf
					rs.movenext
					loop
					
					
				    If setTotal and setTotalBottom Then
					rd.Filter = "colShow = 'A' or colShow = 'B'"
					If rd.RecordCount > 0 Then  %>
					<tr class="TablasNoticias">
						<td colspan="<%=repColCount-hCols%>" class="RepTotalCel" height="4">&nbsp;</td>
					</tr>
					<% ShowRepTotal "B"
					End If %>
				<% End If %>
				</table>
			</div>
			</td>
		</tr>
	</table>
</div>

<% If userType = "C" and Not viewRepPDF Then %>
<form action="report.asp" method="post" name="frmReload">
<% For each item in Request.Form
If item <> "btnReload" Then %>
<input type="hidden" name="<%=item%>" value="<%=Request(item)%>">
<% End If
Next
For each item in Request.QueryString %>
<input type="hidden" name="<%=item%>" value="<%=Request(item)%>">
<% Next %>
<div align="left">
	<table border="0" cellpadding="0" width="93%" id="tblRepButtons">
		<tr>
			<td>
			<input type="submit" value="<%=getviewReportLngStr("DtxtUpdate")%>" name="btnReload" style="width: 100px"></td>
		</tr>
		<tr>
			<% If HaveVals Then %><td>
			<input type="button" value="<%=getviewReportLngStr("LtxtQueryAgain")%>" name="B2" onclick="javascript:<% If userType = "V" Then %>Pic('portal/viewRepVals.asp?rsIndex=<%=Request("rsIndex")%>', 368, 402, 'Yes', 'no')<% ElseIf userType = "C" Then %>	doMyLink('viewRepValsC.asp', 'rsIndex=<%=Request("rsIndex")%>', '');<% End If %>" style="width: 100px"></td><% End If %>
		</tr>
	</table>
</div>
<input type="hidden" name="Excell" value="N">
</form>
<% End If %>

<% Else %>
<div align="center"><font color="red" face="Verdana" size="2"><p><%=getviewReportLngStr("LtxtQueryErr")%></p><p><font color="black"><%=Err.Description%></font></p></font></div>
<% End If %>
<%
Function doGetRSColCSS(ByVal ColSum, ByVal Row)
	LastColorID = -1
	colFormat = ""
	If Not Row and Not ColSum Then
		rc.Filter = "((ColName = '" & Replace(colName, "'", "''") & "' and ApplyToCol = '' and ApplyToRow = 'N') or ApplyToCol = '" & Replace(colName, "'", "''") & "')"
	ElseIf Row Then
		rc.Filter = "ApplyToRow = 'Y'"
	ElseIf ColSum Then
		rc.Filter = "(ColName = '" & Replace(colName, "'", "''") & "' and ApplyToCol = '') and ApplyToRow = 'N'"
	End If
	If rc.recordcount > 0 Then
		rc.movefirst
		do while not rc.eof
			If CInt(LastColorID) <> rc("ColorID") Then
			colFormat = AnalizeFormat(rc("ColorID"), rc("LineID"), CStr(rc("ColName")), colFormat, rc("colOp"), rc("colOpBy"), _
						rc("colValue"), rc("colType"), rc("colValDate"), rc("FontBlink"))
			End If
		rc.movenext
		loop
	End If
	doGetRSColCSS = colFormat 
End Function

Function AnalizeFormat(ByVal ColorID, ByVal LineID, ByVal ColName, ByVal curFormat, ByVal colOp, ByVal colOpBy, ByVal colValue, ByVal colType, ByVal colValueDate, ByVal FontBlink)
	colFormat = curFormat
	If colOp <> "N" and colOp <> "NN" Then 
		Select Case colOpBy
			Case "F"
				If IsNull(rs(CStr(colValue))) Then
					AnalizeFormat = colFormat
					Exit Function
				End If
				Select Case colType
					Case "N" 
						compVal = CDbl(rs(CStr(colValue)))
					Case Else
						compVal = CStr(rs(CStr(colValue)))
				End Select
			Case Else
				Select Case colType 
					Case "D" 
						If IsNull(colValueDate) Then
							AnalizeFormat = colFormat
							Exit Function
						End If
						compVal = colValueDate
					Case "N" 
						If IsNull(colValue) Then
							AnalizeFormat = colFormat
							Exit Function
						End If
						compVal = CDbl(colValue)
					Case Else
						If IsNull(colValue) Then
							AnalizeFormat = colFormat
							Exit Function
						End If
						compVal = CStr(colValue)
				End Select
		End Select
	End If
	Select Case repCols(i, 2)
		Case "Y" 
			If IsNull(strVal) Then
				AnalizeFormat = colFormat
				Exit Function
			End If
			Select Case colType
				Case "N" 
					mainVal = CDbl(strVal)
				Case "D" 
					mainVal = strVal
				Case Else
					mainVal = CStr(strVal)
			End Select
		Case Else
			If IsNull(rs(ColName)) Then
				AnalizeFormat = colFormat
				Exit Function
			End If
			Select Case colType
				Case "N" 
					mainVal = CDbl(rs(ColName))
				Case "D" 
					mainVal = rs(ColName)
				Case Else
					mainVal = CStr(rs(ColName)) 
			End Select
	End Select

	apFormat = False
	Select Case colOp
		Case "="
			apFormat = mainVal = compVal
		Case "<>"
			apFormat = mainVal <> compVal
		Case ">"
			apFormat = mainVal > compVal
		Case "<"
			apFormat = mainVal < compVal
		Case ">="
			apFormat = mainVal >= compVal
		Case "<="
			apFormat = mainVal <= compVal
		Case "N"
			apFormat = IsNull(mainVal)
		Case "NN"
			apFormat = Not IsNull(mainVal)
	End Select
	If apFormat Then 
		DoColorRowBlink = FontBlink = "Y"
		colFormat = colFormat & " rs_" & ColorID & "_" & LineID
		LastColorID = ColorID
	End If

	AnalizeFormat = colFormat
End Function 


Sub LoadViewRepParams
	If rVal.recordcount > 0 Then rVal.movefirst
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	Select Case userType
		Case "C"
			cmd("@CardCode") = Session("UserName")
		Case "V"
			cmd("@SlpCode") = Session("vendid")
	End Select
	
	If rsTop Then cmd("@top") = CInt(Request("varTop"))
	
	do while not rVal.eof
		varIndex = rVal("varIndex")
		varVar = "@" & rVal("varVar")
		Select Case rVal("varType")
			Case "CL"
				cmd(varVar) = Request("var" & varIndex)
			Case Else
				If Request("var" & varIndex) <> "" Then
					Select Case rVal("varDataType")
						Case "nvarchar"
							cmd(varVar) = Request("var" & varIndex)
						Case "datetime"
							cmd(varVar) = SaveCmdDate(Request("var" & varIndex))
						Case "numeric"
							cmd(varVar) = CDbl(getNumericOut(Request("var" & varIndex)))
						Case "int"
							cmd(varVar) = CLng(Request("var" & varIndex))
					End Select
				End If
		End Select
	rVal.movenext
	loop
End Sub

Sub ShowRepTotal(ByVal Pos) %><tr class="TablasNoticias"><% 
	For i = 0 to repColCount - 1
		rd.Filter = "colName = '" & Replace(rs(i).Name, "'", "''") & "'"
		Blank = True
		colLink = False
		If not rd.Eof Then
			ShowCol = repCols(i, 0) <> "H"
			BlankCol = ShowCol and rd("colShow") <> "A" and rd("colShow") <> Pos
   			colLink = repCols(i, 4) <> "N" or repCols(i, 0) = "EML"
		Else
			ShowCol = True
		End if
		If ShowCol Then
			If Not BlankCol Then %>
			<td <% If Not rd.Eof Then %>class="rscol_<%=repCols(i, 13)%>" <% If colLink and ShowLinks Then %>colspan="2"<% End If %><% End If %>>
			<% 	
				Field = rt(rs(i).name)
				strVal = Field
				If Not rd.Eof and Not IsNull(Field) and Field <> "" and Request("Excell") <> "Y" Then 
					strVal = GetRepFormatVal(Field, repCols(i, 0), True)
				End If
				Response.Write strVal%></td><% Else 
%><td <% If Not rd.Eof Then %><% If colLink and ShowLinks Then %>colspan="2"<% End If %><% End If %>>&nbsp;</td>
<% End If
   End If
   next %></tr><% 
End Sub
Function GetRepFormatVal(ByVal strVal, ByVal colFormat, ByVal IsTotal)
	strGetRepFormatVal = strVal
	Select Case colFormat
		Case "R"
			strGetRepFormatVal= FormatNumber(CDbl(strVal),myApp.RateDec)
		Case "S"
			strGetRepFormatVal= FormatNumber(CDbl(strVal),myApp.SumDec)
		Case "P"
			strGetRepFormatVal= FormatNumber(CDbl(strVal),myApp.PriceDec)
		Case "Q"
			strGetRepFormatVal= FormatNumber(CDbl(strVal),myApp.QtyDec)
		Case "%"
			strGetRepFormatVal= FormatNumber(CDbl(strVal),myApp.PercentDec)
		Case "M"
			GetRepFormatVal = FormatNumber(CDbl(strVal),myApp.MeasureDec)
		Case "DT1", "DT2", "DT3", "DT4"
				Select Case CInt(Right(colFormat,1))
					Case 1
						strGetRepFormatVal= FormatDateTime(strVal, 1)
					Case 2
						strGetRepFormatVal= FormatDate(strVal, true)
					Case 3
						strGetRepFormatVal= FormatTime(strVal)
					Case 4
						strGetRepFormatVal = FormatDateTime(strVal, 4)
				End Select
		Case "IS"
			If Not IsTotal Then If Not IsNull(rs("NumInSale")) Then strGetRepFormatVal = NumInItem(CDbl(strVal), CDbl(rs("NumInSale"))) Else strGetRepFormatVal = ""
		Case "IB"
			If Not IsTotal Then If Not IsNull(rs("NumInBuy")) Then strGetRepFormatVal = NumInItem(CDbl(strVal), CDbl(rs("NumInBuy"))) Else strGetRepFormatVal = ""
		Case "IM"
			If Not IsTotal Then strGetRepFormatVal = "<img src=""" & curUrl & "pic.aspx?filename=" & strVal & "&MaxSize=80&dbName=" & Session("olkdb") & """>"
	End Select
	GetRepFormatVal = strGetRepFormatVal
End Function

Function NumInItem(ByVal Qty, ByVal QtyInItem)
	P1 = Fix(Qty/QtyInItem)
	P2 = CInt((Qty/QtyInItem-Fix(Qty/QtyInItem))*QtyInItem)
	RetVal = P1
	If QtyInItem <> 1 Then RetVal = RetVal & "(" & P2 & ")"
	NumInItem = RetVal
End Function

If Session("RetVal") <> "" Then 
	itemCmd = "a"
ElseIf Request("itemCmd") = "" Then 
	itemCmd = "d"
Else 
	itemCmd = Request("itemCmd")
End If
If userType = "V" Then itemCmd = UCase(itemCmd)
%>
<script language="javascript">
var imgAddPath = '<%=imgAddPath%>';
var UserType = '<%=userType%>';
var repItemCmd = '<%=itemCmd%>';
var UserName = '<%=Replace(Session("UserName"), "'", "\'")%>';
var DtxtRestData = '<%=getviewReportLngStr("DtxtRestData")%>';
var LtxtDynObjErr = '<%=getviewReportLngStr("LtxtDynObjErr")%>';
var itemSmallRep = '<%=JBool(Request("itemSmallRep") = "Y")%>';
var Refresh = <%=Refresh%>;
var doRepLegend = <%=JBool(doRepLegend)%>;
</script>
<script language="javascript" src="<%=imgAddPath%>portal/viewReport.js"></script>
<%
rd.Filter = "linkType = 'O'"
If rd.recordcount > 0 Then %>
<form target="_blank" method="post" name="frmViewDetail" action="">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="">
<input type="hidden" name="ViewOnly" value="Y">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="sourceDoc" value="">
<input type="hidden" name="DocNum" value="">
<input type="hidden" name="cmd" value="">
<input type="hidden" name="document" value="C">
<input type="hidden" name="isEntry" value="">
<input type="hidden" name="LinkRep" value="Y">
<input type="hidden" name="high" value="">
<input type="hidden" name="c1" value="">
<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "C" Then %>OITM.ItemCode<% Else %>ItemName<% End If %>">
</form>
<% If userType = "V" Then %><!--#include file="../itemDetails.inc"--><% End If %>
<% End If
retVal = ""
isLoadRec = Request("loadRec") <> ""
For each itm in Request.Form
	If itm <> "Item" and itm <> "AddPath" and itm <> "T1" and itm <> "precio" and itm <> "SaleType" and _
		itm <> "Locked" and itm <> "WithoutPList" and itm <> "redir" and itm <> "err" and itm <> "errMInv" and itm <> "DocFlowErr" and itm <> "retURL" and _
		(not isLoadRec or isLoadRec and itm <> "loadRec" and itm <> "Qty" and itm <> "Price" and itm <> "SaleType" and itm <> "ItmEntry" and itm <> "RecType") Then
		If retVal <> "" Then retVal = retVal & "{y}"
		retVal = retVal & itm & "{i}" & Server.HTMLEncode(Request(itm))
	End If
Next
For each itm in Request.QueryString
	If itm <> "Item" and itm <> "AddPath" and itm <> "T1" and itm <> "precio" and itm <> "SaleType" and _
		itm <> "Locked" and itm <> "WithoutPList" and itm <> "redir" and itm <> "err" and itm <> "errMInv" and itm <> "DocFlowErr" and itm <> "retURL" and _
		(not isLoadRec or isLoadRec and itm <> "loadRec" and itm <> "Qty" and itm <> "Price" and itm <> "SaleType" and itm <> "ItmEntry" and itm <> "RecType") Then
		If retVal <> "" Then retVal = retVal & "{y}"
		retVal = retVal & itm & "{i}" & Server.HTMLEncode(Request(itm))
	End If
Next 

If Action0 or Action1 or Action2 or Action3 or Action5 or Action6 or Action7 Then %>
<form name="frmGoAction" action="execAction.asp" method="post">
<input type="hidden" name="ID" value="">
<input type="hidden" name="ObjectCode" value="">
<input type="hidden" name="Entry" value="">
<input type="hidden" name="Series" value="">
<input type="hidden" name="cmd" value="execAction">
<input type="hidden" name="retVal" value="<%=retVal%>">
</form>
<% End If %>
<% If Action4 Then %>
<form name="frmGoAddItem" action="cart/addCartSubmitM.asp" method="post">
<input type="hidden" name="Item" value="">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="T1" value="1">
<input type="hidden" name="precio" value="">
<input type="hidden" name="SaleType" value="">
<input type="hidden" name="Locked" value="">
<input type="hidden" name="WhsCode" value="">
<input type="hidden" name="WithoutPList" value="">
<input type="hidden" name="redir" value="report">
<input type="hidden" name="retVal" value="<%=retVal%>">
<input type="hidden" name="DocConf" value="">
</form>
<% End If
If Action5 Then %>
<form name="frmGoAddWish" action="wishlist/wlSubmit.asp" method="post">
<input type="hidden" name="Item" value="">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="redir" value="report">
<input type="hidden" name="retVal" value="<%=retVal%>">
</form>
<% End If
rd.Filter = "linkType = 'R'"
do while not rd.eof
rlVars = "" %>
<form method="post" <% If rd("linkPopup") = "Y" Then %>target="_blank"<% End If %> action="<% If Request("itemSmallRep") <> "Y" Then %><% If Request("pop") = "Y" or rd("linkPopup") = "Y" Then %>viewReportPrint.asp<% ELse %>report.asp<% End If %><% Else %>viewReportPrint.asp<% End If %>" name="frmRep<%=rd.bookmark%>">
<% If rd("linkPopup") = "Y" and Request("itemSmallRep") <> "Y" or Request("pop") = "Y" Then %>
<input type="hidden" name="pop" value="Y"><% End If %>
<input type="hidden" name="cmd" value="<% If userType = "V" Then %>report<% ElseIf userType = "C" Then %>viewRep<% End If %>">
<input type="hidden" name="rsIndex" value="<%=rd("linkObject")%>">
<input type="hidden" name="itemSmallRep" value="<% If rd("linkPopup") = "Y" and Request("itemSmallRep") <> "Y" Then %>Y<% Else %><%=Request("itemSmallRep")%><% End If %>">
<% rl.Filter = "colName = '" & rd("colName") & "'"
do while not rl.eof
If rl("valBy") = "F" Then
	If rlVars <> "" Then rlVars = rlVars & ", "
	rlVars = rlVars & "v" & rl("varId")
End If %>
<input type="hidden" name="var<%=rl("varIndex")%>" value="<% If rl("valBy") = "V" Then Response.Write rl("valValue") %>">
<% If rl("varType") = "DD" or rl("varType") = "L" Then %>
<input type="hidden" name="var<%=rl("varIndex")%>Desc" value="<% If rl("valBy") = "V" Then Response.Write rl("valValue") %>">
<% End If %>
<% rl.movenext
loop %>
</form>
<% rl.Filter = "colName = '" & rd("colName") & "' and valBy = 'F'" %>
<script language="javascript">
function goRep<%=rd.bookmark%>(<%=rlVars%>)
{
	<% do while not rl.eof %>
	document.frmRep<%=rd.bookmark%>.var<%=rl("varIndex")%>.value = v<%=rl("varId")%>;
	<% If rl("varType") = "DD" or rl("varType") = "L" Then %>
	document.frmRep<%=rd.bookmark%>.var<%=rl("varIndex")%>Desc.value = v<%=rl("varId")%>;
	<% End If %>
	<% rl.movenext
	loop %>
	document.frmRep<%=rd.bookmark%>.submit();
}
</script>
<% 
rd.movenext
loop %>
