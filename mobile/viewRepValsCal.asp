<!--#include file="lang/viewRepValsCal.asp" -->
<%
Response.Buffer = true
Response.Addheader "Pragma","no-cache"
'============================================================================================
' VARIABLES
'============================================================================================
'Dim i
scriptName = Request.ServerVariables("SCRIPT_NAME")

imgPath = "images/"
%>

<link rel="stylesheet" href="Reportes/style.css">
<% 
monthNames = array("", getviewRepValsCalLngStr("DtxtMonthJanuary"), getviewRepValsCalLngStr("DtxtMonthFebruary"), getviewRepValsCalLngStr("DtxtMonthMarch"), getviewRepValsCalLngStr("DtxtMonthApril"), getviewRepValsCalLngStr("DtxtMonthMay"), getviewRepValsCalLngStr("DtxtMonthJune"), getviewRepValsCalLngStr("DtxtMonthJuly"), getviewRepValsCalLngStr("DtxtMonthAugust"), getviewRepValsCalLngStr("DtxtMonthSeptember"), getviewRepValsCalLngStr("DtxtMonthOctober"), getviewRepValsCalLngStr("DtxtMonthNovember"), getviewRepValsCalLngStr("DtxtMonthDecember"))
dayNames = array("", getviewRepValsCalLngStr("DtxtSmallDayMonday"), getviewRepValsCalLngStr("DtxtSmallDayTuesday"), getviewRepValsCalLngStr("DtxtSmallDayWednesday"), getviewRepValsCalLngStr("DtxtSmallDayThursday"), getviewRepValsCalLngStr("DtxtSmallDayFriday"), getviewRepValsCalLngStr("DtxtSmallDaySaturday"), getviewRepValsCalLngStr("DtxtSmallDaySunday")) 
txtToday = getviewRepValsCalLngStr("DtxtToday")

Select Case Request("cmd")
	Case "viewRepValsCal"
		cmd = "viewRepVals"
		sql = "select IsNull(alterRSName, rsName) rsName, IsNull(alterVarName, varName) varName " & _
		"from OLKRS T0 " & _
		"inner join OLKRSVars T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = " & Request("editVar") & " " & _
		"left outer join OLKRSAlterNames T2 on T2.rsIndex = T0.rsIndex and T2.LanID = " & Session("LanID") & " " & _
		"left outer join OLKRSVarsAlterNames T3 on T3.rsIndex = T1.rsIndex and T3.varIndex = T1.varIndex and T3.LanID = " & Session("LanID") & " " & _
		"where T0.rsIndex = " & Request.Form("rsIndex")
	Case "adSearchValsCal"
		cmd = "adSearch"
		sql = "select IsNull(T2.alterName, T0.Name) Name, IsNull(T3.alterName, T1.Name) varName " & _
		"from OLKCustomSearch T0 " & _
		"inner join OLKCustomSearchVars T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.VarID = " & Request("editVar") & " " & _
		"left outer join OLKCustomSearchAlterNames T2 on T2.ObjectCode = T0.ObjectCode and T2.ID = T0.ID and T2.LanID = " & Session("LanID") & " " & _
		"left outer join OLKCustomSearchVarsAlterNames T3 on T3.ObjectCode = T0.ObjectCode and T3.ID = T1.ID and T3.VarID = T1.VarID and T3.LanID = " & Session("LanID") & " " & _
		"where T0.ObjectCode = " & Request("adObjID") & " and T0.ID = " & Request.Form("ID")
End Select
set rs = conn.execute(sql)
set rd = Server.CreateObject("ADODB.RecordSet")
rsName = rs(0)
varName = rs(1)
rs.close %>
<table border="0" cellspacing="0" width="100%" id="table1">
	<tr class="TblTltMnu">
		<td colspan="2"><img border="0" src="images/arrow_menu.gif" width="9" height="6">&nbsp;<%=Server.HTMLEncode(rsName)%> - <%=Server.HTMLEncode(varName)%></td>
	</tr>
	<form method="POST" name="frmViewRep" action="operaciones.asp">
	<tr class="TblAfueraMnu">
		<td colspan="2" align="center">
		<% 
		If Request("d") = "" Then
			calVal = Request("var" & Request("editVar"))
			If calVal <> "" Then
				calVald = Mid(calVal, InStr(myApp.DateFormat, "dd"), 2)
				calValm = Mid(calVal, InStr(myApp.DateFormat, "MM"), 2)
				calValy = Mid(calVal, InStr(myApp.DateFormat, "yyyy"), 4)
			End If
		Else
				calVald = Request("d")
				calValm = Request("m")
				calValy = Request("y")
		End If %>
		<%
		' call calendar
		makeCalendar calVald,calValm,calValy,""
		%>
		</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td colspan="2">
		<p align="center">
		</td>
	</tr>
		<% 	For each itm in Request.Form
		If itm <> "d" and itm <> "m" and itm <> "y" and itm <> "l" and itm <> "s" Then %>
		<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
		<% 	End If
		Next %>
	<tr class="TblAfueraMnu">
		<td>
		<p align="left">
		<% 	strDate = myApp.DateFormat
			strDate = Replace(strDate, "yyyy", selYear)
			strDate = Replace(strDate, "MM", Right("0" & selMonth, 2))
			strDate = Replace(strDate, "dd", Right("0" & selDay, 2)) %>
		<input type="submit" name="btnSubmit" value="<%=getviewRepValsCalLngStr("DtxtAccept")%>" onclick="javascript:document.frmViewRep.cmd.value='<%=cmd%>';document.frmViewRep.isSubmit.value='R';document.frmViewRep.var<%=Request("editVar")%>.value='<%=strDate%>';"></td>
		<td>
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		<input type="submit" name="btnCancel" value="<%=getviewRepValsCalLngStr("DtxtCancel")%>" onclick="javascript:document.frmViewRep.cmd.value='<%=cmd%>';document.frmViewRep.isSubmit.value='R';"></td>
	</tr>
	</form>
</table>
<%
'============================================================================================
' FUNCTIONS
'============================================================================================

' renders monthselection
function monthSelect()
	dim selMonthShowPrev, selYearShowPrev
	selMonthShowPrev = selMonth
	selYearShowPrev = selYear
	If selMonth <= 1 Then
		selMonthShowPrev = 13
		selYearShowPrev = selYear - 1
	End If
	
	with Response
		.Write "<a href=""javascript:goDay(" & selDay & "," & selMonthShowPrev - 1 & "," & selYearShowPrev & ")"">"
		.write "<img src=""" & imgPath & Session("rtl") & "flechaselec_back.gif"" border=""0"" alt=""" & monthNames(selMonth - 1)
		.Write """></a>" & chr(10)
		.Write "<select name=""m"" class=""month"" onChange=""javascript:goDay(" & selDay & ",this.value," & selYear & ")"">" & chr(10)
	end with
	
	for i = 1 to Ubound(monthNames)
		if i = int(selMonth) then 
			tmpStr = " selected"
		else
			tmpStr = ""
		end if
		Response.Write "<option value=""" & i & """" & tmpStr & ">" & monthNames(i) & "&nbsp;</option>" & chr(10)
	next	
	Response.Write "</select>" & chr(10)
end function
'============================================================================================

' renders yearselection
function yearSelect()
	Response.Write "<select name=""y"" class=""month"" onChange=""javascript:goDay(" & selDay & "," & selMonth  & ",this.value)"">" & chr(10)
	for i = year(now) - 5 to year(now) + 5
		if i = int(selYear) then 
			tmpStr = " selected"
		else
			tmpStr = ""
		end if
		Response.Write "<option value=""" & i & """" & tmpStr & ">" & i & "</option>" & chr(10)
	next
	
	' lets check if the next month value is more than 12
	' if, then show 1 month of next year
	dim selMonthShowNext, selYearShowNext
	selMonthShowNext = selMonth + 1
	selYearShowNext = selYear
	if selMonthShowNext < 1 then 
		selMonthShowNext = 12
		selYearShowNext = selYear - 1
	end if
	if selMonthShowNext > 12 then 
		selMonthShowNext = 1
		selYearShowNext = selYear + 1
	end if
	
	' write end for select, next month-link and hidden for language-selection
	with response
		.Write "</select>" & chr(10)
		.Write "<a href=""javascript:goDay(" & selDay & "," & selMonthShowNext & "," & selYearShowNext & ")"">"
		.write "<img src=""" & imgPath & Session("rtl") & "flechaselec.gif"" border=""0"" alt=""" & monthNames(selMonthShowNext)
		.Write """></a>" & chr(10) 		
	end with
end function
'============================================================================================

' calculate dates
function dayToShow(byval selDayTmp, byval selMonthTmp, byval selYearTmp)
selDay = selDayTmp : selMonth = selMonthTmp : selYear = selYearTmp

' if day selection values are empty, use current date
if len(SelDay) = 0 then SelDay = day(now)
if len(SelYear) = 0 then SelYear = year(now)
if len(selMonth) = 0 then selMonth = month(now)

' if month is 0, then show month 12 of previous year
if selMonth < 1 then
	selMonth = 12
	selYear = selYear - 1
end if

' if month is over 12, then show 1 month of next year
If selMonth > 12 then
	selMonth = 1
	selyear = selYear + 1
end if

' temporary date for date calculations
dim tmpDate
tmpDate = CDate("1/" & selMonth & "/" & selYear)

' how many days are in this month
lastDay = day(DateSerial(Year(tmpDate), Month(tmpDate) + 1, 0))

' what is the weekday of first day in this month
firstWeekDay = weekday(DateSerial(Year(tmpDate), Month(tmpDate), 0)+1, 2)

' check if selected date is valuable
if selDay < 1 then selDay = 1
if int(selDay) > int(lastDay) then selDay = lastDay
end Function
'============================================================================================

' render calendar
function makeCalendar(byval selDayTmp, byval selMonthTmp, byval selYearTmp, byval linkDays)
	
	Dim arrlinkDay
	arrLinkDay = split(linkDays, ",")	
	dayToShow selDayTmp, selMonthTmp, selYearTmp
	dim tmpDayInt, tmpDayInt2
	tmpDayint = 0	
	' render javascript 
	with response
		'.Write "<form method=""get"" action=""" & Request.ServerVariables("SCRIPT_NAME") & """>"
		.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""monthtableouter""><tr><td>"
		.Write "<table border=""0"" cellpadding=""1"" cellspacing=""1"" class=""monthtableinner"">" & chr(10)
		.Write "<tr>" & chr(10) & "<td class=""title"" colspan=""7""><div align=""center"" class=""title"">"
		.Write monthNames(selMonth)& " " & selYear & "</div></td>" & chr(10) & "</tr>" & chr(10)
		.Write "<tr>"
	end with
	
	for i = 1 to UBound(dayNames) 
		with response
			.Write "<td class=""days"" width=""13%""><div align=center class=""days"">"
			.Write dayNames(i)
			.Write "</td>"
		end with		
	next
	Response.Write "</tr>"
	
	dim ib, thisDay
	'do while tmpDayInt2 < lastDay
	for ib = 1  to 6
	Response.Write "<tr>" & chr(10)
		for i = 1 to 7			
		
		if int(tmpDayInt2) = int(day(now)) - 1 and int(month(now)) = int(selMonth) and int(year(now)) = int(selYear) then 
			thisDay = " id=""today"""
		else
			thisDay = ""
		end if			
			
			if tmpDayInt2 = selDay - 1 and firstWeekday > tmpDayInt + 1 then
				Response.Write "<td class=""normalday"" width=""14%""><div align=""center"" class=""day"">"
			elseif tmpDayInt2 = selDay - 1 then
				Response.Write "<td class=""selectedday"" width=""14%""" & thisDay & "><div align=""center"" class=""day"">"
			else
				' check if day is weekend
				if i = 6 or i = 7 then
					Response.Write "<td class=""weekend"" width=""14%""" & thisDay & "><div align=""center"" class=""day"">" 
				else
					Response.Write "<td class=""normalday"" width=""14%""" & thisDay & "><div align=""center"" class=""day"">" 
				end if
			end if		
			
			
			if firstWeekday > tmpDayInt + 1 or lastDay < tmpDayInt2 + 1 then
				Response.Write "&nbsp;"
			else
				tmpDayInt2 = tmpDayInt2 + 1				
				'dim dayFound, arrI
				'dayFound = false
				'for arrI = Lbound(arrLinkDay) to UBound(arrLinkDay)
				'	if int(tmpDayInt2) = int(arrLinkDay(arrI)) then dayFound = true
				'next
				if CInt(tmpDayInt2) <> CInt(selDay) then
					with response
						.Write "<a href=""javascript:goDay(" & tmpDayInt2 & "," 
						.write selMonth & "," & selYear & ")"">" & tmpDayInt2 & "</a>"
					
					end with
				'	dayFound = false
				else
					Response.Write tmpDayInt2
				end if								
			end if
				
			Response.Write "</td>" & chr(10)
			tmpDayInt = tmpDayInt + 1
		next		
	Response.Write "</tr>" & chr(10)
	'loop
	next
	
	with Response
		.Write "<tr>" & chr(10) & "<td class=""title"" colspan=""7"" nowrap>" & chr(10)
		.Write "<div align=""center"" class=""title"">"
		.Write monthSelect()
		.Write "&nbsp;"
		.write yearSelect()
		.Write "</div>" & chr(10) & "</td>"
		.Write chr(10) & "</tr>"
 		.Write todaySelect()
		.Write "</table></td></tr></table>"
		.Write "<input type=""hidden"" name=""d"" value="""">"
		.Write "<input type=""hidden"" name=""l"" value=""" & Request("l") & """>"
		.Write "<script language=""javascript"">function goDay(d, m, y) { "
		.Write "var myForm = document.forms[0];"
		.Write "myForm.d.value = d; myForm.m.value = m; myForm.y.value = y;"
		.Write "myForm.submit();"
		.Write "}</script>"
	end with
end function

function todaySelect
	todayD = Day(Now())
	todayM = Month(Now())
	todayY = Year(Now())
	If Not (CInt(selYear) = CInt(todayY) and CInt(selMonth) = CInt(todayM) and CInt(selDay) = CInt(todayD)) Then
		todaySelect = "<tr><td align=""center"" colspan=""7""><a href=""javascript:goDay(" & Right("0" & Day(Now()), 2) & "," & Right("0" & Month(Now()), 2) & "," & Year(Now()) & ");"">" & txtToday & "</a></td></tr>"
	Else
		todaySelect = "<tr><td align=""center"" colspan=""7"">" & txtToday & "</td></tr>"
	End If
end function

'============================================================================================
' that's it, the end
%>