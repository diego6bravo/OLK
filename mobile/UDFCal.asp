<!--#include file="lang/UDFCal.asp" -->
<%
Response.Buffer = true
Response.Addheader "Pragma","no-cache"
'============================================================================================
' VARIABLES
'============================================================================================
'Dim myToday ' this day
'Dim i
'Dim SelMonth, selDay, selYear
'Dim tmpStr
'Dim monthNames, dayNames
'Dim lastDay
'Dim tmpLang
'dim firstWeekDay
'Dim scriptName
scriptName = Request.ServerVariables("SCRIPT_NAME")
'Dim imgPath
imgPath = "images/"

txtToday = getUDFCalLngStr("DtxtToday")
%>

<link rel="stylesheet" href="Reportes/style.css">
<% 
monthNames = array("", getUDFCalLngStr("DtxtMonthJanuary"), getUDFCalLngStr("DtxtMonthFebruary"), getUDFCalLngStr("DtxtMonthMarch"), getUDFCalLngStr("DtxtMonthApril"), getUDFCalLngStr("DtxtMonthMay"), getUDFCalLngStr("DtxtMonthJune"), getUDFCalLngStr("DtxtMonthJuly"), getUDFCalLngStr("DtxtMonthAugust"), getUDFCalLngStr("DtxtMonthSeptember"), getUDFCalLngStr("DtxtMonthOctober"), getUDFCalLngStr("DtxtMonthNovember"), getUDFCalLngStr("DtxtMonthDecember"))
dayNames = array("", getUDFCalLngStr("DtxtSmallDayMonday"), getUDFCalLngStr("DtxtSmallDayTuesday"), getUDFCalLngStr("DtxtSmallDayWednesday"), getUDFCalLngStr("DtxtSmallDayThursday"), getUDFCalLngStr("DtxtSmallDayFriday"), getUDFCalLngStr("DtxtSmallDaySaturday"), getUDFCalLngStr("DtxtSmallDaySunday")) 


Select Case Request("returnCmd")
	Case "cartopt"
		Title = getUDFCalLngStr("LtxtShopCart")
		TableID = "OINV"
	Case "cartEditLine"
		Title = getUDFCalLngStr("LtxtShopCart")
		TableID = "INV1"
	Case "newClientUDF"
		Title = getUDFCalLngStr("DtxtBP")
		TableID = "OCRD"
	Case "activityUDF"
		Title = getUDFCalLngStr("DtxtActivity")
		TableID = "OCLG"
	Case "newClientContact"
		Title = getUDFCalLngStr("DtxtContact")
		TableID = "OCPR"
	Case "newClientAddress"
		Title = getUDFCalLngStr("DtxtAddress")
		TableID = "CRD1"
End Select

If Request("System") <> "Y" Then
	sql = "select IsNull(T1.AlterDescr, T0.Descr) Descr from CUFD T0 " & _
	"left outer join OLKCUFDAlterNames T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID and T1.LanID = " & Session("LanID") & " " & _
	"where T0.TableID = '" & TableID & "' and T0.AliasID = N'" & Request("editVar") & "'"
	set rs = conn.execute(sql)
	set rd = Server.CreateObject("ADODB.RecordSet")
	Descr = rs(0)
	rs.close
Else
	Select Case TableID
		Case "OINV"
			Select Case Request("editVar")
				Case "DocDate"
					Descr = "DocDate"
				Case "DocDueDate"
					Descr = "DocDueDate"
			End Select
	End Select
End If %>
<table border="0" cellspacing="0" width="100%" id="table1">
	<tr class="TblTltMnu">
		<td colspan="2"><img border="0" src="images/arrow_menu.gif" width="9" height="6">&nbsp;<%=Title%> - <%=Descr%></td>
	</tr>
	<form method="POST" name="frmUFDCal" action="operaciones.asp">
	<tr class="TblAfueraMnu">
		<td colspan="2" align="center">
		<% 
		If Request("d") = "" Then
			If Request("System") <> "Y" Then
				calVal = Request("U_" & Request("editVar"))
			Else
				calVal = Request(Request("editVar"))
			End If
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
		<% End If
		Next %>
	<tr class="TblAfueraMnu">
		<td>
		<p align="left">
		<input type="submit" name="btnSubmit" value="<%=getUDFCalLngStr("DtxtAccept")%>" onclick="javascript:calOk()"></td>
		<td>
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		<input type="submit" name="btnCancel" value="<%=getUDFCalLngStr("DtxtCancel")%>" onclick="javascript:calCancel();"></td>
	</tr>
	</form>
</table>
<% strDate = myApp.DateFormat
strDate = Replace(strDate, "yyyy", selYear)
strDate = Replace(strDate, "MM", Right("0" & selMonth, 2))
strDate = Replace(strDate, "dd", Right("0" & selDay, 2)) %>
<script type="text/javascript">
function calOk()
{
	document.frmUFDCal.cmd.value='<%=Request("returnCmd")%>';
	document.frmUFDCal.<% If Request("System") <> "Y" Then %>U_<% End If %><%=Request("editVar")%>.value='<%=strDate%>';
	<% If TableID = "OINV" and Request("System") = "Y" and Request("editVar") = "DocDate" Then
	Select Case CInt(Request("ObjType"))
		Case 17 %>
		document.frmUFDCal.DocDueDate.value=''; 
	<%	Case 23
		strCalc = selMonth & "/" & selDay & "/" & selYear
		dateCalc = DateAdd("d",30,strCalc)
		strCalc = Right("0" & Day(dateCalc), 2) & "/" & Right("0" & Month(dateCalc), 2) & "/" & Year(dateCalc) %>
		document.frmUFDCal.DocDueDate.value = '<%=strCalc%>';<%
	End Select
	End If %>
}
function calCancel()
{
	document.frmUFDCal.cmd.value='<%=Request("returnCmd")%>';
}
</script>
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

SetLocale(6154)

dim tmpDate
tmpDate = CDate(selMonth & "/01/" & selYear)

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