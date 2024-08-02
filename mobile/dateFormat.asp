<% addLngPathStr = "" %>
<!--#include file="lang/dateFormat.asp" -->

<%

Function doFormatDate(varDate)
	Select Case Session("myLng")
		Case "en"
			myFormat = "%D, %M %DD, %YY"
		Case "es"
			myFormat = "%D, %DD de %M de %YY"
		Case "pt"
			myFormat = "%D, %DD de %M de %YY"
		Case "he"
			myFormat = "&#8207;&#1497;&#1493;&#1501; %D %DD %M %YY"
	End Select
	myFormat = Replace(myFormat, "%DD", Day(varDate))
	myFormat = Replace(myFormat, "%YY", Year(varDate))
	Select Case Month(varDate)
		Case 1
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthJanuary"))
		Case 2
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthFebruary"))
		Case 3
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthMarch"))
		Case 4
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthApril"))
		Case 5
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthMay"))
		Case 6
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthJune"))
		Case 7
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthJuly"))
		Case 8
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthAugust"))
		Case 9
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthSeptember"))
		Case 10
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthOctober"))
		Case 11
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthNovember"))
		Case 12
			myFormat = Replace(myFormat, "%M", getdateFormatLngStr("DtxtMonthDecember"))
	End Select
	Select Case weekday(varDate)
		Case 1
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDaySunday"))
		Case 2
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDayMonday"))
		Case 3
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDayTuesday"))
		Case 4
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDayWednesday"))
		Case 5
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDayThursday"))
		Case 6
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDayFriday"))
		Case 7
			myFormat = Replace(myFormat, "%D", getdateFormatLngStr("DtxtDaySaturday"))
	End Select
	doFormatDate = myFormat
End Function

%>