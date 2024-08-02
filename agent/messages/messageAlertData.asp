<!--#include file="../myHTMLENcode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<% 
set rs = Server.CreateObject("ADODB.RecordSet")

If userType = "V" Then 
	user = Session("vendid")
	MainDoc = "ventas"
ElseIf userType = "C" Then
	user = Session("UserName")
	MainDoc = "clientes"
End If 
set rs = Server.CreateObject("ADODB.recordset")

sql = "SELECT Count('') " & _
	  "FROM olkomsg T0 " & _
	  "INNER JOIN  olkmsg1 T1 ON T1.OlkLog = T0.OlkLog " & _
	  "where OlkUserType = '" & userType & "' and OlkUser = N'" & saveHTMLDecode(user, False) & "' and OlkStatus = 'N' "
set rs = conn.execute(sql)
msgCount = rs(0)
rs.close

sql = "SELECT top 10 T0.OlkLog, CONVERT(nvarchar(60), OlkSubject) AS OlkSubject, " & _
	  "OlkUrgent, OlkStatus, Right(Convert(nvarchar(20),OlkDate,100), 7) AS Time, OLKCommon.dbo.DBOLKDateFormat" & Session("ID") & "(OlkDate) AS Date, Case When DateDiff(minute,OlkDate,getdate()) <= 5 Then 'Y' Else 'N' End IsNew, " & _
	  "Pop, T0.OLKUFromType, T2.CardType " & _
	  "FROM olkomsg T0 " & _
	  "INNER JOIN  olkmsg1 T1 ON T1.OlkLog = T0.OlkLog " & _
	  "left outer join OCRD T2 on T2.CardCode = T0.OlkUFrom and T0.OLKUFromType = 'C' " & _
	  "where OlkUserType = '" & userType & "' and OlkUser = N'" & saveHTMLDecode(user, False) & "' and OlkStatus = 'N' " & _
	  "order by T0.OlkDate desc "
rs.open sql, conn, 3, 1
sql = ""

msgStr = ""
do while not rs.eof
	If msgStr <> "" Then msgStr = msgStr & "{R}"
	msgStr = msgStr & rs(0) & "{S}" & rs(1) & "{S}" & rs(2) & "{S}" & rs(3) & "{S}" & rs(4) & "{S}" & rs(5) & "{S}" & rs(6) & "{S}" & rs(7) & "{S}" & rs(8) & "{S}" & rs(9)
	
	If rs("Pop") = "Y" Then
		If sql <> "" Then sql = sql & ", "
		sql = sql & rs("OlkLog")
	End If
rs.movenext
loop

If sql <> "" Then
	sql = "update OLKMSG1 set Pop = 'N' where OlkLog in (" & sql & ") and OlkUserType = '" & userType & "' and OlkUser = N'" & saveHTMLDecode(user, False) & "'"
	conn.execute(sql)
End If

Response.Write msgCount & "{C}" & msgStr %>
