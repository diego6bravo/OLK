<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Response.Expires = 0
set rs = Server.CreateObject("ADODB.RecordSet")

Dim myAut
set myAut = New clsAuthorization

          

%>
<!--#include file="taskMonitorData.asp"-->
<!--#include file="authorizationClass.asp"-->
<%


strReturnValue = ""
arrValues = Split(Request("GetID"), ",")
For i = 0 to UBound(arrValues)
	
	values = Split(arrValues(i), "|")
	vType = values(0)
	vID = values(1)
	
	If strReturnValue <> "" Then strReturnValue = strReturnValue & "{S}"
	
	strReturnValue = strReturnValue & vType & "|" & vID & "|"
	
	Select Case vType
		Case "S"
			Select Case CInt(vID)
				Case 0
					strReturnValue = strReturnValue & GetTaskMonitorInfo(1)
				Case 1
					strReturnValue = strReturnValue & GetTaskMonitorInfo(2)
				Case 2
					strReturnValue = strReturnValue & GetTaskMonitorInfo(3)
				Case 3
					nextAct = GetTaskMonitorInfo(4)
					strReturnValue = strReturnValue & nextAct(1) & "|" & nextAct(0) & "|" & nextAct(2)
				Case 4
					strReturnValue = strReturnValue & GetTaskMonitorInfo(5)
				Case 5
					openPolls = GetTaskMonitorInfo(6)
					strReturnValue = strReturnValue & rs("Count") & " (" & rs("Pending") & ")"
				Case 6
					openOffers = GetTaskMonitorInfo(7)
					strReturnValue = strReturnValue & rs(0) & " / " & rs(1)
				Case 7
					strReturnValue = strReturnValue & GetTaskMonitorInfo(8)
				Case 8
					strReturnValue = strReturnValue & GetTaskMonitorInfo(9)
				Case 9
					strReturnValue = strReturnValue & GetTaskMonitorInfo(10)
				Case 10
					strReturnValue = strReturnValue & GetTaskMonitorInfo(11)
				Case 11
					strReturnValue = strReturnValue & GetTaskMonitorInfo(12)
				Case 12
					strReturnValue = strReturnValue & GetTaskMonitorInfo(13)
				Case 13
					strReturnValue = strReturnValue & GetTaskMonitorInfo(14)
				Case 14
					strReturnValue = strReturnValue & GetTaskMonitorInfo(15)
				Case 15
					strReturnValue = strReturnValue & GetTaskMonitorInfo(16)
				Case 16
					strReturnValue = strReturnValue & GetTaskMonitorInfo(17)
				Case 17
					strReturnValue = strReturnValue & GetTaskMonitorInfo(18)
			End Select
		Case "U"
			sql = "select Query from OLKInformer where Type = 'U' and ID = " & vID
			set rs = conn.execute(sql)
			sql = 	"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
					"declare @LanID int set @LanID = " & Session("LanID") & " " & _
					"select (" & rs(0) & ")"
			sql = QueryFunctions(sql)
			set rs = conn.execute(sql)
			If Not rs.Eof Then
				strReturnValue = strReturnValue & rs(0)
			End If
	End Select
Next

Response.Write strReturnValue
%>