<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

retVal = ""

Select Case Request.Form("Type")
	Case "S"
		cmd.CommandText = "DBOLKGetActivitySubjects" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@Type") = CInt(Request.Form("Value"))
		rs.open cmd, , 3, 1

		do while not rs.eof
			If retVal <> "" Then retVal = retVal & "{S}"
			retVal = retVal & rs(0) & "{C}" & rs(1)
		rs.movenext
		loop
	Case "CTel"
		cmd.CommandText = "DBOLKGetContactPhone" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@Code") = CInt(Request.Form("Code"))
		cmd("@FieldID") = "Tel1"
		rs.open cmd, , 3, 1
		
		retVal = rs(0)
	Case "Cnt"
		cmd.CommandText = "DBOLKGetCountryStates" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@Code") = Request.Form("Code")
		rs.open cmd, , 3, 1

		do while not rs.eof
			If retVal <> "" Then retVal = retVal & "{S}"
			retVal = retVal & rs(0) & "{C}" & rs(1)
		rs.movenext
		loop
     Case "Dur"
		cmd.CommandText = "DBOLKGetActivityDur"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("ActRetVal")
		cmd("@Recontact") = SaveCmdDate(Request("Recontact"))
		cmd("@BeginTime") = SaveCmdTime(Request("BeginTime"))
		cmd("@endDate") = SaveCmdDate(Request("endDate"))
		cmd("@ENDTime") = SaveCmdTime(Request("ENDTime"))
		cmd.execute()
		retVal = cmd("@Duration") & "{S}" & cmd("@DurType")
	Case "Doc"
		cmd.CommandText = "DBOLKGetDocEntry" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ObjCode") = CInt(Request("DocType"))
		cmd("@CardCode") = Session("UserName")
		cmd("@SearchInt") = Request("Value")
		set rs = cmd.execute()
		If Not rs.Eof Then
			Response.Write "ok|" & rs(0) & "|" & rs(1)
		Else
			Response.Write "err"
		End If
End Select

Response.Write retVal
%>