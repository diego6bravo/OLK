<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->

<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc" -->
<%

LogNum = Session("ItmRetVal")
CmdType = Request.Form("CmdType")
LineID = Request.Form("LineID")

Select Case CmdType
	Case "IsValid"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItmValData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		Response.Write rs(0) & "{S}" & rs(1)
	Case "AddCmbComp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItmAddCmbComp" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		retVal = rs(0) & "{S}"
		
		rs.close
		set rs = Server.CreateObject("ADODB.RecordSet")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetPriceList" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then retVal = retVal & "{L}"
			retVal = retVal & rs(0) & "{D}" & rs(1)
		rs.movenext
		loop
		
		retVal = retVal & "{S}"
		
		rs.close
		set rs = Server.CreateObject("ADODB.RecordSet")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then retVal = retVal & "{W}"
			retVal = retVal & rs(0) & "{D}" & rs(1)
		rs.movenext
		loop
		
		Response.Write retVal
	Case "DelCmbComp"
		cmd.CommandText = "DBOLKItmDelCmbComp" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@LineID") = LineID
		cmd.execute()
		
		Response.Write "ok"
End Select



%>