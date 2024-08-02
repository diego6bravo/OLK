<!--#include file="myHTMLEncode.asp"-->
<% If Session("VendId") = "" Then response.redirect "default.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
set rs = server.createobject("ADODB.RecordSet")
PassDesc = Request("PassDesc") = "Y"

sql = ""

searchStr = saveHTMLDecode(Replace(Request("searchStr"),"*","%"), False)

set rs = Server.CreateObject("ADODB.RecordSet")
Select Case Request("Type")
	Case "DocLink"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSearchBPDoc"
		cmd.Parameters.Refresh()
		cmd("@dbID") = Session("ID")
		cmd("@CardCode") = Session("UserName")
		cmd("@searchStr") = Request("DocNum")
		cmd("@DocType") = Request("DocType")
		set rs = cmd.execute()
	Case "Crd"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("searchStr")
		cmd("@Type") = 2
		cmd("@SlpCode") = Session("vendid")
		set rs = cmd.execute()
	Case "Territory"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("searchStr")
		cmd("@Type") = 1
		cmd("@SlpCode") = Session("vendid")
		set rs = cmd.execute()
End Select

WriteSep = False
If not rs.eof Then
	do while not rs.eof
		If WriteSep Then Response.Write "{S}"
		For i = 0 to rs.Fields.Count - 1
			If i > 0 Then Response.Write "{C}"
			Response.Write rs(i)
		Next
		WriteSep = True
	rs.movenext
	loop
Else
	Response.Write "{NoData}"
End If

conn.close
set rs = nothing
set rd = nothing %>