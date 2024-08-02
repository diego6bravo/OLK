<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
          
If Request("delSerial") = "" Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAddIOSerial" & Session("ID")
	cmd.ActiveConnection = connCommon
	cmd.Parameters.Refresh
	cmd("@ObjectCode") = Session("ObjCode")
	cmd("@Type") = Session("Type")
	cmd("@LogNum") = Session("IORetVal")
	cmd("@WhsCode") = saveHTMLDecode(Session("bodega"), True)
	cmd("@ItemCode") = saveHTMLDecode(Request("ItemCode"), True)
	cmd("@SuppSerial") = saveHTMLDecode(Request("txtSerNum"), True)
	cmd.execute
	If Request("retAddPack") = "Y" and cmd("@Completed") = "Y" Then
		Response.Redirect "../operaciones.asp?cmd=invChkInOutAddByPack&txtOrderNum=" & Request("txtOrderNum") & "&confirm=" & Server.HTMLEncode(Request("ItemCode"))
	Else
		response.redirect "../operaciones.asp?cmd=invChkInOutCheckSerial&txtOrderNum=" & Request("txtOrderNum") & "&ItemCode=" & Server.HTMLEncode(Request("ItemCode")) & "&retVal=" & cmd("@RetVal") & "&SuppSerial=" & Request("txtSerNum") & "&retAddPack=" & Request("retAddPack")
	End If
	conn.close
Else
	sql = "declare @LogNum int set @LogNum = " & Session("IORetVal") & " " & _
			"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Request("ItemCode"), False) & "' " & _
			"delete R3_ObsCommon..DOC4 where LogNum = @LogNum and LineNum in (select LineNum from R3_ObsCommon..DOC1 where LogNum = @LogNum and ItemCode = @ItemCode and WhsCode = N'" & Session("bodega") & "') and LineNum2 = " & Request("delSerial")
	conn.execute(sql)
	response.redirect "../operaciones.asp?cmd=invChkInOutCheckSerial&txtOrderNum=" & Request("txtOrderNum") & "&ItemCode=" & Server.HTMLEncode(Request("ItemCode")) & "&ViewAll=" & Request("ViewAll") & "&retAddPack=" & Request("retAddPack")
End If

%>